package main

// Conversion engine for the native Windows UI. Word remains an out-of-process
// COM dependency. All PDFs are staged before publication, so Cancel never
// destroys pre-existing PDFs and never leaves new PDFs from this run behind.

import (
	"bufio"
	"bytes"
	"context"
	"embed"
	"encoding/base64"
	"encoding/binary"
	"errors"
	"fmt"
	"io"
	"os"
	"os/exec"
	"path/filepath"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
	"unicode/utf16"
)

//go:embed word_worker.ps1
var resources embed.FS

var wordStartupGate = make(chan struct{}, 1)

const (
	wordStartTimeout  = 30 * time.Second
	documentTimeout   = 5 * time.Minute
	workerStopTimeout = 4 * time.Second
)

type job struct {
	name      string
	original  string
	working   string
	output    string
	stagedPDF string
}

type outcome struct {
	job job
	err error
}

type batchResult struct {
	total     int
	converted int
	failed    int
	elapsed   time.Duration
	canceled  bool
	err       error
}

type wordWorker struct {
	cmd      *exec.Cmd
	stdin    io.WriteCloser
	lines    chan string
	finished chan struct{}
	waitErr  error
	stderr   *bytes.Buffer
	pid      int          // PID of the Word COM instance created by this worker, not any user Word.
	process  *trackedWord // Pinned process handle prevents accidental PID reuse.
	sync.Once
}

func encodedCommand(script string) string {
	chars := utf16.Encode([]rune(script))
	raw := make([]byte, 2*len(chars))
	for i, c := range chars {
		binary.LittleEndian.PutUint16(raw[2*i:], c)
	}
	return base64.StdEncoding.EncodeToString(raw)
}

func powershell(script string) (string, error) {
	cmd := exec.Command("powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", encodedCommand(script))
	hideChildWindow(cmd)
	var out, errBuf bytes.Buffer
	cmd.Stdout = &out
	cmd.Stderr = &errBuf
	if err := cmd.Run(); err != nil {
		return "", fmt.Errorf("PowerShell: %w: %s", err, strings.TrimSpace(errBuf.String()))
	}
	return out.String(), nil
}

// This operation changes the user's DEFAULT PRINTER temporarily, then restores
// it even if conversion fails or is canceled. Do not run multiple batches.
func selectPrinter() (func() error, error) {
	const setup = `$ErrorActionPreference='Stop'
function E([string]$s){if($null -eq $s){return ''};[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s))}
$p=@(Get-CimInstance Win32_Printer)
$d=$p|Where-Object {$_.Default}|Select-Object -First 1
$t=$p|Where-Object {$_.Name -eq 'Microsoft Print to PDF'}|Select-Object -First 1
Write-Output ('D='+(E $d.Name))
$c='0'
if(($null -ne $t) -and ($null -ne $d) -and ($d.Name -ne 'Microsoft Print to PDF')){
  $w=New-Object -ComObject WScript.Network
  $w.SetDefaultPrinter('Microsoft Print to PDF')
  Start-Sleep -Milliseconds 350
  $c='1'
}
Write-Output ('C='+$c)`
	output, err := powershell(setup)
	noop := func() error { return nil }
	if err != nil {
		return noop, err
	}
	values := map[string]string{}
	for _, line := range strings.Split(output, "\n") {
		line = strings.TrimSpace(line)
		if i := strings.IndexByte(line, '='); i >= 0 {
			values[line[:i]] = line[i+1:]
		}
	}
	if values["C"] != "1" {
		return noop, nil
	}
	old, err := base64.StdEncoding.DecodeString(values["D"])
	if err != nil || len(old) == 0 {
		return noop, errors.New("could not save the previous default printer")
	}
	restore := `$ErrorActionPreference='Stop'
$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('` + base64.StdEncoding.EncodeToString(old) + `'))
$w=New-Object -ComObject WScript.Network
$w.SetDefaultPrinter($n)`
	return func() error { _, e := powershell(restore); return e }, nil
}

// parseWordReady tolerates a zero HWND from a hidden Word COM server.
// A fallback PID is provided only when PowerShell verified that it is a new
// COM-automation WINWORD process. Missing PID must not prevent conversion.
func parseWordReady(line string) (uintptr, int, error) {
	parts := strings.Split(line, "|")
	if len(parts) < 2 || parts[0] != "READY" {
		return 0, 0, fmt.Errorf("unexpected Word readiness message: %s", line)
	}
	hwnd, err := strconv.ParseUint(parts[1], 10, 64)
	if err != nil {
		return 0, 0, fmt.Errorf("invalid Word window handle: %w", err)
	}
	pid := 0
	if len(parts) >= 3 && parts[2] != "" {
		pid, err = strconv.Atoi(parts[2])
		if err != nil || pid < 0 {
			return 0, 0, fmt.Errorf("invalid Word process ID in readiness message")
		}
	}
	return uintptr(hwnd), pid, nil
}

func startWord(ctx context.Context) (*wordWorker, error) {
	select {
	case wordStartupGate <- struct{}{}:
		defer func() { <-wordStartupGate }()
	case <-ctx.Done():
		return nil, ctx.Err()
	}

	if err := ctx.Err(); err != nil {
		return nil, err
	}
	script, err := resources.ReadFile("word_worker.ps1")
	if err != nil {
		return nil, err
	}
	cmd := exec.Command("powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", encodedCommand(string(script)))
	hideChildWindow(cmd)
	stdin, err := cmd.StdinPipe()
	if err != nil {
		return nil, err
	}
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		stdin.Close()
		return nil, err
	}
	stderr := &bytes.Buffer{}
	cmd.Stderr = stderr
	if err = cmd.Start(); err != nil {
		stdin.Close()
		return nil, err
	}
	w := &wordWorker{cmd: cmd, stdin: stdin, stderr: stderr, lines: make(chan string, 8), finished: make(chan struct{})}
	go func() {
		s := bufio.NewScanner(stdout)
		s.Buffer(make([]byte, 4096), 2*1024*1024)
		for s.Scan() {
			w.lines <- strings.TrimSpace(s.Text())
		}
		close(w.lines)
	}()
	go func() { w.waitErr = cmd.Wait(); close(w.finished) }()
	return awaitWordReady(ctx, w)
}

// Keep the startup handshake alive on Cancel so Word can report its owned PID.
// Killing PowerShell before READY would discard that process identity.
func awaitWordReady(ctx context.Context, w *wordWorker) (*wordWorker, error) {
	timer := time.NewTimer(wordStartTimeout)
	defer timer.Stop()
	cancelSignal := ctx.Done()
	for {
		select {
		case line, ok := <-w.lines:
			if ok && strings.HasPrefix(line, "OWNED|") {
				if err := w.trackStartupPID("READY|" + strings.TrimPrefix(line, "OWNED|")); err != nil {
					w.shutdown()
					return nil, err
				}
				if err := ctx.Err(); err != nil && w.process != nil {
					w.forceStop()
					return nil, err
				}
				continue
			}
			if ok && strings.HasPrefix(line, "READY|") {
				if err := w.trackStartupPID(line); err != nil {
					w.shutdown()
					return nil, err
				}
				if err := ctx.Err(); err != nil {
					w.shutdown()
					return nil, err
				}
				return w, nil
			}
			w.forceStop()
			return nil, fmt.Errorf("Word did not become ready: %s", line)
		case <-w.finished:
			w.forceStop()
			return nil, fmt.Errorf("Word exited on startup: %v %s", w.waitErr, w.stderr.String())
		case <-timer.C:
			w.forceStop()
			return nil, errors.New("Word startup timed out")
		case <-cancelSignal:
			if w.process != nil {
				w.forceStop()
				return nil, ctx.Err()
			}
			// Wait within the existing startup deadline; READY precedes shutdown.
			cancelSignal = nil
		}
	}
}

// Both OWNED and READY use the worker's independently verified new-process PID.
// READY additionally cross-checks HWND when Word supplies one. Keep a single
// pinned handle, and discard it if the final readiness identity disagrees.
func (w *wordWorker) trackStartupPID(line string) error {
	hwnd, candidate, err := parseWordReady(line)
	if err != nil {
		return err
	}
	if candidate <= 0 || (hwnd != 0 && wordPID(hwnd) != candidate) {
		w.releaseWord()
		w.pid = 0
		return nil
	}
	if w.pid > 0 && w.pid != candidate {
		return errors.New("Word startup process identity changed")
	}
	if w.pid != candidate {
		w.releaseWord()
		w.pid = candidate
	}
	if w.process == nil {
		w.process, _ = openTrackedWord(candidate)
	}
	return nil
}

// waitWord observes the pinned OS process handle acquired at startup.
// It cannot target a different process after Windows recycles a PID.
func (w *wordWorker) waitWord(timeout time.Duration) bool {
	if w.process == nil {
		return true
	}
	return w.process.wait(timeout)
}

func (w *wordWorker) killWord() {
	if w.process != nil && !w.waitWord(0) {
		_ = w.process.kill()
	}
}

func (w *wordWorker) releaseWord() {
	if w.process != nil {
		w.process.close()
		w.process = nil
	}
}

func (w *wordWorker) killPowerShell() {
	if w.cmd != nil && w.cmd.Process != nil {
		_ = w.cmd.Process.Kill()
	}
}

func (w *wordWorker) forceStop() {
	if w == nil {
		return
	}
	w.Do(func() {
		defer w.releaseWord()
		// Terminate this worker's Word instance before its owning PowerShell
		// (which may be blocked on a COM call). Never enumerate/kill user Word.
		w.killWord()
		w.killPowerShell()
		if w.stdin != nil {
			_ = w.stdin.Close()
		}
		select {
		case <-w.finished:
		case <-time.After(workerStopTimeout):
		}
		_ = w.waitWord(workerStopTimeout)
	})
}

func (w *wordWorker) shutdown() {
	if w == nil {
		return
	}
	w.Do(func() {
		defer w.releaseWord()
		if w.stdin != nil {
			_, _ = io.WriteString(w.stdin, "QUIT\n")
			_ = w.stdin.Close()
		}
		select {
		case <-w.finished:
		case <-time.After(workerStopTimeout):
			// PowerShell/COM is stuck; terminate the tracked worker processes.
			w.killWord()
			w.killPowerShell()
			select {
			case <-w.finished:
			case <-time.After(time.Second):
			}
		}
		// PowerShell may finish its Word.Quit() call while WINWORD.EXE stays
		// alive in the background. Give Word a grace period, then kill ONLY
		// the tracked owned process if it has not exited on its own.
		if !w.waitWord(1500 * time.Millisecond) {
			w.killWord()
			_ = w.waitWord(workerStopTimeout)
		}
	})
}

func (w *wordWorker) convert(ctx context.Context, j job) error {
	if err := ctx.Err(); err != nil {
		return err
	}
	line := base64.StdEncoding.EncodeToString([]byte(j.working)) + "|" + base64.StdEncoding.EncodeToString([]byte(j.stagedPDF)) + "\n"
	if _, err := io.WriteString(w.stdin, line); err != nil {
		return err
	}
	timer := time.NewTimer(documentTimeout)
	defer timer.Stop()
	select {
	case msg, ok := <-w.lines:
		if !ok {
			return errors.New("Word worker stopped unexpectedly")
		}
		if strings.HasPrefix(msg, "F|") {
			d, _ := base64.StdEncoding.DecodeString(strings.TrimPrefix(msg, "F|"))
			return fmt.Errorf("Word COM: %s", d)
		}
		fields := strings.SplitN(msg, "|", 7)
		if len(fields) < 7 || fields[0] != "R" {
			return fmt.Errorf("unexpected Word response: %s", msg)
		}
		if fields[2] == "1" {
			return nil
		}
		d, _ := base64.StdEncoding.DecodeString(fields[6])
		return fmt.Errorf("Word export: %s", d)
	case <-timer.C:
		w.forceStop()
		return errors.New("Word took more than 5 minutes to process a document")
	case <-ctx.Done():
		w.forceStop()
		return ctx.Err()
	}
}

func collectDocs(dir string) ([]job, error) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, err
	}
	jobs := make([]job, 0)
	for _, e := range entries {
		if e.IsDir() || strings.HasPrefix(e.Name(), "~$") || !strings.EqualFold(filepath.Ext(e.Name()), ".docx") {
			continue
		}
		info, err := e.Info()
		if err != nil || !info.Mode().IsRegular() {
			continue
		}
		path := filepath.Join(dir, e.Name())
		jobs = append(jobs, job{name: e.Name(), original: path, working: path, output: strings.TrimSuffix(path, filepath.Ext(path)) + ".pdf"})
	}
	sort.Slice(jobs, func(i, j int) bool { return strings.ToLower(jobs[i].name) < strings.ToLower(jobs[j].name) })
	return jobs, nil
}

func volumeC(path string) bool { return strings.EqualFold(filepath.VolumeName(path), "C:") }
func allocateCache() (string, error) {
	candidates := []string{os.TempDir(), filepath.Join(os.Getenv("LOCALAPPDATA"), "Temp")}
	for _, dir := range candidates {
		if dir == "" || !volumeC(dir) {
			continue
		}
		if p, err := os.MkdirTemp(dir, "DOCX_to_PDF_*"); err == nil {
			return p, nil
		}
	}
	return "", errors.New("no writable temporary folder on C:")
}
func copyFile(from, to string) error {
	a, err := os.Open(from)
	if err != nil {
		return err
	}
	defer a.Close()
	b, err := os.OpenFile(to, os.O_CREATE|os.O_EXCL|os.O_WRONLY, 0600)
	if err != nil {
		return err
	}
	_, e := io.Copy(b, a)
	ce := b.Close()
	if e != nil {
		return e
	}
	return ce
}

func tryJob(ctx context.Context, j job, w **wordWorker) error {
	var last error
	for attempt := 0; attempt < 2; attempt++ {
		if err := ctx.Err(); err != nil {
			return err
		}
		if *w == nil {
			nw, err := startWord(ctx)
			if err != nil {
				last = err
				continue
			}
			*w = nw
		}
		last = (*w).convert(ctx, j)
		if last == nil {
			return nil
		}
		(*w).shutdown()
		*w = nil
	}
	return last
}

// The worker goroutines are joined before the staging directory is removed.
func processJobs(ctx context.Context, jobs []job, words []*wordWorker, progress func(int, int)) ([]job, int) {
	return processJobsWithStartup(ctx, jobs, words, len(words), 0, nil, progress)
}

// Additional workers join the same queue while existing workers convert. The
// initializer's WaitGroup reservation is taken before any worker is launched,
// so Wait cannot finish while another worker is still being initialized.
func processJobsWithStartup(ctx context.Context, jobs []job, words []*wordWorker, limit int, startup time.Duration, start func(context.Context) (*wordWorker, error), progress func(int, int)) ([]job, int) {
	queue := make(chan job, len(jobs))
	results := make(chan outcome, len(jobs))
	for _, j := range jobs {
		queue <- j
	}
	close(queue)
	var wg sync.WaitGroup
	var performance sync.Mutex
	var service time.Duration
	completed := 0
	growing := start != nil && limit > len(words)
	if growing {
		wg.Add(1)
	}
	launch := func(worker *wordWorker) {
		wg.Add(1)
		go func(w *wordWorker) {
			defer wg.Done()
			defer func() {
				if w != nil {
					w.shutdown()
				}
			}()
			for j := range queue {
				if ctx.Err() != nil {
					return
				}
				began := time.Now()
				err := tryJob(ctx, j, &w)
				if growing && err == nil {
					performance.Lock()
					service += time.Since(began)
					completed++
					performance.Unlock()
				}
				results <- outcome{job: j, err: err}
			}
		}(worker)
	}
	for _, w := range words {
		launch(w)
	}
	if growing {
		go func() {
			defer wg.Done()
			for active := len(words); active < limit; active++ {
				if ctx.Err() != nil || len(queue) == 0 {
					return
				}
				performance.Lock()
				worthwhile := worthStartingWorker(len(queue), active, completed, service, startup)
				performance.Unlock()
				if !worthwhile {
					return
				}
				began := time.Now()
				w, err := start(ctx)
				startup = time.Since(began)
				if err != nil || w == nil {
					return
				}
				if ctx.Err() != nil {
					w.shutdown()
					return
				}
				launch(w)
			}
		}()
	}
	go func() { wg.Wait(); close(results) }()
	successful := make([]job, 0, len(jobs))
	failures := 0
	done := 0
	for r := range results {
		if r.err == nil {
			successful = append(successful, r.job)
		} else if !errors.Is(r.err, context.Canceled) {
			failures++
		}
		done++
		progress(done, len(jobs))
	}
	return successful, failures
}

// Use observed conversion and initialization latency to avoid starting a worker
// after the remaining queue would already have drained. The calibrated batch
// limit remains the upper bound; this guard can only reduce it.
func worthStartingWorker(queued, active, completed int, service, startup time.Duration) bool {
	if queued <= 0 || active <= 0 {
		return false
	}
	if completed <= 0 || startup <= 0 {
		return true
	}
	remaining := float64(queued) * service.Seconds() / float64(completed) / float64(active)
	return remaining > startup.Seconds()*1.25
}

type publication struct{ target, backup string }

func uniqueBackupPath(target string) (string, error) {
	f, err := os.CreateTemp(filepath.Dir(target), ".docx2pdf-backup-*.pdf")
	if err != nil {
		return "", err
	}
	name := f.Name()
	if err := f.Close(); err != nil {
		return "", err
	}
	if err := os.Remove(name); err != nil {
		return "", err
	}
	return name, nil
}
func rollback(published []publication) error {
	var issues []string
	for i := len(published) - 1; i >= 0; i-- {
		p := published[i]
		if e := os.Remove(p.target); e != nil && !os.IsNotExist(e) {
			issues = append(issues, e.Error())
		}
		if p.backup != "" {
			if e := os.Rename(p.backup, p.target); e != nil {
				issues = append(issues, e.Error())
			}
		}
	}
	if len(issues) > 0 {
		return errors.New(strings.Join(issues, "; "))
	}
	return nil
}

// Commit only after conversion completes. Restore any existing outputs on cancel
// or failed publication. No existing PDF is deleted merely because Cancel was clicked.
func publishJobs(ctx context.Context, jobs []job) (int, error) {
	published := make([]publication, 0, len(jobs))
	commit := false
	defer func() {
		if commit {
			for _, p := range published {
				if p.backup != "" {
					_ = os.Remove(p.backup)
				}
			}
		}
	}()
	undo := func(err error) (int, error) {
		if e := rollback(published); e != nil {
			return 0, fmt.Errorf("%w; restoration problem: %v", err, e)
		}
		return 0, err
	}
	for _, j := range jobs {
		if ctx.Err() != nil {
			return undo(ctx.Err())
		}
		tmp, err := os.CreateTemp(filepath.Dir(j.output), ".docx2pdf-new-*.pdf")
		if err != nil {
			return undo(err)
		}
		name := tmp.Name()
		if err = tmp.Close(); err != nil {
			os.Remove(name)
			return undo(err)
		}
		_ = os.Remove(name) // copyFile requires exclusive creation.
		if err = copyFile(j.stagedPDF, name); err != nil {
			_ = os.Remove(name)
			return undo(err)
		}
		if ctx.Err() != nil {
			_ = os.Remove(name)
			return undo(ctx.Err())
		}
		p := publication{target: j.output}
		if _, e := os.Stat(j.output); e == nil {
			p.backup, err = uniqueBackupPath(j.output)
			if err != nil {
				_ = os.Remove(name)
				return undo(err)
			}
			if err = os.Rename(j.output, p.backup); err != nil {
				_ = os.Remove(name)
				return undo(err)
			}
		} else if !os.IsNotExist(e) {
			_ = os.Remove(name)
			return undo(e)
		}
		// Once an old PDF has been renamed, journal the output immediately.
		published = append(published, p)
		if err = os.Rename(name, j.output); err != nil {
			_ = os.Remove(name)
			return undo(err)
		}
	}
	if ctx.Err() != nil {
		return undo(ctx.Err())
	}
	commit = true
	return len(published), nil
}

func convertBatch(ctx context.Context, dir string, progress func(int, int)) (res batchResult) {
	start := time.Now()
	defer func() { res.elapsed = time.Since(start) }()
	docs, err := collectDocs(dir)
	if err != nil {
		res.err = err
		return
	}
	res.total = len(docs)
	if len(docs) == 0 {
		res.err = errors.New("this folder contains no DOCX documents")
		return
	}
	if ctx.Err() != nil {
		res.canceled = true
		return
	}
	restore, perr := selectPrinter()
	// Continue with the existing printer if switching is not possible.
	if perr != nil {
		restore = func() error { return nil }
	}
	defer func() {
		if e := restore(); e != nil {
			if res.err == nil {
				res.err = fmt.Errorf("could not restore default printer: %w", e)
			}
		}
	}()
	cache, err := allocateCache()
	if err != nil {
		res.err = err
		return
	}
	defer os.RemoveAll(cache)
	for i := range docs {
		if ctx.Err() != nil {
			res.canceled = true
			return
		}
		docs[i].stagedPDF = filepath.Join(cache, fmt.Sprintf("%06d.pdf", i))
		if !volumeC(dir) {
			docs[i].working = filepath.Join(cache, fmt.Sprintf("%06d.docx", i))
			if err = copyFile(docs[i].original, docs[i].working); err != nil {
				res.err = fmt.Errorf("copy DOCX: %w", err)
				return
			}
		}
	}
	if ctx.Err() != nil {
		res.canceled = true
		return
	}
	firstStart := time.Now()
	first, err := startWord(ctx)
	if err != nil {
		if ctx.Err() != nil {
			res.canceled = true
		} else {
			res.err = err
		}
		return
	}
	startup := time.Since(firstStart)
	target := adaptiveWorkerLimit(len(docs))
	// Additional COM activations require independently tracked ownership. Keep
	// the existing single worker usable if Word cannot establish that identity.
	if first.process == nil {
		target = 1
	}
	startAdditional := func(ctx context.Context) (*wordWorker, error) {
		if !roomForAdditionalWord() {
			return nil, errors.New("insufficient memory for another Word worker")
		}
		w, e := startWord(ctx) // Includes the shared cancelable startup gate.
		if e != nil {
			return nil, e
		}
		if w.process == nil {
			w.shutdown()
			return nil, errors.New("additional Word worker ownership was not verified")
		}
		return w, nil
	}
	succeeded, failed := processJobsWithStartup(ctx, docs, []*wordWorker{first}, target, startup, startAdditional, progress)
	if ctx.Err() != nil {
		res.canceled = true
		return
	}
	res.failed = failed
	count, err := publishJobs(ctx, succeeded)
	if err != nil {
		if ctx.Err() != nil {
			res.canceled = true
		} else {
			res.err = err
		}
		return
	}
	res.converted = count
	return
}
