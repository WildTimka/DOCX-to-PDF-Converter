// DOCX to PDF for Windows x64.
// Source reconstructed from the released executable's embedded PowerShell worker
// and observed behaviour. It is NOT claimed to be the original Go source code.
package main

import (
	"bufio"
	"bytes"
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
	"strings"
	"sync"
	"time"
	"unicode/utf16"
)

//go:embed word_worker.ps1
var resources embed.FS

const (
	wordStartTimeout = 30 * time.Second
	documentTimeout  = 5 * time.Minute
	closeTimeout     = 5 * time.Second
	parallelMinFiles = 11
)

type job struct {
	name, original, working, output string
}

type outcome struct {
	job job
	err error
}

type wordWorker struct {
	cmd      *exec.Cmd
	stdin    io.WriteCloser
	lines    chan string
	finished chan error
	stderr   *bytes.Buffer
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
	var out, errBuf bytes.Buffer
	cmd.Stdout = &out
	cmd.Stderr = &errBuf
	if err := cmd.Run(); err != nil {
		return "", fmt.Errorf("PowerShell: %w: %s", err, strings.TrimSpace(errBuf.String()))
	}
	return out.String(), nil
}

// The printer selection is intentionally system-wide and temporary. The original
// default name is held until workers finish and restored even on ordinary errors.
func selectPrinter() (func() error, error) {
	const setup = `$ErrorActionPreference='Stop'
function E([string]$s){if($null -eq $s){return ''};[Convert]::ToBase64String([Text.Encoding]::UTF8.GetBytes($s))}
$p=@(Get-CimInstance Win32_Printer)
$d=$p|Where-Object {$_.Default}|Select-Object -First 1
$t=$p|Where-Object {$_.Name -eq 'Microsoft Print to PDF'}|Select-Object -First 1
Write-Output ('D='+(E $d.Name))
Write-Output ('T='+$(if($null -ne $t){'1'}else{'0'}))
$c='0'
if(($null -ne $t) -and ($null -ne $d) -and ($d.Name -ne 'Microsoft Print to PDF')){
  $w=New-Object -ComObject WScript.Network
  $w.SetDefaultPrinter('Microsoft Print to PDF')
  Start-Sleep -Milliseconds 350
  $c='1'
}
Write-Output ('C='+$c)`
	output, err := powershell(setup)
	if err != nil {
		return func() error { return nil }, err
	}
	values := map[string]string{}
	for _, line := range strings.Split(output, "\n") {
		line = strings.TrimSpace(line)
		if p := strings.IndexByte(line, '='); p >= 0 {
			values[line[:p]] = line[p+1:]
		}
	}
	if values["C"] != "1" {
		return func() error { return nil }, nil
	}
	old, err := base64.StdEncoding.DecodeString(values["D"])
	if err != nil || len(old) == 0 {
		return func() error { return nil }, errors.New("cannot decode default printer name")
	}
	restore := `$ErrorActionPreference='Stop'
$n=[Text.Encoding]::UTF8.GetString([Convert]::FromBase64String('` + base64.StdEncoding.EncodeToString(old) + `'))
$w=New-Object -ComObject WScript.Network
$w.SetDefaultPrinter($n)
Start-Sleep -Milliseconds 250`
	return func() error { _, err := powershell(restore); return err }, nil
}

func startWord() (*wordWorker, error) {
	script, err := resources.ReadFile("word_worker.ps1")
	if err != nil {
		return nil, err
	}
	cmd := exec.Command("powershell.exe", "-NoLogo", "-NoProfile", "-NonInteractive", "-STA", "-EncodedCommand", encodedCommand(string(script)))
	stdin, err := cmd.StdinPipe()
	if err != nil {
		return nil, err
	}
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		_ = stdin.Close()
		return nil, err
	}
	stderr := &bytes.Buffer{}
	cmd.Stderr = stderr
	if err := cmd.Start(); err != nil {
		_ = stdin.Close()
		return nil, err
	}
	worker := &wordWorker{cmd: cmd, stdin: stdin, lines: make(chan string, 8), finished: make(chan error, 1), stderr: stderr}
	go func() {
		scanner := bufio.NewScanner(stdout)
		scanner.Buffer(make([]byte, 4096), 2*1024*1024)
		for scanner.Scan() {
			worker.lines <- strings.TrimSpace(scanner.Text())
		}
		close(worker.lines)
	}()
	go func() { worker.finished <- cmd.Wait() }()
	timeout := time.NewTimer(wordStartTimeout)
	defer timeout.Stop()
	select {
	case line, ok := <-worker.lines:
		if ok && line == "READY" {
			return worker, nil
		}
		worker.shutdown()
		return nil, fmt.Errorf("Word not ready: %s", line)
	case err := <-worker.finished:
		worker.shutdown()
		return nil, fmt.Errorf("Word exited while starting: %v (%s)", err, stderr.String())
	case <-timeout.C:
		worker.shutdown()
		return nil, errors.New("Word startup timed out after 30 seconds")
	}
}

func (w *wordWorker) shutdown() {
	if w == nil {
		return
	}
	if w.stdin != nil {
		_, _ = io.WriteString(w.stdin, "QUIT\n")
		_ = w.stdin.Close()
	}
	// There is one dedicated goroutine consuming Wait, so it is safe to
	// wait on finished here. On timeout kill the controlling PowerShell.
	select {
	case <-w.finished:
	case <-time.After(closeTimeout):
		if w.cmd.Process != nil {
			_ = w.cmd.Process.Kill()
		}
		select {
		case <-w.finished:
		case <-time.After(time.Second):
		}
	}
}

func (w *wordWorker) convert(path string) error {
	if w == nil {
		return errors.New("Word not started")
	}
	line := base64.StdEncoding.EncodeToString([]byte(path)) + "\n"
	if _, err := io.WriteString(w.stdin, line); err != nil {
		return fmt.Errorf("Word worker input: %w", err)
	}
	timer := time.NewTimer(documentTimeout)
	defer timer.Stop()
	select {
	case msg, ok := <-w.lines:
		if !ok {
			return errors.New("Word worker stopped unexpectedly")
		}
		if strings.HasPrefix(msg, "F|") {
			data, _ := base64.StdEncoding.DecodeString(strings.TrimPrefix(msg, "F|"))
			return fmt.Errorf("Word COM failed: %s", string(data))
		}
		fields := strings.SplitN(msg, "|", 7)
		if len(fields) < 7 || fields[0] != "R" {
			return fmt.Errorf("unexpected Word response: %s", msg)
		}
		if fields[2] == "1" {
			return nil
		}
		data, err := base64.StdEncoding.DecodeString(fields[6])
		if err != nil {
			return errors.New("Word export failed; invalid error response")
		}
		return fmt.Errorf("Word export failed: %s", string(data))
	case <-timer.C:
		return errors.New("document processing timed out (5 minutes)")
	}
}

func collectDocs(dir string) ([]job, error) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return nil, err
	}
	result := []job{}
	for _, e := range entries {
		if !e.Type().IsRegular() && !e.Type().IsDir() { // follow ordinary files only
			continue
		}
		if e.IsDir() {
			continue
		}
		name := e.Name()
		if strings.HasPrefix(name, "~$") || !strings.EqualFold(filepath.Ext(name), ".docx") {
			continue
		}
		path := filepath.Join(dir, name)
		result = append(result, job{name: name, original: path, working: path, output: strings.TrimSuffix(path, filepath.Ext(path)) + ".pdf"})
	}
	sort.Slice(result, func(i, k int) bool { return strings.ToLower(result[i].name) < strings.ToLower(result[k].name) })
	return result, nil
}

func volumeC(path string) bool { return strings.EqualFold(filepath.VolumeName(path), "C:") }

func allocateCache() (string, error) {
	candidates := []string{os.TempDir(), filepath.Join(os.Getenv("LOCALAPPDATA"), "Temp")}
	for _, candidate := range candidates {
		if candidate == "" || !volumeC(candidate) {
			continue
		}
		dir, err := os.MkdirTemp(candidate, "DOCX_to_PDF_*")
		if err == nil {
			return dir, nil
		}
	}
	return "", errors.New("no writable temporary directory on C: for non-C: source files")
}

func copyFile(source, target string) error {
	in, err := os.Open(source)
	if err != nil {
		return err
	}
	defer in.Close()
	out, err := os.OpenFile(target, os.O_CREATE|os.O_TRUNC|os.O_WRONLY, 0600)
	if err != nil {
		return err
	}
	_, copyErr := io.Copy(out, in)
	closeErr := out.Close()
	if copyErr != nil {
		return copyErr
	}
	return closeErr
}

func stageJobs(docs []job, dir string) ([]job, error) {
	staged := make([]job, 0, len(docs))
	for _, j := range docs {
		cachePath := filepath.Join(dir, j.name)
		if err := copyFile(j.original, cachePath); err != nil {
			return nil, fmt.Errorf("copy %q: %w", j.name, err)
		}
		j.working = cachePath
		staged = append(staged, j)
	}
	return staged, nil
}

func publishPDF(j job) error {
	from := strings.TrimSuffix(j.working, filepath.Ext(j.working)) + ".pdf"
	if j.working == j.original {
		return nil
	}
	// Publish via temporary file on the destination volume. This avoids
	// exposing a partially-copied PDF if a network write fails.
	tmp, err := os.CreateTemp(filepath.Dir(j.output), ".docx2pdf-*.pdf")
	if err != nil {
		return err
	}
	name := tmp.Name()
	_ = tmp.Close()
	defer os.Remove(name)
	if err = copyFile(from, name); err != nil {
		return err
	}
	// Windows rename does not overwrite an existing file. This is the only
	// non-atomic portion when overwriting an older PDF.
	if err = os.Remove(j.output); err != nil && !os.IsNotExist(err) {
		return err
	}
	return os.Rename(name, j.output)
}

func tryJob(j job, w **wordWorker) error {
	var last error
	for attempt := 0; attempt < 2; attempt++ {
		if *w == nil {
			newWorker, err := startWord()
			if err != nil {
				last = err
				continue
			}
			*w = newWorker
		}
		last = (*w).convert(j.working)
		if last == nil {
			return publishPDF(j)
		}
		// A worker is restarted after failed requests to isolate COM errors.
		(*w).shutdown()
		*w = nil
	}
	return last
}

func processJobs(jobs []job, words []*wordWorker) (int, int) {
	queue := make(chan job, len(jobs))
	results := make(chan outcome, len(jobs))
	for _, j := range jobs {
		queue <- j
	}
	close(queue)
	var wg sync.WaitGroup
	for _, word := range words {
		wg.Add(1)
		go func(w *wordWorker) {
			defer wg.Done()
			defer func() {
				if w != nil {
					w.shutdown()
				}
			}()
			for j := range queue {
				err := tryJob(j, &w)
				results <- outcome{job: j, err: err}
			}
		}(word)
	}
	go func() { wg.Wait(); close(results) }()
	success, failures, done := 0, 0, 0
	for r := range results {
		done++
		if r.err != nil {
			failures++
			fmt.Printf("[%d/%d] ОШИБКА %s: %v\n", done, len(jobs), r.job.name, r.err)
		} else {
			success++
			fmt.Printf("[%d/%d] OK  %s\n", done, len(jobs), r.job.name)
		}
	}
	if done < len(jobs) {
		failures += len(jobs) - done
		fmt.Printf("Не обработано файлов: %d\n", len(jobs)-done)
	}
	return success, failures
}

func run() error {
	start := time.Now()
	path, err := os.Executable()
	if err != nil {
		return err
	}
	directory := filepath.Dir(path)
	docs, err := collectDocs(directory)
	if err != nil {
		return err
	}
	fmt.Printf("Найдено DOCX: %d\n", len(docs))
	if len(docs) == 0 {
		return nil
	}
	fmt.Println("Подготовка...")
	restore, err := selectPrinter()
	if err != nil {
		fmt.Printf("Предупреждение: принтер не переключён: %v\n", err)
	}
	defer func() {
		if err := restore(); err != nil {
			fmt.Printf("Не удалось восстановить принтер: %v\n", err)
		}
	}()

	if !volumeC(directory) {
		tempDir, err := allocateCache()
		if err != nil {
			return err
		}
		defer os.RemoveAll(tempDir)
		docs, err = stageJobs(docs, tempDir)
		if err != nil {
			return err
		}
	}
	first, err := startWord()
	if err != nil {
		return fmt.Errorf("Microsoft Word не запустился: %w", err)
	}
	words := []*wordWorker{first}
	if len(docs) >= parallelMinFiles {
		second, err := startWord()
		if err == nil {
			words = append(words, second)
		} else {
			fmt.Println("Второй Word недоступен, продолжаем обработку одним экземпляром.")
		}
	}
	fmt.Println("Конвертация:")
	success, failures := processJobs(docs, words)
	elapsed := time.Since(start).Round(time.Second)
	fmt.Printf("\nГотово. Успешно: %d, ошибок: %d.\n", success, failures)
	fmt.Printf("Общее время работы: %02d:%02d\n", int(elapsed.Minutes()), int(elapsed.Seconds())%60)
	return nil
}

func main() {
	if err := run(); err != nil {
		fmt.Printf("Ошибка: %v\n", err)
	}
	fmt.Print("\nНажмите Enter для закрытия...")
	_, _ = bufio.NewReader(os.Stdin).ReadString('\n')
}
