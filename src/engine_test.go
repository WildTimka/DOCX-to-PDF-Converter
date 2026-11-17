package main

import (
	"context"
	"os"
	"os/exec"
	"path/filepath"
	"strings"
	"testing"
	"time"
)

func mustWrite(t *testing.T, path, data string) {
	t.Helper()
	if err := os.WriteFile(path, []byte(data), 0600); err != nil {
		t.Fatal(err)
	}
}

func TestCollectDocs(t *testing.T) {
	dir := t.TempDir()
	for _, n := range []string{"b.DOCX", "A.docx", "~$temp.docx", "other.txt"} {
		mustWrite(t, filepath.Join(dir, n), "x")
	}
	docs, err := collectDocs(dir)
	if err != nil {
		t.Fatal(err)
	}
	if len(docs) != 2 || docs[0].name != "A.docx" || docs[1].name != "b.DOCX" {
		t.Fatalf("unexpected files: %+v", docs)
	}
}
func TestPublishPreservesExistingBeforeCommit(t *testing.T) {
	dir := t.TempDir()
	old := filepath.Join(dir, "a.pdf")
	stage := filepath.Join(dir, "converted.pdf")
	mustWrite(t, old, "existing")
	mustWrite(t, stage, "converted")
	ctx, cancel := context.WithCancel(context.Background())
	cancel()
	count, err := publishJobs(ctx, []job{{output: old, stagedPDF: stage}})
	if count != 0 || err == nil {
		t.Fatalf("expected cancel: count=%d err=%v", count, err)
	}
	b, _ := os.ReadFile(old)
	if string(b) != "existing" {
		t.Fatalf("existing PDF changed: %q", b)
	}
}
func TestPublishOverwritesAfterAllComplete(t *testing.T) {
	dir := t.TempDir()
	old := filepath.Join(dir, "old.pdf")
	newPath := filepath.Join(dir, "new.pdf")
	src1 := filepath.Join(dir, "src1")
	src2 := filepath.Join(dir, "src2")
	mustWrite(t, old, "old")
	mustWrite(t, src1, "fresh")
	mustWrite(t, src2, "second")
	count, err := publishJobs(context.Background(), []job{{output: old, stagedPDF: src1}, {output: newPath, stagedPDF: src2}})
	if err != nil || count != 2 {
		t.Fatalf("count=%d err=%v", count, err)
	}
	for path, want := range map[string]string{old: "fresh", newPath: "second"} {
		b, e := os.ReadFile(path)
		if e != nil || string(b) != want {
			t.Fatalf("%s => %q, %v", path, b, e)
		}
	}
	files, _ := os.ReadDir(dir)
	for _, f := range files {
		if strings.HasPrefix(f.Name(), ".docx2pdf-") {
			t.Errorf("leftover file %s", f.Name())
		}
	}
}
func TestPublishRollbackOnFailure(t *testing.T) {
	dir := t.TempDir()
	orig := filepath.Join(dir, "existing.pdf")
	another := filepath.Join(dir, "another.pdf")
	src := filepath.Join(dir, "converted.pdf")
	mustWrite(t, orig, "untouched original")
	mustWrite(t, src, "new")
	count, err := publishJobs(context.Background(), []job{{output: orig, stagedPDF: src}, {output: another, stagedPDF: filepath.Join(dir, "missing.pdf")}})
	if err == nil || count != 0 {
		t.Fatalf("expected rollback: count %d error %v", count, err)
	}
	b, _ := os.ReadFile(orig)
	if string(b) != "untouched original" {
		t.Fatalf("rollback damaged original: %q", b)
	}
	if _, err := os.Stat(another); !os.IsNotExist(err) {
		t.Fatalf("unexpected extra PDF: %v", err)
	}
}

// Helper process used to simulate a WINWORD.EXE that ignored Word.Quit().
// The parent test owns the helper PID and must terminate only that process.
func TestTrackedWordHelper(t *testing.T) {
	if os.Getenv("DOCX_TO_PDF_TRACKED_WORD_HELPER") != "1" {
		return
	}
	time.Sleep(30 * time.Second)
}

func TestShutdownTerminatesOnlyTrackedWord(t *testing.T) {
	cmd := exec.Command(os.Args[0], "-test.run=^TestTrackedWordHelper$")
	cmd.Env = append(os.Environ(), "DOCX_TO_PDF_TRACKED_WORD_HELPER=1")
	if err := cmd.Start(); err != nil {
		t.Fatal(err)
	}
	defer func() { _ = cmd.Process.Kill() }()
	exited := make(chan struct{})
	go func() { _, _ = cmd.Process.Wait(); close(exited) }()
	owned, err := openTrackedWord(cmd.Process.Pid)
	if err != nil {
		t.Fatal(err)
	}
	finished := make(chan struct{})
	close(finished) // Simulate exited PowerShell while Word remains running.
	w := &wordWorker{finished: finished, process: owned}
	if w.waitWord(0) {
		t.Fatal("test worker exited before cleanup started")
	}
	w.shutdown()
	select {
	case <-exited:
	case <-time.After(2 * time.Second):
		t.Fatal("orphan process was not terminated by worker shutdown")
	}
}

// Hidden Word automation may expose no top-level window. This must not be a
// fatal startup condition; an independently verified COM PID may be used.
func TestParseWordReadyWithoutWindow(t *testing.T) {
	tests := []struct {
		message string
		hwnd    uintptr
		pid     int
	}{
		{"READY|0|7654", 0, 7654},
		{"READY|0|0", 0, 0},
		{"READY|0", 0, 0},
		{"READY|123456|7654", 123456, 7654},
	}
	for _, tc := range tests {
		hwnd, pid, err := parseWordReady(tc.message)
		if err != nil || hwnd != tc.hwnd || pid != tc.pid {
			t.Errorf("parseWordReady(%q)=(%d,%d,%v), want (%d,%d,nil)", tc.message, hwnd, pid, err, tc.hwnd, tc.pid)
		}
	}
	if _, _, err := parseWordReady("INVALID|0|123"); err == nil {
		t.Fatal("unexpectedly accepted an invalid readiness message")
	}
}
