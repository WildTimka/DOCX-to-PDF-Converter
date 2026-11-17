package main

import (
	"bytes"
	"context"
	"errors"
	"testing"
	"time"
)

func TestStartupAdmissionRemainsCancelable(t *testing.T) {
	wordStartupGate <- struct{}{}
	defer func() { <-wordStartupGate }()
	ctx, cancel := context.WithCancel(context.Background())
	done := make(chan error, 1)
	go func() {
		w, e := startWord(ctx)
		if w != nil {
			w.shutdown()
		}
		done <- e
	}()
	cancel()
	select {
	case e := <-done:
		if !errors.Is(e, context.Canceled) {
			t.Fatalf("expected cancellation without launching Word: %v", e)
		}
	case <-time.After(time.Second):
		t.Fatal("startup admission blocked cancellation")
	}
}

type startupTestInput struct{ done chan struct{} }

func (s *startupTestInput) Write(p []byte) (int, error) { return len(p), nil }
func (s *startupTestInput) Close() error                { close(s.done); return nil }
func TestCancelWaitsForStartupIdentityAndCleanup(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	finished := make(chan struct{})
	input := &startupTestInput{done: finished}
	w := &wordWorker{stdin: input, lines: make(chan string, 1), finished: finished, stderr: &bytes.Buffer{}}
	result := make(chan error, 1)
	go func() { _, err := awaitWordReady(ctx, w); result <- err }()
	w.lines <- "OWNED|0|0"
	cancel()
	select {
	case err := <-result:
		t.Fatalf("initializer returned before readiness/cleanup: %v", err)
	case <-time.After(20 * time.Millisecond):
	}
	w.lines <- "READY|0|0"
	select {
	case err := <-result:
		if !errors.Is(err, context.Canceled) {
			t.Fatalf("expected canceled result: %v", err)
		}
	case <-time.After(time.Second):
		t.Fatal("readiness did not finish cancellation")
	}
	select {
	case <-finished:
	default:
		t.Fatal("worker cleanup was not joined")
	}
}
