package main

import (
	"bytes"
	"context"
	"errors"
	"fmt"
	"strings"
	"sync"
	"testing"
	"time"
)

type poolTestInput struct {
	lines    chan string
	finished chan struct{}
	delay    time.Duration
	once     sync.Once
}

func (s *poolTestInput) Write(p []byte) (int, error) {
	if !strings.HasPrefix(string(p), "QUIT") {
		go func() {
			select {
			case <-time.After(s.delay):
				select {
				case s.lines <- "R|c2FtcGxl|1|1|1|1|":
				case <-s.finished:
				}
			case <-s.finished:
			}
		}()
	}
	return len(p), nil
}
func (s *poolTestInput) Close() error { s.once.Do(func() { close(s.finished) }); return nil }
func poolTestWorker(delay time.Duration) *wordWorker {
	input := &poolTestInput{lines: make(chan string, 1), finished: make(chan struct{}), delay: delay}
	return &wordWorker{stdin: input, lines: input.lines, finished: input.finished, stderr: &bytes.Buffer{}}
}
func poolTestJobs(n int) []job {
	jobs := make([]job, n)
	for i := range jobs {
		jobs[i].name = fmt.Sprint(i)
	}
	return jobs
}
func TestGrowingPoolKeepsQueueAndFallsBack(t *testing.T) {
	first := poolTestWorker(2 * time.Millisecond)
	called := false
	successful, failed := processJobsWithStartup(context.Background(), poolTestJobs(20), []*wordWorker{first}, 4, time.Millisecond, func(context.Context) (*wordWorker, error) {
		called = true
		return nil, errors.New("initialization failed")
	}, func(int, int) {})
	if !called || failed != 0 || len(successful) != 20 {
		t.Fatalf("fallback lost work: called=%v success=%d failed=%d", called, len(successful), failed)
	}
	seen := map[string]bool{}
	for _, j := range successful {
		if seen[j.name] {
			t.Fatal("job converted twice")
		}
		seen[j.name] = true
	}
	select {
	case <-first.finished:
	default:
		t.Fatal("fallback worker not joined")
	}
}
func TestCancellationJoinsInitializerAndReturnedWorker(t *testing.T) {
	ctx, cancel := context.WithCancel(context.Background())
	defer cancel()
	first := poolTestWorker(time.Second)
	entered := make(chan struct{})
	release := make(chan struct{})
	returned := poolTestWorker(time.Millisecond)
	done := make(chan struct{})
	go func() {
		processJobsWithStartup(ctx, poolTestJobs(20), []*wordWorker{first}, 4, time.Millisecond, func(context.Context) (*wordWorker, error) { close(entered); <-release; return returned, nil }, func(int, int) {})
		close(done)
	}()
	select {
	case <-entered:
	case <-time.After(time.Second):
		t.Fatal("initializer did not start")
	}
	cancel()
	select {
	case <-done:
		t.Fatal("pool returned before initializer cleanup")
	case <-time.After(20 * time.Millisecond):
	}
	close(release)
	select {
	case <-done:
	case <-time.After(time.Second):
		t.Fatal("pool did not join canceled initialization")
	}
	for _, w := range []*wordWorker{first, returned} {
		select {
		case <-w.finished:
		default:
			t.Fatal("a canceled worker was not closed")
		}
	}
}
func TestStartupCostGuardAvoidsAnIdleWorker(t *testing.T) {
	if worthStartingWorker(1, 3, 10, time.Second, 2*time.Second) {
		t.Fatal("started a worker for a queue that drains before initialization")
	}
	if !worthStartingWorker(100, 3, 10, 40*time.Second, 2*time.Second) {
		t.Fatal("rejected useful initialization for substantial remaining work")
	}
	if worthStartingWorker(0, 1, 0, 0, time.Second) {
		t.Fatal("empty queue started a worker")
	}
}

func TestGrowingPoolDistributesWorkAndJoinsEveryWorker(t *testing.T) {
	workers := []*wordWorker{poolTestWorker(10 * time.Millisecond)}
	starts := 0
	lastProgress := 0
	ok, failed := processJobsWithStartup(context.Background(), poolTestJobs(60), workers, 3, time.Millisecond,
		func(context.Context) (*wordWorker, error) {
			starts++
			w := poolTestWorker(10 * time.Millisecond)
			workers = append(workers, w)
			return w, nil
		}, func(done, total int) {
			if done != lastProgress+1 || total != 60 {
				t.Errorf("invalid progress: done=%d previous=%d total=%d", done, lastProgress, total)
			}
			lastProgress = done
		})
	if starts != 2 || failed != 0 || len(ok) != 60 || lastProgress != 60 {
		t.Fatalf("growth: starts=%d failed=%d successes=%d progress=%d", starts, failed, len(ok), lastProgress)
	}
	seen := map[string]bool{}
	for _, j := range ok {
		if seen[j.name] {
			t.Fatalf("duplicate job %s", j.name)
		}
		seen[j.name] = true
	}
	for _, w := range workers {
		select {
		case <-w.finished:
		default:
			t.Fatal("pool returned before closing a worker")
		}
	}
}
