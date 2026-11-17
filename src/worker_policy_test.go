package main

import "testing"

func TestCalibratedPolicyBoundaries(t *testing.T) {
	for _, c := range []struct{ n, want int }{
		{-1, 0}, {0, 0}, {1, 1}, {2, 2}, {3, 3}, {5, 3}, {6, 4},
		{20, 4}, {25, 4}, {29, 4}, {30, 5}, {59, 5}, {60, 6},
		{79, 6}, {80, 7}, {120, 7}, {176, 7}, {10000, 7},
	} {
		if got := calibratedWorkerTarget(c.n); got != c.want {
			t.Errorf("documents=%d: got %d want %d", c.n, got, c.want)
		}
	}
}

func TestCalibratedPolicyIsBoundedAndMonotone(t *testing.T) {
	previous := 0
	for n := 1; n <= 10000; n++ {
		got := calibratedWorkerTarget(n)
		if got < previous || got < 1 || got > 7 || got > n {
			t.Fatalf("invalid target %d for %d documents (previous %d)", got, n, previous)
		}
		previous = got
	}
}
