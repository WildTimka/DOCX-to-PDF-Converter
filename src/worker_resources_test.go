package main

import "testing"

func TestWorkerResourceAdmission(t *testing.T) {
	cases := []struct {
		name         string
		target, cpus int
		available    uint64
		known        bool
		want         int
	}{
		{"empty", 0, 8, 16 << 30, true, 0},
		{"negative", -1, 8, 16 << 30, true, 0},
		{"cpu", 8, 4, 16 << 30, true, 4},
		{"one CPU", 8, 1, 16 << 30, true, 1},
		{"unknown CPU", 8, 0, 16 << 30, true, 1},
		{"memory", 8, 8, 4 << 30, true, 4},
		{"exact one slot", 8, 8, wordMemoryReserve + wordMemoryBudget, true, 1},
		{"below reserve", 8, 8, 1 << 30, true, 1},
		{"unknown memory", 8, 8, 16 << 30, false, 1},
		{"target preserved", 3, 8, 16 << 30, true, 3},
		{"large memory no overflow", 8, 8, ^uint64(0), true, 8},
	}
	for _, c := range cases {
		t.Run(c.name, func(t *testing.T) {
			if got := limitWordWorkers(c.target, c.cpus, c.available, c.known); got != c.want {
				t.Fatalf("got %d, want %d", got, c.want)
			}
		})
	}
}
