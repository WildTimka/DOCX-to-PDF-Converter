package main

import "runtime"

// Calibrated on the supplied 176-document workload, i7-6700K / 8 logical CPUs.
// Use the fewest workers within 3% of the measured best, keeping targets
// monotone. These tested points are not universal exact break-even boundaries.
// RELEASE_NOTES_v1.1.0.md summarizes the evidence and limits.
func calibratedWorkerTarget(documents int) int {
	switch {
	case documents <= 0:
		return 0
	case documents == 1:
		return 1
	case documents == 2:
		return 2
	case documents < 6:
		return 3
	case documents < 30:
		return 4
	case documents < 60:
		return 5
	case documents < 80:
		return 6
	default:
		return 7
	}
}

func adaptiveWorkerLimit(documents int) int {
	available, known := availableMemoryBytes()
	return limitWordWorkers(calibratedWorkerTarget(documents), runtime.NumCPU(), available, known)
}

