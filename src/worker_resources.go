package main

const (
	wordMemoryReserve = uint64(2 << 30)
	wordMemoryBudget  = uint64(512 << 20)
)

// Admission limits are conservative safety bounds, not fitted throughput estimates.
// The existing first worker remains usable even under low/unknown memory.
func limitWordWorkers(target, logicalCPUs int, available uint64, known bool) int {
	if target <= 0 {
		return 0
	}
	if logicalCPUs < 1 || !known || available < wordMemoryReserve+wordMemoryBudget {
		return 1
	}
	if target > logicalCPUs {
		target = logicalCPUs
	}
	slots := (available - wordMemoryReserve) / wordMemoryBudget
	if slots < uint64(target) {
		target = int(slots)
	}
	return target
}

func roomForAdditionalWord() bool {
	available, known := availableMemoryBytes()
	return known && available >= wordMemoryReserve+wordMemoryBudget
}
