//go:build !windows

package main

func availableMemoryBytes() (uint64, bool) { return 0, false }
