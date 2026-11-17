//go:build !windows

package main

func wordPID(uintptr) int { return 0 }
