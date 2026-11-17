//go:build windows

package main

import (
	"os/exec"
	"syscall"
)

// A GUI application must not open a separate console for every PowerShell worker.
func hideChildWindow(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x08000000}
}
