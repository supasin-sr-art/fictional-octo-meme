//go:build !windows

package main

import "os/exec"

func configureHidden(cmd *exec.Cmd) {}
func killProcessTree(cmd *exec.Cmd) {
	if cmd != nil && cmd.Process != nil {
		_ = cmd.Process.Kill()
	}
}

func killExcelProcess(pid int) {}
