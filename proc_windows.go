//go:build windows

package main

import (
	"os/exec"
	"strconv"
	"strings"
	"syscall"
)

func configureHidden(cmd *exec.Cmd) {
	cmd.SysProcAttr = &syscall.SysProcAttr{HideWindow: true, CreationFlags: 0x08000000}
}

func killProcessTree(cmd *exec.Cmd) {
	if cmd == nil || cmd.Process == nil {
		return
	}
	killer := exec.Command("taskkill.exe", "/PID", strconv.Itoa(cmd.Process.Pid), "/T", "/F")
	configureHidden(killer)
	_ = killer.Run()
	_ = cmd.Process.Kill()
}

func killExcelProcess(pid int) {
	if pid <= 0 {
		return
	}
	check := exec.Command("tasklist.exe", "/FI", "PID eq "+strconv.Itoa(pid), "/FO", "CSV", "/NH")
	configureHidden(check)
	out, err := check.Output()
	if err != nil || !strings.Contains(strings.ToUpper(string(out)), "EXCEL.EXE") {
		return
	}
	killer := exec.Command("taskkill.exe", "/PID", strconv.Itoa(pid), "/T", "/F")
	configureHidden(killer)
	_ = killer.Run()
}
