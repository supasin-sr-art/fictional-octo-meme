package main

import (
	"net/http"
	"net/http/httptest"
	"os"
	"os/exec"
	"path/filepath"
	"testing"
	"time"
)

func TestValidateEtlText(t *testing.T) {
	good := "1.7240899843:HP651249:HLABOR\n2.7240895649:HP651248:HLABOR"
	if err := validateEtlText(good); err != nil {
		t.Fatalf("valid ETL rejected: %v", err)
	}
	if err := validateEtlText(""); err != nil {
		t.Fatalf("empty ETL rejected: %v", err)
	}
	if err := validateEtlText("1.bad:HP1:HLABOR"); err == nil {
		t.Fatal("invalid ETL accepted")
	}
	if err := validateEtlText("1.7240899843:HP1:HLABOR\n2.7240899843:HP2:HLABOR"); err == nil {
		t.Fatal("duplicate ETL accepted")
	}
}

func TestCanonicalUploadNames(t *testing.T) {
	cases := map[string]string{
		"master":    "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก_WEB.xlsx",
		"issue":     "เช็คสถานะ ISSUE_WEB.xlsx",
		"daily":     "รายงานงานประกันแรงงานต่างด้าว_WEB.xlsx",
		"m190":      "M190027_PRD008_Premium_by_Policy_WEB.xlsx",
		"sm":        "ข้อมูลไม่สมบูรณ์_WEB.xlsx",
		"blacklist": "Blacklist_WEB.xls",
	}
	for field, want := range cases {
		got := canonicalUploadName(field, "sample.xls")
		if got != want {
			t.Fatalf("%s: got %q want %q", field, got, want)
		}
	}
}

func TestParseWorkflowSummary(t *testing.T) {
	lines := []string{"log", "{", `  "DateStart": "2026-06-01",`, `  "M190PropIdRows": 123,`, `  "PvProposalFilterApplied": true`, "}"}
	s := parseWorkflowSummary(lines)
	if s["DateStart"] != "2026-06-01" || s["M190PropIdRows"] != "123" || s["PvProposalFilterApplied"] != "true" {
		t.Fatalf("unexpected summary: %#v", s)
	}
}

func TestReadTSVHeaderOnly(t *testing.T) {
	d := t.TempDir()
	p := filepath.Join(d, "empty.tsv")
	header := "ProposalID\tPolicy\tTotalPremium\tAgingDays\n"
	if err := os.WriteFile(p, []byte(header), 0644); err != nil {
		t.Fatal(err)
	}
	rows, err := readTSV(p)
	if err != nil {
		t.Fatal(err)
	}
	if len(rows) != 0 {
		t.Fatalf("got %d rows", len(rows))
	}
}

func TestCancelDeleteRunDeletesOnlyTarget(t *testing.T) {
	base := t.TempDir()
	targetDir := filepath.Join(base, "target")
	otherDir := filepath.Join(base, "other")
	if err := os.MkdirAll(targetDir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(otherDir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(targetDir, "temp.txt"), []byte("x"), 0644); err != nil {
		t.Fatal(err)
	}
	if err := os.WriteFile(filepath.Join(otherDir, "keep.txt"), []byte("y"), 0644); err != nil {
		t.Fatal(err)
	}

	s := &Server{
		runs: map[string]*Run{
			"target": {ID: "target", Status: "failed", Dir: targetDir},
			"other":  {ID: "other", Status: "completed", Dir: otherDir},
		},
		activeCmds: map[string]*exec.Cmd{}, cancelledRuns: map[string]bool{}, deletedRuns: map[string]bool{},
	}
	req := httptest.NewRequest(http.MethodPost, "/api/cancel-delete-run?id=target", nil)
	w := httptest.NewRecorder()
	s.handleCancelDeleteRun(w, req)
	if w.Code != http.StatusOK {
		t.Fatalf("status=%d body=%s", w.Code, w.Body.String())
	}

	deadline := time.Now().Add(2 * time.Second)
	for time.Now().Before(deadline) {
		if _, err := os.Stat(targetDir); os.IsNotExist(err) {
			break
		}
		time.Sleep(20 * time.Millisecond)
	}
	if _, err := os.Stat(targetDir); !os.IsNotExist(err) {
		t.Fatalf("target dir still exists: %v", err)
	}
	if _, err := os.Stat(filepath.Join(otherDir, "keep.txt")); err != nil {
		t.Fatalf("other run was affected: %v", err)
	}
	if s.runs["target"] != nil {
		t.Fatal("target remains in history")
	}
	if s.runs["other"] == nil {
		t.Fatal("other run removed")
	}
}

func TestRunDisplayName(t *testing.T) {
	started := time.Date(2026, time.June, 23, 8, 33, 55, 0, time.Local)
	run := &Run{StartedAt: started}
	want := "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก 2026-06-23"
	if got := runDisplayName(run); got != want {
		t.Fatalf("got %q want %q", got, want)
	}
}

func TestCleanupExpiredRunsKeepsRecentRun(t *testing.T) {
	base := t.TempDir()
	expiredDir := filepath.Join(base, "expired")
	recentDir := filepath.Join(base, "recent")
	if err := os.MkdirAll(expiredDir, 0755); err != nil {
		t.Fatal(err)
	}
	if err := os.MkdirAll(recentDir, 0755); err != nil {
		t.Fatal(err)
	}

	s := &Server{
		runs: map[string]*Run{
			"expired": {ID: "expired", StartedAt: time.Now().Add(-15 * 24 * time.Hour), Dir: expiredDir},
			"recent":  {ID: "recent", StartedAt: time.Now().Add(-13 * 24 * time.Hour), Dir: recentDir},
		},
		activeCmds: map[string]*exec.Cmd{}, cancelledRuns: map[string]bool{}, deletedRuns: map[string]bool{},
	}
	if got := s.cleanupExpiredRuns(); got != 1 {
		t.Fatalf("deleted=%d want 1", got)
	}
	if s.runs["expired"] != nil {
		t.Fatal("expired run remains in history")
	}
	if s.runs["recent"] == nil {
		t.Fatal("recent run was removed")
	}
	if _, err := os.Stat(expiredDir); !os.IsNotExist(err) {
		t.Fatalf("expired dir still exists: %v", err)
	}
	if _, err := os.Stat(recentDir); err != nil {
		t.Fatalf("recent dir was affected: %v", err)
	}
}
