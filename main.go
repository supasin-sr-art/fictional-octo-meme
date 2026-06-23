package main

import (
	"archive/zip"
	"bufio"
	"context"
	"embed"
	"encoding/csv"
	"encoding/json"
	"errors"
	"fmt"
	"io"
	"io/fs"
	"log"
	"mime/multipart"
	"net"
	"net/http"
	"net/url"
	"os"
	"os/exec"
	"path/filepath"
	"runtime"
	"sort"
	"strconv"
	"strings"
	"sync"
	"time"
)

//go:embed web/* workflow.ps1 extract_result.ps1
var embedded embed.FS

const (
	runRetentionDays = 14
	runRetention     = runRetentionDays * 24 * time.Hour
)

type DetailRow struct {
	ProposalID    string  `json:"proposalId"`
	Policy        string  `json:"policy"`
	CertificateNo string  `json:"certificateNo"`
	AgencyCode    string  `json:"agencyCode"`
	Mticode       string  `json:"mticode"`
	AgencyName    string  `json:"agencyName"`
	RequestCode   string  `json:"requestCode"`
	EmployeeName  string  `json:"employeeName"`
	AlienCode     string  `json:"alienCode"`
	AlienName     string  `json:"alienName"`
	TotalPremium  float64 `json:"totalPremium"`
	CreateDate    string  `json:"createDate"`
	SourceStatus  string  `json:"sourceStatus"`
	EPropID       string  `json:"ePropId"`
	Discount      string  `json:"discount"`
	PendingStatus string  `json:"pendingStatus"`
	AgingDays     int     `json:"agingDays"`
	PendingRange  string  `json:"pendingRange"`
	Incomplete    string  `json:"incomplete"`
	Blacklist     string  `json:"blacklist"`
	MenuE         string  `json:"menuE"`
}

type Run struct {
	ID         string            `json:"id"`
	Status     string            `json:"status"`
	Message    string            `json:"message"`
	Progress   int               `json:"progress"`
	StartedAt  time.Time         `json:"startedAt"`
	FinishedAt *time.Time        `json:"finishedAt,omitempty"`
	Logs       []string          `json:"logs,omitempty"`
	Summary    map[string]string `json:"summary,omitempty"`
	Rows       []DetailRow       `json:"-"`
	Files      map[string]string `json:"files,omitempty"`
	Error      string            `json:"error,omitempty"`
	Dir        string            `json:"-"`
	InputNames map[string]string `json:"inputNames,omitempty"`
}

type RunHistoryItem struct {
	Run
	DisplayName      string    `json:"displayName"`
	ExpiresAt        time.Time `json:"expiresAt"`
	RemainingSeconds int64     `json:"remainingSeconds"`
	RetentionDays    int       `json:"retentionDays"`
}

type AppConfig struct {
	Language   string `json:"language"`
	Theme      string `json:"theme"`
	StorageDir string `json:"storageDir"`
}

type Server struct {
	mu            sync.RWMutex
	runs          map[string]*Run
	configMu      sync.RWMutex
	config        AppConfig
	cmdMu         sync.Mutex
	activeCmds    map[string]*exec.Cmd
	cancelledRuns map[string]bool
	deletedRuns   map[string]bool
	baseDir       string
	workflowPath  string
	extractPath   string
	port          int
}

func main() {
	baseDir := appDataDir()
	if err := os.MkdirAll(filepath.Join(baseDir, "Runs"), 0o755); err != nil {
		log.Fatal(err)
	}
	workflowPath := filepath.Join(baseDir, "workflow.ps1")
	extractPath := filepath.Join(baseDir, "extract_result.ps1")
	mustExtract("workflow.ps1", workflowPath)
	mustExtract("extract_result.ps1", extractPath)

	s := &Server{runs: map[string]*Run{}, config: loadAppConfig(baseDir), activeCmds: map[string]*exec.Cmd{}, cancelledRuns: map[string]bool{}, deletedRuns: map[string]bool{}, baseDir: baseDir, workflowPath: workflowPath, extractPath: extractPath}
	s.loadHistory()
	s.cleanupExpiredRuns()
	go s.runRetentionCleanupLoop()

	ln, port, err := listenAvailable(8765, 20)
	if err != nil {
		log.Fatal(err)
	}
	s.port = port
	_ = os.WriteFile(filepath.Join(baseDir, "server-port.txt"), []byte(strconv.Itoa(port)), 0o644)

	mux := http.NewServeMux()
	mux.HandleFunc("/api/health", s.handleHealth)
	mux.HandleFunc("/api/run", s.handleRun)
	mux.HandleFunc("/api/status", s.handleStatus)
	mux.HandleFunc("/api/result", s.handleResult)
	mux.HandleFunc("/api/history", s.handleHistory)
	mux.HandleFunc("/api/download", s.handleDownload)
	mux.HandleFunc("/api/open-output", s.handleOpenOutput)
	mux.HandleFunc("/api/cancel", s.handleCancel)
	mux.HandleFunc("/api/delete-run", s.handleDeleteRun)
	mux.HandleFunc("/api/cancel-delete-run", s.handleCancelDeleteRun)
	mux.HandleFunc("/api/shutdown", s.handleShutdown)
	mux.HandleFunc("/api/settings", s.handleSettings)
	mux.HandleFunc("/api/select-storage", s.handleSelectStorage)
	mux.HandleFunc("/api/open-storage", s.handleOpenStorage)
	mux.HandleFunc("/api/system-status", s.handleSystemStatus)
	mux.HandleFunc("/api/manual-status", s.handleManualStatus)
	mux.HandleFunc("/api/manual", s.handleManualDownload)
	mux.HandleFunc("/api/reset", s.handleResetProgram)

	webFS, err := fs.Sub(embedded, "web")
	if err != nil {
		log.Fatal(err)
	}
	mux.Handle("/", http.FileServer(http.FS(webFS)))

	url := fmt.Sprintf("http://127.0.0.1:%d", port)
	if runtime.GOOS == "windows" {
		go func() {
			time.Sleep(600 * time.Millisecond)
			_ = exec.Command("rundll32", "url.dll,FileProtocolHandler", url).Start()
		}()
	}
	log.Printf("BLACKWOLF Web V2.6.5 Run Retention Engine started: %s", url)
	srv := &http.Server{Handler: securityHeaders(mux), ReadHeaderTimeout: 20 * time.Second}
	if err := srv.Serve(ln); err != nil && !errors.Is(err, http.ErrServerClosed) {
		log.Fatal(err)
	}
}

func mustExtract(name, path string) {
	b, err := embedded.ReadFile(name)
	if err != nil {
		log.Fatal(err)
	}
	if err := os.WriteFile(path, b, 0o644); err != nil {
		log.Fatal(err)
	}
}
func appDataDir() string {
	if v := os.Getenv("LOCALAPPDATA"); v != "" {
		return filepath.Join(v, "BLACKWOLF_WEB_V2")
	}
	exe, _ := os.Executable()
	return filepath.Join(filepath.Dir(exe), "BLACKWOLF_WEB_DATA")
}
func listenAvailable(start, attempts int) (net.Listener, int, error) {
	for p := start; p < start+attempts; p++ {
		if ln, err := net.Listen("tcp", fmt.Sprintf("127.0.0.1:%d", p)); err == nil {
			return ln, p, nil
		}
	}
	return nil, 0, fmt.Errorf("ไม่พบพอร์ตว่างสำหรับเปิด BLACKWOLF Web")
}
func securityHeaders(next http.Handler) http.Handler {
	return http.HandlerFunc(func(w http.ResponseWriter, r *http.Request) {
		w.Header().Set("X-Content-Type-Options", "nosniff")
		w.Header().Set("X-Frame-Options", "DENY")
		w.Header().Set("Referrer-Policy", "no-referrer")
		w.Header().Set("Cache-Control", "no-store")
		next.ServeHTTP(w, r)
	})
}

func defaultAppConfig() AppConfig {
	return AppConfig{Language: "th", Theme: "light", StorageDir: ""}
}

func normalizeAppConfig(cfg AppConfig) AppConfig {
	if cfg.Language != "en" {
		cfg.Language = "th"
	}
	if cfg.Theme != "dark" {
		cfg.Theme = "light"
	}
	cfg.StorageDir = strings.TrimSpace(cfg.StorageDir)
	return cfg
}

func loadAppConfig(baseDir string) AppConfig {
	cfg := defaultAppConfig()
	b, err := os.ReadFile(filepath.Join(baseDir, "settings.json"))
	if err == nil {
		_ = json.Unmarshal(b, &cfg)
	}
	return normalizeAppConfig(cfg)
}

func (s *Server) saveConfigLocked() error {
	s.config = normalizeAppConfig(s.config)
	b, err := json.MarshalIndent(s.config, "", "  ")
	if err != nil {
		return err
	}
	return os.WriteFile(filepath.Join(s.baseDir, "settings.json"), b, 0o644)
}

func (s *Server) configSnapshot() AppConfig {
	s.configMu.RLock()
	defer s.configMu.RUnlock()
	return s.config
}

func writableDirectory(path string) error {
	path = strings.TrimSpace(path)
	if path == "" {
		return fmt.Errorf("ไม่ได้เลือกโฟลเดอร์")
	}
	if err := os.MkdirAll(path, 0o755); err != nil {
		return err
	}
	probe := filepath.Join(path, fmt.Sprintf(".blackwolf-write-test-%d.tmp", time.Now().UnixNano()))
	if err := os.WriteFile(probe, []byte("ok"), 0o600); err != nil {
		return err
	}
	return os.Remove(probe)
}

func (s *Server) handleSettings(w http.ResponseWriter, r *http.Request) {
	switch r.Method {
	case http.MethodGet:
		cfg := s.configSnapshot()
		jsonResponse(w, map[string]any{"ok": true, "settings": cfg, "baseDir": s.baseDir})
	case http.MethodPost:
		var input AppConfig
		dec := json.NewDecoder(io.LimitReader(r.Body, 1<<20))
		if err := dec.Decode(&input); err != nil {
			jsonError(w, "อ่านการตั้งค่าไม่สำเร็จ: "+err.Error(), http.StatusBadRequest)
			return
		}
		input = normalizeAppConfig(input)
		if input.StorageDir != "" {
			if err := writableDirectory(input.StorageDir); err != nil {
				jsonError(w, "โฟลเดอร์จัดเก็บใช้งานไม่ได้: "+err.Error(), http.StatusBadRequest)
				return
			}
		}
		s.configMu.Lock()
		s.config = input
		err := s.saveConfigLocked()
		s.configMu.Unlock()
		if err != nil {
			jsonError(w, "บันทึกการตั้งค่าไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
			return
		}
		jsonResponse(w, map[string]any{"ok": true, "settings": input})
	default:
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
	}
}

func (s *Server) handleSelectStorage(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	if runtime.GOOS != "windows" {
		jsonError(w, "การเลือกโฟลเดอร์รองรับเฉพาะ Windows", http.StatusBadRequest)
		return
	}
	powerShellPath, ok := findWindowsPowerShell()
	if !ok {
		jsonError(w, "ไม่พบ Windows PowerShell สำหรับเปิดหน้าต่างเลือกโฟลเดอร์", http.StatusInternalServerError)
		return
	}
	script := `[Console]::OutputEncoding = [System.Text.Encoding]::UTF8
Add-Type -AssemblyName System.Windows.Forms
$dialog = New-Object System.Windows.Forms.FolderBrowserDialog
$dialog.Description = 'เลือกโฟลเดอร์จัดเก็บผลลัพธ์ BLACKWOLF'
$dialog.ShowNewFolderButton = $true
if ($dialog.ShowDialog() -eq [System.Windows.Forms.DialogResult]::OK) {
    Write-Output $dialog.SelectedPath
    exit 0
}
exit 2`
	cmd := exec.Command(powerShellPath, "-NoLogo", "-NoProfile", "-NonInteractive", "-Sta", "-ExecutionPolicy", "Bypass", "-Command", script)
	configureHidden(cmd)
	out, err := cmd.Output()
	if err != nil {
		if exitErr, ok := err.(*exec.ExitError); ok && exitErr.ExitCode() == 2 {
			jsonResponse(w, map[string]any{"ok": true, "cancelled": true})
			return
		}
		jsonError(w, "เปิดหน้าต่างเลือกโฟลเดอร์ไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
		return
	}
	selected := strings.TrimSpace(strings.TrimPrefix(string(out), "\ufeff"))
	if err := writableDirectory(selected); err != nil {
		jsonError(w, "โฟลเดอร์จัดเก็บใช้งานไม่ได้: "+err.Error(), http.StatusBadRequest)
		return
	}
	s.configMu.Lock()
	s.config.StorageDir = selected
	err = s.saveConfigLocked()
	s.configMu.Unlock()
	if err != nil {
		jsonError(w, "บันทึกโฟลเดอร์ไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
		return
	}
	jsonResponse(w, map[string]any{"ok": true, "path": selected})
}

func (s *Server) handleOpenStorage(w http.ResponseWriter, r *http.Request) {
	cfg := s.configSnapshot()
	path := cfg.StorageDir
	if path == "" {
		path = filepath.Join(s.baseDir, "Runs")
	}
	if err := os.MkdirAll(path, 0o755); err != nil {
		jsonError(w, "เปิดโฟลเดอร์ไม่ได้: "+err.Error(), http.StatusInternalServerError)
		return
	}
	if runtime.GOOS == "windows" {
		_ = exec.Command("explorer.exe", path).Start()
	}
	jsonResponse(w, map[string]any{"ok": true, "path": path})
}

func (s *Server) findManualPDF() string {
	candidates := []string{}
	if exe, err := os.Executable(); err == nil {
		candidates = append(candidates, filepath.Join(filepath.Dir(exe), "BLACKWOLF_User_Manual.pdf"))
	}
	candidates = append(candidates, filepath.Join(s.baseDir, "BLACKWOLF_User_Manual.pdf"))
	for _, candidate := range candidates {
		if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
			return candidate
		}
	}
	return ""
}

func (s *Server) handleManualStatus(w http.ResponseWriter, r *http.Request) {
	path := s.findManualPDF()
	jsonResponse(w, map[string]any{"ok": true, "available": path != "", "filename": filepath.Base(path)})
}

func (s *Server) handleManualDownload(w http.ResponseWriter, r *http.Request) {
	path := s.findManualPDF()
	if path == "" {
		jsonError(w, "ยังไม่ได้เพิ่มไฟล์คู่มือ BLACKWOLF_User_Manual.pdf", http.StatusNotFound)
		return
	}
	w.Header().Set("Content-Disposition", "attachment; filename=BLACKWOLF_User_Manual.pdf")
	w.Header().Set("Content-Type", "application/pdf")
	http.ServeFile(w, r, path)
}

func (s *Server) handleSystemStatus(w http.ResponseWriter, r *http.Request) {
	ready, checks := engineHealth()
	cfg := s.configSnapshot()
	s.mu.RLock()
	totalRuns := len(s.runs)
	activeRuns := 0
	for _, run := range s.runs {
		if run.Status == "running" || run.Status == "queued" {
			activeRuns++
		}
	}
	s.mu.RUnlock()
	storage := cfg.StorageDir
	if storage == "" {
		storage = filepath.Join(s.baseDir, "Runs")
	}
	storageWritable := writableDirectory(storage) == nil
	jsonResponse(w, map[string]any{
		"ok": true, "engineReady": ready, "checks": checks,
		"storagePath": storage, "customStorage": cfg.StorageDir != "", "storageWritable": storageWritable,
		"manualAvailable": s.findManualPDF() != "", "activeRuns": activeRuns, "totalRuns": totalRuns,
		"version": "2.6.5", "baseDir": s.baseDir,
	})
}

func (s *Server) handleResetProgram(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	// หยุดเฉพาะ Process ที่ BLACKWOLF เปิด แล้วล้างเฉพาะข้อมูลภายใน AppData ของโปรแกรม
	s.cmdMu.Lock()
	cmds := make([]*exec.Cmd, 0, len(s.activeCmds))
	for id, cmd := range s.activeCmds {
		s.cancelledRuns[id] = true
		s.deletedRuns[id] = true
		cmds = append(cmds, cmd)
	}
	s.cmdMu.Unlock()
	for _, cmd := range cmds {
		killProcessTree(cmd)
	}
	s.mu.Lock()
	s.runs = map[string]*Run{}
	s.mu.Unlock()
	_ = os.RemoveAll(filepath.Join(s.baseDir, "Runs"))
	if err := os.MkdirAll(filepath.Join(s.baseDir, "Runs"), 0o755); err != nil {
		jsonError(w, "รีเซ็ตโฟลเดอร์ Run ไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
		return
	}
	s.configMu.Lock()
	s.config = defaultAppConfig()
	err := s.saveConfigLocked()
	s.configMu.Unlock()
	if err != nil {
		jsonError(w, "รีเซ็ตการตั้งค่าไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
		return
	}
	jsonResponse(w, map[string]any{"ok": true, "message": "รีเซ็ตโปรแกรมเรียบร้อยแล้ว"})
}

func (s *Server) handleHealth(w http.ResponseWriter, r *http.Request) {
	ready, checks := engineHealth()
	jsonResponse(w, map[string]any{
		"ok":       true,
		"ready":    ready,
		"version":  "2.6.5",
		"engine":   "BLACKWOLF Local Excel Engine",
		"port":     s.port,
		"platform": runtime.GOOS,
		"checks":   checks,
	})
}

func engineHealth() (bool, map[string]any) {
	checks := map[string]any{
		"windows":        runtime.GOOS == "windows",
		"powershell":     false,
		"excel":          false,
		"powershellPath": "",
		"excelDetection": "",
	}
	if runtime.GOOS != "windows" {
		return false, checks
	}

	powerShellPath, powerShellOK := findWindowsPowerShell()
	if powerShellOK {
		checks["powershell"] = true
		checks["powershellPath"] = powerShellPath
	}

	// ตรวจ Registry ทั้ง Office 64-bit และ Office 32-bit รวมถึง App Paths
	registryChecks := [][]string{
		{"query", `HKCR\Excel.Application\CLSID`},
		{"query", `HKCR\Excel.Application\CLSID`, "/reg:64"},
		{"query", `HKCR\Excel.Application\CLSID`, "/reg:32"},
		{"query", `HKLM\SOFTWARE\Classes\Excel.Application\CLSID`, "/reg:64"},
		{"query", `HKLM\SOFTWARE\Classes\Excel.Application\CLSID`, "/reg:32"},
		{"query", `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`, "/reg:64"},
		{"query", `HKLM\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`, "/reg:32"},
		{"query", `HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\App Paths\excel.exe`},
	}
	for _, args := range registryChecks {
		if runHiddenCommand(4*time.Second, "reg.exe", args...) {
			checks["excel"] = true
			checks["excelDetection"] = "registry"
			break
		}
	}

	// บางเครื่องติดตั้ง Office แบบ Click-to-Run แต่ Registry อยู่คนละมุมมอง
	// จึงตรวจตำแหน่ง EXCEL.EXE ที่พบบ่อยเพิ่มเติม
	if !checks["excel"].(bool) {
		for _, candidate := range commonExcelPaths() {
			if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
				checks["excel"] = true
				checks["excelDetection"] = candidate
				break
			}
		}
	}

	// วิธีสุดท้าย: ทดสอบสร้าง Excel COM จริงแบบชั่วคราว
	// ใช้เฉพาะเมื่อ Registry และตำแหน่งไฟล์ตรวจไม่พบ
	if !checks["excel"].(bool) && powerShellOK {
		probe := `$excel = $null
try {
    $type = [type]::GetTypeFromProgID('Excel.Application')
    if ($null -eq $type) { exit 1 }
    $excel = [Activator]::CreateInstance($type)
    if ($null -eq $excel) { exit 1 }
    $excel.DisplayAlerts = $false
    $excel.Quit()
    [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel)
    $excel = $null
    [GC]::Collect()
    [GC]::WaitForPendingFinalizers()
    exit 0
}
catch {
    if ($null -ne $excel) {
        try { $excel.Quit() } catch {}
        try { [void][Runtime.InteropServices.Marshal]::FinalReleaseComObject($excel) } catch {}
    }
    exit 1
}`
		if runHiddenCommand(20*time.Second, powerShellPath,
			"-NoLogo", "-NoProfile", "-NonInteractive", "-Sta",
			"-ExecutionPolicy", "Bypass", "-Command", probe) {
			checks["excel"] = true
			checks["excelDetection"] = "com-probe"
		}
	}

	ready := checks["windows"].(bool) && checks["powershell"].(bool) && checks["excel"].(bool)
	return ready, checks
}

func findWindowsPowerShell() (string, bool) {
	if path, err := exec.LookPath("powershell.exe"); err == nil {
		return path, true
	}

	windowsDir := strings.TrimSpace(os.Getenv("SystemRoot"))
	if windowsDir == "" {
		windowsDir = strings.TrimSpace(os.Getenv("WINDIR"))
	}
	if windowsDir == "" {
		windowsDir = `C:\Windows`
	}

	candidates := []string{
		filepath.Join(windowsDir, "System32", "WindowsPowerShell", "v1.0", "powershell.exe"),
		filepath.Join(windowsDir, "Sysnative", "WindowsPowerShell", "v1.0", "powershell.exe"),
		filepath.Join(windowsDir, "SysWOW64", "WindowsPowerShell", "v1.0", "powershell.exe"),
	}
	for _, candidate := range candidates {
		if info, err := os.Stat(candidate); err == nil && !info.IsDir() {
			return candidate, true
		}
	}
	return "", false
}

func commonExcelPaths() []string {
	roots := []string{
		os.Getenv("ProgramFiles"),
		os.Getenv("ProgramFiles(x86)"),
	}
	officeDirs := []string{"Office16", "Office15", "Office14"}
	paths := make([]string, 0, len(roots)*len(officeDirs)*2)
	for _, root := range roots {
		root = strings.TrimSpace(root)
		if root == "" {
			continue
		}
		for _, officeDir := range officeDirs {
			paths = append(paths,
				filepath.Join(root, "Microsoft Office", "root", officeDir, "EXCEL.EXE"),
				filepath.Join(root, "Microsoft Office", officeDir, "EXCEL.EXE"),
			)
		}
	}
	return paths
}

func runHiddenCommand(timeout time.Duration, name string, args ...string) bool {
	ctx, cancel := context.WithTimeout(context.Background(), timeout)
	defer cancel()
	cmd := exec.CommandContext(ctx, name, args...)
	configureHidden(cmd)
	err := cmd.Run()
	return err == nil && ctx.Err() == nil
}

func (s *Server) handleRun(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", 405)
		return
	}
	r.Body = http.MaxBytesReader(w, r.Body, 2<<30)
	if err := r.ParseMultipartForm(128 << 20); err != nil {
		jsonError(w, "อ่านไฟล์ไม่สำเร็จ: "+err.Error(), 400)
		return
	}

	id := time.Now().Format("20060102-150405") + "-" + strconv.FormatInt(time.Now().UnixNano()%100000, 10)
	dir := filepath.Join(s.baseDir, "Runs", id)
	inputDir := filepath.Join(dir, "input")
	outputDir := filepath.Join(dir, "output")
	if err := os.MkdirAll(inputDir, 0o755); err != nil {
		jsonError(w, err.Error(), 500)
		return
	}
	if err := os.MkdirAll(outputDir, 0o755); err != nil {
		jsonError(w, err.Error(), 500)
		return
	}

	paths := map[string]string{}
	inputNames := map[string]string{}
	for _, field := range []string{"master", "issue", "daily", "m190"} {
		p, name, err := saveUpload(r.MultipartForm, field, inputDir)
		if err != nil {
			os.RemoveAll(dir)
			jsonError(w, err.Error(), 400)
			return
		}
		paths[field], inputNames[field] = p, name
	}
	for _, field := range []string{"sm", "blacklist"} {
		p, name, err := saveUploadOptional(r.MultipartForm, field, inputDir)
		if err != nil {
			os.RemoveAll(dir)
			jsonError(w, err.Error(), 400)
			return
		}
		paths[field], inputNames[field] = p, name
	}
	etlText := strings.TrimSpace(r.FormValue("etlText"))
	if err := validateEtlText(etlText); err != nil {
		os.RemoveAll(dir)
		jsonError(w, err.Error(), 400)
		return
	}
	etlPath := filepath.Join(inputDir, "ETL.txt")
	if err := os.WriteFile(etlPath, append([]byte{0xEF, 0xBB, 0xBF}, []byte(etlText)...), 0o644); err != nil {
		os.RemoveAll(dir)
		jsonError(w, err.Error(), 500)
		return
	}
	paths["etl"], inputNames["etl"] = etlPath, "ETL.txt"

	run := &Run{ID: id, Status: "queued", Message: "รับไฟล์ครบแล้ว", Progress: 5, StartedAt: time.Now(), Logs: []string{"รับไฟล์ครบแล้ว"}, Summary: map[string]string{}, Files: map[string]string{}, Dir: dir, InputNames: inputNames}
	s.mu.Lock()
	s.runs[id] = run
	s.persistRunLocked(run)
	s.mu.Unlock()
	go s.processRun(run, paths, inputDir, outputDir)
	jsonResponse(w, map[string]any{"ok": true, "runId": id})
}

func saveUpload(form *multipart.Form, field, dir string) (string, string, error) {
	list := form.File[field]
	if len(list) == 0 {
		return "", "", fmt.Errorf("กรุณาเลือกไฟล์ %s", field)
	}
	return saveFileHeader(list[0], field, dir)
}
func saveUploadOptional(form *multipart.Form, field, dir string) (string, string, error) {
	list := form.File[field]
	if len(list) == 0 || list[0].Filename == "" {
		return "", "", nil
	}
	return saveFileHeader(list[0], field, dir)
}
func saveFileHeader(h *multipart.FileHeader, field, dir string) (string, string, error) {
	ext := strings.ToLower(filepath.Ext(h.Filename))
	if field == "blacklist" {
		if ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" {
			return "", "", fmt.Errorf("ไฟล์ Blacklist ต้องเป็น .xls, .xlsx หรือ .xlsm")
		}
	} else if ext != ".xlsx" {
		return "", "", fmt.Errorf("ไฟล์ %s ต้องเป็น .xlsx", field)
	}
	src, err := h.Open()
	if err != nil {
		return "", "", err
	}
	defer src.Close()
	safe := canonicalUploadName(field, h.Filename)
	dstPath := filepath.Join(dir, safe)
	dst, err := os.Create(dstPath)
	if err != nil {
		return "", "", err
	}
	_, err = io.Copy(dst, src)
	cerr := dst.Close()
	if err == nil {
		err = cerr
	}
	if err != nil {
		return "", "", err
	}
	return dstPath, h.Filename, nil
}
func sanitizeFilename(name string) string {
	name = filepath.Base(name)
	return strings.NewReplacer("..", "_", "/", "_", "\\", "_", ":", "_", "*", "_", "?", "_", "\"", "_", "<", "_", ">", "_", "|", "_").Replace(name)
}

func canonicalUploadName(field, original string) string {
	ext := strings.ToLower(filepath.Ext(original))
	switch field {
	case "master":
		return "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก_WEB.xlsx"
	case "issue":
		return "เช็คสถานะ ISSUE_WEB.xlsx"
	case "daily":
		return "รายงานงานประกันแรงงานต่างด้าว_WEB.xlsx"
	case "m190":
		return "M190027_PRD008_Premium_by_Policy_WEB.xlsx"
	case "sm":
		return "ข้อมูลไม่สมบูรณ์_WEB.xlsx"
	case "blacklist":
		if ext != ".xls" && ext != ".xlsx" && ext != ".xlsm" {
			ext = ".xlsx"
		}
		return "Blacklist_WEB" + ext
	default:
		return field + "_" + sanitizeFilename(original)
	}
}

func validateEtlText(text string) error {
	if strings.TrimSpace(text) == "" {
		return nil
	}
	seen := map[string]int{}
	bad := []string{}
	for i, raw := range strings.Split(strings.ReplaceAll(text, "\r\n", "\n"), "\n") {
		line := strings.TrimSpace(raw)
		if line == "" {
			continue
		}
		firstDot := strings.Index(line, ".")
		if firstDot <= 0 || firstDot == len(line)-1 {
			bad = append(bad, fmt.Sprintf("บรรทัด %d", i+1))
			continue
		}
		seq := strings.TrimSpace(line[:firstDot])
		rest := strings.TrimSpace(line[firstDot+1:])
		parts := strings.Split(rest, ":")
		if _, err := strconv.Atoi(seq); err != nil || len(parts) < 3 || strings.TrimSpace(parts[0]) == "" || strings.TrimSpace(parts[1]) == "" || strings.TrimSpace(parts[2]) == "" {
			bad = append(bad, fmt.Sprintf("บรรทัด %d", i+1))
			continue
		}
		propID := strings.TrimSpace(parts[0])
		for _, ch := range propID {
			if ch < '0' || ch > '9' {
				bad = append(bad, fmt.Sprintf("บรรทัด %d", i+1))
				propID = ""
				break
			}
		}
		if propID == "" {
			continue
		}
		if prior, ok := seen[propID]; ok {
			return fmt.Errorf("ETL มี PropID ซ้ำ %s ที่บรรทัด %d และ %d", propID, prior, i+1)
		}
		seen[propID] = i + 1
	}
	if len(bad) > 0 {
		if len(bad) > 8 {
			bad = append(bad[:8], "...")
		}
		return fmt.Errorf("รูปแบบ ETL ไม่ถูกต้อง: %s (รูปแบบต้องเป็น No.PropID:Policy:Group)", strings.Join(bad, ", "))
	}
	return nil
}

func (s *Server) copyResultsToSelectedStorage(runID string, files map[string]string) (string, error) {
	cfg := s.configSnapshot()
	if strings.TrimSpace(cfg.StorageDir) == "" {
		return "", nil
	}
	if err := writableDirectory(cfg.StorageDir); err != nil {
		return "", err
	}
	targetDir := filepath.Join(cfg.StorageDir, "BLACKWOLF_"+sanitizeFilename(runID))
	if err := os.MkdirAll(targetDir, 0o755); err != nil {
		return "", err
	}
	for _, source := range files {
		if source == "" {
			continue
		}
		if info, err := os.Stat(source); err != nil || info.IsDir() {
			continue
		}
		target := filepath.Join(targetDir, filepath.Base(source))
		if err := copyFile(source, target); err != nil {
			return targetDir, err
		}
	}
	return targetDir, nil
}

func (s *Server) processRun(run *Run, paths map[string]string, inputDir, outputDir string) {
	s.updateRun(run.ID, func(r *Run) {
		r.Status = "running"
		r.Message = "กำลังสำรองไฟล์ต้นฉบับ"
		r.Progress = 10
	})
	backupMaster := filepath.Join(outputDir, "Backup_Master_Original"+strings.ToLower(filepath.Ext(paths["master"])))
	backupIssue := filepath.Join(outputDir, "Backup_ISSUE_Original"+strings.ToLower(filepath.Ext(paths["issue"])))
	if err := copyFile(paths["master"], backupMaster); err != nil {
		s.failRun(run.ID, err)
		return
	}
	if err := copyFile(paths["issue"], backupIssue); err != nil {
		s.failRun(run.ID, err)
		return
	}

	largeDataMode := false
	workflowTimeout := 30 * time.Minute
	extractTimeout := 30 * time.Minute
	if info, statErr := os.Stat(paths["daily"]); statErr == nil && info.Size() >= 50<<20 {
		largeDataMode = true
		workflowTimeout = 120 * time.Minute
		extractTimeout = 60 * time.Minute
	}
	s.updateRun(run.ID, func(r *Run) {
		if largeDataMode {
			r.Message = "กำลังตรวจและประมวลผล Workflow แบบ Large Data"
			r.Logs = append(r.Logs, "เปิด Large Data Mode: รองรับ Daily Report ระดับ 500,000+ แถว และขยายเวลาประมวลผลสูงสุด 120 นาที")
		} else {
			r.Message = "กำลังตรวจและประมวลผล Workflow ด้วย Excel"
		}
		r.Progress = 18
	})
	workflowArgs := []string{"-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", s.workflowPath, "-InputDir", inputDir, "-OutputDir", outputDir}
	lines, err := s.runPowerShell(run.ID, workflowArgs, 18, 72, workflowTimeout)
	if err != nil {
		msg := lastUsefulLine(lines)
		if errors.Is(err, errRunCancelled) {
			s.cancelRun(run.ID)
			return
		}
		if strings.Contains(err.Error(), "ใช้เวลาเกินเวลาที่กำหนด") {
			msg = err.Error()
		} else if msg == "" {
			msg = err.Error()
		}
		_ = os.WriteFile(filepath.Join(outputDir, "error.txt"), []byte(msg), 0o644)
		s.failRun(run.ID, errors.New(msg))
		return
	}
	workflowSummary := parseWorkflowSummary(lines)

	masterOut, err := findWorkflowMaster(outputDir)
	if err != nil {
		s.failRun(run.ID, err)
		return
	}
	candidate := filepath.Join(outputDir, "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก V.2.5.3.xlsx")
	if !strings.EqualFold(masterOut, candidate) {
		if err := os.Rename(masterOut, candidate); err != nil {
			if err = copyFile(masterOut, candidate); err != nil {
				s.failRun(run.ID, err)
				return
			}
		}
	}
	updatedIssue := filepath.Join(outputDir, "เช็คสถานะ ISSUE.xlsx")
	if err := copyFile(paths["issue"], updatedIssue); err != nil {
		s.failRun(run.ID, err)
		return
	}
	s.updateRun(run.ID, func(r *Run) {
		r.Message = "กำลังสร้าง Dashboard, CSV และ Audit"
		r.Progress = 78
	})
	extractArgs := []string{"-NoLogo", "-NoProfile", "-ExecutionPolicy", "Bypass", "-File", s.extractPath, "-WorkbookPath", candidate, "-OutputDir", outputDir}
	extractLines, err := s.runPowerShell(run.ID, extractArgs, 78, 92, extractTimeout)
	if err != nil {
		if errors.Is(err, errRunCancelled) {
			s.cancelRun(run.ID)
			return
		}
		msg := lastUsefulLine(extractLines)
		if msg == "" {
			msg = err.Error()
		}
		s.failRun(run.ID, errors.New(msg))
		return
	}

	summaryPath := filepath.Join(outputDir, "summary.txt")
	summary, err := readSummary(summaryPath)
	if err != nil {
		s.failRun(run.ID, err)
		return
	}
	for k, v := range workflowSummary {
		summary[k] = v
	}
	if err := writeSummary(summaryPath, summary); err != nil {
		s.failRun(run.ID, err)
		return
	}
	_ = appendWorkflowAudit(filepath.Join(outputDir, "Audit_Report.txt"), workflowSummary)
	rows, err := readTSV(filepath.Join(outputDir, "pending_detail.tsv"))
	if err != nil {
		s.failRun(run.ID, err)
		return
	}

	fileMap := map[string]string{
		"master": candidate, "issue": updatedIssue, "csv": filepath.Join(outputDir, "Pending_Detail.csv"), "audit": filepath.Join(outputDir, "Audit_Report.txt"),
		"summary": filepath.Join(outputDir, "summary.txt"), "backupMaster": backupMaster, "backupIssue": backupIssue,
	}
	storageCopyPath, storageCopyErr := s.copyResultsToSelectedStorage(run.ID, fileMap)
	if storageCopyPath != "" {
		summary["SelectedStoragePath"] = storageCopyPath
	}
	now := time.Now()
	s.updateRun(run.ID, func(r *Run) {
		r.Status = "completed"
		r.Message = "ประมวลผลสำเร็จ พร้อมตรวจสอบและดาวน์โหลด"
		r.Progress = 100
		r.Summary = summary
		r.Rows = rows
		r.Files = fileMap
		r.FinishedAt = &now
		r.Logs = append(r.Logs, "สร้าง Master ฉบับอัปเดต, เช็คสถานะ ISSUE, CSV และ Audit สำเร็จ")
		if storageCopyPath != "" && storageCopyErr == nil {
			r.Logs = append(r.Logs, "คัดลอกผลลัพธ์ไปยังโฟลเดอร์ที่เลือก: "+storageCopyPath)
		}
		if storageCopyErr != nil {
			r.Logs = append(r.Logs, "WARNING: คัดลอกผลลัพธ์ไปโฟลเดอร์ที่เลือกไม่สำเร็จ: "+storageCopyErr.Error())
		}
	})
}

var errRunCancelled = errors.New("ผู้ใช้ยกเลิกการรัน")

func (s *Server) runPowerShell(id string, args []string, from, to int, timeoutLimit time.Duration) ([]string, error) {
	cmd := exec.Command("powershell.exe", args...)
	configureHidden(cmd)
	stdout, err := cmd.StdoutPipe()
	if err != nil {
		return nil, err
	}
	stderr, err := cmd.StderrPipe()
	if err != nil {
		return nil, err
	}
	s.cmdMu.Lock()
	if s.deletedRuns[id] {
		s.cmdMu.Unlock()
		return nil, errRunCancelled
	}
	if err = cmd.Start(); err != nil {
		s.cmdMu.Unlock()
		return nil, fmt.Errorf("เปิด Local Engine ไม่สำเร็จ: %w", err)
	}
	s.activeCmds[id] = cmd
	delete(s.cancelledRuns, id)
	s.cmdMu.Unlock()
	defer func() {
		s.cmdMu.Lock()
		delete(s.activeCmds, id)
		s.cmdMu.Unlock()
	}()

	started := time.Now()
	var mu sync.Mutex
	lines := []string{}
	stageMessage := "กำลังประมวลผลด้วย Excel Engine"
	lastHeartbeatMinute := -1
	var wg sync.WaitGroup
	scan := func(rd io.Reader, isErr bool) {
		defer wg.Done()
		sc := bufio.NewScanner(rd)
		sc.Buffer(make([]byte, 64*1024), 4*1024*1024)
		for sc.Scan() {
			line := strings.TrimSpace(sc.Text())
			if line == "" {
				continue
			}
			mu.Lock()
			lines = append(lines, line)
			mu.Unlock()
			prefix := ""
			if isErr {
				prefix = "ERROR: "
			}
			progress, message, hasProgress := parseProgressLine(line)
			if hasProgress && message != "" {
				mu.Lock()
				stageMessage = message
				mu.Unlock()
			}
			s.updateRun(id, func(r *Run) {
				if hasProgress {
					if progress < from {
						progress = from
					}
					if progress > to {
						progress = to
					}
					r.Progress = progress
					if message != "" {
						r.Message = message
					}
				} else {
					r.Logs = append(r.Logs, prefix+line)
					if len(r.Logs) > 500 {
						r.Logs = r.Logs[len(r.Logs)-500:]
					}
				}
			})
		}
	}
	wg.Add(2)
	go scan(stdout, false)
	go scan(stderr, true)

	waitCh := make(chan error, 1)
	go func() { waitCh <- cmd.Wait() }()
	ticker := time.NewTicker(5 * time.Second)
	defer ticker.Stop()
	timeout := time.NewTimer(timeoutLimit)
	defer timeout.Stop()

	for {
		select {
		case waitErr := <-waitCh:
			wg.Wait()
			s.cmdMu.Lock()
			cancelled := s.cancelledRuns[id]
			delete(s.cancelledRuns, id)
			s.cmdMu.Unlock()
			if cancelled {
				return lines, errRunCancelled
			}
			return lines, waitErr
		case <-ticker.C:
			elapsed := time.Since(started).Round(time.Second)
			mu.Lock()
			base := stageMessage
			mu.Unlock()
			minute := int(elapsed / time.Minute)
			s.updateRun(id, func(r *Run) {
				if r.Status != "running" {
					return
				}
				r.Message = fmt.Sprintf("%s · ผ่านไป %s", base, formatElapsed(elapsed))
				if minute > 0 && minute != lastHeartbeatMinute {
					r.Logs = append(r.Logs, fmt.Sprintf("Excel Engine ยังทำงานอยู่ — ผ่านไป %s", formatElapsed(elapsed)))
					if len(r.Logs) > 500 {
						r.Logs = r.Logs[len(r.Logs)-500:]
					}
				}
			})
			if minute > 0 {
				lastHeartbeatMinute = minute
			}
		case <-timeout.C:
			killProcessTree(cmd)
			waitErr := <-waitCh
			s.cmdMu.Lock()
			delete(s.cancelledRuns, id)
			s.cmdMu.Unlock()
			wg.Wait()
			_ = waitErr
			return lines, fmt.Errorf("Excel Engine ใช้เวลาเกินเวลาที่กำหนด (%s) ระบบหยุดการรันเพื่อป้องกันการค้าง กรุณาปิด Excel ที่เปิดอยู่แล้วลองใหม่", timeoutLimit.Round(time.Minute))
		}
	}
}

func formatElapsed(d time.Duration) string {
	total := int(d.Seconds())
	if total < 0 {
		total = 0
	}
	return fmt.Sprintf("%02d:%02d", total/60, total%60)
}

func parseProgressLine(line string) (int, string, bool) {
	parts := strings.SplitN(line, "|", 3)
	if len(parts) != 3 || strings.ToUpper(strings.TrimSpace(parts[0])) != "PROGRESS" {
		return 0, "", false
	}
	p, err := strconv.Atoi(strings.TrimSpace(parts[1]))
	if err != nil {
		return 0, "", false
	}
	return p, strings.TrimSpace(parts[2]), true
}

func parseWorkflowSummary(lines []string) map[string]string {
	joined := strings.Join(lines, "\n")
	start := strings.Index(joined, "{")
	end := strings.LastIndex(joined, "}")
	out := map[string]string{}
	if start < 0 || end <= start {
		return out
	}
	var raw map[string]any
	if err := json.Unmarshal([]byte(joined[start:end+1]), &raw); err != nil {
		return out
	}
	for k, v := range raw {
		switch x := v.(type) {
		case string:
			out[k] = x
		case float64:
			if x == float64(int64(x)) {
				out[k] = strconv.FormatInt(int64(x), 10)
			} else {
				out[k] = strconv.FormatFloat(x, 'f', -1, 64)
			}
		case bool:
			out[k] = strconv.FormatBool(x)
		default:
			b, _ := json.Marshal(x)
			out[k] = string(b)
		}
	}
	return out
}

func writeSummary(path string, summary map[string]string) error {
	keys := make([]string, 0, len(summary))
	for k := range summary {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	var b strings.Builder
	for _, k := range keys {
		fmt.Fprintf(&b, "%s=%s\r\n", k, summary[k])
	}
	return os.WriteFile(path, append([]byte{0xEF, 0xBB, 0xBF}, []byte(b.String())...), 0o644)
}

func appendWorkflowAudit(path string, summary map[string]string) error {
	if len(summary) == 0 {
		return nil
	}
	f, err := os.OpenFile(path, os.O_APPEND|os.O_WRONLY|os.O_CREATE, 0o644)
	if err != nil {
		return err
	}
	defer f.Close()
	_, err = fmt.Fprintln(f, "\r\nWORKFLOW RECONCILIATION")
	if err != nil {
		return err
	}
	for _, k := range []string{"DateStart", "DateEnd", "ReportLatestDate", "ReportRowsAfterDateStatusFilter", "M190PropIdRows", "EtlPropIdRows", "IssuedRowsRemoved", "PendingRowsWrittenToData", "IssueDataRowsWritten", "IssueCheckRowsWritten", "IssueM190RowsWritten", "IssueEtlRowsWritten", "SmPropIdsWritten", "BlPropIdsWritten", "MenuEStatusesWritten", "PvDataRowsCopiedToPvFinal"} {
		if v, ok := summary[k]; ok {
			if _, err = fmt.Fprintf(f, "%s: %s\r\n", k, v); err != nil {
				return err
			}
		}
	}
	return nil
}

func lastUsefulLine(lines []string) string {
	for i := len(lines) - 1; i >= 0; i-- {
		v := strings.TrimSpace(lines[i])
		if v != "" && !strings.HasPrefix(v, "{") && !strings.HasPrefix(v, "}") {
			return v
		}
	}
	return ""
}
func findWorkflowMaster(dir string) (string, error) {
	entries, err := os.ReadDir(dir)
	if err != nil {
		return "", err
	}
	var newest string
	var mt time.Time
	for _, e := range entries {
		if e.IsDir() {
			continue
		}
		n := strings.ToLower(e.Name())
		if !strings.HasSuffix(n, ".xlsx") || strings.Contains(n, "backup_") || strings.Contains(n, "candidate_master") {
			continue
		}
		info, _ := e.Info()
		if newest == "" || info.ModTime().After(mt) {
			newest = filepath.Join(dir, e.Name())
			mt = info.ModTime()
		}
	}
	if newest == "" {
		return "", errors.New("ไม่พบไฟล์ Master ผลลัพธ์จาก Workflow")
	}
	return newest, nil
}
func copyFile(src, dst string) error {
	in, err := os.Open(src)
	if err != nil {
		return err
	}
	defer in.Close()
	out, err := os.Create(dst)
	if err != nil {
		return err
	}
	_, err = io.Copy(out, in)
	cerr := out.Close()
	if err == nil {
		err = cerr
	}
	return err
}

func (s *Server) updateRun(id string, fn func(*Run)) {
	s.mu.Lock()
	if r := s.runs[id]; r != nil {
		fn(r)
		s.persistRunLocked(r)
	}
	s.mu.Unlock()
}
func (s *Server) failRun(id string, err error) {
	now := time.Now()
	s.updateRun(id, func(r *Run) {
		r.Status = "failed"
		r.Message = "ประมวลผลไม่สำเร็จ"
		r.Error = err.Error()
		r.Progress = 100
		r.FinishedAt = &now
		r.Logs = append(r.Logs, "ERROR: "+err.Error())
	})
}

func (s *Server) cancelRun(id string) {
	now := time.Now()
	s.updateRun(id, func(r *Run) {
		r.Status = "cancelled"
		r.Message = "หยุดการรันแล้ว"
		r.Error = ""
		r.Progress = 100
		r.FinishedAt = &now
		r.Logs = append(r.Logs, "หยุดการรันเรียบร้อยแล้ว")
	})
}
func (s *Server) handleStatus(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	s.mu.RLock()
	run := s.runs[id]
	if run == nil {
		s.mu.RUnlock()
		jsonError(w, "ไม่พบ Run", 404)
		return
	}
	cp := *run
	cp.Logs = append([]string(nil), run.Logs...)
	cp.Rows = nil
	s.mu.RUnlock()
	jsonResponse(w, cp)
}
func (s *Server) handleResult(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	page, _ := strconv.Atoi(r.URL.Query().Get("page"))
	if page < 1 {
		page = 1
	}
	size, _ := strconv.Atoi(r.URL.Query().Get("pageSize"))
	if size < 1 || size > 500 {
		size = 50
	}
	search := strings.ToLower(strings.TrimSpace(r.URL.Query().Get("search")))
	status := r.URL.Query().Get("status")
	s.mu.RLock()
	run := s.runs[id]
	if run == nil {
		s.mu.RUnlock()
		jsonError(w, "ไม่พบ Run", 404)
		return
	}
	rows := append([]DetailRow(nil), run.Rows...)
	summary := cloneMap(run.Summary)
	s.mu.RUnlock()
	filtered := make([]DetailRow, 0, len(rows))
	for _, row := range rows {
		if status != "" && status != "all" && row.PendingStatus != status {
			continue
		}
		if search != "" {
			hay := strings.ToLower(strings.Join([]string{row.ProposalID, row.Policy, row.CertificateNo, row.AgencyName, row.EmployeeName, row.AlienCode, row.AlienName, row.PendingStatus}, " "))
			if !strings.Contains(hay, search) {
				continue
			}
		}
		filtered = append(filtered, row)
	}
	total := len(filtered)
	start := (page - 1) * size
	if start > total {
		start = total
	}
	end := start + size
	if end > total {
		end = total
	}
	jsonResponse(w, map[string]any{"ok": true, "summary": summary, "rows": filtered[start:end], "total": total, "page": page, "pageSize": size})
}
func (s *Server) handleHistory(w http.ResponseWriter, r *http.Request) {
	deleted := s.cleanupExpiredRuns()
	now := time.Now()
	s.mu.RLock()
	list := make([]RunHistoryItem, 0, len(s.runs))
	for _, v := range s.runs {
		cp := *v
		cp.Rows = nil
		cp.Logs = nil
		expiresAt := runExpiresAt(v)
		remaining := int64(expiresAt.Sub(now).Seconds())
		if remaining < 0 {
			remaining = 0
		}
		list = append(list, RunHistoryItem{
			Run:              cp,
			DisplayName:      runDisplayName(v),
			ExpiresAt:        expiresAt,
			RemainingSeconds: remaining,
			RetentionDays:    runRetentionDays,
		})
	}
	s.mu.RUnlock()
	sort.Slice(list, func(i, j int) bool { return list[i].StartedAt.After(list[j].StartedAt) })
	jsonResponse(w, map[string]any{"ok": true, "runs": list, "retentionDays": runRetentionDays, "deletedExpiredRuns": deleted})
}
func (s *Server) handleDownload(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	key := r.URL.Query().Get("file")
	s.mu.RLock()
	run := s.runs[id]
	if run == nil {
		s.mu.RUnlock()
		jsonError(w, "ไม่พบ Run", 404)
		return
	}
	p := run.Files[key]
	s.mu.RUnlock()
	if p == "" {
		jsonError(w, "ไม่พบไฟล์ผลลัพธ์", 404)
		return
	}
	if _, err := os.Stat(p); err != nil {
		jsonError(w, "ไฟล์ถูกย้ายหรือลบแล้ว", 404)
		return
	}
	name := filepath.Base(p)
	w.Header().Set("Content-Disposition", fmt.Sprintf("attachment; filename=%q; filename*=UTF-8''%s", "BLACKWOLF_Result"+filepath.Ext(name), url.PathEscape(name)))
	http.ServeFile(w, r, p)
}
func (s *Server) handleOpenOutput(w http.ResponseWriter, r *http.Request) {
	id := r.URL.Query().Get("id")
	s.mu.RLock()
	run := s.runs[id]
	s.mu.RUnlock()
	if run == nil {
		jsonError(w, "ไม่พบ Run", 404)
		return
	}
	dir := filepath.Join(run.Dir, "output")
	if runtime.GOOS == "windows" {
		_ = exec.Command("explorer.exe", dir).Start()
	}
	jsonResponse(w, map[string]any{"ok": true, "path": dir})
}
func (s *Server) handleCancel(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	id := strings.TrimSpace(r.URL.Query().Get("id"))
	if id == "" {
		jsonError(w, "ไม่พบ Run ID", http.StatusBadRequest)
		return
	}
	s.cmdMu.Lock()
	cmd := s.activeCmds[id]
	if cmd != nil {
		s.cancelledRuns[id] = true
	}
	s.cmdMu.Unlock()
	if cmd == nil {
		jsonError(w, "Run นี้ไม่ได้กำลังทำงาน", http.StatusConflict)
		return
	}
	s.updateRun(id, func(run *Run) {
		run.Message = "กำลังหยุด Excel Engine"
		run.Logs = append(run.Logs, "ผู้ใช้สั่งหยุดการรัน")
	})
	killProcessTree(cmd)
	jsonResponse(w, map[string]any{"ok": true})
}

func (s *Server) handleDeleteRun(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	id := strings.TrimSpace(r.URL.Query().Get("id"))
	if id == "" {
		jsonError(w, "ไม่พบ Run ID", http.StatusBadRequest)
		return
	}
	s.cmdMu.Lock()
	_, active := s.activeCmds[id]
	s.cmdMu.Unlock()
	if active {
		jsonError(w, "Run นี้กำลังทำงาน กรุณากดหยุดการรันก่อน", http.StatusConflict)
		return
	}
	s.mu.Lock()
	run := s.runs[id]
	if run != nil {
		delete(s.runs, id)
	}
	s.mu.Unlock()
	if run == nil {
		jsonError(w, "ไม่พบ Run", http.StatusNotFound)
		return
	}
	if err := os.RemoveAll(run.Dir); err != nil {
		jsonError(w, "ลบข้อมูล Run ไม่สำเร็จ: "+err.Error(), http.StatusInternalServerError)
		return
	}
	jsonResponse(w, map[string]any{"ok": true})
}

func (s *Server) handleCancelDeleteRun(w http.ResponseWriter, r *http.Request) {
	if r.Method != http.MethodPost {
		http.Error(w, "method not allowed", http.StatusMethodNotAllowed)
		return
	}
	id := strings.TrimSpace(r.URL.Query().Get("id"))
	if id == "" {
		jsonError(w, "ไม่พบ Run ID", http.StatusBadRequest)
		return
	}

	// Remove only the selected Run from history. Original files outside the
	// BLACKWOLF run folder are never deleted.
	s.mu.Lock()
	run := s.runs[id]
	if run != nil {
		delete(s.runs, id)
	}
	s.mu.Unlock()
	if run == nil {
		jsonError(w, "ไม่พบ Run", http.StatusNotFound)
		return
	}

	// A tombstone prevents the background goroutine from starting another
	// PowerShell/Excel stage after the user has deleted this Run.
	s.cmdMu.Lock()
	s.deletedRuns[id] = true
	cmd := s.activeCmds[id]
	wasActive := cmd != nil
	if wasActive {
		s.cancelledRuns[id] = true
	}
	s.cmdMu.Unlock()

	_ = os.WriteFile(filepath.Join(run.Dir, ".delete-pending"), []byte(time.Now().Format(time.RFC3339)), 0o644)
	if cmd != nil {
		if excelPID := readProcessID(filepath.Join(run.Dir, "output", "excel.pid")); excelPID > 0 {
			killExcelProcess(excelPID)
		}
		killProcessTree(cmd)
	}
	go s.removeRunDirWithRetry(id, run.Dir)
	jsonResponse(w, map[string]any{"ok": true, "cancelled": wasActive, "deletedRunId": id})
}

func readProcessID(path string) int {
	b, err := os.ReadFile(path)
	if err != nil {
		return 0
	}
	pid, err := strconv.Atoi(strings.TrimSpace(string(b)))
	if err != nil || pid <= 0 {
		return 0
	}
	return pid
}

func (s *Server) removeRunDirWithRetry(id, dir string) {
	// Give the target process tree a moment to release Excel/workbook handles.
	for i := 0; i < 40; i++ {
		s.cmdMu.Lock()
		_, active := s.activeCmds[id]
		s.cmdMu.Unlock()
		if !active {
			break
		}
		time.Sleep(250 * time.Millisecond)
	}
	for i := 0; i < 12; i++ {
		if err := os.RemoveAll(dir); err == nil {
			return
		}
		time.Sleep(time.Duration(i+1) * 300 * time.Millisecond)
	}
}

func (s *Server) handleShutdown(w http.ResponseWriter, r *http.Request) {
	s.cmdMu.Lock()
	cmds := make([]*exec.Cmd, 0, len(s.activeCmds))
	for id, cmd := range s.activeCmds {
		s.cancelledRuns[id] = true
		cmds = append(cmds, cmd)
	}
	s.cmdMu.Unlock()
	for _, cmd := range cmds {
		killProcessTree(cmd)
	}
	jsonResponse(w, map[string]any{"ok": true})
	go func() { time.Sleep(500 * time.Millisecond); os.Exit(0) }()
}

func readSummary(path string) (map[string]string, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	out := map[string]string{}
	sc := bufio.NewScanner(f)
	for sc.Scan() {
		line := strings.TrimSpace(strings.TrimPrefix(sc.Text(), "\ufeff"))
		if line == "" || strings.HasPrefix(line, "#") {
			continue
		}
		p := strings.SplitN(line, "=", 2)
		if len(p) == 2 {
			out[strings.TrimSpace(p[0])] = strings.TrimSpace(p[1])
		}
	}
	return out, sc.Err()
}
func readTSV(path string) ([]DetailRow, error) {
	f, err := os.Open(path)
	if err != nil {
		return nil, err
	}
	defer f.Close()
	rd := csv.NewReader(bufio.NewReader(f))
	rd.Comma = '\t'
	rd.FieldsPerRecord = -1
	rd.LazyQuotes = true
	header, err := rd.Read()
	if err != nil {
		return nil, err
	}
	if len(header) > 0 {
		header[0] = strings.TrimPrefix(header[0], "\ufeff")
	}
	idx := map[string]int{}
	for i, h := range header {
		idx[h] = i
	}
	val := func(rec []string, k string) string {
		if i, ok := idx[k]; ok && i < len(rec) {
			return rec[i]
		}
		return ""
	}
	rows := []DetailRow{}
	for {
		rec, e := rd.Read()
		if errors.Is(e, io.EOF) {
			break
		}
		if e != nil {
			return nil, e
		}
		prem, _ := strconv.ParseFloat(val(rec, "TotalPremium"), 64)
		age, _ := strconv.Atoi(val(rec, "AgingDays"))
		rows = append(rows, DetailRow{ProposalID: val(rec, "ProposalID"), Policy: val(rec, "Policy"), CertificateNo: val(rec, "CertificateNo"), AgencyCode: val(rec, "AgencyCode"), Mticode: val(rec, "Mticode"), AgencyName: val(rec, "AgencyName"), RequestCode: val(rec, "RequestCode"), EmployeeName: val(rec, "EmployeeName"), AlienCode: val(rec, "alienCode"), AlienName: val(rec, "alienNameEn"), TotalPremium: prem, CreateDate: val(rec, "CreateDate"), SourceStatus: val(rec, "SourceStatus"), EPropID: val(rec, "EPropID"), Discount: val(rec, "Discount"), PendingStatus: val(rec, "PendingStatus"), AgingDays: age, PendingRange: val(rec, "PendingRange"), Incomplete: val(rec, "IncompleteStatus"), Blacklist: val(rec, "BlacklistStatus"), MenuE: val(rec, "MenuEProblem")})
	}
	return rows, nil
}
func createZip(out string, files map[string]string) error {
	f, err := os.Create(out)
	if err != nil {
		return err
	}
	zw := zip.NewWriter(f)
	keys := make([]string, 0, len(files))
	for k := range files {
		keys = append(keys, k)
	}
	sort.Strings(keys)
	for _, k := range keys {
		p := files[k]
		if p == "" {
			continue
		}
		in, e := os.Open(p)
		if e != nil {
			zw.Close()
			f.Close()
			return e
		}
		w, e := zw.Create(filepath.Base(p))
		if e == nil {
			_, e = io.Copy(w, in)
		}
		in.Close()
		if e != nil {
			zw.Close()
			f.Close()
			return e
		}
	}
	if err = zw.Close(); err != nil {
		f.Close()
		return err
	}
	return f.Close()
}
func runDisplayName(run *Run) string {
	if run == nil || run.StartedAt.IsZero() {
		return "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก"
	}
	return "เช็คกรมธรรม์ต่างด้าวที่ยังไม่ออก " + run.StartedAt.Local().Format("2006-01-02")
}

func runExpiresAt(run *Run) time.Time {
	if run == nil || run.StartedAt.IsZero() {
		return time.Now().Add(runRetention)
	}
	return run.StartedAt.Add(runRetention)
}

func (s *Server) cleanupExpiredRuns() int {
	now := time.Now()

	// Snapshot active process IDs first so cleanup never removes a Run that is
	// currently controlled by this server process.
	s.cmdMu.Lock()
	active := make(map[string]bool, len(s.activeCmds))
	for id, cmd := range s.activeCmds {
		if cmd != nil {
			active[id] = true
		}
	}
	s.cmdMu.Unlock()

	expired := make([]*Run, 0)
	s.mu.Lock()
	for id, run := range s.runs {
		if run == nil || active[id] || run.StartedAt.IsZero() {
			continue
		}
		if !now.Before(runExpiresAt(run)) {
			delete(s.runs, id)
			expired = append(expired, run)
		}
	}
	s.mu.Unlock()

	for _, run := range expired {
		_ = os.RemoveAll(run.Dir)
	}
	return len(expired)
}

func (s *Server) runRetentionCleanupLoop() {
	ticker := time.NewTicker(30 * time.Minute)
	defer ticker.Stop()
	for range ticker.C {
		s.cleanupExpiredRuns()
	}
}

func cloneMap(in map[string]string) map[string]string {
	out := map[string]string{}
	for k, v := range in {
		out[k] = v
	}
	return out
}
func (s *Server) persistRunLocked(r *Run) {
	meta := *r
	meta.Rows = nil
	b, _ := json.MarshalIndent(meta, "", "  ")
	_ = os.WriteFile(filepath.Join(r.Dir, "run.json"), b, 0o644)
}
func (s *Server) loadHistory() {
	dirs, _ := os.ReadDir(filepath.Join(s.baseDir, "Runs"))
	for _, d := range dirs {
		if !d.IsDir() {
			continue
		}
		runDir := filepath.Join(s.baseDir, "Runs", d.Name())
		if _, err := os.Stat(filepath.Join(runDir, ".delete-pending")); err == nil {
			_ = os.RemoveAll(runDir)
			continue
		}
		p := filepath.Join(runDir, "run.json")
		b, e := os.ReadFile(p)
		if e != nil {
			continue
		}
		var r Run
		if json.Unmarshal(b, &r) != nil {
			continue
		}
		r.Dir = runDir
		if r.Status == "completed" {
			r.Rows, _ = readTSV(filepath.Join(r.Dir, "output", "pending_detail.tsv"))
		}
		s.runs[r.ID] = &r
	}
}
func jsonResponse(w http.ResponseWriter, v any) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	_ = json.NewEncoder(w).Encode(v)
}
func jsonError(w http.ResponseWriter, msg string, code int) {
	w.Header().Set("Content-Type", "application/json; charset=utf-8")
	w.WriteHeader(code)
	_ = json.NewEncoder(w).Encode(map[string]any{"ok": false, "error": msg})
}
