BLACKWOLF Web V2.6.5 Run Retention — Source Code

ไฟล์หลัก
- main.go: Backend/API, Settings, Storage picker, Manual, Status, Reset
- web/index.html: หน้าเว็บและหน้า Settings
- web/app.js: ภาษา ธีม สถานะ ปุ่มต่าง ๆ
- web/styles.css: Light/Dark Theme และ Settings UI
- workflow.ps1: Excel Workflow Engine
- extract_result.ps1: สร้างผลลัพธ์ Dashboard/CSV/Audit

ทดลองจาก Source
1. ติดตั้ง Go 1.23+
2. เปิด Command Prompt ในโฟลเดอร์ Source
3. รัน:
   go test ./...
   go run .

Build Windows
   go build -trimpath -ldflags="-s -w" -o BLACKWOLF_Web_Server.exe .

เพิ่มคู่มือ PDF ภายหลัง
- ตั้งชื่อ BLACKWOLF_User_Manual.pdf
- วางข้าง BLACKWOLF_Web_Server.exe
- ปุ่มดาวน์โหลดใน Settings จะเปิดใช้งานอัตโนมัติ


V2.6.5 Run History Retention
- Run History displays a friendly Thai report name instead of the internal Run ID.
- Internal Run files are retained for 14 days and deleted automatically after expiry.
- Downloaded files and files copied to external storage are not deleted.
