@echo off
powershell -NoProfile -Command ^
  "$conn = Get-NetTCPConnection -LocalPort 5000 -EA SilentlyContinue;" ^
  "if ($conn) {" ^
  "  Stop-Process -Id $conn.OwningProcess -Force -EA SilentlyContinue;" ^
  "  Write-Host 'SWT CAC Builder stopped.';" ^
  "} else {" ^
  "  Write-Host 'No server running on port 5000.';" ^
  "}"
