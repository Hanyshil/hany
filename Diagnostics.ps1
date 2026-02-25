# Diagnostics Script - Report Only
Write-Host "--- מתחיל ניתוח מערכת עמוק (מצב אבחון בלבד) ---" -ForegroundColor Cyan
# 1. בדיקת מעבד (CPU) - תהליכים שצורכים הכי הרבה משאבים
Write-Host "`n[1/5] מנתח עומס מעבד..." -ForegroundColor Yellow
$TopCPU = Get-Process | Sort-Object CPU -Descending | Select-Object -First 5 Name, @{Name="CPU(s)"; Expression={$_.CPU}}, Description
$TopCPU | Format-Table -AutoSize
# 2. בדיקת זיכרון (RAM)
Write-Host "[2/5] בודק ניצול זיכרון..." -ForegroundColor Yellow
$OS = Get-CimInstance Win32_OperatingSystem
$TotalRAM = [Math]::Round($OS.TotalVisibleMemorySize / 1MB, 2)
$FreeRAM = [Math]::Round($OS.FreePhysicalMemory / 1MB, 2)
$UsedPercent = [Math]::Round(($TotalRAM - $FreeRAM) / $TotalRAM * 100, 2)
Write-Host "סה''כ זיכרון: $TotalRAM GB"
Write-Host "זיכרון פנוי: $FreeRAM GB"
Write-Host "אחוז ניצול: $UsedPercent%"
$TopMem = Get-Process | Sort-Object WorkingSet -Descending | Select-Object -First 5 Name, @{Name="RAM(MB)"; Expression={[Math]::Round($_.WorkingSet / 1MB,2)}}
$TopMem | Format-Table -AutoSize
# 3. בדיקת שטח דיסק (Disk Space)
Write-Host "[3/5] סורק כוננים ושטח אחסון..." -ForegroundColor Yellow
Get-PSDrive -PSProvider FileSystem | Select-Object Name, @{Name="Free(GB)"; Expression={[Math]::Round($_.Free / 1GB,2)}}, @{Name="Used(GB)"; Expression={[Math]::Round($_.Used / 1GB,2)}} | Format-Table -AutoSize
# 4. תוכנות שעולות עם הדלקת המחשב (Startup)
Write-Host "[4/5] בודק יישומים שמופעלים אוטומטית..." -ForegroundColor Yellow
Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location | Format-Table -AutoSize
# 5. בדיקת בריאות הדיסק (S.M.A.R.T)
Write-Host "[5/5] בודק תקינות חומרה (דיסק קשיח)..." -ForegroundColor Yellow
Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus | Format-Table -AutoSize
Write-Host "`n--- האבחון הושלם ---" -ForegroundColor Cyan
Write-Host "שים לב: אם אחוז הניצול ב-RAM מעל 80% או הדיסק מלא מעל 90%, זו כנראה סיבת האיטיות." -ForegroundColor White
