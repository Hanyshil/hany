# Diagnostics Script - Report Only
Write-Host "--- מתחיל ניתוח מערכת עמוק (מצב אבחון בלבד) ---" -ForegroundColor Cyan

# FIX 1: חותמת זמן
Write-Host "תאריך ושעה: $(Get-Date -Format 'dd/MM/yyyy HH:mm:ss')" -ForegroundColor White

# FIX 2: בדיקת הרשאות מנהל
$IsAdmin = ([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
if (-not $IsAdmin) {
    Write-Host "אזהרה: הסקריפט לא רץ כמנהל. חלק מהנתונים עלולים להיות חסרים." -ForegroundColor Red
}

# 1. בדיקת מעבד (CPU) - FIX 3: שימוש ב-Get-Counter לעומס CPU אמיתי ולא מצטבר
Write-Host "`n[1/5] מנתח עומס מעבד (דגימה של שנייה אחת)..." -ForegroundColor Yellow
try {
    $CPUCounters = Get-Counter '\Process(*)\% Processor Time' -SampleInterval 1 -MaxSamples 1 -ErrorAction Stop
    $TopCPU = $CPUCounters.CounterSamples |
        Where-Object { $_.InstanceName -notin @('_total', 'idle') } |
        Sort-Object CookedValue -Descending |
        Select-Object -First 5 @{Name="Name"; Expression={$_.InstanceName}}, @{Name="CPU%"; Expression={[Math]::Round($_.CookedValue, 1)}}
    $TopCPU | Format-Table -AutoSize
} catch {
    Write-Host "שגיאה בקריאת נתוני CPU: $_" -ForegroundColor Red
}

# 2. בדיקת זיכרון (RAM)
Write-Host "[2/5] בודק ניצול זיכרון..." -ForegroundColor Yellow
$OS = Get-CimInstance Win32_OperatingSystem

# FIX 4: המרה מפורשת KB -> GB (חלוקה ב-1024 פעמיים)
$TotalRAM = [Math]::Round($OS.TotalVisibleMemorySize / 1024 / 1024, 2)
$FreeRAM  = [Math]::Round($OS.FreePhysicalMemory      / 1024 / 1024, 2)

# FIX 5: הגנה מפני חלוקה באפס
if ($TotalRAM -gt 0) {
    $UsedPercent = [Math]::Round(($TotalRAM - $FreeRAM) / $TotalRAM * 100, 2)
} else {
    $UsedPercent = 0
    Write-Host "אזהרה: לא ניתן לחשב אחוז ניצול RAM." -ForegroundColor Red
}

Write-Host "סה''כ זיכרון: $TotalRAM GB"
Write-Host "זיכרון פנוי: $FreeRAM GB"
Write-Host "אחוז ניצול: $UsedPercent%"

$TopMem = Get-Process | Sort-Object WorkingSet -Descending |
    Select-Object -First 5 Name, @{Name="RAM(MB)"; Expression={[Math]::Round($_.WorkingSet / 1MB, 2)}}
$TopMem | Format-Table -AutoSize

# 3. בדיקת שטח דיסק (Disk Space) - FIX 6: סינון כוננים עם ערכי null
Write-Host "[3/5] סורק כוננים ושטח אחסון..." -ForegroundColor Yellow
Get-PSDrive -PSProvider FileSystem |
    Where-Object { $_.Free -ne $null -and $_.Used -ne $null } |
    Select-Object Name,
        @{Name="Free(GB)"; Expression={[Math]::Round($_.Free / 1GB, 2)}},
        @{Name="Used(GB)"; Expression={[Math]::Round($_.Used / 1GB, 2)}} |
    Format-Table -AutoSize

# 4. תוכנות שעולות עם הדלקת המחשב (Startup)
Write-Host "[4/5] בודק יישומים שמופעלים אוטומטית..." -ForegroundColor Yellow
Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location | Format-Table -AutoSize

# 5. בדיקת בריאות הדיסק (S.M.A.R.T)
Write-Host "[5/5] בודק תקינות חומרה (דיסק קשיח)..." -ForegroundColor Yellow
try {
    Get-PhysicalDisk | Select-Object FriendlyName, HealthStatus, OperationalStatus | Format-Table -AutoSize
} catch {
    Write-Host "שגיאה: נדרשות הרשאות מנהל לקריאת נתוני דיסק." -ForegroundColor Red
}

Write-Host "`n--- האבחון הושלם ---" -ForegroundColor Cyan
Write-Host "שים לב: אם אחוז הניצול ב-RAM מעל 80% או הדיסק מלא מעל 90%, זו כנראה סיבת האיטיות." -ForegroundColor White
