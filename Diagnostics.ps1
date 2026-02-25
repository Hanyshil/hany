# Diagnostics Script
param(
    [switch]$KillHighMemory,
    [int]$ThresholdMB = 500
)

# רשימת תהליכי מערכת מוגנים שלא ייסגרו בשום מצב
$ProtectedProcesses = @(
    'system', 'idle', 'svchost', 'lsass', 'winlogon', 'csrss',
    'smss', 'wininit', 'services', 'explorer', 'dwm',
    'powershell', 'pwsh', 'taskhostw', 'spoolsv', 'audiodg'
)

if ($KillHighMemory) {
    Write-Host "--- מצב סגירת תהליכים כבד-זיכרון (סף: ${ThresholdMB}MB) ---" -ForegroundColor Cyan
} else {
    Write-Host "--- מתחיל ניתוח מערכת עמוק (מצב אבחון בלבד) ---" -ForegroundColor Cyan
}

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

# 6. סגירת תהליכים כבדי-זיכרון (רק עם הפרמטר -KillHighMemory)
if ($KillHighMemory) {
    Write-Host "`n[6/6] מחפש תהליכים שגוזלים מעל ${ThresholdMB}MB..." -ForegroundColor Red

    $HeavyProcs = Get-Process |
        Where-Object { ($_.WorkingSet / 1MB) -gt $ThresholdMB -and $_.Name -notin $ProtectedProcesses } |
        Sort-Object WorkingSet -Descending

    if ($HeavyProcs.Count -eq 0) {
        Write-Host "לא נמצאו תהליכים מעל הסף. אין מה לסגור." -ForegroundColor Green
    } else {
        Write-Host "נמצאו $($HeavyProcs.Count) תהליכים מעל הסף:`n" -ForegroundColor Yellow
        foreach ($proc in $HeavyProcs) {
            $ramMB = [Math]::Round($proc.WorkingSet / 1MB, 1)
            $answer = Read-Host "סגור '$($proc.Name)' (PID $($proc.Id), ${ramMB}MB)? [y/n]"
            if ($answer -eq 'y') {
                try {
                    Stop-Process -Id $proc.Id -Force -ErrorAction Stop
                    Write-Host "  >> '$($proc.Name)' נסגר בהצלחה." -ForegroundColor Green
                } catch {
                    Write-Host "  >> שגיאה בסגירת '$($proc.Name)': $_" -ForegroundColor Red
                }
            } else {
                Write-Host "  >> דילוג על '$($proc.Name)'." -ForegroundColor Gray
            }
        }
    }
    Write-Host "`n--- סיום מצב סגירת תהליכים ---" -ForegroundColor Cyan
}
