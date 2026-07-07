# publish.ps1 — הדרך היחידה לפרסם שינוי בסקיל 720-script-qa.
# עושה בפקודה אחת: בדיקת CHANGELOG → סנכרון לעותק הגלובלי → commit+push → zip טרי.
# שימוש (מתוך תיקיית הסקיל):  .\publish.ps1 -Message "מה השתנה"
param(
    [Parameter(Mandatory = $true)][string]$Message
)
$ErrorActionPreference = "Stop"
$dev  = $PSScriptRoot
$glob = Join-Path $env:USERPROFILE ".claude\skills\720-script-qa"
$zip  = Join-Path (Split-Path (Split-Path (Split-Path $dev))) "720-script-qa.zip"  # QA_skill\720-script-qa.zip

# 1. CHANGELOG חייב לכלול את תאריך היום — בלי רישום אין פרסום
$today = Get-Date -Format "yyyy-MM-dd"
$changelog = Get-Content (Join-Path $dev "CHANGELOG.md") -Raw -Encoding UTF8
if ($changelog -notmatch [regex]::Escape("## $today")) {
    Write-Error "CHANGELOG.md לא עודכן להיום ($today). הוסיפו סעיף '## $today' עם השינוי — ואז פרסמו."
}

# 2. תזכורת כיול (הכיול עצמו רץ ע\"י סוכן — ר' SKILL.md, רתמת-כיול)
Write-Host "תזכורת: אם שונו בדיקות/references — להריץ את רתמת-הכיול (tests/calibration) לפני פרסום." -ForegroundColor Yellow

# 3. סנכרון dev → global (תוכן הסקיל בלבד, בלי .git)
if (-not (Test-Path $glob)) { New-Item -ItemType Directory -Path $glob -Force | Out-Null }
foreach ($item in 'SKILL.md','README.md','CHANGELOG.md','publish.ps1','references','scripts','examples','tests') {
    $p = Join-Path $dev $item
    if (Test-Path $p) { Copy-Item -LiteralPath $p -Destination $glob -Recurse -Force }
}
Write-Host "sync -> $glob" -ForegroundColor Green

# 4. git commit + push (אם יש מה)
Push-Location $dev
try {
    git add -A
    $pending = git status --porcelain
    if ($pending) {
        git commit -m $Message
        if ($LASTEXITCODE -ne 0) { throw "git commit נכשל" }
        git push
        if ($LASTEXITCODE -ne 0) { throw "git push נכשל" }
        Write-Host "pushed: $Message" -ForegroundColor Green
    } else {
        Write-Host "אין שינויים ל-commit (רק סנכרון+zip)." -ForegroundColor Yellow
    }
} finally { Pop-Location }

# 5. zip טרי לאדמין (SKILL.md בשורש הארכיון)
$stage = Join-Path $env:TEMP ("pkg720-" + [guid]::NewGuid().ToString("N").Substring(0,8))
New-Item -ItemType Directory -Path $stage -Force | Out-Null
foreach ($item in 'SKILL.md','README.md','CHANGELOG.md','references','scripts','examples','tests') {
    $p = Join-Path $dev $item
    if (Test-Path $p) { Copy-Item -LiteralPath $p -Destination $stage -Recurse -Force }
}
if (Test-Path $zip) { Remove-Item -LiteralPath $zip -Force }
Compress-Archive -Path (Join-Path $stage '*') -DestinationPath $zip -Force
Remove-Item -LiteralPath $stage -Recurse -Force
Write-Host ("zip: {0} ({1:N1} KB)" -f $zip, ((Get-Item $zip).Length / 1KB)) -ForegroundColor Green
Write-Host "פורסם בהצלחה." -ForegroundColor Green
