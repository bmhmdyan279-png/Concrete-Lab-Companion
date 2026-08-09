# verify.ps1 — Automatic integrity check for Concrete Lab Companion
# Run from the repository root (where build.py is)

$ErrorActionPreference = "Stop"
$repoRoot = $PSScriptRoot
Set-Location $repoRoot

Write-Host "══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Concrete Lab Companion — Verification Suite" -ForegroundColor Cyan
Write-Host "══════════════════════════════════════════════"

$pass = 0
$fail = 0

function Test-Check {
    param([string]$name, [bool]$condition, [string]$detail)
    if ($condition) {
        Write-Host "  ✅ $name" -ForegroundColor Green
        $script:pass++
    } else {
        Write-Host "  ❌ $name — $detail" -ForegroundColor Red
        $script:fail++
    }
}

# 1. Old files removed
Test-Check "Old build script removed" (-not (Test-Path "build_workbook.py")) "build_workbook.py still exists"
Test-Check "Excel folder removed" (-not (Test-Path "excel")) "excel/ folder still present"
Test-Check "Releases folder removed" (-not (Test-Path "releases")) "releases/ folder still present"

# 2. Required new files exist
$requiredFiles = @(
    "build.py", "config.yaml", "requirements.txt", "README.md",
    "CHANGELOG.md", "LICENSE", "CITATION.cff", "CONTRIBUTING.md",
    ".gitignore", "landing/index.html", "validation/errata.yaml",
    ".github/workflows/build.yml", ".github/ISSUE_TEMPLATE/bug_report.md",
    ".github/ISSUE_TEMPLATE/feature_request.md", ".github/PULL_REQUEST_TEMPLATE.md"
)
foreach ($f in $requiredFiles) {
    Test-Check "File exists: $f" (Test-Path $f) "Missing file"
}

# 3. Check config.yaml key name corrected
$cfg = Get-Content config.yaml -Raw
Test-Check "calc_fill present in config.yaml" ($cfg -match "calc_fill:") "Still has calculation_fill?"

# 4. Build the workbook (if not already built)
Write-Host "`n🔨 Building workbook..." -ForegroundColor Yellow
# جایگزین کردن بلوک فعلی با این:
try {
    $buildOutput = & python build.py 2>&1
    $buildSuccess = ($LASTEXITCODE -eq 0) -or ($buildOutput -match "✅ Build complete!")
} catch {
    $buildSuccess = $false
}
Test-Check "Build script runs successfully" $buildSuccess "build.py failed. See output: $buildOutput"

# 5. Verify built sheets match config
if (Test-Path "output/Concrete_Lab_Companion_v1.0.0.xlsx") {
    try {
        $wb = New-Object -ComObject Excel.Application
        $wb.Visible = $false
        $workbook = $wb.Workbooks.Open((Resolve-Path "output/Concrete_Lab_Companion_v1.0.0.xlsx").Path)
        $sheetNames = $workbook.Sheets | ForEach-Object { $_.Name }
        $workbook.Close($false)
        $wb.Quit()
        [System.Runtime.Interopservices.Marshal]::ReleaseComObject($wb) | Out-Null

        # Read expected sheets from config
        $yamlContent = Get-Content config.yaml -Raw
        # Very basic YAML extraction for key_sheets list
        $inKeySheets = $false
        $expectedSheets = @()
        foreach ($line in ($yamlContent -split "`n")) {
            if ($line -match "key_sheets:") { $inKeySheets = $true; continue }
            if ($inKeySheets) {
                if ($line -match "^\s+- ""(.+)""") {
                    $expectedSheets += $matches[1]
                } elseif ($line -match "^\s+- \w") {
                    # maybe another list item without quotes? skip
                } else {
                    # end of list
                    break
                }
            }
        }
        $missing = $expectedSheets | Where-Object { $_ -notin $sheetNames }
        Test-Check "All expected sheets present in built file" ($missing.Count -eq 0) "Missing sheets: $($missing -join ', ')"
    } catch {
        Test-Check "Excel COM check" $false "Could not open workbook (maybe Excel not installed?). Skipping sheet check."
    }
} else {
    Test-Check "Built workbook exists" $false "output/Concrete_Lab_Companion_v1.0.0.xlsx not found"
}

# 6. Hash consistency
$hashFile = "output/Concrete_Lab_Companion_v1.0.0.xlsx.sha256"
if (Test-Path $hashFile) {
    $hashFromFile = (Get-Content $hashFile -Raw).Split(" ")[0]
    $htmlHash = (Select-String -Path landing/index.html -Pattern '[a-f0-9]{64}' -AllMatches).Matches.Value | Select-Object -First 1
    Test-Check "SHA-256 hash matches in index.html" ($hashFromFile -eq $htmlHash) "Hash in landing page differs from build. File: $hashFromFile, HTML: $htmlHash"
} else {
    Test-Check "Hash file exists" $false "output/*.sha256 not found. Run build.py first."
}

# 7. Golden cases validation
$goldenDir = "validation/golden_cases"
if (Test-Path $goldenDir) {
    $jsonFiles = Get-ChildItem $goldenDir -Filter *.json
    $validCount = 0
    foreach ($f in $jsonFiles) {
        try {
            $json = Get-Content $f.FullName -Raw | ConvertFrom-Json
            if ($json.test_id -and $json.inputs -and $json.expected) {
                $validCount++
            }
        } catch {
            Write-Host "    ❌ Invalid JSON: $($f.Name)" -ForegroundColor Red
            $script:fail++
        }
    }
    Test-Check "All golden cases valid JSON" ($validCount -eq $jsonFiles.Count) "Some files are invalid or missing fields"
    Write-Host "    ($validCount/$($jsonFiles.Count) valid)" -ForegroundColor DarkGray
} else {
    Test-Check "Golden cases directory exists" $false "$goldenDir missing"
}

# 8. Errata YAML structure
$errata = Get-Content validation/errata.yaml -Raw
$errataValid = $errata -match "errata:" -and $errata -match "id:" -and $errata -match "status:"
Test-Check "errata.yaml structure looks correct" $errataValid "Check file manually"

# 9. .gitignore should exclude .xlsx
$gitignore = Get-Content .gitignore -Raw
Test-Check ".gitignore excludes *.xlsx" ($gitignore -match "\*\.xlsx") "*.xlsx not in gitignore; output may be committed"

# 10. GitHub Actions workflow present
Test-Check "CI workflow exists" (Test-Path ".github/workflows/build.yml") "CI missing"

# Summary
Write-Host "`n══════════════════════════════════════════════" -ForegroundColor Cyan
Write-Host "  Results: $pass passed, $fail failed" -ForegroundColor $(if ($fail -eq 0) { "Green" } else { "Red" })
Write-Host "══════════════════════════════════════════════" -ForegroundColor Cyan

if ($fail -gt 0) {
    exit 1
} else {
    Write-Host "All checks passed! 🎉" -ForegroundColor Green
    exit 0
}