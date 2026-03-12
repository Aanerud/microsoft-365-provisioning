<#
.SYNOPSIS
    Pretty-prints the TextCraft Europe CSV with readable formatting.
.DESCRIPTION
    Reads textcraft-europe.csv and displays each person as a formatted card
    with colour-coded sections for identity, location, org, skills, and more.
.PARAMETER CsvPath
    Path to the CSV file. Defaults to config\textcraft-europe.csv relative to repo root.
.PARAMETER Filter
    Optional wildcard filter applied to name, department or jobTitle.
.PARAMETER Brief
    Show a compact summary table instead of full cards.
#>
param(
    [string]$CsvPath,
    [string]$Filter,
    [switch]$Brief
)

# --- resolve path ---
if (-not $CsvPath) {
    $CsvPath = Join-Path $PSScriptRoot "..\config\textcraft-europe.csv"
}
if (-not (Test-Path $CsvPath)) {
    Write-Host "File not found: $CsvPath" -ForegroundColor Red
    exit 1
}

$people = Import-Csv $CsvPath

# --- optional filter ---
if ($Filter) {
    $people = $people | Where-Object {
        $_.name       -like "*$Filter*" -or
        $_.department -like "*$Filter*" -or
        $_.jobTitle   -like "*$Filter*"
    }
}

$total = ($people | Measure-Object).Count

# --- helper: clean array-style strings like "['a','b']" -> "a, b" ---
function Clean-List([string]$raw) {
    if (-not $raw -or $raw -eq '[]') { return '' }
    $raw = $raw -replace "^\['" , ''
    $raw = $raw -replace "'\]$", ''
    $raw = $raw -replace "','", ', '
    $raw = $raw -replace "'", ''
    return $raw
}

# --- helper: wrap long text at a given width ---
function Wrap-Text([string]$text, [int]$width = 80, [string]$indent = '    ') {
    if (-not $text) { return @() }
    $words = $text -split '\s+'
    $lines = @()
    $line  = ''
    foreach ($w in $words) {
        if (($line.Length + $w.Length + 1) -gt $width) {
            $lines += "$indent$line"
            $line = $w
        } elseif ($line) {
            $line = "$line $w"
        } else {
            $line = $w
        }
    }
    if ($line) { $lines += "$indent$line" }
    return $lines
}

# --- helper: print a labelled field ---
function Print-Field([string]$label, [string]$value, [ConsoleColor]$color = 'Gray') {
    if (-not $value -or $value -eq '[]') { return }
    $pad = ' ' * [Math]::Max(0, 20 - $label.Length)
    Write-Host "  $label$pad" -ForegroundColor $color -NoNewline
    Write-Host $value
}

# --- helper: print a labelled multi-line field ---
function Print-ListField([string]$label, [string]$raw, [ConsoleColor]$color = 'Gray') {
    $cleaned = Clean-List $raw
    if (-not $cleaned) { return }
    $pad = ' ' * [Math]::Max(0, 20 - $label.Length)
    Write-Host "  $label$pad" -ForegroundColor $color -NoNewline
    Write-Host $cleaned
}

function Print-LongField([string]$label, [string]$value, [ConsoleColor]$color = 'Gray') {
    if (-not $value) { return }
    Write-Host "  $label" -ForegroundColor $color
    $wrapped = Wrap-Text $value 90 '    '
    foreach ($l in $wrapped) { Write-Host $l }
}

# ================================================================
#  BRIEF MODE - compact table
# ================================================================
if ($Brief) {
    Write-Host ""
    Write-Host "  TextCraft Europe - $total people" -ForegroundColor Cyan
    Write-Host ("  " + ('-' * 96)) -ForegroundColor DarkGray

    $fmt = "  {0,-25} {1,-30} {2,-18} {3,-20}"
    Write-Host ($fmt -f 'NAME','JOB TITLE','DEPARTMENT','LOCATION') -ForegroundColor Yellow
    Write-Host ("  " + ('-' * 96)) -ForegroundColor DarkGray

    foreach ($p in $people) {
        $loc = "$($p.city), $($p.country)"
        Write-Host ($fmt -f $p.name, $p.jobTitle, $p.department, $loc)
    }
    Write-Host ""
    return
}

# ================================================================
#  FULL CARD MODE
# ================================================================
$separator = '=' * 100
$thinSep   = '-' * 96

Write-Host ""
Write-Host "  $separator" -ForegroundColor DarkCyan
Write-Host "   TextCraft Europe - $total People" -ForegroundColor Cyan
Write-Host "  $separator" -ForegroundColor DarkCyan
Write-Host ""

$i = 0
foreach ($p in $people) {
    $i++

    # -- header --
    Write-Host ("  +" + ('-' * 98) + "+") -ForegroundColor DarkYellow
    Write-Host "  |  " -ForegroundColor DarkYellow -NoNewline
    Write-Host "$($p.name)" -ForegroundColor White -NoNewline
    Write-Host " - " -NoNewline -ForegroundColor DarkGray
    Write-Host "$($p.jobTitle)" -ForegroundColor Yellow -NoNewline
    $fill = 93 - $p.name.Length - $p.jobTitle.Length
    if ($fill -lt 0) { $fill = 0 }
    Write-Host ((' ' * $fill) + '|') -ForegroundColor DarkYellow
    Write-Host ("  +" + ('-' * 98) + "+") -ForegroundColor DarkYellow

    # -- identity --
    Write-Host "  IDENTITY" -ForegroundColor Cyan
    Print-Field 'Email'          $p.email             Cyan
    Print-Field 'Employee ID'    $p.employeeId        Cyan
    Print-Field 'Employee Type'  $p.employeeType      Cyan
    Print-Field 'Hire Date'      $p.employeeHireDate  Cyan

    # -- organization --
    Write-Host "  ORGANIZATION" -ForegroundColor Green
    Print-Field 'Department'     $p.department         Green
    Print-Field 'Role'           $p.role               Green
    Print-Field 'Company'        $p.companyName        Green
    Print-Field 'Manager'        $p.ManagerEmail       Green
    Print-Field 'VTeam'          $p.VTeam              Green
    Print-Field 'Cost Center'    $p.CostCenter         Green
    Print-Field 'Project Code'   $p.ProjectCode        Green
    Print-Field 'Benefit Plan'   $p.BenefitPlan        Green
    Print-Field 'Building Access'$p.BuildingAccess     Green

    # -- location --
    $addr = (@($p.streetAddress, $p.city, $p.state, $p.country, $p.postalCode) |
            Where-Object { $_ }) -join ', '
    Write-Host "  LOCATION" -ForegroundColor Magenta
    Print-Field 'Office'         $p.officeLocation     Magenta
    Print-Field 'Address'        $addr                 Magenta
    Print-Field 'Usage Location' $p.usageLocation      Magenta
    Print-Field 'Pref. Language' $p.preferredLanguage  Magenta
    Print-Field 'Mobile'         $p.mobilePhone        Magenta
    Print-ListField 'Business Phones' $p.businessPhones Magenta

    # -- professional --
    Write-Host "  PROFESSIONAL" -ForegroundColor Yellow
    Print-Field 'Writing Style'   $p.WritingStyle       Yellow
    Print-Field 'Specialization'  $p.Specialization     Yellow
    Print-ListField 'Skills'            $p.skills             Yellow
    Print-ListField 'Languages'         $p.languages          Yellow
    Print-ListField 'Certifications'    $p.certifications     Yellow
    Print-ListField 'Awards'            $p.awards             Yellow
    Print-ListField 'Education'         $p.educationalActivities Yellow

    # -- content --
    $hasContent = $p.aboutMe -or $p.interests -or $p.projects -or
                  $p.responsibilities -or $p.publications -or $p.patents
    if ($hasContent) {
        Write-Host "  CONTENT" -ForegroundColor DarkCyan
        Print-LongField  'About Me'         $p.aboutMe           DarkCyan
        Print-ListField  'Interests'        $p.interests         DarkCyan
        Print-ListField  'Projects'         $p.projects          DarkCyan
        Print-ListField  'Responsibilities' $p.responsibilities  DarkCyan
        Print-ListField  'Publications'     $p.publications      DarkCyan
        Print-ListField  'Patents'          $p.patents           DarkCyan
    }

    # -- divider --
    if ($i -lt $total) {
        Write-Host ""
        Write-Host "  $thinSep" -ForegroundColor DarkGray
        Write-Host ""
    }
}

Write-Host ""
Write-Host "  $separator" -ForegroundColor DarkCyan
Write-Host "   End - $total people listed" -ForegroundColor Cyan
Write-Host "  $separator" -ForegroundColor DarkCyan
Write-Host ""
