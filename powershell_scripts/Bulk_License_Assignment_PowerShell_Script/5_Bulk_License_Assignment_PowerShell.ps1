<#
AJ Labs PowerShell Script for Bulk License Assignment (aj-labs.com)

- Input columns expected: userPrincipalName, usageLocation, SkuPartNumber
- For each row:
    * Update-MgUser -UsageLocation
    * Assign license from SkuPartNumber via Set-MgUserLicense
- Output file: Output_Log_License_DD_MMM_YYYY_XX.xlsx (never overwrites)

-Create Entra ID users from Excel
- PowerShell 7+ recommended
- Modules (install once if missing):
      Install-Module Microsoft.Graph -Scope CurrentUser -Force
      Install-Module ImportExcel   -Scope CurrentUser -Force

-Permissions: When prompted by Connect-MgGraph, consent to: User.ReadWrite.All
#>



# ====== LOCATION SETTINGS (EDIT THESE) ======
$ExcelFolder   = "E:\AJ Labs 2\Advanced Microsoft Teams Training\Training Docs\PowerShell\Administration via PowerShell"
$ExcelFileName = "AJ_EntraID_License_Management_ Template.xlsx"   # <-- note the space before 'Template'
$TenantId      = ""   # Optional tenant GUID to reduce prompts
# ============================================

# ---- Behavior flags ----
$DryRun = $false    # Set $true to preview without making changes
$ThrottleEvery = 15 # sleep a bit after every N processed rows (avoids throttling)
$ThrottleSecs  = 2

# ---- Ensure folder exists ----
if (-not (Test-Path -Path $ExcelFolder -PathType Container)) {
    New-Item -Path $ExcelFolder -ItemType Directory -Force | Out-Null
}
$ExcelPath = Join-Path -Path $ExcelFolder -ChildPath $ExcelFileName

# ---- Modules ----
Import-Module ImportExcel -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop
Import-Module Microsoft.Graph.Users.Actions -ErrorAction Stop
Import-Module Microsoft.Graph.Identity.DirectoryManagement -ErrorAction Stop

# ---- Graph sign-in (single prompt guard) ----
$neededScopes = @('User.ReadWrite.All','Directory.Read.All')
$needConnect  = $true
try {
    $ctx = Get-MgContext -ErrorAction Stop
    if ($ctx -and $ctx.Account -and $ctx.Scopes) {
        $hasAll = $true
        foreach ($s in $neededScopes) { if (-not ($ctx.Scopes -contains $s)) { $hasAll = $false; break } }
        if ($hasAll) { $needConnect = $false }
    }
} catch { $needConnect = $true }

if ($needConnect) {
    Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
    if ([string]::IsNullOrWhiteSpace($TenantId)) {
        Connect-MgGraph -Scopes $neededScopes -NoWelcome | Out-Null
    } else {
        Connect-MgGraph -Scopes $neededScopes -TenantId $TenantId -NoWelcome | Out-Null
    }
}

# ========= READ EXCEL USING DISPLAYED TEXT (no coercion) =========
try {
    $ResolvedPath = (Resolve-Path -Path $ExcelPath -ErrorAction Stop).Path
} catch {
    throw "Excel file not found: $ExcelPath`nPlace it there or update `$ExcelFolder / `$ExcelFileName."
}

$pkg = Open-ExcelPackage -Path $ResolvedPath -ErrorAction Stop
try {
    $ws = $pkg.Workbook.Worksheets | Where-Object { -not $_.Hidden } | Select-Object -First 1
    if (-not $ws) { throw "No visible worksheet found in $ResolvedPath" }

    $rowCount = $ws.Dimension.End.Row
    $colCount = $ws.Dimension.End.Column
    if ($rowCount -lt 2 -or $colCount -lt 1) { throw "No data rows found in $ResolvedPath" }

    # Headers (Row 1) as displayed text
    $InputHeaders = @()
    for ($c=1; $c -le $colCount; $c++) { $InputHeaders += $ws.Cells[1,$c].Text }

    # Data rows (Text)
    $rows = New-Object System.Collections.Generic.List[object]
    for ($r=2; $r -le $rowCount; $r++) {
        $obj = [ordered]@{}
        for ($c=1; $c -le $colCount; $c++) {
            $hdr = $InputHeaders[$c-1]
            $obj[$hdr] = $ws.Cells[$r,$c].Text
        }
        $rows.Add([pscustomobject]$obj) | Out-Null
    }
} finally {
    Close-ExcelPackage -ExcelPackage $pkg -NoSave:$true
}
if (-not $rows -or $rows.Count -eq 0) { throw "No data rows found in $ResolvedPath" }
# ================================================================

# ---- Helper: case/space-insensitive header match ----
function Header-Is {
    param([string]$Header, [string[]]$Variants)
    $h = ($Header -replace '\s','').ToLower()
    foreach ($v in $Variants) { if ($h -eq (($v -replace '\s','').ToLower())) { return $true } }
    return $false
}

# ---- Column variants (tolerant) ----
$MAP = @{
  userPrincipalName = @('userPrincipalName','UserPrincipalName','UPN','user principal name')
  usageLocation     = @('usageLocation','UsageLocation','Usage Location','CountryCode')
  SkuPartNumber     = @('SkuPartNumber','SKU','Sku','Sku Part Number','License','Plan','Sku Part')
}

function Get-Col {
  param([Parameter(Mandatory=$true)] $Row, [Parameter(Mandatory=$true)] [string[]] $HeaderVariants)
  foreach ($h in $HeaderVariants) {
    $prop = $Row.PSObject.Properties | Where-Object { $_.Name -eq $h }
    if ($prop) { $val = $prop.Value; if ($null -ne $val -and "$val".Trim() -ne '') { return "$val" } }
    $prop2 = $Row.PSObject.Properties | Where-Object {
      ($_.Name -replace '\s','').ToLower() -eq ($h -replace '\s','').ToLower()
    }
    if ($prop2) { $val2 = $prop2.Value; if ($null -ne $val2 -and "$val2".Trim() -ne '') { return "$val2" } }
  }
  return $null
}

# ---- Utilities ----
function Normalize-UsageLocation($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return $null }
    $t = $val.ToString().Trim()
    return $t.Substring(0, [Math]::Min(2, $t.Length)).ToUpperInvariant()
}
function Get-ExcelColumnName { param([int]$Index)
    if ($Index -lt 1) { throw "Index must be 1 or greater." }
    $name = ""
    while ($Index -gt 0) { $rem = ($Index - 1) % 26; $name = [char](65 + $rem) + $name; $Index = [math]::Floor(($Index - 1) / 26) }
    return $name
}
function Needs-Apostrophe {
    param([string]$Header, [string]$Value)
    if ([string]::IsNullOrEmpty($Value)) { return $false }
    # Start with 0 / + / any non-alphanumeric (saves leading zeros, formulas, etc.)
    return ($Value -match '^(0|\+|[^A-Za-z0-9])')
}

# ---- Cache tenant SKUs (SkuPartNumber -> SkuId) ----
Write-Host "Loading tenant subscribed SKUs..." -ForegroundColor Cyan
$skuMap = @{}
try {
    $subs = Get-MgSubscribedSku -All
    foreach ($s in $subs) {
        if ($s.SkuPartNumber -and $s.SkuId) {
            $skuMap[$s.SkuPartNumber.ToUpperInvariant()] = $s.SkuId
        }
    }
} catch {
    throw "Failed to read tenant subscribed SKUs: $($_.Exception.Message)"
}

# ---- Process rows ----
Write-Host "Processing $($rows.Count) row(s)..." -ForegroundColor Cyan
$logRows = New-Object System.Collections.Generic.List[object]
$processed = 0

foreach ($row in $rows) {
    # Start log row as exact copy of input
    $log = [ordered]@{}
    foreach ($h in $InputHeaders) { $log[$h] = "$($row.$h)" }
    $log['Status'] = ''   # fill at end

    $upn  = Get-Col $row $MAP.userPrincipalName
    $loc  = Get-Col $row $MAP.usageLocation
    $skuP = Get-Col $row $MAP.SkuPartNumber

    if ([string]::IsNullOrWhiteSpace($upn)) {
        $log['Status'] = "Skipped - Missing userPrincipalName"
        $logRows.Add([pscustomobject]$log) | Out-Null
        continue
    }

    $messages = @()

    try {
        if ($DryRun) {
            $messages += "[DRY-RUN] Would update UsageLocation to '$loc' and assign '$skuP'"
        } else {
            # 1) Set UsageLocation if provided
            if (-not [string]::IsNullOrWhiteSpace($loc)) {
                $normLoc = Normalize-UsageLocation $loc
                if ($normLoc) {
                    Update-MgUser -UserId $upn -UsageLocation $normLoc -ErrorAction Stop
                    $messages += "UsageLocation set to '$normLoc'"
                }
            }

            # 2) Assign license if SkuPartNumber provided
            if (-not [string]::IsNullOrWhiteSpace($skuP)) {
                $key = $skuP.Trim().ToUpperInvariant()
                if (-not $skuMap.ContainsKey($key)) {
                    throw "SkuPartNumber '$skuP' not found in tenant SubscribedSkus"
                }
                $skuId = $skuMap[$key]

                # Optional: avoid duplicate assignment
                $assigned = @()
                try {
                    $u = Get-MgUser -UserId $upn -Property AssignedLicenses -ErrorAction Stop
                    if ($u -and $u.AssignedLicenses) { $assigned = $u.AssignedLicenses.SkuId }
                } catch {}

                if ($assigned -and ($assigned -contains $skuId)) {
                    $messages += "License '$skuP' already assigned"
                } else {
                    Set-MgUserLicense -UserId $upn -AddLicenses @(@{SkuId=$skuId}) -RemoveLicenses @() -ErrorAction Stop
                    $messages += "License '$skuP' assigned"
                }
            }
        }

        if ($messages.Count -gt 0) {
            if ($messages -join '; ' -match 'assigned') {
                $log['Status'] = "License Assigned; " + ($messages -join '; ')
            } else {
                $log['Status'] = ($messages -join '; ')
            }
        } else {
            $log['Status'] = "No changes"
        }

    } catch {
        $log['Status'] = "Error: $($_.Exception.Message)"
    }

    $logRows.Add([pscustomobject]$log) | Out-Null

    $processed++
    if (-not $DryRun -and $ThrottleEvery -gt 0 -and ($processed % $ThrottleEvery -eq 0)) {
        Start-Sleep -Seconds $ThrottleSecs
    }
}

# ---- Build output filename: Output_Log_License_DD_MMM_YYYY_XX.xlsx ----
$todayToken = (Get-Date).ToString("dd_MMM_yyyy").ToUpper()
$baseName   = "Output_Log_License_{0}" -f $todayToken
$existing = Get-ChildItem -Path $ExcelFolder -Filter ("{0}_*.xlsx" -f $baseName) -ErrorAction SilentlyContinue
$nextIndex = 1
if ($existing) {
    $max = 0
    foreach ($f in $existing) {
        if ($f.BaseName -match "^$([regex]::Escape($baseName))_(\d{2})$") {
            $i = [int]$matches[1]
            if ($i -gt $max) { $max = $i }
        }
    }
    $nextIndex = $max + 1
}
$indexToken = "{0:00}" -f $nextIndex
$logXlsx = Join-Path -Path $ExcelFolder -ChildPath ("{0}_{1}.xlsx" -f $baseName, $indexToken)

# ---- Write Output Excel (Text columns; conditional apostrophe) ----
$orderedCols = @($InputHeaders + 'Status')

$pkgOut = Open-ExcelPackage -Path $logXlsx -Create
try {
    $wsOut = Add-Worksheet -ExcelPackage $pkgOut -WorksheetName 'Log'

    # Set headers and force Text on all columns
    for ($c=1; $c -le $orderedCols.Count; $c++) {
        $hdr = $orderedCols[$c-1]
        $wsOut.Cells[1,$c].Value = $hdr
        $colLetter = Get-ExcelColumnName $c
        $wsOut.Cells["$($colLetter):$($colLetter)"].Style.Numberformat.Format = '@'
    }

    # Write rows (prefix apostrophe only when needed)
    $r = 2
    foreach ($lr in $logRows) {
        for ($c=1; $c -le $orderedCols.Count; $c++) {
            $hdr = $orderedCols[$c-1]
            $val = ""
            if ($lr.PSObject.Properties.Name -contains $hdr) { $val = "$($lr.$hdr)" }

            if (Needs-Apostrophe -Header $hdr -Value $val) {
                $wsOut.Cells[$r,$c].Value = "'$val"
            } else {
                $wsOut.Cells[$r,$c].Value = $val
            }
        }
        $r++
    }

    Close-ExcelPackage -ExcelPackage $pkgOut   # save
    Write-Host "Detailed log written to: $logXlsx" -ForegroundColor Cyan
} catch {
    try { Close-ExcelPackage -ExcelPackage $pkgOut -NoSave:$true } catch {}
    throw
}
