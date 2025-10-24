<#
AJ Labs PowerShell Script for Entra ID Bulk User Creation (aj-labs.com)

-Create Entra ID users from Excel
- PowerShell 7+ recommended
- Modules (install once if missing):
      Install-Module Microsoft.Graph -Scope CurrentUser -Force
      Install-Module ImportExcel   -Scope CurrentUser -Force

-Permissions: When prompted by Connect-MgGraph, consent to: User.ReadWrite.All

- Excel headers (case-insensitive ok)
  Mandatory:
    userPrincipalName, displayName, firstName, lastName, MailNickname, Initial password
  Optional:
    userType, mail, accountEnabled, usageLocation, jobTitle, department, companyName,
    employeeID, employeeType, employeeHireDate, officeLocation,
    streetAddress, city, state, postalCode, country,
    telephoneNumber, mobilePhone
#>

# ====== LOCATION SETTINGS (EDIT THESE) ======
$ExcelFolder   = "E:\AJ Labs 2\Advanced Microsoft Teams Training\Training Docs\PowerShell\Administration via PowerShell" #Update your file location here
$ExcelFileName = "AJ_EntraID_User_Create_Template.xlsx" #Update your file name here
$TenantId      = ""   # Optional tenant GUID to reduce extra prompts
# ============================================

# Behavior flags
$ForceChangePassword = $true
$DryRun = $false  # Set $true to inspect/validate before actually creating

# Ensure folder exists
if (-not (Test-Path -Path $ExcelFolder -PathType Container)) {
    New-Item -Path $ExcelFolder -ItemType Directory -Force | Out-Null
}
$ExcelPath = Join-Path -Path $ExcelFolder -ChildPath $ExcelFileName

# -------- Modules --------
Import-Module ImportExcel -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# -------- Graph sign-in (single prompt guard) --------
$neededScopes = @('User.ReadWrite.All')
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
    throw "Excel file not found: $ExcelPath`nPlace the file there or update `$ExcelFolder / `$ExcelFileName."
}

$pkg = Open-ExcelPackage -Path $ResolvedPath -ErrorAction Stop
try {
    $ws = $pkg.Workbook.Worksheets | Where-Object { -not $_.Hidden } | Select-Object -First 1
    if (-not $ws) { throw "No visible worksheet found in $ResolvedPath" }

    $rowCount = $ws.Dimension.End.Row
    $colCount = $ws.Dimension.End.Column
    if ($rowCount -lt 2 -or $colCount -lt 1) { throw "No data rows found in $ResolvedPath" }

    $InputHeaders = @()
    for ($c=1; $c -le $colCount; $c++) { $InputHeaders += $ws.Cells[1,$c].Text }

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

# -------- Mandatory headers (variants) --------
$MANDATORY_MAP = @{
  userPrincipalName = @('userPrincipalName','UserPrincipalName','UPN','user principal name')
  displayName       = @('displayName','DisplayName','Display Name','name')
  firstName         = @('firstName','FirstName','First Name','GivenName','Given Name')
  lastName          = @('lastName','LastName','Last Name','Surname')
  MailNickname      = @('MailNickname','mailNickname','Mail Nickname','Alias')
  InitialPassword   = @('Initial password','Initial Password','initialPassword','Password')
}

# -------- Optional headers (variants) --------
$OPTIONAL_MAP = @{
  userType        = @('userType','UserType')
  mail            = @('mail','Mail','Email','Email Address')
  accountEnabled  = @('accountEnabled','AccountEnabled','Account Enabled','Enabled')
  usageLocation   = @('usageLocation','UsageLocation','Usage Location','CountryCode')
  jobTitle        = @('jobTitle','JobTitle','Job Title','Title')
  department      = @('department','Department')
  companyName     = @('companyName','CompanyName','Company Name','Organization')
  employeeID      = @('employeeID','employeeId','EmployeeID','Employee Id','Employee ID')
  employeeType     = @('employeeType','EmployeeType','Employee Type')
  employeeHireDate = @('employeeHireDate','EmployeeHireDate','HireDate','Hire Date','Date of Joining','DOJ')
  officeLocation  = @('officeLocation','OfficeLocation','Office Location','Location')
  streetAddress   = @('streetAddress','StreetAddress','Street Address','Address')
  city            = @('city','City')
  state           = @('state','State','Province','State/Province')
  postalCode      = @('postalCode','PostalCode','Postal Code','Zip','Zip Code','PIN')
  country         = @('country','Country','Country/Region')
  telephoneNumber = @('telephoneNumber','TelephoneNumber','Telephone Number','Business Phone','Phone')
  mobilePhone     = @('mobilePhone','MobilePhone','Mobile Phone','Cell','Cell Phone')
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

function Convert-ToBool($val, $default=$true) {
    if ($null -eq $val -or ($val -is [string] -and $val.Trim() -eq '')) { return $default }
    switch -Regex ($val.ToString()) {
        '^(true|1|yes|y)$'  { return $true }
        '^(false|0|no|n)$'  { return $false }
        default             { return $default }
    }
}
function Normalize-UsageLocation($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return $null }
    $t = $val.ToString().Trim()
    return $t.Substring(0, [Math]::Min(2, $t.Length)).ToUpperInvariant()
}
function Build-BusinessPhones($val) {
    if ([string]::IsNullOrWhiteSpace($val)) { return @() }
    return ($val -split '[;,]') | ForEach-Object { $_.Trim() } | Where-Object { $_ -ne '' }
}
function Resolve-HireDateIso($value) {
    if ($null -eq $value) { return $null }
    $s = "$value".Trim()
    if ($s -eq '') { return $null }
    if ($s -match '^\d{4}-\d{2}-\d{2}$') { return $s }
    $dnum = 0.0
    if ([double]::TryParse($s, [System.Globalization.NumberStyles]::Float, [System.Globalization.CultureInfo]::InvariantCulture, [ref]$dnum)) {
        try { return ([DateTime]::FromOADate($dnum)).ToString('yyyy-MM-dd') } catch {}
    }
    $dt = [datetime]::MinValue
    if ([DateTime]::TryParse($s, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
        return $dt.ToString('yyyy-MM-dd')
    }
    $fmts = @('yyyy-MM-dd','dd-MM-yyyy','MM-dd-yyyy','dd/MM/yyyy','MM/dd/yyyy','yyyy/MM/dd','dd.M.yyyy','M/d/yyyy','d/M/yyyy')
    foreach ($f in $fmts) {
        $dt = [datetime]::MinValue
        if ([DateTime]::TryParseExact($s, $f, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString('yyyy-MM-dd')
        }
    }
    return $s
}
function Get-ExcelColumnName { param([int]$Index)
    if ($Index -lt 1) { throw "Index must be 1 or greater." }
    $name = ""
    while ($Index -gt 0) { $rem = ($Index - 1) % 26; $name = [char](65 + $rem) + $name; $Index = [math]::Floor(($Index - 1) / 26) }
    return $name
}
function Header-Is {
    param([string]$Header, [string[]]$Variants)
    $h = ($Header -replace '\s','').ToLower()
    foreach ($v in $Variants) { if ($h -eq (($v -replace '\s','').ToLower())) { return $true } }
    return $false
}

# -------- Main --------
$results = New-Object System.Collections.Generic.List[Object]
$logRows = New-Object System.Collections.Generic.List[Object]
$createdCount = 0

Write-Host "Processing $($rows.Count) row(s) from Excel..." -ForegroundColor Cyan

foreach ($row in $rows) {

    # Exact copy of input row
    $log = [ordered]@{}
    foreach ($h in $InputHeaders) { $log[$h] = "$($row.$h)" }
    $log['userIDString'] = ''  # renamed from NewUserId
    $log['Comments']     = ''

    # Mandatory
    $upn          = Get-Col $row $MANDATORY_MAP.userPrincipalName
    $displayName  = Get-Col $row $MANDATORY_MAP.displayName
    $firstName    = Get-Col $row $MANDATORY_MAP.firstName
    $lastName     = Get-Col $row $MANDATORY_MAP.lastName
    $mailNickname = Get-Col $row $MANDATORY_MAP.MailNickname
    $initPwd      = Get-Col $row $MANDATORY_MAP.InitialPassword

    $errList = @()
    if ([string]::IsNullOrWhiteSpace($upn))          { $errList += 'userPrincipalName' }
    if ([string]::IsNullOrWhiteSpace($displayName))  { $errList += 'displayName' }
    if ([string]::IsNullOrWhiteSpace($firstName))    { $errList += 'firstName' }
    if ([string]::IsNullOrWhiteSpace($lastName))     { $errList += 'lastName' }
    if ([string]::IsNullOrWhiteSpace($mailNickname)) { $errList += 'MailNickname' }
    if ([string]::IsNullOrWhiteSpace($initPwd))      { $errList += 'Initial password' }

    if ($errList.Count -gt 0) {
        $results.Add([pscustomobject]@{ UserPrincipalName=$upn; Status="Skipped - Missing: $($errList -join ', ')"; NewUserId=$null; Message="" })
        $log['Comments'] = "Skipped - Missing: $($errList -join ', ')"
        $logRows.Add([pscustomobject]$log) | Out-Null
        continue
    }

    # Optional (strings)
    $userType         = Get-Col $row $OPTIONAL_MAP.userType
    $mail             = Get-Col $row $OPTIONAL_MAP.mail
    $accountEnabled   = Get-Col $row $OPTIONAL_MAP.accountEnabled
    $usageLocation    = Get-Col $row $OPTIONAL_MAP.usageLocation
    $jobTitle         = Get-Col $row $OPTIONAL_MAP.jobTitle
    $department       = Get-Col $row $OPTIONAL_MAP.department
    $companyName      = Get-Col $row $OPTIONAL_MAP.companyName
    $employeeID       = Get-Col $row $OPTIONAL_MAP.employeeID
    $employeeType     = Get-Col $row $OPTIONAL_MAP.employeeType
    $employeeHireDate = Get-Col $row $OPTIONAL_MAP.employeeHireDate
    $officeLocation   = Get-Col $row $OPTIONAL_MAP.officeLocation
    $streetAddress    = Get-Col $row $OPTIONAL_MAP.streetAddress
    $city             = Get-Col $row $OPTIONAL_MAP.city
    $state            = Get-Col $row $OPTIONAL_MAP.state
    $postalCode       = Get-Col $row $OPTIONAL_MAP.postalCode
    $country          = Get-Col $row $OPTIONAL_MAP.country
    $telephoneNumber  = Get-Col $row $OPTIONAL_MAP.telephoneNumber
    $mobilePhone      = Get-Col $row $OPTIONAL_MAP.mobilePhone

    # PasswordProfile
    $passwordProfile = @{ Password="$initPwd"; ForceChangePasswordNextSignIn=$ForceChangePassword }

    # Graph params
    $params = @{
        UserPrincipalName = "$upn".Trim()
        DisplayName       = "$displayName".Trim()
        GivenName         = "$firstName".Trim()
        Surname           = "$lastName".Trim()
        MailNickname      = "$mailNickname".Trim()
        PasswordProfile   = $passwordProfile
        AccountEnabled    = Convert-ToBool $accountEnabled $true
        UserType          = if ($userType) { "$userType".Trim() } else { "Member" }
    }

    if ($mail)                { $params['Mail']            = "$mail".Trim() }
    if ($usageLocation)       { $params['UsageLocation']   = (Normalize-UsageLocation $usageLocation) }
    if ($jobTitle)            { $params['JobTitle']        = "$jobTitle" }
    if ($department)          { $params['Department']      = "$department" }
    if ($companyName)         { $params['CompanyName']     = "$companyName" }
    if ($employeeID)          { $params['EmployeeId']      = "$employeeID" }
    if ($employeeType)        { $params['EmployeeType']    = "$employeeType" }

    $hireIso = if ($employeeHireDate) { Resolve-HireDateIso $employeeHireDate } else { $null }
    if ($hireIso) { $params['EmployeeHireDate'] = $hireIso }

    if ($officeLocation)      { $params['OfficeLocation']  = "$officeLocation" }
    if ($streetAddress)       { $params['StreetAddress']   = "$streetAddress" }
    if ($city)                { $params['City']            = "$city" }
    if ($state)               { $params['State']           = "$state" }
    if ($postalCode)          { $params['PostalCode']      = "$postalCode" }
    if ($country)             { $params['Country']         = "$country" }
    if ($mobilePhone)         { $params['MobilePhone']     = "$mobilePhone" }

    $bp = Build-BusinessPhones $telephoneNumber
    if ($bp.Count -gt 0)      { $params['BusinessPhones']  = $bp }

    # Idempotency
    $exists = $null
    try { $exists = Get-MgUser -UserId "$upn" -ErrorAction Stop } catch {}

    if ($exists) {
        Write-Host "Skipping existing user: $upn" -ForegroundColor Yellow
        $results.Add([pscustomobject]@{ UserPrincipalName=$upn; Status="Skipped - Exists"; NewUserId=$exists.Id; Message="" })
        $log['userIDString'] = "$($exists.Id)"
        $log['Comments']     = "Skipped - Exists"
        $logRows.Add([pscustomobject]$log) | Out-Null
        continue
    }

    if ($DryRun) {
        Write-Host "`n[DRY-RUN] Preview for $upn" -ForegroundColor Yellow
        $params.GetEnumerator() | Sort-Object Key | Format-Table -AutoSize
        $results.Add([pscustomobject]@{ UserPrincipalName=$upn; Status="DRY-RUN"; NewUserId=$null; Message="" })
        $log['Comments']  = "DRY-RUN"
        $logRows.Add([pscustomobject]$log) | Out-Null
        continue
    }

    try {
        $new = New-MgUser @params -ErrorAction Stop
        $newUserId = $new.Id
        Write-Host "Created: $upn  (Id: $newUserId)" -ForegroundColor Green
        $results.Add([pscustomobject]@{ UserPrincipalName=$upn; Status="Created"; NewUserId=$newUserId; Message="" })
        $log['userIDString'] = "$newUserId"
        $log['Comments']     = "User Created"
        $logRows.Add([pscustomobject]$log) | Out-Null
        $createdCount++; if ($createdCount % 10 -eq 0) { Start-Sleep -Seconds 3 }
    } catch {
        $err = $_.Exception.Message
        Write-Warning "Failed: $upn => $err"
        $results.Add([pscustomobject]@{ UserPrincipalName=$upn; Status="Failed"; NewUserId=$null; Message=$err })
        $log['userIDString'] = ''
        $log['Comments']     = $err
        $logRows.Add([pscustomobject]$log) | Out-Null
    }
}

# -------- Save Output_log only (no CSV summary) --------

# Build filename: Output_log_DD_MMM_YYYY_XX.xlsx (MMM uppercase)
$todayToken = (Get-Date).ToString("dd_MMM_yyyy").ToUpper()
$baseName   = "Output_log_{0}" -f $todayToken
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

# Output columns: input headers + userIDString + Comments
$orderedCols = @($InputHeaders + 'userIDString' + 'Comments')

# Build Output_log with Text format on all columns
$pkgOut = Open-ExcelPackage -Path $logXlsx -Create
try {
    $wsOut = Add-Worksheet -ExcelPackage $pkgOut -WorksheetName 'Log'

    # Headers and column text format
    for ($c=1; $c -le $orderedCols.Count; $c++) {
        $hdr = $orderedCols[$c-1]
        $wsOut.Cells[1,$c].Value = $hdr
        $colLetter = Get-ExcelColumnName $c
        $wsOut.Cells["$($colLetter):$($colLetter)"].Style.Numberformat.Format = '@'  # Text
    }

    # Decide when to prefix apostrophe
    function Needs-Apostrophe {
        param([string]$Header, [string]$Value)
        if ([string]::IsNullOrEmpty($Value)) { return $false }
        # Always for phone columns
        if (Header-Is -Header $Header -Variants $OPTIONAL_MAP.telephoneNumber) { return $true }
        if (Header-Is -Header $Header -Variants $OPTIONAL_MAP.mobilePhone)     { return $true }
        # Otherwise, only if starts with 0 / + / non-alphanumeric
        return ($Value -match '^(0|\+|[^A-Za-z0-9])')
    }

    # Write rows
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

    Close-ExcelPackage -ExcelPackage $pkgOut   # saves by default
    Write-Host "Detailed log written to: $logXlsx" -ForegroundColor Cyan
} catch {
    try { Close-ExcelPackage -ExcelPackage $pkgOut -NoSave:$true } catch {}
    throw
}

# Final on-screen table
$results | Format-Table -AutoSize
