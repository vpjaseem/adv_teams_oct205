<# 
AJ Labs PowerShell Script for Entra ID Bulk User Creation / Update (Idempotent)

- Creates Entra ID users from Excel if missing
- If user already exists (matched by userPrincipalName), only changed properties are updated
- IMPORTANT: When updating existing users, the password is NOT changed

Requirements:
  PowerShell 7+
  Modules (install once if missing):
      Install-Module Microsoft.Graph -Scope CurrentUser -Force
      Install-Module ImportExcel   -Scope CurrentUser -Force
Permissions:
  When prompted by Connect-MgGraph, consent to: User.ReadWrite.All

Excel headers (case-insensitive ok)
  Mandatory:
    userPrincipalName, displayName, firstName, lastName, MailNickname, Initial password
  Optional:
    userType, mail, accountEnabled, usageLocation, jobTitle, department, companyName,
    employeeID, employeeType, employeeHireDate, officeLocation, streetAddress, city, state,
    postalCode, country, telephoneNumber, mobilePhone
#>

[CmdletBinding()]
param(
    [string]$ExcelFolder   = (Get-Location).Path,
    [string]$ExcelFileName = 'AJ_EntraID_User_Create_Template.xlsx',
    [string]$TenantId      = '',
    [bool]  $ForceChangePassword = $true
)

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

# -------- Paths --------
$ExcelPath = Join-Path -Path $ExcelFolder -ChildPath $ExcelFileName

# -------- Modules --------
Import-Module ImportExcel -ErrorAction Stop
Import-Module Microsoft.Graph.Users -ErrorAction Stop

# -------- Graph sign-in: ALWAYS prompt (force fresh interactive login) --------
$neededScopes = @('User.ReadWrite.All')
try { Disconnect-MgGraph -ErrorAction SilentlyContinue } catch {}
Write-Host "Signing in to Microsoft Graph..." -ForegroundColor Cyan
if ([string]::IsNullOrWhiteSpace($TenantId)) {
    Connect-MgGraph -Scopes $neededScopes -NoWelcome | Out-Null
} else {
    Connect-MgGraph -Scopes $neededScopes -TenantId $TenantId -NoWelcome | Out-Null
}

# ========= READ EXCEL USING DISPLAYED TEXT (no coercion) =========
try {
    $ResolvedPath = (Resolve-Path -Path $ExcelPath -ErrorAction Stop).Path
} catch {
    throw ("Excel file not found: {0}" -f $ExcelPath)
}

$pkg  = Open-ExcelPackage -Path $ResolvedPath
$rows = New-Object System.Collections.Generic.List[object]

try {
    $wsAll = $pkg.Workbook.Worksheets
    if (-not $wsAll -or $wsAll.Count -lt 1) {
        throw ("No worksheet found in {0}" -f $ResolvedPath)
    }

    # Prefer first visible sheet with data; then any sheet with data; then first sheet (1-based)
    $ws = $wsAll | Where-Object { $_.Hidden -eq 'Visible' -and $_.Dimension } | Select-Object -First 1
    if (-not $ws) { $ws = $wsAll | Where-Object { $_.Dimension } | Select-Object -First 1 }
    if (-not $ws) { $ws = $wsAll[1] }  # EPPlus is 1-based

    if (-not $ws) { throw ("No worksheet could be selected in {0}" -f $ResolvedPath) }
    $dim = $ws.Dimension
    if (-not $dim) { throw ("No data range found in {0}" -f $ResolvedPath) }

    # Read header row as text
    $headers = @()
    for ($c=$dim.Start.Column; $c -le $dim.End.Column; $c++) {
        $headers += $ws.Cells[1,$c].Text
    }

    # Keep original input headers for log
    $InputHeaders = $headers | ForEach-Object { $_ }

    # Read each row into PSCustomObject using .Text
    for ($r = $dim.Start.Row + 1; $r -le $dim.End.Row; $r++) {
        $obj = [ordered]@{}
        for ($c=$dim.Start.Column; $c -le $dim.End.Column; $c++) {
            $hdr = $ws.Cells[1,$c].Text
            if ([string]::IsNullOrWhiteSpace($hdr)) { continue }
            $obj[$hdr] = $ws.Cells[$r,$c].Text
        }
        # Skip empty rows (all values null/blank)
        $allBlank = $true
        foreach ($v in $obj.Values) { if (-not [string]::IsNullOrWhiteSpace("$v")) { $allBlank = $false; break } }
        if (-not $allBlank) {
            $rows.Add([pscustomobject]$obj) | Out-Null
        }
    }
} finally {
    Close-ExcelPackage -ExcelPackage $pkg -NoSave:$true
}
if (-not $rows -or $rows.Count -eq 0) { throw ("No data rows found in {0}" -f $ResolvedPath) }
# ================================================================

# -------- Header maps (case-insensitive) --------
$MANDATORY_MAP = @{
  userPrincipalName = @('userPrincipalName','UserPrincipalName','UPN','user principal name')
  displayName       = @('displayName','DisplayName','Display Name','name')
  firstName         = @('firstName','FirstName','First Name','GivenName','Given Name')
  lastName          = @('lastName','LastName','Last Name','Surname')
  MailNickname      = @('MailNickname','mailNickname','Mail Nickname','Alias')
  InitialPassword   = @('Initial password','Initial Password','initialPassword','Password')
}
$OPTIONAL_MAP = @{
  userType        = @('userType','UserType')
  mail            = @('mail','Mail','Email','Email Address')
  accountEnabled  = @('accountEnabled','AccountEnabled','Account Enabled','Enabled')
  usageLocation   = @('usageLocation','UsageLocation','Usage Location','CountryCode')
  jobTitle        = @('jobTitle','JobTitle','Job Title','Title')
  department      = @('department','Department')
  companyName     = @('companyName','CompanyName','Company Name','Organization')
  employeeID      = @('employeeID','employeeId','EmployeeID','Employee Id','Employee ID')
  employeeType    = @('employeeType','EmployeeType')
  employeeHireDate= @('employeeHireDate','EmployeeHireDate','HireDate','Hire Date')
  officeLocation  = @('officeLocation','OfficeLocation','Office Location','Office')
  streetAddress   = @('streetAddress','StreetAddress','Street','Address')
  city            = @('city','City','Town')
  state           = @('state','State','Province','StateOrProvince')
  postalCode      = @('postalCode','PostalCode','Zip','ZIP','ZIP Code','Zip Code')
  country         = @('country','Country','CountryOrRegion','Country/Region')
  telephoneNumber = @('telephoneNumber','TelephoneNumber','Phone','BusinessPhone','Business Phone')
  mobilePhone     = @('mobilePhone','MobilePhone','Mobile','CellPhone','Cell Phone')
}

function Header-Is {
    param([string]$Header, [string[]]$Variants)
    $h = ($Header -replace '\s','').ToLower()
    foreach ($v in $Variants) { if ($h -eq (($v -replace '\s','').ToLower())) { return $true } }
    return $false
}

function Get-Col {
    param(
        [pscustomobject]$Row,
        [string[]]$Variants
    )
    foreach ($prop in $Row.PSObject.Properties.Name) {
        if (Header-Is $prop $Variants) { return $Row.$prop }
    }
    return $null
}

function Convert-ToBool {
    param([string]$value, [bool]$default=$false)
    if ([string]::IsNullOrWhiteSpace($value)) { return $default }
    $s = "$value".Trim().ToLower()
    switch ($s) {
        'true' {'True'; break}
        'false' {'False'; break}
        'yes' {'True'; break}
        'no' {'False'; break}
        '1' {'True'; break}
        '0' {'False'; break}
        default { return $default }
    }
}

function Normalize-UsageLocation {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $cc = "$value".Trim().ToUpper()
    if ($cc.Length -gt 2) { $cc = $cc.Substring(0,2) }
    return $cc
}

function Resolve-HireDateIso {
    param([string]$value)
    if ([string]::IsNullOrWhiteSpace($value)) { return $null }
    $s = "$value".Trim()
    if ($s -match '^\d{4}-\d{2}-\d{2}$') { return $s }
    # try OADate / numeric
    $dnum = 0.0
    if ([double]::TryParse($s, [ref]$dnum)) {
        try { return ([DateTime]::FromOADate($dnum)).ToString('yyyy-MM-dd') } catch {}
    }
    # normal parse
    $dt = [datetime]::MinValue
    if ([DateTime]::TryParse($s, [ref]$dt)) { return $dt.ToString('yyyy-MM-dd') }
    # exacts
    $fmts = @('yyyy-MM-dd','dd-MM-yyyy','MM-dd-yyyy','dd/MM/yyyy','MM/dd/yyyy','yyyy/MM/dd','dd.M.yyyy','M/d/yyyy','d/M/yyyy')
    foreach ($f in $fmts) {
        $dt = [datetime]::MinValue
        if ([DateTime]::TryParseExact($s, $f, [System.Globalization.CultureInfo]::InvariantCulture, [System.Globalization.DateTimeStyles]::None, [ref]$dt)) {
            return $dt.ToString('yyyy-MM-dd')
        }
    }
    return $s
}

# -------- Validate presence of mandatory columns --------
$allHeaders = $InputHeaders
foreach ($mKey in $MANDATORY_MAP.Keys) {
    $found = $false
    foreach ($hdr in $allHeaders) { if (Header-Is $hdr $MANDATORY_MAP[$mKey]) { $found = $true; break } }
    if (-not $found) { throw ("Missing mandatory column for '{0}' (variants: {1})" -f $mKey, ($MANDATORY_MAP[$mKey] -join ', ')) }
}

# -------- Output scaffolding --------
$results = New-Object System.Collections.Generic.List[object]

$todayToken = (Get-Date).ToString('dd_MMM_yyyy')
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

# We'll build log rows as clean PSCustomObjects, then Export-Excel once.
$logRows = New-Object System.Collections.Generic.List[object]

# -------- Row processing --------
$rowIndex = 0
foreach ($row in $rows) {
    $rowIndex++

    try {
        # MANDATORY
        $upn           = Get-Col $row $MANDATORY_MAP.userPrincipalName
        $displayName   = Get-Col $row $MANDATORY_MAP.displayName
        $firstName     = Get-Col $row $MANDATORY_MAP.firstName
        $lastName      = Get-Col $row $MANDATORY_MAP.lastName
        $mailNickname  = Get-Col $row $MANDATORY_MAP.MailNickname
        $initPwd       = Get-Col $row $MANDATORY_MAP.InitialPassword

        if ([string]::IsNullOrWhiteSpace($upn) -or [string]::IsNullOrWhiteSpace($displayName) -or
            [string]::IsNullOrWhiteSpace($firstName) -or [string]::IsNullOrWhiteSpace($lastName) -or
            [string]::IsNullOrWhiteSpace($mailNickname)) {
            throw ("Row {0}: missing mandatory values." -f $rowIndex)
        }

        # OPTIONAL
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

        # Normalize
        $accountEnabledBool = Convert-ToBool $accountEnabled $true
        $usageLocation2     = Normalize-UsageLocation $usageLocation
        $hireIso            = if ($employeeHireDate) { Resolve-HireDateIso $employeeHireDate } else { $null }

        # -------- Build create params --------
        $passwordProfile = @{ Password="$initPwd"; ForceChangePasswordNextSignIn=$ForceChangePassword }

        $createParams = @{
            UserPrincipalName = "$upn".Trim()
            DisplayName       = "$displayName".Trim()
            GivenName         = "$firstName".Trim()
            Surname           = "$lastName".Trim()
            MailNickname      = "$mailNickname".Trim()
            PasswordProfile   = $passwordProfile
            AccountEnabled    = $accountEnabledBool
            UserType          = if ($userType) { "$userType".Trim() } else { "Member" }
        }
        if ($mail)              { $createParams['Mail']             = "$mail".Trim() }
        if ($usageLocation2)    { $createParams['UsageLocation']    = $usageLocation2 }
        if ($jobTitle)          { $createParams['JobTitle']         = "$jobTitle" }
        if ($department)        { $createParams['Department']       = "$department" }
        if ($companyName)       { $createParams['CompanyName']      = "$companyName" }
        if ($employeeID)        { $createParams['EmployeeId']       = "$employeeID" }
        if ($employeeType)      { $createParams['EmployeeType']     = "$employeeType" }
        if ($hireIso)           { $createParams['EmployeeHireDate'] = $hireIso }
        if ($officeLocation)    { $createParams['OfficeLocation']   = "$officeLocation" }
        if ($streetAddress)     { $createParams['StreetAddress']    = "$streetAddress" }
        if ($city)              { $createParams['City']             = "$city" }
        if ($state)             { $createParams['State']            = "$state" }
        if ($postalCode)        { $createParams['PostalCode']       = "$postalCode" }
        if ($country)           { $createParams['Country']          = "$country" }
        if ($mobilePhone)       { $createParams['MobilePhone']      = "$mobilePhone" }
        if ($telephoneNumber)   { $createParams['BusinessPhones']   = @("$telephoneNumber") }

        # -------- Existence check --------
        $existing = $null
        try {
            $existing = Get-MgUser -Filter "userPrincipalName eq '$($upn.Replace("'","''"))'" -ConsistencyLevel eventual -CountVariable null -ErrorAction Stop
        } catch {
            $existing = Get-MgUser -Filter "userPrincipalName eq '$($upn.Replace("'","''"))'" -ErrorAction Stop
        }

        $existingCount = if ($existing) { ($existing | Measure-Object).Count } else { 0 }

        if ($existingCount -eq 0) {
            # -------- Create --------
            $newUser = New-MgUser @createParams
            $uid     = if ($newUser.Id) { $newUser.Id } else { '' }

            $results.Add([pscustomobject]@{
                userPrincipalName = $upn
                Action            = 'Created'
                userIDString      = $uid
                Comments          = 'User created successfully.'
            }) | Out-Null
        }
        else {
            # -------- Update (idempotent) --------
            $u = $existing | Select-Object -First 1
            $updateParams = @{}

            function _needsUpdateStr { param($cur,$new) return ((("$cur") -ne ("$new")) -and -not [string]::IsNullOrWhiteSpace("$new")) }
            function _needsUpdateBool { param($cur,$new) return ($null -ne $new -and [bool]$cur -ne [bool]$new) }

            if (_needsUpdateStr $u.DisplayName  ($createParams.DisplayName))   { $updateParams['DisplayName']   = $createParams.DisplayName }
            if (_needsUpdateStr $u.GivenName    ($createParams.GivenName))     { $updateParams['GivenName']     = $createParams.GivenName }
            if (_needsUpdateStr $u.Surname      ($createParams.Surname))       { $updateParams['Surname']       = $createParams.Surname }
            if (_needsUpdateStr $u.MailNickname ($createParams.MailNickname))  { $updateParams['MailNickname']  = $createParams.MailNickname }
            if (_needsUpdateStr $u.Mail         ($createParams.Mail))          { $updateParams['Mail']          = $createParams.Mail }
            if ($createParams.ContainsKey('UsageLocation')) {
                $curUL = ($u.UsageLocation | ForEach-Object { "$_" }) -join ''
                if (_needsUpdateStr $curUL $createParams.UsageLocation) { $updateParams['UsageLocation'] = $createParams.UsageLocation }
            }
            if (_needsUpdateStr $u.JobTitle     ($createParams.JobTitle))      { $updateParams['JobTitle']      = $createParams.JobTitle }
            if (_needsUpdateStr $u.Department   ($createParams.Department))    { $updateParams['Department']    = $createParams.Department }
            if (_needsUpdateStr $u.CompanyName  ($createParams.CompanyName))   { $updateParams['CompanyName']   = $createParams.CompanyName }
            if (_needsUpdateStr $u.EmployeeId   ($createParams.EmployeeId))    { $updateParams['EmployeeId']    = $createParams.EmployeeId }
            if (_needsUpdateStr $u.EmployeeType ($createParams.EmployeeType))  { $updateParams['EmployeeType']  = $createParams.EmployeeType }
            if ($createParams.ContainsKey('EmployeeHireDate')) {
                $curHire = $null
                if ($u.AdditionalProperties -and $u.AdditionalProperties.ContainsKey('employeeHireDate')) { $curHire = "$($u.AdditionalProperties['employeeHireDate'])" }
                elseif ($u.PSObject.Properties.Name -contains 'EmployeeHireDate') { $curHire = "$($u.EmployeeHireDate)" }
                $curHireShort = if ($curHire) { (Get-Date $curHire).ToString('yyyy-MM-dd') } else { $null }
                if (_needsUpdateStr $curHireShort $createParams.EmployeeHireDate) { $updateParams['EmployeeHireDate'] = $createParams.EmployeeHireDate }
            }

            if (_needsUpdateStr $u.OfficeLocation ($createParams.OfficeLocation)) { $updateParams['OfficeLocation'] = $createParams.OfficeLocation }
            if (_needsUpdateStr $u.StreetAddress  ($createParams.StreetAddress))  { $updateParams['StreetAddress']  = $createParams.StreetAddress }
            if (_needsUpdateStr $u.City           ($createParams.City))           { $updateParams['City']           = $createParams.City }
            if (_needsUpdateStr $u.State          ($createParams.State))          { $updateParams['State']          = $createParams.State }
            if (_needsUpdateStr $u.PostalCode     ($createParams.PostalCode))     { $updateParams['PostalCode']     = $createParams.PostalCode }
            if (_needsUpdateStr $u.Country        ($createParams.Country))        { $updateParams['Country']        = $createParams.Country }

            # Telephones
            if ($createParams.ContainsKey('BusinessPhones')) {
                $curBiz = if ($u.BusinessPhones) { $u.BusinessPhones[0] } else { $null }
                $newBiz = $createParams.BusinessPhones[0]
                if (_needsUpdateStr $curBiz $newBiz) { $updateParams['BusinessPhones'] = @($newBiz) }
            }
            if (_needsUpdateStr $u.MobilePhone ($createParams.MobilePhone)) { $updateParams['MobilePhone'] = $createParams.MobilePhone }

            # AccountEnabled can be toggled
            if (_needsUpdateBool $u.AccountEnabled $accountEnabledBool) { $updateParams['AccountEnabled'] = $accountEnabledBool }

            # DO NOT update: PasswordProfile for existing users
            # DO NOT update: UserType (immutable Member/Guest)
            # DO NOT update: UserPrincipalName in this script

            if ($updateParams.Count -gt 0) {
                Update-MgUser -UserId $u.Id @updateParams
                $results.Add([pscustomobject]@{
                    userPrincipalName = $upn
                    Action            = 'Updated'
                    userIDString      = $u.Id
                    Comments          = ('Updated fields: ' + (($updateParams.Keys) -join ', '))
                }) | Out-Null
            } else {
                $results.Add([pscustomobject]@{
                    userPrincipalName = $upn
                    Action            = 'NoChange'
                    userIDString      = $u.Id
                    Comments          = 'No differences detected.'
                }) | Out-Null
            }
        }
    }
    catch {
        $results.Add([pscustomobject]@{
            userPrincipalName = (Get-Col $row $MANDATORY_MAP.userPrincipalName)
            Action            = 'Error'
            userIDString      = ''
            Comments          = $_.Exception.Message
        }) | Out-Null
    }

    # ---- Build a log row mirroring the input + action summary ----
    $log = [ordered]@{}
    foreach ($hdr in $InputHeaders) {
        $val = $row.$hdr
        $log[$hdr] = if ($null -eq $val) { '' } else { "$val" }
    }
    $resForUpn = $results | Where-Object { $_.userPrincipalName -eq $upn } | Select-Object -First 1
    $log['userIDString'] = if ($resForUpn) { $resForUpn.userIDString } else { '' }
    $log['Comments']     = if ($resForUpn) { ("{0}: {1}" -f $resForUpn.Action, $resForUpn.Comments) } else { 'No action row found' }
    $logRows.Add([pscustomobject]$log) | Out-Null
}

# -------- Write Output Log (xlsx) in one shot --------
$logRows | Export-Excel -Path $logXlsx -WorksheetName 'Log' -TableName 'Log' -AutoSize -ClearSheet
Write-Host ("Detailed log written to: {0}" -f $logXlsx) -ForegroundColor Cyan

# -------- Final on-screen table --------
$results | Format-Table -AutoSize
