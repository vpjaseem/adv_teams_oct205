# ===== Settings =====
$ExcelFolder = "E:\AJ Labs 2\Advanced Microsoft Teams Training\Training Docs\PowerShell\Administration via PowerShell"  # Update your file location here
$ExcelFile   = "AJ_EntraID_License_Export_All.xlsx"
$XlsxPath    = [System.IO.Path]::Combine($ExcelFolder, $ExcelFile)

# ===== Prereqs =====
# Install-Module Microsoft.Graph -Scope CurrentUser -Force
# Install-Module ImportExcel   -Scope CurrentUser -Force
Import-Module ImportExcel
Connect-MgGraph -Scopes "User.Read.All","Directory.Read.All" | Out-Null

# Ensure target folder exists (absolute path)
[void][System.IO.Directory]::CreateDirectory($ExcelFolder)

# ===== Build fast lookup for SkuId -> SkuPartNumber =====
# Normalize GUID keys to lowercase strings to avoid mismatches
$skuIndex = @{}
Get-MgSubscribedSku | ForEach-Object {
    $key = ($_.SkuId).ToString().ToLower()
    $skuIndex[$key] = $_.SkuPartNumber
}

# ===== Get users & flatten AssignedLicenses =====
$users = Get-MgUser -All -Property Id,DisplayName,UserPrincipalName,AssignedLicenses

$result = foreach ($u in $users) {
    $assigned = $u.AssignedLicenses.SkuId
    if ($assigned) {
        foreach ($skuId in $assigned) {
            $skuKey = ($skuId).ToString().ToLower()
            [PSCustomObject]@{
                UserPrincipalName = $u.UserPrincipalName
                DisplayName       = $u.DisplayName
                SkuId             = $skuId
                SkuPartNumber     = $skuIndex[$skuKey]  # now resolves correctly
            }
        }
    } else {
        [PSCustomObject]@{
            UserPrincipalName = $u.UserPrincipalName
            DisplayName       = $u.DisplayName
            SkuId             = $null
            SkuPartNumber     = $null
        }
    }
}

# ===== Export to Excel (absolute path) =====
$result |
  Sort-Object UserPrincipalName, SkuPartNumber |
  Export-Excel -Path $XlsxPath -WorksheetName 'UserLicenses' `
               -AutoSize -TableName 'UserLicenses' -FreezeTopRow -BoldTopRow -ClearSheet

# Optional per-user summary sheet (SKUs comma-separated). Delete this block if not needed.
$perUser = $result | Group-Object UserPrincipalName | ForEach-Object {
    [PSCustomObject]@{
        UserPrincipalName = $_.Name
        DisplayName       = ($_.Group | Select-Object -First 1 -ExpandProperty DisplayName)
        SKUs              = ($_.Group | Where-Object {$_.SkuPartNumber} |
                              Select-Object -ExpandProperty SkuPartNumber -Unique) -join ','
    }
}
$perUser |
  Sort-Object UserPrincipalName |
  Export-Excel -Path $XlsxPath -WorksheetName 'PerUserSummary' `
               -AutoSize -TableName 'PerUserSummary' -FreezeTopRow -BoldTopRow

Write-Host "Exported: $XlsxPath"
