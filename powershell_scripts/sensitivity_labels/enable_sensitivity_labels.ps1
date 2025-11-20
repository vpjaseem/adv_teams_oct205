Install-Module Microsoft.Graph -Scope AllUsers
Install-Module Microsoft.Graph.Beta -Scope AllUsers

Connect-MgGraph -Scopes "Directory.ReadWrite.All"


$grpUnifiedSetting = Get-MgBetaDirectorySetting | Where-Object { $_.Values.Name -eq "EnableMIPLabels" }
$grpUnifiedSetting.Values

$params = @{
     Values = @(
 	    @{
 		    Name = "EnableMIPLabels"
 		    Value = "True"
 	    }
     )
}

Update-MgBetaDirectorySetting -DirectorySettingId $grpUnifiedSetting.Id -BodyParameter $params

$Setting = Get-MgBetaDirectorySetting -DirectorySettingId $grpUnifiedSetting.Id
$Setting.Values
