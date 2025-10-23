#Adding IIS and PHP Modules

##Install IIS
Install-WindowsFeature web-server

# Installing optional applications

## Download, and Install Google Chrome
$Path = $env:TEMP; $Installer = "chrome_installer.exe"; Invoke-WebRequest "http://dl.google.com/chrome/install/375.126/chrome_installer.exe" -OutFile $Path\$Installer; Start-Process -FilePath $Path\$Installer -Args "/silent /install" -Verb RunAs -Wait; Remove-Item $Path\$Installer

## Download, and Install Notepad++
$Path = $env:TEMP; $Installer = "notepad_pp_installer.exe"; Invoke-WebRequest "https://github.com/notepad-plus-plus/notepad-plus-plus/releases/download/v8.8.7/npp.8.8.7.Installer.x64.exe" -OutFile $Path\$Installer; Start-Process -FilePath $Path\$Installer /S -NoNewWindow -Wait -PassThru; Remove-Item $Path\$Installer

## Download, and Install WinRAR
$Path = $env:TEMP; $Installer = "winRAR.exe"; Invoke-WebRequest "https://www.win-rar.com/fileadmin/winrar-versions/winrar/winrar-x64-620.exe" -OutFile $Path\$Installer; Start-Process -FilePath $Path\$Installer /S -NoNewWindow -Wait -PassThru; Remove-Item $Path\$Installer

## Download, and Install WinSCP
$Path = $env:TEMP; $Installer = "winSCP.exe"; Invoke-WebRequest "https://raw.githubusercontent.com/vpjaseem/Cloud-Computing/refs/heads/main/Azure/Downloads/WinSCP-6.5.4-Setup.exe" -OutFile $Path\$Installer
cd $env:TEMP 
.\winSCP.exe /VERYSILENT /NORESTART /ALLUSERS
#Remove-Item $Path\$Installer -Force

## Download, and Install PuTTY
$Path = $env:TEMP; $Installer = "PuTTY.msi"; Invoke-WebRequest "https://the.earth.li/~sgtatham/putty/latest/w64/putty-64bit-0.83-installer.msi" -OutFile $Path\$Installer; Start-Process -FilePath $Path\$Installer /qn

## Download, and Install Mozilla
$Path = $env:TEMP; $Installer = "Mozilla.exe"; Invoke-WebRequest "https://cdn.stubdownloader.services.mozilla.com/builds/firefox-stub/en-US/win/cdbfd26f68e0ce385a351852f04c300602c1fe8acf27e69c47eca4e79c447e2b/Firefox%20Installer.exe" -OutFile $Path\$Installer; Start-Process -FilePath $Path\$Installer -Args "/s" -Verb RunAs -Wait; 
