#!pwsh

#!ps
#timeout=300000
#maxlength=100000


$fileExtension = ".ps1"
$className = "Microsoft.PowerShellScript.1"
$desiredOpeningCommand = "pwsh.exe -ExecutionPolicy Bypass -File `"`%1`" %~2"


$registryPathOfUserKey  = Join-Path "HKEY_USERS" $((Get-Item -Path "registry::HKLM\Software\Microsoft\Windows\CurrentVersion\Authentication\LogonUI").GetValue("LastLoggedOnUserSID"))
$registryPathOfFileExtsKeyForExtension = Join-Path $registryPathOfUserKey "Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\$fileExtension"
$registryPathOfClassesKeyForExtension = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$fileExtension"
$registryPathOfClassesKeyForClass = "HKEY_LOCAL_MACHINE\SOFTWARE\Classes\$className"


Get-Item "registry::$registryPathOfFileExtsKeyForExtension"
Get-Item "registry::$registryPathOfClassesKeyForExtension"

$pathOfTempFile = join-path $env:TEMP ((New-Guid).Guid)
reg export $registryPathOfFileExtsKeyForExtension $pathOfTempFile
Write-Output "FileExtsKeyForExtension: "
Get-Content -Raw -Path $pathOfTempFile 

$pathOfTempFile = join-path $env:TEMP ((New-Guid).Guid)
reg export $registryPathOfClassesKeyForExtension $pathOfTempFile
Write-Output "ClassesKeyForExtension: "
Get-Content -Raw -Path $pathOfTempFile 

$pathOfTempFile = join-path $env:TEMP ((New-Guid).Guid)
reg export $registryPathOfClassesKeyForClass $pathOfTempFile
Write-Output "ClassesKeyForClass: "
Get-Content -Raw -Path $pathOfTempFile 

$registryPathOfCommandKeyForOpeningTheClass = Join-Path $registryPathOfClassesKeyForClass "Shell\Open\Command"



Set-ItemProperty -Path "registry::$registryPathOfCommandKeyForOpeningTheClass" -Name '(Default)'  -Type "String"  -Value $desiredOpeningCommand

Set-ExecutionPolicy Bypass -force

$pathOfTempFile = join-path $env:TEMP ((New-Guid).Guid)
reg export $registryPathOfClassesKeyForClass $pathOfTempFile
Write-Output "ClassesKeyForClass: "
Get-Content -Raw -Path $pathOfTempFile 


& {
    #ensure that chocoloatey is installed
    Set-ExecutionPolicy Bypass -Scope Process -Force; [System.Net.ServicePointManager]::SecurityProtocol = [System.Net.ServicePointManager]::SecurityProtocol -bor 3072; iex ((New-Object System.Net.WebClient).DownloadString('https://community.chocolatey.org/install.ps1'))
    choco upgrade --acceptlicense --confirm chocolatey --timeout 240

    #ensure that autohotkey is installed
    choco install --acceptlicense -y autohotkey --timeout 240
    choco upgrade --force --acceptlicense -y autohotkey --timeout 240

    #ensure that pwsh is installed
    choco install --acceptlicense -y pwsh --timeout 240
    choco upgrade --acceptlicense -y pwsh --timeout 240
}