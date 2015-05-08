param(
  [string] $TargetDir = $(throw "TargetDir is required!")
)
$path = Split-Path -parent $MyInvocation.MyCommand.Definition  
$helpAsm = "$($TargetDir)\Lapointe.PowerShell.MamlGenerator.dll"
$cmdletAsm = "$($TargetDir)\Lapointe.SharePoint.PowerShell.dll"
Write-Host "Help generation work path: $path"
Write-Host "Help generation maml assembly path: $helpAsm"
Write-Host "Help generation cmdlet assembly path: $cmdletAsm"

Start-Process "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\gacutil.exe" -ArgumentList "/uf","Lapointe.PowerShell.MamlGenerator"
Start-Process "C:\Program Files (x86)\Microsoft SDKs\Windows\v7.0A\bin\gacutil.exe" -ArgumentList "/uf","Lapointe.SharePoint.PowerShell"

Write-Host "Loading help assembly..."
[System.Reflection.Assembly]::LoadFrom($helpAsm)
Write-Host "Loading cmdlet assembly..."
$asm = [System.Reflection.Assembly]::LoadFrom($cmdletAsm)
$asm
Write-Host "Generating help..."
[Lapointe.PowerShell.MamlGenerator.CmdletHelpGenerator]::GenerateHelp($asm, "$path\POWERSHELL\Help", $true)

