cls
function Get-ScriptDirectory {
	$Invocation = (Get-Variable MyInvocation -Scope 1).Value
	Split-Path $Invocation.MyCommand.Path
}
Set-Location (Get-ScriptDirectory)
$solutionFolder = Resolve-Path "."
$msbuild = "C:\Windows\Microsoft.NET\Framework64\v4.0.30319\MSBuild.exe"
$rootBuildFolder = Resolve-Path ".\Packages"
$outDir = "$rootBuildFolder"
mkdir $outDir -Force | Out-Null
$projects = @{
				"$solutionFolder\Lapointe.SharePoint2010.PowerShell\Lapointe.SharePoint2010.PowerShell.csproj" = @("ReleaseMOSS", "ReleaseFoundation")
				"$solutionFolder\Lapointe.SharePoint2013.PowerShell\Lapointe.SharePoint2013.PowerShell.csproj" = @("ReleaseMOSS", "ReleaseFoundation")
			}
foreach ($project in $projects.Keys) {
	Write-Host "Building $project..." -ForegroundColor Blue
	foreach ($config in $projects[$project]) {
		$version = "SP2010"
		if ($project.Contains("Lapointe.SharePoint2013.PowerShell.csproj")) { $version = "SP2013" }
		Write-Host "Building $config..." -ForegroundColor Blue
		del "$outDir\$version\$config\*.wsp" -Force -ErrorAction SilentlyContinue
		&$msbuild $project /v:m /t:Rebuild /t:Package /p:Configuration="$config" /p:OutDir="$outDir\$version\$config"
		del "$outDir\$version\$config\*.dll" -Force -ErrorAction SilentlyContinue
		del "$outDir\$version\$config\*.pdb" -Force -ErrorAction SilentlyContinue
        del "$outDir\$version\$config\*.config" -Force -ErrorAction SilentlyContinue
	}
}
