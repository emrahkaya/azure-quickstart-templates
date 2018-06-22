# Downloads the Visual Studio Team Services Build Agent and installs on the new machine
# and registers with the Visual Studio Team Services account and build agent pool

# Enable option

[CmdletBinding()]
Param
(
	[Parameter(Mandatory=$true)]
	[string]$VSTSAccount,

	[Parameter(Mandatory=$true)]
	[string]$PersonalAccessToken,

	[Parameter(Mandatory=$true)]
	[string]$AgentName,

	[Parameter(Mandatory=$true)]
	[string]$PoolName,

	[Parameter(Mandatory=$true)]
	[int]$AgentCount,

	[Parameter(Mandatory=$true)]
	[string]$AdminUser,

	[Parameter(Mandatory=$true)]
	[switch]$WillUpdateModules,

	[Parameter(Mandatory=$true)]
	[array]$Modules
)
$VerbosePreference = "Continue"

Write-Verbose "Entering InstallVSOAgent.ps1"

$currentLocation = Split-Path -parent $MyInvocation.MyCommand.Definition
Write-Verbose "Current folder: $currentLocation"

#Create a temporary directory where to download from VSTS the agent package (vsts-agent.zip) and then launch the configuration.
$agentTempFolderName = Join-Path $env:temp ([System.IO.Path]::GetRandomFileName())
New-Item -ItemType Directory -Force -Path $agentTempFolderName
Write-Verbose "Temporary Agent download folder: $agentTempFolderName"

$serverUrl = "https://$VSTSAccount.visualstudio.com"
Write-Verbose "Server URL: $serverUrl"

$retryCount = 3
$retries = 1
Write-Verbose "Downloading Agent install files"
[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
do
{
  try
  {
	Write-Verbose "Trying to download for the $retries . time"

    $latestRelease = Invoke-RestMethod -Uri "https://api.github.com/repos/Microsoft/vsts-agent/releases"
	$latestRelease = $latestRelease | Where-Object assets -ne $null | Sort-Object created_at -Descending | Select-Object -First 1
    $latestReleaseDownloadUrl = ($latestRelease.assets | ? { $_.name -like "*win-x64*" }).browser_download_url
	Write-Verbose "Trying to get from $latestReleaseDownloadUrl and saving it as $agentTempFolderName\agent.zip"
    Invoke-WebRequest -Uri $latestReleaseDownloadUrl -Method Get -OutFile "$agentTempFolderName\agent.zip"
    Write-Verbose "Downloaded agent successfully on attempt $retries"
    break
  }
  catch
  {
    $exceptionText = ($_ | Out-String).Trim()
    Write-Verbose "Exception occured downloading agent: $exceptionText in try number $retries"
    $retries++
    Start-Sleep -Seconds 30 
  }
} 
while ($retries -le $retryCount)

for ($i=0; $i -lt $AgentCount; $i++)
{
	$Agent = ($AgentName + "-" + $i)

	# Construct the agent folder under the main (hardcoded) C: drive.
	$agentInstallationPath = Join-Path "C:" $Agent

	# Create the directory for this agent.
	New-Item -ItemType Directory -Force -Path $agentInstallationPath

	# Set the current directory to the agent dedicated one previously created.
	Push-Location -Path $agentInstallationPath
	
	Write-Verbose "Extracting the zip file for the agent:$Agent to $agentInstallationPath"
	$destShellFolder = (new-object -com shell.application).namespace("$agentInstallationPath")
	$destShellFolder.CopyHere((new-object -com shell.application).namespace("$agentTempFolderName\agent.zip").Items(),16)

	# Removing the ZoneIdentifier from files downloaded from the internet so the plugins can be loaded
	# Don't recurse down _work or _diag, those files are not blocked and cause the process to take much longer
	Write-Verbose "Unblocking files"
	Get-ChildItem -Recurse -Path $agentInstallationPath | Unblock-File | out-null

	# Retrieve the path to the config.cmd file.
	$agentConfigPath = [System.IO.Path]::Combine($agentInstallationPath, 'config.cmd')
	Write-Verbose "Agent config.cmd Location = $agentConfigPath"
	if (![System.IO.File]::Exists($agentConfigPath))
	{
		Write-Error "File not found: $agentConfigPath"
		return
	}

	# Call the agent with the configure command and all the options (this creates the settings file) without prompting
	# the user or blocking the cmd execution
	Write-Verbose "Configuring agent '$($Agent)'"		
	.\config.cmd --unattended --url $serverUrl --auth PAT --token $PersonalAccessToken --pool $PoolName --agent $Agent --runasservice
	
	Write-Verbose "Agent install output: $LASTEXITCODE"
	
	Pop-Location
}

if ($WillUpdateModules)
{
	Write-Verbose "Updating modules.."

	# Adding new Path to PSModulePath environment variable
	$CurrentValue = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
	[Environment]::SetEnvironmentVariable("PSModulePath", $CurrentValue + ";C:\Modules", "Machine")
	$NewValue = [Environment]::GetEnvironmentVariable("PSModulePath", "Machine")
	Write-Verbose "new Path is: $($NewValue)"

	# Creating new Path
	if (!(Test-Path -Path C:\Modules -ErrorAction SilentlyContinue))
	{	New-Item -ItemType Directory -Name Modules -Path C:\ }

	# Installing New Modules and Removing Old
	Foreach ($Module in $Modules)
	{	Find-Module -Name $Module.Name -RequiredVersion $Module.Version -Repository PSGallery | Save-Module -Path C:\Modules	}

	$DefaultModules = "PowerShellGet", "PackageManagement","Pester"

	Foreach ($Module in $DefaultModules)
	{
		if ($tmp = Get-Module $Module -ErrorAction SilentlyContinue) {	Remove-Module $Module -Force	}
		Find-Module -Name $Module -Repository PSGallery | Install-Module -Force -Confirm:$false -SkipPublisherCheck
	}

	# Uninstalling old Azure PowerShell Modules
	$programName = "Microsoft Azure PowerShell"
	$app = Get-WmiObject -Class Win32_Product -Filter "Name Like '$($programName)%'"
	$app.Uninstall()

	Write-Verbose "Modules' update complete."
}

Write-Verbose "Exiting InstallVSTSAgent.ps1"
