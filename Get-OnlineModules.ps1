<#
.SYNOPSIS
    Saves PowerShell Modules for import to diconnect networks
.DESCRIPTION
    Run this script on a internet connected system
    this script will download latest nuget assembly with packagemanagement modules
    plus any additional module found. Required for disconnected system
.NOTES
    https://docs.microsoft.com/en-us/powershell/gallery/psget/repository/bootstrapping_nuget_proivder_and_exe
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,HelpMessage='Specify modules to download. for multiple, separate by commas')]
	[string[]]$OnlineModules,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Remove older modules if found')]
	[switch]$RemoveOld = $true,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Install modules on online system as well')]
	[switch]$Install = $true,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Re-Download modules if exist')]
	[switch]$Refresh = $false
)
##*===============================================
##* VARIABLE DECLARATION
##*===============================================
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName

#Get required folder and File paths
[string]$DownloadedModulesPath = Join-Path -Path $scriptRoot -ChildPath 'DownloadedModules'

If(!$OnlineModules){
    $OnlineModules = "PowerShellGet","VMware.PowerCLI","PowervRA","PowervRO"
}
##*===============================================
#See if system is conencted to the internet
$internetConnected = Test-NetConnection www.powershellgallery.com -CommonTCPPort HTTP -InformationLevel Quiet -WarningAction SilentlyContinue

If($internetConnected)
{
    $Nuget = Install-PackageProvider Nuget –force –verbose
    $NuGetAssemblyVersion = $($Nuget).version
    Write-Host "INSTALLED: Nuget [$NuGetAssemblyVersion] was installed" -ForegroundColor Green
    #Copy Nuget prereqs
    $NuGetAssemblyDestPath = Get-ChildItem "$DownloadedModulesPath\\nuget" -Filter *.dll -Recurse
    $NuGetAssemblySourcePath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget"
    If ($NuGetAssemblyDestPath.FullName)
    {
        If($Refresh){
            Write-Host "BACKUP: Copying nuget Assembly [$NuGetAssemblyVersion] from $NuGetAssemblySourcePath" -ForegroundColor Gray
            Copy-Item "$NuGetAssemblySourcePath\$NuGetAssemblyVersion\Microsoft.PackageManagement.NuGetProvider.dll" $NuGetAssemblyDestPath.FullName -Force -ErrorAction SilentlyContinue
        }
        Else{
            Write-Host "FOUND: Nuget [$NuGetAssemblyVersion] already copied" -ForegroundColor Green
        }
    }
    Else{
        Write-Host "BACKUP: Copying nuget Assembly [$NuGetAssemblyVersion] from $NuGetAssemblySourcePath" -ForegroundColor Gray
        New-Item "$DownloadedModulesPath\nuget\$NuGetAssemblyVersion" -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        Copy-Item "$NuGetAssemblySourcePath\$NuGetAssemblyVersion\Microsoft.PackageManagement.NuGetProvider.dll" "$DownloadedModulesPath\nuget\$NuGetAssemblyVersion" -ErrorAction SilentlyContinue
    }

    #loop through each module
    Foreach ($Module in $OnlineModules){
        
        #get the module if found online
        $ModuleFound = Find-Module $Module
        If($ModuleFound)
        {
            [string]$ModuleVersion = $ModuleFound.Version
            [string]$ModuleName = $ModuleFound.Name
            
            #If specified, remove older modules in DownloadedModule directory if found
            If($RemoveOld)
            {
                $LikeModulesExist = Get-ChildItem $DownloadedModulesPath -Directory | Where-Object {$_.FullName -match "$ModuleName" -and $_.FullName -notmatch "$ModuleName-$ModuleVersion"} | foreach {
                        $_ | Remove-Item -Force -Recurse
                        Write-host "REMOVED: $($_.FullName)" -ForegroundColor DarkYellow
                    }
            }


            #Check to see it module is already downloaded
            If(Test-Path "$DownloadedModulesPath\$ModuleName-$ModuleVersion")
            {
                #If specified, Re-Download modules 
                If($Refresh)
                {
                    Write-Host "BACKUP: $ModuleName [$ModuleVersion] found but will be re-downloaded..." -ForegroundColor Gray
                    Save-Module -Name $ModuleName -Path $DownloadedModulesPath\$ModuleName-$ModuleVersion -Force
                }
                Else{
                    Write-Host "FOUND: $ModuleName [$ModuleVersion] already downloaded" -ForegroundColor Green
                }
            }
            Else{
                Write-Host "BACKUP: $ModuleName [$ModuleVersion] not found, downloading for offline install" -ForegroundColor Gray
                New-Item "$DownloadedModulesPath\$ModuleName-$ModuleVersion" -ItemType Directory | Out-Null
                Save-Module -Name $ModuleName -Path $DownloadedModulesPath\$ModuleName-$ModuleVersion
            }


            #If specified, Install modules on local system as well 
            If($Install)
            {
                Write-Host "INSTALL: $Module [$ModuleVersion] will be installed locally as well, please wait..." -ForegroundColor Yellow
                Install-Module $Module -AllowClobber -SkipPublisherCheck -Force
            }
        }
        Else{
            Write-Host "WARNING: $Module was not found online" -ForegroundColor Yellow
        }

    } #End Loop

}
Else{
    Write-Host "ERROR: Unable to connect to the internet to grab modules" -ForegroundColor Red
    throw $_.error
}