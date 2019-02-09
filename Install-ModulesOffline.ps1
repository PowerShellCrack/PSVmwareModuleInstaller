<#
.SYNOPSIS
    Installs PowerShell Modules on diconnect networks
.DESCRIPTION
    this script will install latest nuget assembly with packagemanagement modules
    plus any additional module found. Required for disconnected system
.NOTES
    https://docs.microsoft.com/en-us/powershell/gallery/psget/repository/bootstrapping_nuget_proivder_and_exe
#>
[CmdletBinding()]
Param (
	[Parameter(Mandatory=$false,Position=0,HelpMessage='Specify modules to install. for multiple, separate by commas, for all type All or do not specify')]
	[string[]]$InstallModules,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Where to load the modules. CurrentLocation = Default: Load from script modules directory; 
                                                                                   UserModulePath = Copy module to user PSModulePath;
                                                                                   SystemModulePath = Copy module to Program Files Directory')]
	[ValidateSet("CurrentLocation","UserModulePath","SystemModulePath")]
    [string] $SkopePath = 'CurrentLocation'
)


##*=============================================
##* FUNCTIONS
##*=============================================


##*===============================================
##* VARIABLE DECLARATION
##*===============================================
## Variables: Script Name and Script Paths
[string]$scriptPath = $MyInvocation.MyCommand.Definition
[string]$scriptName = [IO.Path]::GetFileNameWithoutExtension($scriptPath)
[string]$scriptFileName = Split-Path -Path $scriptPath -Leaf
[string]$scriptRoot = Split-Path -Path $scriptPath -Parent
[string]$invokingScript = (Get-Variable -Name 'MyInvocation').Value.ScriptName
#  Get the invoking script directory
If ($invokingScript) {
	#  If this script was invoked by another script
	[string]$scriptParentPath = Split-Path -Path $invokingScript -Parent
}
Else {
	#  If this script was not invoked by another script, fall back to the directory one level above this script
	[string]$scriptParentPath = (Get-Item -LiteralPath $scriptRoot).Parent.FullName
}

#Get relative folder and File paths
[string]$ModulesPath = Join-Path -Path $scriptRoot -ChildPath 'Modules'
[string]$BinPath = Join-Path -Path $scriptRoot -ChildPath 'Bin'
[string]$ScriptsPath = Join-Path -Path $scriptRoot -ChildPath 'Scripts'

#Get all paths to PowerShell Modules
$UserModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$home*"}
$AllUsersModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:ProgramFiles\WindowsPowerShell*"}
$SystemModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:windir*"}

#find profile module
$PowerShellNoISEProfile = $profile -replace "ISE",""


#Install Nuget prereq
$NuGetAssemblySourcePath = Get-ChildItem "$BinPath\nuget" -Recurse -Filter *.dll
If($NuGetAssemblySourcePath){
    $NuGetAssemblyVersion = [System.Diagnostics.FileVersionInfo]::GetVersionInfo($NuGetAssemblySourcePath.FullName).FileVersion
    $NuGetAssemblyDestPath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\nuget\$NuGetAssemblyVersion"
    If (!(Test-Path $NuGetAssemblyDestPath)){
        Write-Host "Copying nuget Assembly ($NuGetAssemblyVersion) to $NuGetAssemblyDestPath" -ForegroundColor Cyan
        New-Item $NuGetAssemblyDestPath -ItemType Directory -ErrorAction SilentlyContinue
        Copy-Item -Path $NuGetAssemblySourcePath.FullName -Destination $NuGetAssemblyDestPath –Recurse -ErrorAction SilentlyContinue
    }
}


# Get Modules in modules folder and whats installed
If ($InstallModules){
    $query = $InstallModules
    $AlreadyInstalledModules = Get-Module -Name $InstallModules -ListAvailable
}
Else{
    $query = '(\d+\.)(\d+\.)(\d+\.)(\d)' #Regex filter for 4 digit version number (eg: 1.0.0.1)
    $AlreadyInstalledModules = Get-Module -ListAvailable
}
$Modules = Get-ChildItem -Path $ModulesPath -Recurse | Where-Object { $_.Name -match $query} | % {Get-ChildItem -Path $_.FullName -Filter *.psd1}



# Collect Modules with dependecies
$ModuleObject = @()
foreach($module in $modules)
{
    "Collecting manifest information for module: {0}" -f $module.BaseName | Write-Host -ForegroundColor Cyan
    $Manifest = Test-ModuleManifest $module.FullName -ErrorAction SilentlyContinue -WarningAction SilentlyContinue

    #clear list, build collection object
    $Dependencies = @()
    $Dependencies = New-Object System.Collections.Generic.List[System.Object]

    #parse all required modules from manifest
    If($Manifest.RequiredModules)
    {
        ForEach($RequiredModule in $Manifest.RequiredModules)
        {
            If($Manifest.RequiredModules.Count -eq 1){"Found {1} dependency for module: {0}" -f $Module.BaseName, $Manifest.RequiredModules.Count | Write-Host -ForegroundColor Gray}
            If($Manifest.RequiredModules.Count -gt 1){"Found {1} dependencies for module: {0}" -f $Module.BaseName, $Manifest.RequiredModules.Count | Write-Host -ForegroundColor Gray}

            "Collecting module dependency details on: {0}" -f $RequiredModule.Name | Write-Host -ForegroundColor DarkCyan
            $RequiredModuleDir = Get-ChildItem -Path $ModulesPath -Filter *.psd1 -Recurse | Where {$_.BaseName -eq $RequiredModule.Name} | Select -First 1
            
            #if Module is found
            If($RequiredModuleDir){
                [string]$RequiredModuleVersion = (Test-ModuleManifest $RequiredModuleDir -ErrorAction SilentlyContinue).Version
                $RootPath = Split-Path $RequiredModuleDir.Directory -Resolve

                #If there is a dependencies add its path to an collection
                If($RequiredModuleDir){
                    $Dependencies.Add($RequiredModuleDir.FullName)
                    #$Dependencies.Add($RootPath)
                }
            }
            Else{
                "Missing [{1}] dependency for module: {0}" -f $Module.BaseName, $Manifest.RequiredModules | Write-Host -ForegroundColor Gray
            }
            
        }
        [int32]$DependenciesCount = $Dependencies.Count
    }
    Else{
        "No dependencies found for module: {0}" -f $Module.BaseName | Write-Host -ForegroundColor Gray
        [int32]$DependenciesCount = 0
    }

    # Build an object that consists of module dependecy information 
    $ModuleObject += new-object psobject -property @{
        ModuleName=$module.BaseName
        ModuleVersion=$Manifest.Version.ToString()
        ModuleFolder=$module.DirectoryName
        ModulePath=$module.FullName
        ModuleSkope=$SkopePath
        DependencyPath=$Dependencies
        DependencyCount=$DependenciesCount
    }


    #get the max amount of dependecies to determine install loop 
    #------------------------------------------------------------
    $MaxDependencies = $ModuleObject.DependencyCount | Measure-Object -Maximum
}













#build User Powershell profile directory
#https://blogs.technet.microsoft.com/heyscriptingguy/2012/05/21/understanding-the-six-powershell-profiles/
If(!(Test-Path $UserModulePath)){
    If($profile -match "ISE"){
        New-Item -Path $PowerShellNoISEProfile -ItemType File -ErrorAction SilentlyContinue | Out-null
    }
    Else{
        New-Item -Path $profile -ItemType File -ErrorAction SilentlyContinue | Out-null
    }
    New-Item -Path $UserModulePath -ItemType Directory -ErrorAction SilentlyContinue | Out-null
    New-Item -Path $UserScriptsPath -ItemType Directory -ErrorAction SilentlyContinue | Out-null
}

#TESTS - Remove when done
$InstallModules = 'VMware.PowerCLI'
$Module = 'VMware.PowerCLI'
$UserModulePath = "C:\Users\tracyr.ctr\Documents\WindowsPowerShell\Modules"

(Test-ModuleManifest  $ModulesInstalledWrong.FullName[0] -ErrorAction Ignore).RequiredModules
#TESTS - Remove when done

Foreach($Module in $InstallModules){
    #Naming Schema for a downloade module is: modulepath\modulename\modulename
    $DownloadedModule = Get-ChildItem "$ModulesPath" -Directory | Where-Object {$_.FullName -match "$Module"}
    If($DownloadedModule){
        #Get Downloaded Module Version. Get-OnlineModules name formated is modulename-moduleversion
        $DownloadedModuleVersion = ($DownloadedModule.Name).split("-")[1]

        $Global:AllModuleObject = @()       
        #$FindAllUserModules = Get-ChildItem "$UserModulePath" -Filter *.psd1 -Recurse
        #$ModulesInstalledWrong = $FindAllUserModules | Where-Object { $_.DirectoryName -match "$Module" }

        #Find user modules that were installed wrong; collect them and their dependencies for processing later
        Collect-Modules -ModulePath "$UserModulePath\$Module" -WhereID InvalidUserDirectory

        #Find user modules installed; collect them for processing later
        Collect-Modules -ModulePath "$UserModulePath\$Module\$Module" -WhereID UserDirectory -NotVersion $DownloadedModuleVersion -AppendObject

        #Check AllUsers path
        Collect-Modules -ModulePath $AllUsersModulePath -WhereID AllUserDirectory -NotVersion $DownloadedModuleVersion -AppendObject

        #Check System path
        Collect-Modules -ModulePath $SystemModulePath -WhereID SystemDirectory -NotVersion $DownloadedModuleVersion -AppendObject
        
        If( ($FoundAction -eq "Ignore") -and ($FoundUserManifest) ){
            Write-Host "WARNING: Found older modules. Triggered not to remove modules; may cause conflicts" -ForegroundColor DarkMagenta
        }
        ElseIf($FoundUserManifest){
            Foreach ($UserManifest in $FoundUserManifest){
                #parse module for psd1 file to get versions that are installed (incase the folder is misleading)
                $ParsedUserModule = Test-ModuleManifest $UserManifest.fullname -ErrorAction Ignore
                Write-Host "Found existing module $Module [$($ParsedUserModule.Version)] in directory [$UserModulePath]" -ForegroundColor DarkYellow
            
                #build collection for module dependecies to properly identify the correct ones if multiple version exist
                $ModuleCollection = @()
                $ParsedUserModule.RequiredModules | ForEach-Object {
                    $FullPath = "$UserModulePath\$Module\" + $_.Name + "\"+ $_.Version
                    
                    $Dependency = Test-ModuleManifest $_.RequiredModules -ErrorAction Ignore
                    $ModuleTable = New-Object -TypeName PSObject -Property ([ordered]@{
                            Name    = $_.Name
                            Version = $_.Version
                            FullPath= $Path
                            Dependency = $Dependency
                        })

                        $ModuleCollection += $ModuleTable
                    }

                #deal with main module
                If( ($FoundAction -eq "Uninstall") ){
                    #remove main module just in case
                    Get-InstalledModule -Name $Module -RequiredVersion $ParsedUserModule.Version -ErrorAction SilentlyContinue | Uninstall-Module -Force -ErrorAction SilentlyContinue
                }
                
                If( ($FoundAction -eq "Remove") -or ($FoundAction -eq "Delete") ){
                    #remove main module just in case
                    Write-Host "Removing $Module [$($ParsedUserModule.Version)] in directory [$UserModulePath]" -ForegroundColor DarkYellow
                    Get-InstalledModule -Name $Module -RequiredVersion $ParsedUserModule.Version -ErrorAction SilentlyContinue | Uninstall-Module -Force -ErrorAction SilentlyContinue | Remove-Module -Force
                }

                If($FoundAction -eq "Delete"){ 
                    
                }


                #Deal with dependency modules based on manifest
                Foreach($dependentModule in $ModuleCollection){
                    #delete all modules
                    #remove individual modules depenencies.  
                        
                }
                
            }
        }
        Else{
            Write-Host "INFO: No [$Module] module found." -ForegroundColor Gray
        }


    }
    Else{
        Write-Host "WARNING: $Module was not found in [$UserModulePath]" -ForegroundColor Yellow
    }
    

    
   # Copy-WithProgress -Source $PowerCLINetPath -Destination $UserModulePath -ExcludeType Directories -ProgressDisplayName 'Copying Vmware PowerCLI Modules Files...'
}

#userinstall
Get-Module -ListAvailable | Where-Object ModuleBase -like "C:\Users\tracyr.ctr\Documents\WindowsPowerShell\Modules*" |
Sort-Object -Property Name, Version -Descending |
Get-Unique -PipelineVariable Module |
ForEach-Object {
    if (-not(Test-Path -Path "$($_.ModuleBase)\PSGetModuleInfo.xml")) {
        Find-Module -Name $_.Name -OutVariable Repo -ErrorAction SilentlyContinue |
        Compare-Object -ReferenceObject $_ -Property Name, Version |
        Where-Object SideIndicator -eq '=>' |
        Select-Object -Property Name,
                                Version,
                                @{label='Repository';expression={$Repo.Repository}},
                                @{label='InstalledVersion';expression={$Module.Version}}
    }
    
} #| ForEach-Object {Install-Module -Name $_.Name -Force}