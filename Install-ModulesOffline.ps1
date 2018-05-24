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
	[Parameter(Mandatory=$false,Position=0,HelpMessage='Specify modules to install. for multiple, separate by commas')]
	[string[]]$InstallModules,
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Decide what to do it older modules are found')]
    [ValidateSet("Remove","Uninstall","Delete","Ignore")]
	[string]$FoundAction = 'Delete',
    [Parameter(Mandatory=$false,Position=1,HelpMessage='Where to load the modules. CurrentLocation = Default: Load from script modules directory; 
                                                                                   UserModulePath = Copy module to user PSModulePath;
                                                                                   SystemModulePath = Copy module to Program Files Directory')]
	[ValidateSet("CurrentLocation","UserModulePath","SystemModulePath")]
    [string] $LoadPath = 'CurrentLocation'
)

##*=============================================
##* FUNCTIONS
##*=============================================
Function Parse-Psd1{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$false)]
        [Microsoft.PowerShell.DesiredStateConfiguration.ArgumentToConfigurationDataTransformationAttribute()]
        [hashtable] $data
    ) 
    return $data
}

function ReportStartOfActivity($activity) {
   $script:currentActivity = $activity
   Write-Progress -Activity $loadingActivity -CurrentOperation $script:currentActivity -PercentComplete $script:percentComplete
}

function ReportFinishedActivity() {
   $script:completedActivities++
   $script:percentComplete = (100.0 / $totalActivities) * $script:completedActivities
   $script:percentComplete = [Math]::Min(99, $percentComplete)
   
   Write-Progress -Activity $loadingActivity -CurrentOperation $script:currentActivity -PercentComplete $script:percentComplete
}

# Load modules
function Load-DependencyModules(){
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true)]
        $ModuleName
    )
    ReportStartOfActivity "Searching for $productShortName module components..."

   

   $loaded = Get-Module -Name $ModuleName -ErrorAction Ignore | % {$_.Name}
   $registered = Get-Module -Name $ModuleName -ListAvailable -ErrorAction Ignore | % {$_.Name}
   $notLoaded = $null
   $notLoaded = $registered | ? {$loaded -notcontains $_}
   
   ReportFinishedActivity
   
   foreach ($module in $registered) {
      if ($loaded -notcontains $module) {
		 ReportStartOfActivity "Loading module $module"
         
		 #Import-Module $module
		 Write-Host $module
		 ReportFinishedActivity
      }
   }
}

Function Collect-Modules{
    [CmdletBinding()]
    Param (
        [Parameter(Mandatory=$true,HelpMessage='Please enter the path to the psd1 manifest file')]
		#[ValidateScript({('.psd1' -contains [IO.Path]::GetExtension($_))})]
		[Alias('FilePath')]
		[string]$ModulePath,
        [Parameter(Mandatory=$false)]
        [string]$SearchPath,
        [Parameter(Mandatory=$false)]
        [string]$NotVersion,
        [Parameter(Mandatory=$false)]
        [switch]$AppendObject,
        [Parameter(Mandatory=$false)]
        [boolean]$ReturnObject = $true,
        [Parameter(Mandatory=$false)]
        [switch]$GetDependencies = $true,
        [Parameter(Mandatory=$true)]
        [string]$WhereID

    )
    Begin{
        If($SearchPath){
            If($NotVersion){
                $FoundManifest = Get-ChildItem $ModulePath -Filter *.psd1 -Depth 1 | Where-Object {$_.FullName -notmatch "$NotVersion"}
            }
            Else{
                $FoundManifest = Get-ChildItem $ModulePath -Filter *.psd1 -Depth 1
            }
        }
        Else{
            If($NotVersion){
                $FoundManifest = Get-ChildItem $ModulePath -Filter *.psd1 -Depth 1 | Where-Object {$_.FullName -notmatch "$NotVersion"}
            }
            Else{
                $FoundManifest = Get-ChildItem $ModulePath -Filter *.psd1 -Depth 1
            }
        }
    }
    Process{
        $ModuleObject = @()
        If($FoundManifest){
            $FoundManifest | ForEach{
                $Manifest = Test-ModuleManifest $_.FullName -ErrorAction Ignore
                $Name = Split-Path ("$ModulePath\" + $_.Name) -Leaf
                $FullPath = "$ModulePath\" + $_.BaseName + "\"+ $Manifest.Version

                If($GetDependencies){
                    #clear list, build collection object
                    $DependencyList = @()
                    $DependencyList = New-Object System.Collections.Generic.List[System.Object]
                    $Manifest.RequiredModules | ForEach{
                        #find modules psd1 file
                        $RequiredModulesDir = Get-ChildItem ("$UserModulePath\$Module\" + $_.Name + "\" + $_.Version) -Filter *.psd1 -Depth 1
                        If($RequiredModulesDir){
                            $DependencyList.Add($RequiredModulesDir.DirectoryName)
                        }
                    
                    }
                }
                Else{$Dependencies = 'Excluded'}
                
                $ModuleObject += new-object psobject -property @{
                    Name=$_.BaseName
                    Version=$Manifest.Version
                    FullPath=$FullPath
                    DependencyPath=$DependencyList
                    Where=$WhereID
                }

            }
        }
        Else{
            Write-Host "No Modules found or invalid path [$ModulePath]" -ForegroundColor Yellow
        }
    }
    End{
        If($FoundManifest){
            Write-Host "Global Object is named [Global:ModuleCollection]"
            If($AppendObject){$Global:ModuleCollection += $ModuleObject}
            Else{$Global:ModuleCollection = $ModuleObject}
        
            If($ReturnObject -eq $true){return $Global:ModuleCollection}
            Else{return $true}
        }
        Else{return $false}
    }
}
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
[string]$ModulesPath = Join-Path -Path $scriptRoot -ChildPath 'Modules'
[string]$BinPath = Join-Path -Path $scriptRoot -ChildPath 'Bin'

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