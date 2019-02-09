#######################################################################################################################
# PowerCLI installer - Current user
#######################################################################################################################
$LatestModuleVersion = '10.1.0.8403314'

##*=============================================
##* FUNCTIONS
##*=============================================
function Copy-WithProgress {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $true)]
        [string] $Source,
        [Parameter(Mandatory = $true)]
        [string] $Destination,
        [int] $Gap = 0,
        [int] $ReportGap = 200,
        [ValidateSet("Directories","Files")]
        [string] $ExcludeType,
        [string] $Exclude,
        [string] $ProgressDisplayName
    )
    # Define regular expression that will gather number of bytes copied
    $RegexBytes = '(?<=\s+)\d+(?=\s+)';

    #region Robocopy params
    # MIR = Mirror mode
    # NP  = Don't show progress percentage in log
    # NC  = Don't log file classes (existing, new file, etc.)
    # BYTES = Show file sizes in bytes
    # NJH = Do not display robocopy job header (JH)
    # NJS = Do not display robocopy job summary (JS)
    # TEE = Display log in stdout AND in target log file
    # XF file [file]... :: eXclude Files matching given names/paths/wildcards.
    # XD dirs [dirs]... :: eXclude Directories matching given names/paths.
    $CommonRobocopyParams = '/MIR /NP /NDL /NC /BYTES /NJH /NJS';
    
    switch ($ExcludeType){
        Files { $CommonRobocopyParams += ' /XF {0}' -f $Exclude };
	    Directories { $CommonRobocopyParams += ' /XD {0}' -f $Exclude };
    }
    
    #endregion Robocopy params
    
    #generate log format
    $TimeGenerated = "$(Get-Date -Format HH:mm:ss).$((Get-Date).Millisecond)+000"
    $Line = '<![LOG[{0}]LOG]!><time="{1}" date="{2}" component="{3}" context="" type="{4}" thread="" file="">'

    #region Robocopy Staging
    Write-Verbose -Message 'Analyzing robocopy job ...';
    $StagingLogPath = '{0}\offlinemodules-staging-{1}.log' -f $env:temp, (Get-Date -Format 'yyyy-MM-dd hh-mm-ss');

    $StagingArgumentList = '"{0}" "{1}" /LOG:"{2}" /L {3}' -f $Source, $Destination, $StagingLogPath, $CommonRobocopyParams;
    Write-Verbose -Message ('Staging arguments: {0}' -f $StagingArgumentList);
    Start-Process -Wait -FilePath robocopy.exe -ArgumentList $StagingArgumentList -WindowStyle Hidden;
    # Get the total number of files that will be copied
    $StagingContent = Get-Content -Path $StagingLogPath;
    $TotalFileCount = $StagingContent.Count - 1;

    # Get the total number of bytes to be copied
    [RegEx]::Matches(($StagingContent -join "`n"), $RegexBytes) | % { $BytesTotal = 0; } { $BytesTotal += $_.Value; };
    Write-Verbose -Message ('Total bytes to be copied: {0}' -f $BytesTotal);
    #endregion Robocopy Staging

    #region Start Robocopy
    # Begin the robocopy process
    $RobocopyLogPath = '{0}\offlinemodules-{1}.log' -f $env:temp, (Get-Date -Format 'yyyy-MM-dd hh-mm-ss');
    $ArgumentList = '"{0}" "{1}" /LOG:"{2}" /ipg:{3} {4}' -f $Source, $Destination, $RobocopyLogPath, $Gap, $CommonRobocopyParams;
    Write-Verbose -Message ('Beginning the robocopy process with arguments: {0}' -f $ArgumentList);
    $Robocopy = Start-Process -FilePath robocopy.exe -ArgumentList $ArgumentList -Verbose -PassThru -WindowStyle Hidden;
    Start-Sleep -Milliseconds 100;
    #endregion Start Robocopy

    #region Progress bar loop
    while (!$Robocopy.HasExited) {
        Start-Sleep -Milliseconds $ReportGap;
        $BytesCopied = 0;
        $LogContent = Get-Content -Path $RobocopyLogPath;
        $BytesCopied = [Regex]::Matches($LogContent, $RegexBytes) | ForEach-Object -Process { $BytesCopied += $_.Value; } -End { $BytesCopied; };
        $CopiedFileCount = $LogContent.Count - 1;
        Write-Verbose -Message ('Bytes copied: {0}' -f $BytesCopied);
        Write-Verbose -Message ('Files copied: {0}' -f $LogContent.Count);
        $Percentage = 0;
        if ($BytesCopied -gt 0) {
           $Percentage = (($BytesCopied/$BytesTotal)*100)
        }
        If ($ProgressDisplayName){$ActivityDisplayName = $ProgressDisplayName}Else{$ActivityDisplayName = 'Robocopy'}
        Write-Progress -Activity $ActivityDisplayName -Status ("Copied {0} of {1} files; Copied {2} of {3} bytes" -f $CopiedFileCount, $TotalFileCount, $BytesCopied, $BytesTotal) -PercentComplete $Percentage
    }
    #endregion Progress loop

    #region Function output
    [PSCustomObject]@{
        BytesCopied = $BytesCopied;
        FilesCopied = $CopiedFileCount;
    };
    #endregion Function output
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
        ReportStartOfActivity "Searching for $productShortName module components..."

   

       $loaded = Get-Module -Name $moduleList.ModuleName -ErrorAction Ignore | % {$_.Name}
       $registered = Get-Module -Name $moduleList.ModuleName -ListAvailable -ErrorAction Ignore | % {$_.Name}
       $notLoaded = $null
       $notLoaded = $registered | ? {$loaded -notcontains $_}
   
       ReportFinishedActivity
   
       foreach ($module in $registered) {
          if ($loaded -notcontains $module) {
		     ReportStartOfActivity "Loading module $module"
         
		     Import-Module $module
		 
		     ReportFinishedActivity
          }
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
$UserScriptPath = "$env:USERPROFILE\Documents\WindowsPowerShell\Scripts"
$AllUsersModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:ProgramFiles\WindowsPowerShell*"}
$SystemModulePath = $env:PSModulePath -split ';' | Where {$_ -like "$env:windir*"}

#find profile module

$PowerShellNoISEProfile = $profile -replace "ISE",""


#Install Nuget prereq
$NugetAssembly = Get-ChildItem $BinPath -Recurse -Filter Nuget -Directory
If ($NugetAssembly){
    $NugetAssemblyVersion = (Get-ChildItem $NugetAssembly.FullName).Name
    $NugetAssemblyDestPath = "$env:ProgramFiles\PackageManagement\ProviderAssemblies\Nuget"
    If (!(Test-Path $NugetAssemblyDestPath\$NugetAssemblyVersion)){
        Write-Host "Copying Nuget Assembly ($NugetAssemblyVersion) to $NugetAssemblyDestPath" -ForegroundColor Cyan
        New-Item $NugetAssemblyDestPath -ItemType Directory -ErrorAction SilentlyContinue
        #Copy-Item -Path "$NugetAssemblySourcePath\*" -Destination $NugetAssemblyDestPath –Recurse -ErrorAction SilentlyContinue
        Copy-WithProgress -Source $NugetAssembly.FullName -Destination $NugetAssemblyDestPath -ProgressDisplayName 'Copying Nuget Assembly Files...'
    }
    Else{
        Write-Host "INFO: Nuget assembly is already up-to-date with version [$NugetAssemblyVersion]..." -ForegroundColor DarkGreen
    }
}
Else{
     Write-Host "ERROR: Nuget module was not found..." -ForegroundColor Red
     Exit -1
}

#Find PowerCLI Module
$PowerCLIFolder = $("VMware.PowerCLI")
$DownloadedModule = Get-ChildItem $ModulesPath -Filter *.psd1 -Recurse | Where-Object {$_.FullName -match $LatestModuleVersion -and $_.Name -match 'VMware.PowerCLI.psd1'}
#$DownloadedModule = Get-ChildItem "$ModulesPath" -Directory | Where-Object {$_.FullName -match 'VMware.PowerCLI'}
#get the version based on naming convention
If ($DownloadedModule){
    $Manifest = Test-ModuleManifest $DownloadedModule.FullName -ErrorAction Ignore
    $PowerCLIModuleBasePath = $Manifest.ModuleBase
    $PowerCLIModuleVersion = $Manifest.Version.ToString()
    $PowerCLIModuleSourcePath = "$ModulesPath\$PowerCLIFolder"
    $PowerCLIModuleDestPath = "$UserModulePath\$PowerCLIFolder"

    #copy PowerCLI Modules to User directory if they don't exist ($env:PSModulePath)
    If (!(Test-Path "$PowerCLIModuleDestPath\$PowerCLIFolder\$PowerCLIModuleVersion\VMware.PowerCLI.psd1")){
        Write-Host "Copying VMware PowerCLI [$PowerCLIModuleVersion] module and dependencies files to $PowerCLIModuleDestPath" -ForegroundColor Cyan
        #Create directory if not exists
        New-Item $PowerCLIModuleDestPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        New-Item $UserScriptPath -ItemType Directory -ErrorAction SilentlyContinue | Out-Null
        #copy modules files
        Copy-WithProgress -Source $PowerCLIModuleSourcePath -Destination $PowerCLIModuleDestPath -ExcludeType Directories -Exclude 'Nuget' -ProgressDisplayName 'Copying Vmware PowerCLI Modules Files...'
        #copy scripts for module files
        Copy-Item -Path "$BinPath\$PowerCLIFolder\VMware PowerCLI (32-Bit).lnk" -Destination "$env:USERPROFILE\Desktop" -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "$BinPath\$PowerCLIFolder\VMware PowerCLI.lnk" -Destination "$env:USERPROFILE\Desktop" -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "$BinPath\$PowerCLIFolder\Initialize-PowerCLIEnvironment.ps1" -Destination $UserScriptPath -ErrorAction SilentlyContinue | Out-Null
        Copy-Item -Path "$BinPath\$PowerCLIFolder\VMware.PowerCLI.ico" -Destination $PowerCLIModuleDestPath -ErrorAction SilentlyContinue | Out-Null
    }


    $CopiedPowerCLIModulePSD1Path = "$PowerCLIModuleDestPath\$PowerCLIFolder\$PowerCLIModuleVersion\VMware.PowerCLI.psd1"

    If(Test-Path $CopiedPowerCLIModulePSD1Path){
        $Manifest = Test-ModuleManifest $CopiedPowerCLIModulePSD1Path -ErrorAction Ignore
        $PowerCLIVersion = $Manifest.Version.ToString()
        $moduleList = $Manifest.RequiredModules
    }
    Else{
        Write-Host "ERROR: VMware.PowerCLI Module was not found in user directory..." -ForegroundColor Red
        #Exit -1
    }
}
Else{
     Write-Host "ERROR: PowerCLI version: $LatestModuleVersion, was not found..." -ForegroundColor Red
     Exit -1
}

If($moduleList){
    #Load Module
    $productName = $PowerCLIFolder
    $productShortName = "PowerCLI"

    $loadingActivity = "Loading $productName"
    $script:completedActivities = 0
    $script:percentComplete = 0
    $script:currentActivity = ""
    $script:totalActivities = `
       $moduleList.Count + 1

    Load-DependencyModules
}

If($LASTEXITCODE -gt 0){Break}
#Import-Module $MainModulePSD1

# Update PowerCLI version after snap-in load
$powerCliFriendlyVersion = [VMware.VimAutomation.Sdk.Util10.ProductInfo]::PowerCLIFriendlyVersion
$host.ui.RawUI.WindowTitle = $powerCliFriendlyVersion

$productName = "PowerCLI"

# Launch text
write-host "          Welcome to VMware $productName!"
write-host ""
write-host "Log in to a vCenter Server or ESX host:              " -NoNewLine
write-host "Connect-VIServer" -foregroundcolor yellow
write-host "To find out what commands are available, type:       " -NoNewLine
write-host "Get-VICommand" -foregroundcolor yellow
write-host "To show searchable help for all PowerCLI commands:   " -NoNewLine
write-host "Get-PowerCLIHelp" -foregroundcolor yellow  
write-host "Once you've connected, display all virtual machines: " -NoNewLine
write-host "Get-VM" -foregroundcolor yellow
write-host "If you need more help, visit the PowerCLI community: " -NoNewLine
write-host "Get-PowerCLICommunity" -foregroundcolor yellow
write-host ""
write-host "       Copyright (C) VMware, Inc. All rights reserved."
write-host ""
write-host ""

# CEIP
Try	{
	$configuration = Get-PowerCLIConfiguration -Scope Session

	if ($promptForCEIP -and
		$configuration.ParticipateInCEIP -eq $null -and `
		[VMware.VimAutomation.Sdk.Util10Ps.CommonUtil]::InInteractiveMode($Host.UI)) {

		# Prompt
		$caption = "Participate in VMware Customer Experience Improvement Program (CEIP)"
		$message = `
			"VMware's Customer Experience Improvement Program (`"CEIP`") provides VMware with information " +
			"that enables VMware to improve its products and services, to fix problems, and to advise you " +
			"on how best to deploy and use our products.  As part of the CEIP, VMware collects technical information " +
			"about your organization’s use of VMware products and services on a regular basis in association " +
			"with your organization’s VMware license key(s).  This information does not personally identify " +
			"any individual." +
			"`n`nFor more details: press Ctrl+C to exit this prompt and type `"help about_ceip`" to see the related help article." +
			"`n`nYou can join or leave the program at any time by executing: Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP `$true or `$false. "

		$acceptLabel = "&Join"
		$choices = (
			(New-Object -TypeName "System.Management.Automation.Host.ChoiceDescription" -ArgumentList $acceptLabel,"Participate in the CEIP"),
			(New-Object -TypeName "System.Management.Automation.Host.ChoiceDescription" -ArgumentList "&Leave","Don't participate")
		)
		$userChoiceIndex = $Host.UI.PromptForChoice($caption, $message, $choices, 0)
		
		$participate = $choices[$userChoiceIndex].Label -eq $acceptLabel

		if ($participate) {
         [VMware.VimAutomation.Sdk.Interop.V1.CoreServiceFactory]::CoreService.CeipService.JoinCeipProgram();
      } else {
         Set-PowerCLIConfiguration -Scope User -ParticipateInCEIP $false -Confirm:$false | Out-Null
      }
	}
} Catch {
	# Fail silently
}
# end CEIP

Write-Progress -Activity $loadingActivity -Completed


cd \