<### What is this and who made it?! ####
Author:  Rudy Bankson, Twitter: PHYSX51
This script will create an application and test deployment for the Microsoft 365 Enterprise Apps, Visio, and Project
using the latest files from Microsoft. 

Known Issues:
I use Current Channel. There is a piece of code here that checks to get the latest version/build info for the Current
Channel of the apps. This will be used by SCCM to make the folder and application in SCCM with correct versioning info.
If you want to use another channel, scroll on down until you find this line and replace "Current" with the exact
name of your preferred channel. 
$CurrentChannel = $xmlElm.ReleaseHistory.UpdateChannel | Where-Object Name -EQ 'Current'
Then modify the XML files to also point to your desired channel. I'll make this a variable later... just don't want 
to rework that much yet.


#### VARIABLES YOU NEED TO CHANGE ####>

#  Create this folder and place this PowerShell script in it, along with the XML files. You will need
#  to modify the configuration.xml file to place your ORG NAME in it. It's obvious where. If you don't
#  download the XML files, this script will download my defaults for you, and the only thing wonky
#  will be the ORG NAME showing up as Contoso. This will be the "template" that copies over files
#  to your monthly build.
$TemplateUNC = "\\YourFileServer\software\SCCM\Applications\Microsoft\Office365\OfficeM365_Forever\"  #This script is touchy... put a \ at the end of the UNC or it won't work right.

#  This folder should just be one level up from your "Forever" folder (above) which contains your "templates".
$OfficeRoot = "\\YourFileServer\software\SCCM\Applications\Microsoft\Office365\" #This script is touchy... put a \ at the end of the UNC or it won't work right.

#  This is going to be all of your SCCM site info
$SiteCode = "A01"
$ProviderMachineName = "YourMP.fqdn.com"  # SMS Provider machine name (Your 'SCCM Server')
$DPName = "DPName.YourFQDN.COM"

#  Create a Device Collection. This script will make an available deployment to it when finished.
$TestMachineCollection = "A0100564"

#  In Applications, this is the folder your applications will be created under. I don't make this for you,
#  so be sure to create the folder in SCCM Applications on your own.
$CMConsoleAppFolder = "! Pre-Production"

#  Get your own icons... I can't do this for you because lawyers.
$o365icon = "\\YourFileServer\software\SCCM\Icons\o365.png"
$VisioIcon = "\\YourFileServer\software\SCCM\Icons\Microsoft-Visio.png"
$ProjectIcon = "\\YourFileServer\software\SCCM\Icons\Microsoft-Project.png"

##############################################################################
##############################################################################
######################## DON'T MODIFY AFTER THIS LINE ########################
##############################################################################
##############################################################################

# Variable magic - don't touch this part
$ConfigXML = $TemplateUNC+"configuration.xml"
$DefaultConfigXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/configuration.xml'
$DownloadXML = $TemplateUNC+"download.xml"
$DefaultDownloadXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/download.xml'
$ProjectXML = $TemplateUNC+"project.xml"
$DefaultProjectXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/project.xml'
$RemoveOfficeXML = $TemplateUNC+"removeOffice.xml"
$DefaultRemoveOfficeXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/removeOffice.xml'
$RemoveProjectXML = $TemplateUNC+"removeProject.xml"
$DefaultRemoveProjectXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/removeProject.xml'
$RemoveVisioXML = $TemplateUNC+"removeVisio.xml"
$DefaultRemoveVisioXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/removeVisio.xml'
$SetupEXE = $TemplateUNC+"setup.exe"
$VisioXML = $TemplateUNC+"visio.xml"
$DefaultVisioXML = 'https://raw.githubusercontent.com/rudybankson/M365AppsForever/main/visio.xml'
$CMConsoleAppFolderFinal = $SiteCode+":Application\"+$CMConsoleAppFolder

# Get download path for ODT and set to variable
$PageURL = 'https://www.microsoft.com/en-us/download/confirmation.aspx?id=49117'
$ODTexe = Invoke-RestMethod $PageURL | ForEach-Object { if ($_ -match 'url=(https://.*officedeploymenttool.*\.exe)') { $matches[1] } }


#### This section gets information about the Current Channel release history of Microsoft 365 Enterprise Apps ####
#### The script will download the ReleaseHistory.cab file from MS and extract it to get the ReleaseHistory.xml ####
#### It then parses the XML file to create version & build number variables for the latest Current Channel release ####

# URL of the CAB file to download
$url = "https://officecdn.microsoft.com/pr/wsus/releasehistory.cab"

# Destination directory
$destinationRoot = "C:\Temp\M365EntAppsRelHistory"

# Create a variable to later create a folder in the destination directory with today's date as the name
$destinationDirectory = Join-Path -Path $destinationRoot -ChildPath (Get-Date -Format "yyyy-MM-dd")

# Check if destination directory exists, if not, create it
if (-not (Test-Path $destinationDirectory -PathType Container)) {
    New-Item -Path $destinationDirectory -ItemType Directory | Out-Null
}

# Path to save the CAB file
$cabFilePath = Join-Path -Path $destinationDirectory -ChildPath (Split-Path -Path $url -Leaf)

# Download the CAB file
Invoke-WebRequest -Uri $url -OutFile $cabFilePath

# Check if download was successful
if (-not (Test-Path $cabFilePath -PathType Leaf)) {
    Write-Error "Failed to download the CAB file from $url."
    exit
}

# Extract the contents of the CAB file
$extractdir = $destinationDirectory+"\releasehistory.xml"
cmd.exe /c "expand $cabFilePath -F:releasehistory.xml $extractdir"
Write-Host "CAB file downloaded from $url and extracted to $extractdir"

# Extract info from XML for build process
[xml]$xmlElm = Get-Content -Path $extractdir
$CurrentChannel = $xmlElm.ReleaseHistory.UpdateChannel | Where-Object Name -EQ 'Current'
$CurrentChannelLatest = $CurrentChannel.Update | Where-Object Latest -EQ 'True'

# Set variables for Version, Legacy Version, and Build
$CurrentChannelVersion = $CurrentChannelLatest.Version
$CurrentChannelLegacyVersion = $CurrentChannelLatest.LegacyVersion
$CurrentChannelBuild = $CurrentChannelLatest.Build

#### End section for getting the variables about the Current Channel release ####


# Copy over missing files from Rudy's GitHub and MS
if (-not (Test-Path $ConfigXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultConfigXML -OutFile $ConfigXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $DownloadXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultDownloadXML -OutFile $DownloadXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $ProjectXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultProjectXML -OutFile $ProjectXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $RemoveOfficeXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultRemoveOfficeXML -OutFile $RemoveOfficeXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $RemoveProjectXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultRemoveProjectXML -OutFile $RemoveProjectXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $RemoveVisioXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultRemoveVisioXML -OutFile $RemoveVisioXML | Out-Null
    Start-Sleep -Seconds 3
    }

if (-not (Test-Path $VisioXML -PathType Leaf)) {
    Invoke-WebRequest -Uri $DefaultVisioXML -OutFile $VisioXML | Out-Null
    Start-Sleep -Seconds 3
    }

# Create variables to be used to run ODT... yes, it could be cleaner
$TemplateUNCshort = $TemplateUNC.Substring(0, $TemplateUNC.Length - 1)
$Arguments = '"'+'/quiet /extract:`"'+$TemplateUNCshort+'`""'
$stringArguments = $Arguments.ToString()
$StartProcessODT = "Start-Process "+"$destinationDirectory\ODT.exe"+" -ArgumentList "
$FullODT = $StartProcessODT+$stringArguments+' -Wait'

if (-not (Test-Path $SetupEXE -PathType Leaf)) {
    Invoke-WebRequest -Uri $ODTexe -OutFile "$destinationDirectory\ODT.exe" | Out-Null
    try {
      Invoke-Expression -Command $FullODT
      # Path to the file
$file1 = $TemplateUNC+"configuration-Office365-x64.xml"

# Check if the file exists
if (Test-Path $file1) {
    # If the file exists, delete it
    Remove-Item $file1 -Force
    Write-Output "File deleted successfully."
} else {
    Write-Output "File does not exist."
}

# Path to the file
$file2 = $TemplateUNC+"configuration-Office365-x86.xml"

# Check if the file exists
if (Test-Path $file2) {
    # If the file exists, delete it
    Remove-Item $file2 -Force
    Write-Output "File deleted successfully."
} else {
    Write-Output "File does not exist."
}

# Path to the file
$file3 = $TemplateUNC+"configuration-Office2019Enterprise.xml"

# Check if the file exists
if (Test-Path $file3) {
    # If the file exists, delete it
    Remove-Item $file3 -Force
    Write-Output "File deleted successfully."
} else {
    Write-Output "File does not exist."
}

# Path to the file
$file4 = $TemplateUNC+"configuration-Office2021Enterprise.xml"

# Check if the file exists
if (Test-Path $file4) {
    # If the file exists, delete it
    Remove-Item $file4 -Force
    Write-Output "File deleted successfully."
} else {
    Write-Output "File does not exist."
}


      }
    catch {
      Write-Warning 'Error running ODT, error code:'
      Write-Warning $_
      }
    }


# Creates new folder for next Version of Microsoft 365 Apps for Enterprise
$Version = $CurrentChannelVersion
$NewFolderName=$OfficeRoot+$Version
if (Test-Path -Path $NewFolderName) {Write-Host 'Path already exists - aborting'}
else {
New-Item -Path $NewFolderName -type directory

# Copies XML files from template folder to new folder
Copy-Item -Path $ConfigXML -Destination $NewFolderName
Copy-Item -Path $DownloadXML -Destination $NewFolderName
Copy-Item -Path $ProjectXML -Destination $NewFolderName
Copy-Item -Path $RemoveOfficeXML -Destination $NewFolderName
Copy-Item -Path $RemoveProjectXML -Destination $NewFolderName
Copy-Item -Path $RemoveVisioXML -Destination $NewFolderName
Copy-Item -Path $SetupEXE -Destination $NewFolderName
Copy-Item -Path $VisioXML -Destination $NewFolderName

# Download the new Microsoft 365 Apps for Enterprise binaries
Set-Location $NewFolderName
.\setup.exe /download download.xml | Out-Null

# Creates Application in SCCM for Microsoft 365 Apps for Enterprise

# Sets variables to be used later in the packaging process
$ProductName = 'Microsoft 365 Apps for Enterprise '+$Version
$ProductVersion = $CurrentChannelLegacyVersion
$Manufacturer = 'Microsoft'

# Imports SCCM PowerShell module
# Customizations
$initParams = @{}
#$initParams.Add("Verbose", $true) # Uncomment this line to enable verbose logging
#$initParams.Add("ErrorAction", "Stop") # Uncomment this line to stop the script on any errors

# Import the ConfigurationManager.psd1 module 
if((Get-Module ConfigurationManager) -eq $null) {
    Import-Module "$($ENV:SMS_ADMIN_UI_PATH)\..\ConfigurationManager.psd1" @initParams 
}

# Connect to the site's drive if it is not already present
if((Get-PSDrive -Name $SiteCode -PSProvider CMSite -ErrorAction SilentlyContinue) -eq $null) {
    New-PSDrive -Name $SiteCode -PSProvider CMSite -Root $ProviderMachineName @initParams
}

# Set the current location to be the site code.
Set-Location "$($SiteCode):\" @initParams

# Make some more variables
$SCCMAppName = $ProductName+' '+$ProductVersion
$Detection = New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\O365ProPlusRetail - en-us' -PropertyType Version -ValueName 'DisplayVersion' -Value -ExpectedValue $ProductVersion -ExpressionOperator GreaterEquals

# Create application and deployment type in SCCM
New-CMApplication -Name $SCCMAppName -Description "Microsoft Word, Excel, PowerPoint, Outlook, and other Office apps." -Publisher "Microsoft" -SoftwareVersion $ProductVersion -IconLocationFile $o365icon -AutoInstall $true
Add-CMScriptDeploymentType -ApplicationName $SCCMAppName -DeploymentTypeName $SCCMAppName -InstallCommand 'setup.exe /configure configuration.xml' -UninstallCommand 'setup.exe /configure removeoffice.xml' -AddDetectionClause $Detection -ContentLocation $NewFolderName -InstallationBehaviorType InstallForSystem -EstimatedRuntimeMins 30 -LogonRequirementType WhetherOrNotUserLoggedOn

# Distribute Content to DP specified in variable at top
Start-CMContentDistribution -ApplicationName $SCCMAppName -DistributionPointName $DPName

# Create SCCM deployment to EP test devices
New-CMApplicationDeployment -CollectionId $TestMachineCollection -Name $SCCMAppName -DeployAction Install -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly -AvailableDateTime (get-date) -TimeBaseOn LocalTime -Verbose

# Moves to Pre-Prod Folder
$AppToMove = Get-CMApplication -Name $SCCMAppName
Move-CMObject -FolderPath $CMConsoleAppFolderFinal -InputObject $AppToMove
# ===========================================================================================


# ===========================================================================================
# Create application for Visio
# Make some variables
$ProductName2 = 'Microsoft Visio '+$Version
$SCCMAppName2 = $ProductName2+' '+$ProductVersion
$Detection2 = New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\VisioProRetail - en-us' -PropertyType Version -ValueName 'DisplayVersion' -Value -ExpectedValue $ProductVersion -ExpressionOperator GreaterEquals

# Create Visio application and deployment type in SCCM
New-CMApplication -Name $SCCMAppName2 -Description "Microsoft Visio for Subscription Members." -Publisher "Microsoft" -SoftwareVersion $ProductVersion -IconLocationFile $VisioIcon -AutoInstall $true
Add-CMScriptDeploymentType -ApplicationName $SCCMAppName2 -DeploymentTypeName $SCCMAppName2 -InstallCommand 'setup.exe /configure visio.xml' -UninstallCommand 'setup.exe /configure removevisio.xml' -AddDetectionClause $Detection2 -ContentLocation $NewFolderName -InstallationBehaviorType InstallForSystem -EstimatedRuntimeMins 30 -LogonRequirementType WhetherOrNotUserLoggedOn

# Distribute Content to DP specified in variable at top
Start-CMContentDistribution -ApplicationName $SCCMAppName2 -DistributionPointName $DPName

# Create SCCM deployment to EP test devices
New-CMApplicationDeployment -CollectionId $TestMachineCollection -Name $SCCMAppName2 -DeployAction Install -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly -AvailableDateTime (get-date) -TimeBaseOn LocalTime -Verbose

# Moves to Pre-Prod Folder
$AppToMove2 = Get-CMApplication -Name $SCCMAppName2
Move-CMObject -FolderPath $CMConsoleAppFolderFinal -InputObject $AppToMove2
# ===========================================================================================


# ===========================================================================================
# Create application for Project
# Make some variables
$ProductName3 = 'Microsoft Project '+$Version
$SCCMAppName3 = $ProductName3+' '+$ProductVersion
$Detection3 = New-CMDetectionClauseRegistryKeyValue -Hive LocalMachine -KeyName 'SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall\ProjectProRetail - en-us' -PropertyType Version -ValueName 'DisplayVersion' -Value -ExpectedValue $ProductVersion -ExpressionOperator GreaterEquals

# Create Visio application and deployment type in SCCM
New-CMApplication -Name $SCCMAppName3 -Description "Microsoft Project for Subscription Members." -Publisher "Microsoft" -SoftwareVersion $ProductVersion -IconLocationFile $ProjectIcon -AutoInstall $true
Add-CMScriptDeploymentType -ApplicationName $SCCMAppName3 -DeploymentTypeName $SCCMAppName3 -InstallCommand 'setup.exe /configure project.xml' -UninstallCommand 'setup.exe /configure removeproject.xml' -AddDetectionClause $Detection3 -ContentLocation $NewFolderName -InstallationBehaviorType InstallForSystem -EstimatedRuntimeMins 30 -LogonRequirementType WhetherOrNotUserLoggedOn

# Distribute Content to DP specified in variable at top
Start-CMContentDistribution -ApplicationName $SCCMAppName3 -DistributionPointName $DPName

# Create SCCM deployment to EP test devices
New-CMApplicationDeployment -CollectionId $TestMachineCollection -Name $SCCMAppName3 -DeployAction Install -DeployPurpose Available -UserNotification DisplaySoftwareCenterOnly -AvailableDateTime (get-date) -TimeBaseOn LocalTime -Verbose

# Moves to Pre-Prod Folder
$AppToMove3 = Get-CMApplication -Name $SCCMAppName3
Move-CMObject -FolderPath $CMConsoleAppFolderFinal -InputObject $AppToMove3
# ===========================================================================================


}

