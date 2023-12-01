<#
.SYNOPSIS
Just a quick and dirty check of WMI, if it fails blast the entire thing, and start fresh. Using Microsoft recommended rebuilding with a custom regex to avoid
un-installation .MOF files.

.DESCRIPTION
This script will simply query the WMI table for the computer name, and if this is not successfully returned, it will initiate a delete and rebuild of the WMI database.
The rebuild stops necessary services, then starts with registering the files for SCCM, followed by registering all .MOF, .DLL, & .MFL files afterwords. Once done, this
performs a final WMI check, and should populate the Nexthink engine variable with the hostname of the PC if the repair was successful.

.FUNCTIONALITY
Remediation

.INPUTS
ID  Label                           Description
N/A

.OUTPUTS
ID  Label                           Description
1   WMI                     Should contain the device hostname is WMI has been successfully repaired.

.NOTES
Context:           <InteractiveUser/LocalSystem>
Version:           1.0.0
Original Author:   Brandon Woods
Created:           Wednesday, November 29th 2023, 9:42:57 am
Copyright (C) 2023 Keysight Technologies
File: WMI Test & Repair.ps1
Last Modified: 2023/12/01 08:36:15
Modified By: Brandon Woods
HISTORY:
Date                 By     Comments
2023-12-01 08:32 am	 BW	    V1.0.0  - Rough code, need to add error checking and validation, but this works for now...
#>



$REMOTE_ACTION_DLL_PATH = "$env:NEXTHINK\RemoteActions\nxtremoteactions.dll"

function Add-NexthinkRemoteActionDLL {

    if (-not (Test-Path -Path $REMOTE_ACTION_DLL_PATH)) {
        throw 'Nexthink Remote Action DLL not found. '
    }
    Add-Type -Path $REMOTE_ACTION_DLL_PATH
}

Add-NexthinkRemoteActionDLL
$WMITest = &WMIC.exe computersystem get name
if ($WMITest) {
    $host.ui.WriteLine("WMI responded OK. No Repair needed.")
} else {
    $Host.UI.WriteLine("WMI issues found... beginning repair")

    Stop-Service -Name CcmExec -Force
    Stop-Service -Name VMAuthdService -Force
    Stop-Service -Name winmgmt -Force

    push-location $env:windir\System32\wbem
    remove-Item -Path repository -Recurse -Force

    &regsvr32.exe /s $env:windir\system32\scecli.dll
    &regsvr32.exe /s $env:windir\system32\userenv.dll

    &mofcomp.exe cimwin32.mof
    &mofcomp.exe cimwin32.mfl
    &mofcomp.exe rsop.mof
    &mofcomp.exe rsop.mfl

    $DLLFiles = Get-ChildItem -Recurse | Where-Object {$_ -match '^((?!uninstall|Uninstall).)*.dll$'} |ForEach-Object { $_.FullName }
    $MOFFiles = Get-ChildItem -Recurse | Where-Object {$_ -match '^((?!uninstall|Uninstall).)*.mof$'} |ForEach-Object { $_.FullName }
    $MFLFiles = Get-ChildItem -Recurse | Where-Object {$_ -match '^((?!uninstall|Uninstall).)*.mfl$'} |ForEach-Object { $_.FullName }

    foreach ($File in $DLLFiles) {
        &regsvr32.exe /s $File
    }

    foreach ($File in $MOFFiles) {
        &mofcomp.exe $File
    }

    foreach ($File in $MFLFiles) {
        &mofcomp.exe $File
    }

    Start-Service -Name Winmgmt
    Start-Service -Name VMAuthdService
    Start-Service -Name CcmExec
}

# Checking if it worked - Should display computer name if repair was successful.
$FinalWMITest = &WMIC.exe computersystem get name
[Nxt]::WriteOutputString("WMI", "WMI Hostname Results `r`n $FinalWMITest")