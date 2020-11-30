<#
.SYNOPSIS
Deletes all empty folders in the SCOM console.

Author: Patrick Seidl
        (c) s2 - seidl solutions
Date:   08.04.2019

.DESCRIPTION
Deletes all empty folders in the SCOM console like the ones created during MP creation.

.PARAMETER whatIf
Only lists folders to be deleted. Will not delete any folder.
  
.EXAMPLE 1 (Default)
Delete-EmptyConsoleFolders.ps1

.EXAMPLE 2
Delete-EmptyConsoleFolders.ps1 -whatIf

#>

param([switch]$whatIf)

"Load SDK"
Import-Module OperationsManager
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager.Common.dll") | Out-Null
[System.Reflection.Assembly]::LoadWithPartialName("Microsoft.EnterpriseManagement.OperationsManager.dll") | Out-Null

"Connect to MG"
$UserRegKeyPath = "HKCU:\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\User Settings"
$MachineRegKeyPath = "HKLM:\SOFTWARE\Microsoft\Microsoft Operations Manager\3.0\Machine Settings"
$UserRegValueName = "SDKServiceMachine"
$MachineRegValueName = "DefaultSDKServiceMachine"
$regValue = $null
$managementServer = $null
$regValue = Get-ItemProperty -path:$UserRegKeyPath -name:$UserRegValueName -ErrorAction:SilentlyContinue
if ($regValue -ne $null) {
       $managementServer = $regValue.SDKServiceMachine
}
if ($managementServer -eq $null -or $managementServer.Length -eq 0) {
       $regValue = Get-ItemProperty -path:$MachineRegKeyPath -name:$MachineRegValueName -ErrorAction:SilentlyContinue
       if ($regValue -ne $null) {
           $managementServer = $regValue.DefaultSDKServiceMachine
       }
}
$managementGroup = [Microsoft.EnterpriseManagement.ManagementGroup]::Connect($managementServer)

"Get all folders"
$allFolders = $managementGroup.Presentation.GetFolders() | ? {$_.Identifier -like "*Folder_*"} # automatically created folders start with that ID

"Get folders from unsealed MPs with subfolders"
$usedFolders = $null
$usedFolders = @()
$i = 0
$all = $allFolders.count
foreach ($oneFolder in $allFolders) {
    $i++
    Write-Progress "Get folders with subfolders" -Status "$i of $all" -PercentComplete ($i / $all * 100)
    if ($oneFolder.getmanagementpack().sealed -eq $true) {
        #"Sealed MP"
        $usedFolders += $oneFolder.Name
    } else {
        $subFolder = $oneFolder.GetSubFolders()
        if ($subFolder.count -gt 0) {
            #"Has Subfolders"
            $usedFolders += $oneFolder.Name
        }
    }
}

"Get parent folders for views"
$usedFolders += (($managementGroup.Presentation.GetViews()).GetFolders()).Name

"Get parent folders for dashboards"
$compTypes = $managementGroup.Dashboard.GetComponentTypes()
$i = 0
$all = $compTypes.count
foreach ($compType in $compTypes) {
    $i++
    Write-Progress "Get folders with dashboards" -Status "$i of $all" -PercentComplete ($i / $all * 100)
    $compTypeName = $compType.name
    $compRefParent = ($managementGroup.Dashboard.GetComponentReferences() | ? {$_.name -like "$compTypeName*"}).parent
    $usedFolders += ($managementGroup.Presentation.GetFolders() | ? {$_.Id -eq $compRefParent.Id}).name
}

"Delete unused folders:"
$i = 0
$all = $allFolders.count
foreach ($checkFolder in $allFolders) {
    $i++
    Write-Progress "Delete unused folders" -Status "$i of $all" -PercentComplete ($i / $all * 100)
    if ($usedFolders -notcontains $checkFolder.Name) {
        $checkFolder.DisplayName
        if ($whatIf) {   
            "Testing only, folder will not be deleted"
        } else {
            $checkFolder.Status = [Microsoft.EnterpriseManagement.Configuration.ManagementPackElementStatus]::PendingDelete
            $checkFolder.GetManagementPack().AcceptChanges()
        }
    }
}

"Done"