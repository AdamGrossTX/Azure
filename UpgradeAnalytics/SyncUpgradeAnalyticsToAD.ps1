<#
.SYNOPSIS
Connect to Azure Upgrade Analytics (In Azure Monitor) and import readiness data into AD

.DESCRIPTION
Connect to Azure Upgrade Analytics (In Azure Monitor) and import readiness data into AD

.EXAMPLE
.\SyncUpgradeAnalyticsToAD.ps1 -OUFilter "MyOuName*" -WorkspaceID "23232-3423424312-2423242-23"

.NOTES
Author: Adam Gross
Twitter: @AdamGrossTX
GitHub: AdamGrossTX
Web: http://www.asquaredozen.com

Version: 1.0

Special thanks to everyone who helped point me to ADSI in this thread https://twitter.com/AdamGrossTX/status/1099432202545348609
The OU lookup part of this script uses ther AdsiPS https://github.com/lazywinadmin/AdsiPS by François-Xavier Cat

#>
[cmdletbinding()]
param(
    [string]$OUFilter = "*",
    [string]$WorkspaceID = "", #Upgrade Analytics WorkspaceID
    [switch]$CreateNewContext = $false
)

if (-not (Get-Module -Name Az)) {Install-Module Az -Force}
if (-not (Get-Module -Name AdsiPS)) {Install-Module AdsiPS -Force -AllowClobber; Import-Module AdsiPS -Force}

Add-Type -AssemblyName System.DirectoryServices.AccountManagement

#region Get Upgrade Analytics Data
    If($CreateNewContext.IsPresent) {
        Login-AzAccount
        Save-AzContext -Path "$($PSScriptRoot)\azprofile.json" -Force
    }

    $Query ='UAComputer | project Computer,UpgradeDecision'
    $Hours = 24
    $TimeSpan = (New-TimeSpan -Hours $Hours)
    Import-AzContext -Path "$($PSScriptRoot)\azprofile.json"

    $AzResults = Invoke-AzOperationalInsightsQuery -WorkspaceId $WorkspaceID -Query $Query -Timespan $TimeSpan
    $AzComputerList = $AzResults.Results | Sort-Object Computer
#endregion

#region AD
    $OUList = Get-ADSIOrganizationalUnit -Name "$($OUFilter)"
    
    workflow UpdateMachines {
    param($ADSPathList, $AzComputerList)
        Foreach -parallel ($path in $ADSPathList) {
            InlineScript {
                [adsi]$adsiDevice = $using:path
                $UpgradeDecision = $AzComputerList | Where-Object Computer -eq $adsiDevice.properties.cn | Select-Object -ExpandProperty UpgradeDecision
                If(!($UpgradeDecision)) { $UpgradeDecision = "NoData"}
                $adsiDevice.Put("extensionAttribute9",$UpgradeDecision)
                $adsiDevice.SetInfo()
            }
        }
    }

    ForEach ($OU in $OUList)
    {
        [adsi]$DN = "LDAP://$($OU.DistinguishedName)"
        $searcher=new-object System.DirectoryServices.DirectorySearcher($DN,'objectCategory=computer')
        $Searcher.SizeLimit = 10000
        $DeviceList = $searcher.FindAll()
        Write-Host $DeviceList.Count
        $ADSPathList = $DeviceList | ForEach-Object{$_.properties.adspath}
        UpdateMachines -ADSPathList $ADSPathList -AzComputerList $AzComputerList
    }
#endregion
