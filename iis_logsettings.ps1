<#
.SYNOPSIS
    Script to verify (and optionally set) IIS logging settings
    
.DESCRIPTION
    Script will check IIS logging settings like Log Directory, Encoding, LogFile rollover
    and logfile rotation size.

.PARAMETER Compare
    <optional> Compares all Exchange servers with the expected settings

.PARAMETER Set
    <optional> Configures all servers without expected settings with the correct ones

.PARAMETER Directory
    <optional> Log Directory

.PARAMETER ScheduleRollover
    <optional> Rollover Method Schedule

.PARAMETER MaxSizeRollover
    <optional> Rollover Method MaxSize

.PARAMETER Schedule
    <optional> Schedule Hourly,Daily,Weekly,Monthly
        
.PARAMETER RolloverSizeMB
    <optional> MaxSize Rollover threshold in MB

.PARAMETER Encoding
    <optional> Log Format UTF8/ANSI

.EXAMPLE
    .\iis_logsettings.ps1 [-Set] [-Compare] [-Directory <iis_logpath>] [-ScheduleRollover] [-MaxSizeRollover] [-Schedule <Daily/Hourly/Weekly/Monthly>][-RolloverSize <MB>] [-Encoding <UTF8/ANSI>]

.NOTES
    Steffen Meyer
    Cloud Solution Architect
    Microsoft Deutschland GmbH

    V1.0 - 27.10.2026 - Initial Version
    V1.1 - 28.10.2026 - Added "Set"
    V2.0 - 05.11.2026 - Take into account of Logging configuration of IIS root and all websites, rolloversize param changed from GB to MB
    V2.1 - 10.11.2026 - some issues corrected
    V2.2 - 23.11.2026 - corrected some issues, added Hourly,Daily,Weekly,Monthly rollover methods
    V2.5 - 24.11.2026 - reduced to only siteDefaults, Default Web Site, Exchange Back End
#>

[CmdletBinding(DefaultParameterSetName = 'ExportAll')]
Param(
     [Parameter(Mandatory,Position=0,ParameterSetName = 'CompareSchedule',HelpMessage='Create list of misconfigured server(s)')]
     [Parameter(Mandatory,Position=0,ParameterSetName = 'CompareMaxSize')]
     [Switch]$Compare,
     [Parameter(Mandatory,Position=0,ParameterSetName = 'SetSchedule',HelpMessage='Set expected settings to list of misconfigured server(s)')]
     [Parameter(Mandatory,Position=0,ParameterSetName = 'SetMaxSize')]
     [Switch]$Set,
     [Parameter(Mandatory,Position=1,ParameterSetName = 'CompareSchedule',HelpMessage='IIS LogFile Directory')]
     [Parameter(Mandatory,Position=1,ParameterSetName = 'CompareMaxSize')]
     [Parameter(Mandatory,Position=1,ParameterSetName = 'SetSchedule')]
     [Parameter(Mandatory,Position=1,ParameterSetName = 'SetMaxSize')]
     [String]$Directory,
     [Parameter(Mandatory,Position=2,ParameterSetName = 'CompareSchedule',HelpMessage='RolloverMethod Schedule')]
     [Parameter(Mandatory,Position=2,ParameterSetName = 'SetSchedule')]
     [Switch]$ScheduleRollover,
     [Parameter(Mandatory,Position=2,ParameterSetName = 'CompareMaxSize',HelpMessage='RolloverMethod MaxSize')]
     [Parameter(Mandatory,Position=2,ParameterSetName = 'SetMaxSize')]
     [Switch]$MaxSizeRollover,
     [Parameter(Mandatory,Position=3,ParameterSetName = 'CompareSchedule',HelpMessage='Rollover timeframe')]
     [Parameter(Mandatory,Position=3,ParameterSetName = 'SetSchedule')]
     [ValidateSet('Hourly','Daily','Weekly','Monthly')]
     [String]$Schedule = 'Daily',
     [Parameter(Mandatory,Position=3,ParameterSetName = 'CompareMaxSize',HelpMessage='RolloverSize in MB')]
     [Parameter(Mandatory,Position=3,ParameterSetName = 'SetMaxSize')]
     [ValidateRange(1,4095)]
     [long]$RolloverSizeMB,
     [Parameter(Mandatory=$false,Position=4,ParameterSetName = 'CompareSchedule',HelpMessage='Encoding UTF8/ANSI')]
     [Parameter(Mandatory=$false,Position=4,ParameterSetName = 'CompareMaxSize')]
     [Parameter(Mandatory=$false,Position=4,ParameterSetName = 'SetSchedule')]
     [Parameter(Mandatory=$false,Position=4,ParameterSetName = 'SetMaxSize')]
     [ValidateSet('UTF8','ANSI')]
     [String]$Encoding = 'UTF8'
     )

$version = "V2.5_24.11.2026"
$now = Get-Date

Write-Host "`nScript version: $version"
Write-Host   "Script started: $($now.tostring("dd.MM.yyyy HH:mm:ss"))"

$ScriptPath = $PSScriptRoot

#Check if Exchange SnapIn is available and load it
$SnapIn = "Microsoft.Exchange.Management.PowerShell.SnapIn"

if (!(Get-PSSession | Where-Object ConfigurationName -contains "Microsoft.Exchange"))
{
    if (Get-PSSnapin -Name $SnapIn -Registered -ErrorAction SilentlyContinue)
    {
        Write-Host "`nLoading the Exchange Powershell SnapIn..." -ForegroundColor Yellow
        Add-PSSnapin $SnapIn -ErrorAction SilentlyContinue
        . $env:ExchangeInstallPath\bin\RemoteExchange.ps1
        Connect-ExchangeServer -auto -AllowClobber
    }
    else
    {
        Write-Host "`nExchange Management Tools are not installed. Run the script on a different machine." -ForegroundColor Red
        Return
    }
}

#Detect, where the script is executed
if (!(Get-ExchangeServer -Identity $env:COMPUTERNAME -ErrorAction SilentlyContinue).serverrole -eq "Mailbox")
{
    Write-Host "`nATTENTION: Script is executed on a non-Exchangeserver..." -ForegroundColor Cyan
}

#START SCRIPT

#Create list of servers
$Servers = Get-ExchangeServer | where ServerRole -like "Mailbox" | sort name 

if ($ScheduleRollover.IsPresent)
{
    $RolloverMethod = $Schedule
}
else
{
    $RolloverMethod = 'MaxSize'
}

if ($MaxSizeRollover.IsPresent)
{
    [int64]$RolloverSizeBytes = $RolloverSizeMB * 1MB
}

$ServersCount = $Servers.Count
$Counter = 0

$Results = @()

foreach ($Server in $Servers)
{
    $Counter++

    Write-Progress -Activity "Getting IIS Logging Configuration is in progress..." -Status "Server $Counter of $ServersCount : $Server" -PercentComplete (($Counter / $ServersCount) * 100)
    
    $RemoteResult = Invoke-Command -ComputerName $Server.fqdn -ScriptBlock {
    
        Import-Module WebAdministration

        $Result = [ordered]@{}
        
        $Websites = @('siteDefaults','Default Web Site','Exchange Back End')

        foreach ($website in $Websites)
        {
            if ($website -eq 'siteDefaults')
            {
                $filter = "system.applicationHost/sites/siteDefaults/logFile"
                $prefix = 'Root'
            }
            else
            {
                $filter = "system.applicationHost/sites/site[@name='$website']/logFile"
                $prefix = $website -replace '[^a-zA-Z0-9]', ''
            }
            
            $LogSettings = Get-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter $filter -Name '*'
            
            $Result["${Prefix}Directory"] = $LogSettings.directory
            $Result["${Prefix}LogFormat"] = $LogSettings.logFormat
            $Result["${Prefix}RolloverMethod"] = $LogSettings.period
            $Result["${Prefix}RolloverSize"] = $LogSettings.truncateSize

        }
        #global log encoding setting
        $Encoding = Get-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.applicationHost/log' -Name 'logInUTF8'
        $ActualUTF8 = $Encoding.value

        if ($null -eq $ActualUTF8 -or $ActualUTF8 -isnot [bool])
        {
            Throw "Unexpected datatype for logInUTF8: $(if ($null -eq $ActualUTF8) { 'NULL' } else { $ActualUTF8.GetType().FullName })"
        }

        $Result["Encoding"] = $ActualUTF8

        [PSCustomObject]$Result
    }

    #if at least one of the expected options isn't fulfilled, add to object
    if (($PSCmdlet.ParameterSetName -eq 'ExportAll') -or 
       ((@($RemoteResult.DefaultWebSiteDirectory,$RemoteResult.ExchangeBackEndDirectory) -notlike $Directory) -or 
       (@($RemoteResult.DefaultWebSiteRolloverMethod,$RemoteResult.ExchangeBackEndRolloverMethod) -ne $RolloverMethod) -or 
       (($RolloverMethod -eq 'MaxSize') -and (@($RemoteResult.DefaultWebSiteRolloverSize,$RemoteResult.ExchangeBackEndRolloverSize) -ne $RolloverSizeBytes)) -or 
       ($Encoding -eq 'UTF8' -and !($RemoteResult.Encoding)) -or ($Encoding -eq 'ANSI' -and $RemoteResult.Encoding)))
    {
        $Results += [PSCustomObject]@{
            Server   = $Server
            RootDirectory = $RemoteResult.RootDirectory
            RootLogFormat = $RemoteResult.RootLogFormat
            RootRollOverMethod = $RemoteResult.RootRolloverMethod
            RootRollOverSize = $RemoteResult.RootRolloverSize
            DefaultWebSiteDirectory = $RemoteResult.DefaultWebSiteDirectory
            DefaultWebSiteFormat = $RemoteResult.DefaultWebSiteLogFormat
            DefaultWebSiteRollOverMethod = $RemoteResult.DefaultWebSiteRolloverMethod
            DefaultWebSiteRollOverSize = $RemoteResult.DefaultWebSiteRolloverSize
            ExchangeBackEndDirectory = $RemoteResult.ExchangeBackEndDirectory
            ExchangeBackEndLogFormat = $RemoteResult.ExchangeBackEndLogFormat
            ExchangeBackEndRollOverMethod = $RemoteResult.ExchangeBackEndRolloverMethod
            ExchangeBackEndRollOverSize = $RemoteResult.ExchangeBackEndRolloverSize
            Encoding = if ($RemoteResult.Encoding) { 'UTF-8' } else { 'ANSI' }
        }
    }
}
Write-Progress -Activity "Getting IIS Logging Configuration is in progress..." -Completed

#Output
if ($Results)
{
    #$Results | Format-Table -Property @{n="Server";e={$_.Server};Alignment='Left'},@{n="RootDirectory";e={$_.RootDirectory};Alignment='Left'},@{n="RootRollOverMethod";e={$_.RootRollOverMethod};Alignment='Left'},@{n="RootRollOverSize";e={$_.RootRollOverSize};Alignment='Left'},@{n="DefaultWebSiteDirectory";e={$_.DefaultWebSiteDirectory};Alignment='Left'},@{n="DefaultWebSiteRollOverMethod";e={$_.DefaultWebSiteRollOverMethod};Alignment='Left'},@{n="DefaultWebSiteRollOverSize";e={$_.DefaultWebSiteRollOverSize};Alignment='Left'},@{n="ExchangeBackEndDirectory";e={$_.ExchangeBackEndDirectory};Alignment='Left' },@{n="ExchangeBackEndRollOverMethod";e={$_.ExchangeBackEndRollOverMethod};Alignment='Left'},@{n="ExchangeBackEndRollOverSize";e={$_.ExchangeBackEndRollOverSize};Alignment='Left'},@{n="Encoding";e={$_.Encoding};Alignment='Left'} -AutoSize
    $Results | Format-Table -Property @{n="Server";e={$_.Server};Alignment='Left'},@{n="DefaultWebSiteDirectory";e={$_.DefaultWebSiteDirectory};Alignment='Left'},@{n="DefaultWebSiteRollOverMethod";e={$_.DefaultWebSiteRollOverMethod};Alignment='Left'},@{n="DefaultWebSiteRollOverSize";e={$_.DefaultWebSiteRollOverSize};Alignment='Left'},@{n="ExchangeBackEndDirectory";e={$_.ExchangeBackEndDirectory};Alignment='Left' },@{n="ExchangeBackEndRollOverMethod";e={$_.ExchangeBackEndRollOverMethod};Alignment='Left'},@{n="ExchangeBackEndRollOverSize";e={$_.ExchangeBackEndRollOverSize};Alignment='Left'},@{n="Encoding";e={$_.Encoding};Alignment='Left'} -AutoSize
    $Results | Export-Csv -Path "$($ScriptPath)\iis_logsettings_$($now.tostring("yyyy_MM_dd-HH_mm_ss")).csv" -Delimiter ";" -NoTypeInformation -Encoding UTF8
}
else
{
    Write-Host "`nAll Exchange servers are configured with the expected IIS Log settings." -ForegroundColor Green    
}

#Set configuration
if ($Set -and $Results)
{
    $SetResults = @()

    $Servers = $Results.Server

    $ServersCount = $Servers.Count
    $Counter = 0

    Write-Host "IIS LogFile changes of $($Serverscount) server(s) are needed..." -ForegroundColor Yellow
    
    $Continue = Read-Host "Continue? ( Y / N ): "

    if ($Continue -in 'Y','y')
    {
        foreach ($Server in $Servers)
        {
            $Counter++

            Write-Progress -Activity "Setting IIS Logging Configuration is in progress..." -Status "Server $Counter of $ServersCount : $Server" -PercentComplete (($Counter / $ServersCount) * 100)
        
            $RemoteSetResult = Invoke-Command -ComputerName $Server.fqdn -ScriptBlock {
    
                Import-Module WebAdministration
            
                $Result = [ordered]@{}

                $Websites = @('siteDefaults','Default Web Site','Exchange Back End')

                foreach ($website in $Websites)
                {
                    if ($website -eq 'siteDefaults')
                    {
                        $filter = "system.applicationHost/sites/siteDefaults/logFile"
                        $prefix = 'Root'
                    }
                    else
                    {
                        $filter = "system.applicationHost/sites/site[@name='$website']/logFile"
                        $prefix = $website -replace '[^a-zA-Z0-9]', ''
                    }
                    
                    try
                    {
                        $SetLogSettings = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter $filter -Name 'directory' -Value $using:Directory -ErrorAction Stop
                        $Result["${Prefix}Directory"] = $using:Directory
                    }
                    catch
                    {
                        $Result["${Prefix}Directory"] = "Failure"
                    }
                    if ($using:ScheduleRollover)
                    {
                        try
                        {
                            $SetLogSettings = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter $filter -Name 'period' -Value $using:Schedule -ErrorAction Stop
                            $Result["${Prefix}RolloverMethod"] = $using:Schedule
                        }
                        catch
                        {
                            $Result["${Prefix}RolloverMethod"] = "Failure"
                        }
                    }
                    elseif ($using:MaxSizeRollover)
                    {
                        try
                        {
                            $SetLogSettings = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter $filter -Name 'period' -Value 'MaxSize' -ErrorAction Stop
                            $Result["${Prefix}RolloverMethod"] = "MaxSize"
                        }
                        catch
                        {
                            $Result["${Prefix}RolloverMethod"] = "Failure"
                        }
                        try
                        {
                            $SetLogSettings = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter $filter -Name 'truncateSize' -Value $using:RolloverSizeBytes -ErrorAction Stop
                            $Result["${Prefix}RolloverSize"] = $using:RolloverSizeBytes
                        }
                        catch
                        {
                            $Result["${Prefix}RolloverSize"] = "Failure"
                        }
                    }
                }
                if ($using:Encoding -eq 'UTF8')
                {
                    try
                    {
                        $SetEncoding = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.applicationHost/log' -Name 'logInUTF8' -Value $true -ErrorAction Stop
                        $Result["Encoding"] = "UTF-8"
                    }
                    catch
                    {
                        $Result["Encoding"] = "Failure"
                    }
                }
                else
                {
                    try
                    {
                        $SetEncoding = Set-WebConfigurationProperty -PSPath 'MACHINE/WEBROOT/APPHOST' -Filter 'system.applicationHost/log' -Name 'logInUTF8' -Value $false -ErrorAction Stop
                        $Result["Encoding"] = "ANSI"
                    }
                    catch
                    {
                        $Result["Encoding"] = "Failure"
                    }
                }    
                [PSCustomObject]$Result
                                
            }
                
            $SetResults += [PSCustomObject]@{
                Server   = $Server
                RootDirectory = $RemoteSetResult.RootDirectory
                RootLogFormat = $RemoteSetResult.RootLogFormat
                RootRollOverMethod = $RemoteSetResult.RootRolloverMethod
                RootRollOverSize = $RemoteSetResult.RootRolloverSize
                DefaultWebSiteDirectory = $RemoteSetResult.DefaultWebSiteDirectory
                DefaultWebSiteFormat = $RemoteSetResult.DefaultWebSiteLogFormat
                DefaultWebSiteRollOverMethod = $RemoteSetResult.DefaultWebSiteRolloverMethod
                DefaultWebSiteRollOverSize = $RemoteSetResult.DefaultWebSiteRolloverSize
                ExchangeBackEndDirectory = $RemoteSetResult.ExchangeBackEndDirectory
                ExchangeBackEndLogFormat = $RemoteSetResult.ExchangeBackEndLogFormat
                ExchangeBackEndRollOverMethod = $RemoteSetResult.ExchangeBackEndRolloverMethod
                ExchangeBackEndRollOverSize = $RemoteSetResult.ExchangeBackEndRolloverSize
                Encoding = $RemoteSetResult.Encoding
            }
        }
        Write-Progress -Activity "Setting IIS Logging Configuration is in progress..." -Completed
    
        #Output after SET Operation
        Write-Host "IIS logging configuration of the following server(s) was changed:" -ForegroundColor Yellow
        #$SetResults | Format-Table -Property @{n="Server";e={$_.Server};Alignment='Left'},@{n="RootDirectory";e={$_.RootDirectory};Alignment='Left'},@{n="RootRollOverMethod";e={$_.RootRollOverMethod};Alignment='Left'},@{n="RootRollOverSize";e={$_.RootRollOverSize};Alignment='Left'},@{n="DefaultWebSiteDirectory";e={$_.DefaultWebSiteDirectory};Alignment='Left'},@{n="DefaultWebSiteRollOverMethod";e={$_.DefaultWebSiteRollOverMethod};Alignment='Left'},@{n="DefaultWebSiteRollOverSize";e={$_.DefaultWebSiteRollOverSize};Alignment='Left'},@{n="ExchangeBackEndDirectory";e={$_.ExchangeBackEndDirectory};Alignment='Left' },@{n="ExchangeBackEndRollOverMethod";e={$_.ExchangeBackEndRollOverMethod};Alignment='Left'},@{n="ExchangeBackEndRollOverSize";e={$_.ExchangeBackEndRollOverSize};Alignment='Left'},@{n="Encoding";e={$_.Encoding};Alignment='Left'} -AutoSize
        $SetResults | Format-Table -Property @{n="Server";e={$_.Server};Alignment='Left'},@{n="DefaultWebSiteDirectory";e={$_.DefaultWebSiteDirectory};Alignment='Left'},@{n="DefaultWebSiteRollOverMethod";e={$_.DefaultWebSiteRollOverMethod};Alignment='Left'},@{n="DefaultWebSiteRollOverSize";e={$_.DefaultWebSiteRollOverSize};Alignment='Left'},@{n="ExchangeBackEndDirectory";e={$_.ExchangeBackEndDirectory};Alignment='Left' },@{n="ExchangeBackEndRollOverMethod";e={$_.ExchangeBackEndRollOverMethod};Alignment='Left'},@{n="ExchangeBackEndRollOverSize";e={$_.ExchangeBackEndRollOverSize};Alignment='Left'},@{n="Encoding";e={$_.Encoding};Alignment='Left'} -AutoSize

        $SetResults | Export-Csv -Path "$($ScriptPath)\iis_logsettings_changes_$($now.tostring("yyyy_MM_dd-HH_mm_ss")).csv" -Delimiter ";" -NoTypeInformation -Encoding UTF8
    }
    else
    {
        Write-Host "`nNotice: We made no changes." -ForegroundColor Green
    }
}