<#
.SYNOPSIS
    Generate dynamic html report for WSUS and email
.DESCRIPTION
    This script will generate a list of servers managed by WSUS, create and email a dynamic htm report.
.EXAMPLE
    Run this PS Script on WSUS Server
.INPUTS
    Reads from WSUS Server
.OUTPUTS
    HTML report saved to C:\Reports folder
.NOTES
    9/16/2021 Mark Tellier

    This script runs on WSUS and produces a beautiful dynamic html report for WSUS Security updates.
    However, sending the report to outlook, the embedded report looses some of it's flash.

    LastSyncTime is reported in UTC time, the script converts to local time.
#>

# --- LOAD REQUIRED MODULES
If ( ! (Get-module PoshWSUS )) {
    Import-Module PoshWSUS
}

# --- MAIN VARIABLES
$report = @()
$lineNo = 0
$WSUShost = $env:COMPUTERNAME.ToUpper()
$today = get-date -Format "MM-dd-yyyy"

# --- HTML REPORT
$mainTitle = '<h1>Security Update Report - ACME Infrastructure</h1>'
$footer = "<footer>WSUS Server $WSUShost generated this report on $today</footer>"
$reportOutput = "C:\Reports\WSUS_Report.html"

# --- EMAIL VARIABLES
$smtpServer = "10.10.10.10"
$smtpFrom = "noreploy@acme.com"
$smtpTo = "The Dude <dude@acme.com>"
$smtpSubject = "WSUS ACME Report"

# --- HTML STYLES
$htmlParams = @"
<style>
    * {
        font-family: Arial, Helvetica, sans-serif;
    }
    body {
        background-color: white;
        font-family: Tahoma;
        font-size: 10px;
    }

    footer {
        padding: 20px 0px;
        width: 100%;
        text-align: center;
        font-size: 1.2em;
    }

    /* Table border */
    table {
        margin: auto;
        font-family: Segoe UI;
        font-size: 1.2em;
        box-shadow: 2px 2px 4px #888;
        border: thin ridge grey;
        width: 90%;
    }
    
    /* Table Header */
    th { 
        text-align: center;
        font-family: verdana;
        background: #1565C0;
        color: #FFFFFF;
        max-width: 400px;
        padding: 5px 10px;
        text-align: center;
    }
    
    /* Table cell */
    td {
        font-size: 1.2em;
        padding: 5px 10px;
        color: #000;
    }

    /* Table Row, data color definition */
    tr { background: #b8d1f3; }

    /* Table Odd and Even Row Color definition */
    tr:nth-child(even) {
        background: #f8fafd; }
    tr:nth-child(odd) {
        background: #dae5f4; }

    /* Table Row hover */
    tr:hover { 
    background-color: pink; }

    h1 {
        font-family: Georgia;
        font-size: 4em;
        text-align: center;
        line-height: 2;
        color: #212121;
        text-shadow: 2px 2px 2px #9E9E9E;
    }
    h2 {
        font-size: 3em;
    }
    h3 {
        font-size: 2.5em;
    }

</style>
"@

# --- READ INFO FROM WSUS
Connect-PSWSUSServer -WsusServer $WSUShost -Port 8530
$clientList = Get-PSWSUSUpdateSummaryPerClient | Sort-Object Computer

# --- PROCESS DATA
foreach ( $client in $clientList ) {

    $extInfo = Get-PSWSUSClient -ComputerName $client.Computer
    $lineNo++

    $items = [ordered]@{

        "Index"             = $lineNo
        "Computer"          = $client.Computer.tolower()
        "IP Address"        = $extInfo.IPAddress
        "Operating System"  = $extInfo.OSDescription
        "Last Sync"         = $extInfo.LastSyncTime.tolocaltime()
        "Last Result"       = $extInfo.LastSyncResult
        "Needed"            = $client.NeededCount
        "Downloaded"        = $client.DownloadedCount
        "Failed"            = $client.FailedCount
        "Installed"         = $client.InstalledCount 
        "Pending Reboot"    = $client.PendingReboot
        
    }

    $report += New-Object -TypeName psobject -Property $items

}

# --- OUTPUT TO SCREEN
$report | ConvertTo-Html -Head $htmlParams -PreContent $mainTitle -PostContent $footer | Out-File -FilePath $reportOutput
Invoke-Item $reportOutput

# --- EMAIL REPORT
$sslreport = $report | ConvertTo-Html -Head $htmlParams -PreContent $mainTitle -PostContent $footer | Out-String

Send-MailMessage `
    -SmtpServer $smtpServer `
    -From $smtpFrom `
    -To $smtpTo `
    -Subject $smtpSubject `
    -BodyAsHtml `
    -Body $sslReport `
    -Attachments $reportOutput
