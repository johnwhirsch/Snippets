<#
    .SYNOPSIS
        This script/function will convert a Tenable.sc Vulnerability Detail report to a list of remediation instructions & links.
    .DESCRIPTION
        The report will seperate vulnerabilities by host and then give you a summary of vulnerabilities and detailed list of remediation instructions.
        The script will save a HTML copy of the report in the same directory as the Create-VulnerabilityReport.ps1 (will be change to My Documents\Vulnerability Reports).
        The script will also copy the HTML output of the report to the clipboard to be pasted into a rich text ticketing system.
    .EXAMPLE
        Convert report and use the original CSV file name for the report name:
            Create-VulnerabilityReport -TenableCSV "C:\Path to\Vulnerability Detailed Report.csv"

        Convert report and specify name for the report name:
            Create-VulnerabilityReport -TenableCSV "C:\Path to\Vulnerability Detailed Report.csv" -ReportName "ServerName Vulnerabilities To Address"
#>


function Create-VulnerabilityReport {
    
    param([Parameter(Mandatory=$True)][string]$TenableCSV,[string]$ReportName = (Split-Path $TenableCSV -leaf).replace(".csv","") )

    $TenableResults = Import-Csv $TenableCSV;

    $TenableResults | % { $_.'Vulnerability Priority Rating' = [int]$_.'Vulnerability Priority Rating' }

    $TenableVulnerabilities = $TenableResults | ? { $_.Severity -ne "Info" } `
    | Select-Object Plugin, "Severity", "Vulnerability Priority Rating", "Plugin Name", "DNS Name", "Synopsis", "Solution", "See Also", "First Discovered", "Exploit Ease", "Check Type" `
    | Sort-Object "Vulnerability Priority Rating" -Descending

    $Hosts = $TenableVulnerabilities | Select-Object "DNS Name" -Unique

    $header = @();

    $header += "<style>";
    $header += "body { font-family: Arial, Helvetica, sans-serif; font-size: 13px; }";
    $header += "</style>";

    $body = @();

    $body += "<div>Report Creation Date: $(Get-Date -DisplayHint Date)</div>";

    foreach($HostDNSname in $Hosts){

        $HostVulnerabilities = $TenableVulnerabilities | ? { $_.'DNS Name' -eq $HostDNSname.'DNS Name' }

        $HostInfo = $HostVulnerabilities | Select-Object -First 1
    

        # --------------------------------------------------------------
        # Begin Report Summary section
        # Contains all vulnerabilities listed in the report for current system
        # --------------------------------------------------------------

        $body += "<H1>Vulnerabilities for $($HostInfo.'DNS Name')</H1>";
        $body += "<H2 style=`"margin: 15px 0px 0px; padding: 0px`">Report Summary:</H2>";
        $body += "<table width=`"100%; style=`"border-bottom: 1px solid #000; padding: 10px 0px; margin: 20px 0px; display:block;`"`"><tbody><tr><th style=`"text-align:left;`">Plugin Name</th><th style=`"text-align:left;`">Severity</th><th style=`"text-align:left;`">VPR</th></tr>"
    
        foreach($HostVulnerability in $HostVulnerabilities){ #Summary
            switch($HostVulnerability.'Vulnerability Priority Rating'){
                { $_ -ge 8 }{$bgcolor = "red"} 
                { $_ -ge 5 -and $_ -lt 8 }{$bgcolor = "orange"}
                default {$bgcolor = "yellow"}
            } 
            $body += "<tr style=`"background: $($bgcolor);`"><td><a href=`"#$([uri]::EscapeDataString($HostVulnerability.'Plugin Name'))`" style=`"color: black;`">$($HostVulnerability.'Plugin Name')</td><td>$($HostVulnerability.Severity)</td><td>$($HostVulnerability.'Vulnerability Priority Rating')</td></tr>"
        } #EO ($HostVulnerability in $HostVulnerabilities) summary
    
        $body += "</tbody></table>";


        # --------------------------------------------------------------
        # Begin Remediation section
        # Contains all remediation solutinos for the currrent system
        # --------------------------------------------------------------

        $body += "<H2 style=`"margin: 15px 0px 0px; padding: 0px;`">Remediation Guide:</H2>";
        $body += "<table><tbody style=`"width=100%;`">"

        foreach($HostVulnerability in $HostVulnerabilities){
                
            $body += "<tr style=`"width=100%;`"><td style=`"border-bottom: 1px dotted #000; padding: 15px 0px; width=100%;`">";
            $body += "<a name=`"$([uri]::EscapeDataString($HostVulnerability.'Plugin Name'))`"></a>";
            $body += "<table><tbody style=`"width=100%;`">";
            $body += "<tr><td style=`"padding: 0px 0px 5px;`"><H3 style=`"display:inline`">$($HostVulnerability.'Plugin Name')</H3></td></tr>";
            $body += "<tr><td style=`"padding: 0px 0px 5px;`"><b>Synopsis:</b> $($HostVulnerability.Synopsis)</td></tr>";
            $body += "<tr><td style=`"padding: 0px 0px 5px;`"><b>Tenable Documentation:</b> https://www.tenable.com/plugins/nessus/$($HostVulnerability.Plugin) </td></tr>"
            $body += "<tr><td style=`"padding: 0px 0px 5px;`"><b>Solution:</b><br>";        
            $body += "$($HostVulnerability.Solution)";
            $body += "</td></tr>";
               
            $SeeAlso = $($HostVulnerability.'See Also').Split([Environment]::NewLine)
            if($SeeAlso.Count -gt 1){
                $body += "<tr><td><b>See Also:</b><br>"
                $SeeAlso | % { $SeeAlsoURL = $_.Trim(); $body += "$($SeeAlsoURL) <br>" }
                $body += "</td></tr>";
            }

            $body += "</tbody></table>"; # EO Vulnerability details table
            $body += "</td></tr>"; # EO row in remediation table

        } # EO foreach($HostVulnerability in $HostVulnerabilities) Remediation

        $body += "</tbody></table>"; # EO remediation table

    } # EO foreach($HostDNSname in $Hosts)


    # --------------------------------------------------------------
    # Begin report export
    # This section finds all URLs in the body and turns them into links instead of just plaintext
    # After URLs are converted, the report is exported to a HTML file in the same folder as the script
    # The report is also copied to the users clipboard so it can be pasted into a new incident
    # --------------------------------------------------------------


    [regex]$URLPattern = "(http)(s)?(:\/\/)([^\s,]+)"

    $URLs = $URLPattern.Matches($($body)).Value

    $URLs | Select-Object -Unique | % { $body = $body -Replace [regex]::escape($_),"<a href=`"$($_)`" target=`"_blank`">$($_)</a>" }

    $report = ConvertTo-Html -Body $body -Title "Vulnerability Remediation Instructions" -Head $header    

    $ReportName = $ReportName -replace " ", "_";
    $ReportName = $ReportName -replace ".html", "";
    $ReportName = $ReportName -replace ".htm", "";
    $ReportFilePath = "$($PSScriptRoot)\$($ReportName).html";

    $report | Out-File $ReportFilePath
    
    Set-Clipboard $report -AsHtml

    & 'C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe' @($ReportFilePath)

    Add-Type -AssemblyName System.Windows.Forms 
    $global:balloon = New-Object System.Windows.Forms.NotifyIcon
    $path = (Get-Process -id $pid).Path
    $balloon.Icon = [System.Drawing.Icon]::ExtractAssociatedIcon($path) 
    $balloon.BalloonTipIcon = [System.Windows.Forms.ToolTipIcon]::Info 
    $balloon.BalloonTipTitle = "Report Copied to Clipboard" 
    $balloon.BalloonTipText = "You can now paste the HTML report directly into the incident background box"
    $balloon.Visible = $true 
    $balloon.ShowBalloonTip(5000)

 }
