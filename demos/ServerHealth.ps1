#requires -version 5.1
#requires -module ServerManager,CimCmdlets

<#
ServerHealth.ps1
A script to do a quick health check and create an HTML report for a Windows server.

usage: c:\scripts\serverhealth.ps1 -computer chi-dc04 -path c:\work\dc04-health.html

#>

[cmdletbinding()]
Param(
  [Parameter(Position = 0, Mandatory,HelpMessage = "Enter the name of a Windows-based server.")]
  [ValidateNotNullorEmpty()]
  [Alias("name", "cn")]
  [string]$Computername,
  [Parameter(Position = 1, HelpMessage = "Enter the filename and path for the html report,")]
  [ValidateNotNullorEmpty()]
  [string]$Path = "ServerHealth.html",
  [Parameter(HelpMessage = "Specify a credential object for the remote server.")]
  [Alias("RunAs")]
  [PSCredential]$Credential
)

#region Setup
Write-Verbose "Starting $($MyInvocation.MyCommand)"
#initialize an array for HTML fragments
$fragments = @()
$fragments += "<a href='javascript:toggleAll();' title='Click to toggle all sections'>+/-</a>"

$ReportTitle = "Server Health Report: $($Computername.toUpper())"

#this must be left justified
$head = @"
<Title>$ReportTitle</Title>
<style>
h2 {
width:95%;
background-color:#7BA7C7;
font-family:Tahoma;
font-size:10pt;
font-color:Black;
}
body { background-color:#FFFFFF;
       font-family:Tahoma;
       font-size:10pt; }
td, th { border:1px solid black;
         border-collapse:collapse; }
th { color:white;
     background-color:black; }
table, tr, td, th { padding: 2px; margin: 0px }
tr:nth-child(odd) {background-color: lightgray}
table { width:95%;margin-left:5px; margin-bottom:20px;}
.alert {background-color: red ! important}
.warn {background-color: yellow ! important}
</style>
<script type='text/javascript' src='https://ajax.googleapis.com/ajax/libs/jquery/1.4.4/jquery.min.js'>
</script>
<script type='text/javascript'>
function toggleDiv(divId) {
   `$("#"+divId).toggle();
}
function toggleAll() {
    var divs = document.getElementsByTagName('div');
    for (var i = 0; i < divs.length; i++) {
        var div = divs[i];
        `$("#"+div.id).toggle();
    }
}
</script>
<br>
<H1>$ReportTitle</H1>
"@

#build a hashtable of parameters for New-CimSession
$cimParams = @{
  ErrorAction   = "Stop"
  ErrorVariable = "myErr"
  Computername  = $Computername
}

if ($credential.username) {
  Write-Verbose "Adding a PSCredential for $($Credential.username)"
  $cimParams.Add("Credential", $Credential)
}

#create a CIM Session
Write-Verbose "Creating a CIM Session for $($Computername.toUpper())"
Try {
  $cs = New-CimSession @cimParams
}
Catch {
  Write-Warning "Failed to create CIM session for $($Computername.toUpper())"
  #bail out
  Return
}

#endregion

#region get OS data and uptime
Write-Verbose "Getting OS and uptime"
$os = $cs | Get-CimInstance -ClassName Win32_OperatingSystem
$Text = "Operating System"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

$fragments += $os | Select-Object @{Name = "Computername"; Expression = { $_.CSName } },
@{Name = "Operating System"; Expression = { $_.Caption } },
@{Name = "Service Pack"; Expression = { $_.CSDVersion } }, LastBootUpTime,
@{Name = "Uptime"; Expression = { (Get-Date) - $_.LastBootUpTime } } |
ConvertTo-Html -Fragment
$fragments += "</div>"

#endregion

#region Memory
Write-Verbose "Getting memory usage"
$Text = "Memory Usage"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

[xml]$html = $os |
Select-Object @{Name = "TotalMemoryMB"; Expression = { [int]($_.TotalVisibleMemorySize / 1mb) } },
@{Name = "FreeMemoryMB"; Expression = { [math]::Round($_.FreePhysicalMemory / 1MB, 2) } },
@{Name = "PercentFreeMemory"; Expression = { [math]::Round(($_.FreePhysicalMemory / $_.TotalVisibleMemorySize) * 100, 2) } },
@{Name = "TotalVirtualMemoryMB"; Expression = { [int]($_.TotalVirtualMemorySize / 1mb) } },
@{Name = "FreeVirtualMemoryMB"; Expression = { [math]::Round($_.FreeVirtualMemory / 1MB, 2) } },
@{Name = "PercentFreeVirtualMemory"; Expression = { [math]::Round(($_.FreeVirtualMemory / $_.TotalVirtualMemorySize) * 100, 2) } } |
ConvertTo-Html -Fragment

#parse html to add color attributes
for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
  $class = $html.CreateAttribute("class")
  #check the value of the percent free memory column and assign a class to the row
  if (($html.table.tr[$i].td[2] -as [double]) -le 10) {
    $class.value = "alert"
    $html.table.tr[$i].ChildNodes[2].Attributes.Append($class) | Out-Null
  }
  elseif (($html.table.tr[$i].td[2] -as [double]) -le 20) {
    $class.value = "warn"
    $html.table.tr[$i].ChildNodes[2].Attributes.Append($class) | Out-Null
  }
}
#add the new HTML to the fragment
$fragments += $html.innerXML
$fragments += "</div>"

#endregion

#region top processes

#get top 25 processes sorted by WorkinsetSize

Write-Verbose "Getting Process information"
$Text = "Top 25 Processes"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

$processes = $cs |
Get-CimInstance -ClassName win32_process -Filter "Name <> 'System' AND Name <>'System Idle Process'" |
Sort-Object WorkingSetSize -Descending |
Select-Object -Property ProcessID, Name, HandleCount,
@{Name = "WS(MB)"; Expression = { [math]::Round($_.WorkingSetSize / 1MB, 4) } },
@{Name = "VS(MB)"; Expression = { [math]::Round($_.VirtualSize / 1MB, 4) } },
CommandLine, CreationDate,
@{Name = "Runtime"; Expression = { (Get-Date) - $_.CreationDate } } -First 25

$fragments += $processes | ConvertTo-Html -Fragment
$fragments += "</div>"

#endregion

#region get disk drive status
Write-Verbose "Getting drive status"
$drives = $cs | Get-CimInstance -ClassName Win32_Logicaldisk -Filter "DriveType=3"
$Text = "Drive Utilization"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

[xml]$html = $drives | Select-Object DeviceID,
@{Name = "SizeGB"; Expression = { [int]($_.Size / 1GB) } },
@{Name = "FreeGB"; Expression = { [math]::Round($_.Freespace / 1GB, 4) } },
@{Name = "PercentFree"; Expression = { [math]::Round(($_.freespace / $_.size) * 100, 2) } } |
ConvertTo-Html -Fragment

#parse html to add color attributes
for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
  $class = $html.CreateAttribute("class")
  #check the value of the percent free column and assign a class to the row
  if (($html.table.tr[$i].td[3] -as [double]) -le 10) {
    $class.value = "alert"
    $html.table.tr[$i].ChildNodes[3].Attributes.Append($class) | Out-Null
  }
  elseif (($html.table.tr[$i].td[3] -as [double]) -le 20) {
    $class.value = "warn"
    $html.table.tr[$i].ChildNodes[3].Attributes.Append($class) | Out-Null
  }
}

$fragments += $html.InnerXml
$fragments += "</div>"

Clear-Variable html
#endregion

#region get recent errors and warning in all classic logs
Write-Verbose "Getting recent eventlog errors and warnings"

#any errors or audit failures in the last 24 hours will be displayed in red
#warnings in last 24 hours will be displayed in yellow
#This could be re-written to use Get-WinEvent

$Yesterday = (Get-Date).Date.AddDays(-1)
$after = [System.Management.ManagementDateTimeConverter]::ToDmtfDateTime($yesterday)

#get all event logs with entries
$logs = $cs | Get-CimInstance win32_ntEventlogFile -Filter "NumberOfRecords > 0"
#exclude security log
$Text = "Event Logs"
#$fragments+="<h2>$Text</h2>"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

#process security event log for Audit Failures
$fragments += "<h3>Security</h3>"

[xml]$html = $cs |
Get-CimInstance Win32_NTLogEvent -Filter "Logfile = 'Security' AND Type = 'FailureAudit' AND TimeGenerated >= ""$after""" |
Select-Object -Property TimeGenerated, Type, EventCode, SourceName, Message |
ConvertTo-Html -Fragment

if ($html.table) {
  #if a failure in the last day, display in red
  if (($html.table.tr[$i].td[1] -eq 'FailureAudit') -AND ([datetime]($html.table.tr[$i].td[0]) -gt $yesterday)) {
    $class.value = "alert"
    $html.table.tr[$i].Attributes.Append($class) | Out-Null
  }

  $fragments += $html.InnerXml
}
Else {
  #no recent audit failures
  Write-Verbose "No recent audit failures"
  $fragments += "<p style='color:green;'>No recent audit failures</p>"
}

Clear-Variable html

#process all the other logs
foreach ($log in ($logs | Where-Object logfilename -NE 'Security')) {
  Write-Verbose "Processing event log $($log.LogfileName)"
  $fragments += "<h3>$($log.LogfileName)</h3>"

  [xml]$html = $cs |
  Get-CimInstance Win32_NTLogEvent -Filter "Logfile = ""$($log.logfilename)"" AND Type <> 'Information' AND TimeGenerated >= ""$after""" |
  Select-Object -Property TimeGenerated, Type, EventCode, SourceName, Message |
  ConvertTo-Html -Fragment

  if ($html.table) {
    #color errors in red
    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
      $class = $html.CreateAttribute("class")
      #check the value of the entry type column and assign a class to the row if within the last day
      if (($html.table.tr[$i].td[1] -eq 'Error') -AND ([datetime]($html.table.tr[$i].td[0]) -gt $yesterday)) {
        $class.value = "alert"
        $html.table.tr[$i].Attributes.Append($class) | Out-Null
      }
      elseif (($html.table.tr[$i].td[1] -eq 'Warning') -AND ([datetime]($html.table.tr[$i].td[0]) -gt $yesterday)) {
        $class.value = "warn"
        $html.table.tr[$i].Attributes.Append($class) | Out-Null
      }
    } #for

    $fragments += $html.InnerXml
  }
  else {
    #no errors or warnings
    Write-Verbose "No recent errors or warnings for $($log.logfilename)"
    $fragments += "<p style='color:green;'>No recent errors or warnings</p>"
  }
  Clear-Variable html
} #foreach

$fragments += "</div>"

#endregion

#region get services that should be running but aren't
Write-Verbose "Getting services that should be running but aren't"
$services = $cs | Get-CimInstance -ClassName win32_service -Filter "startmode='Auto' AND state <> 'Running'"
$Text = "Services"
#$fragments+="<h2>$Text</h2>"
$div = $Text.Replace(" ", "_")
$fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

$fragments += $services | Select-Object Name, Displayname, Description, State |
ConvertTo-Html -Fragment
$fragments += "</div>"

#endregion

#region get installed features
Try {
  $featParam = @{
    Computername = $Computername
    ErrorAction  = "Stop"
  }

  if ($Credential.username) {
    $featParam.Add("Credential", $Credential)
  }
  Write-Verbose "Getting installed windows features"
  $features = Get-WindowsFeature @featParam |
  Where-Object installed | Sort-Object Name, Parent |
  Select-Object DisplayName, Name, Description

  $Text = "Installed Windows Features"
  #$fragments+="<h2>$Text</h2>"
  $div = $Text.Replace(" ", "_")
  $fragments += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><h2>$Text</h2></a><div id=""$div"">"

  $fragments += $features | ConvertTo-Html -Fragment
  $fragments += "</div>"

}
Catch {
  Write-Warning "Can't get windows features for $Computername. $($_.Exception.Message)."
}

#endregion

#region creating output
Write-Verbose "Adding footer"
$fragments += "<br><i>Created $(Get-Date)</i>"

#create the HTML report
Write-Verbose "Creating an HTML report"

ConvertTo-Html -Head $head -Title $reportTitle -Body $Fragments | Out-File -FilePath $path -Encoding ascii

Write-Verbose "Saving the HTML report to $Path"

Write-Host "Report saved to $path for $($computername.ToUpper())" -ForegroundColor Green

Write-Verbose "Ending $($MyInvocation.MyCommand)"

#endregion