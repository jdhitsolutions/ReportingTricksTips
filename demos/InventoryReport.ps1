#a reporting wrapper for Get-FeatureInventory

Using Namespace System.Collections.Generic

[cmdletbinding()]
Param(
    [Parameter(Mandatory, Position = 0, HelpMessage = "Enter the name of a Windows server")]
    [string[]]$Computername,
    [Parameter(HelpMessage = "Enter an alternate credential for the remote computer")]
    [PSCredential]$Credential,
    [Parameter(HelpMessage = "Specify the report path")]
    [string]$Path = ".\InventoryReport.html"
)

#dot source the function
. $PSScriptRoot\Get-FeatureInventory.ps1

#define some metadata
$info = [PSCustomObject]@{
    ReportDate    = Get-Date -Format g
    Source        = $env:COMPUTERNAME
    ScriptPath    = $($myinvocation.mycommand).path
    ScriptVersion = "1.1.0"
    Author        = "$env:USERDOMAIN\$env:USERNAME"
}

#initalize a collection
$data = [list[object]]::new()
#initialize a collection for html fragments
$fragments = [list[string]]::new()

$ReportTitle = "Windows Server Feature Inventory Report"

#this must be left justified
$head = @"
<!--
$(($info | Out-String).Trim())
-->
<Title>$ReportTitle</Title>
<style>
h2 {
    width: 95%;
    font-family: Tahoma;
    font-size: 12pt;
    padding-left: 10pt;
}
h3 {
    width: 95%;
    font-family: Tahoma;
    font-size: 10pt;
    padding-left: 10pt;
}
body {
    background-color: #FFFFFF;
    font-family: Tahoma;
    font-size: 10pt;
}
td,th {
    border: 1px solid black;
    border-collapse: collapse;
}
th {
    color: white;
    background-color: rgb(109, 18, 228);
}
table,tr,td,th {
    padding: 2px;
    margin: 0px
}
tr:nth-child(odd) {
    background-color: lightgray
}
table {
    width: 95%;
    margin-left: 5px;
    margin-bottom: 20px;
    font-family: Verdana, Geneva, Tahoma, sans-serif;
    font-size: 8pt;
}
}
.footer {
    font-size: 8pt;
    width: 25%;
}
.footer tr:nth-child(odd) {
    background-color: white
}
.footer td,
tr {
    border-collapse: collapse;
    padding: 0px;
    border: none;
}
</style><br>
<H1>$ReportTitle</H1>
"@

$params = @{}
if ($Credential.username) {
    $params.Add("Credential".Credential)
}

foreach ($computer in $computername) {
    $params["Computername"] = $computer
    $feat = Get-FeatureInventory @params
    if ($feat) {
        $data.AddRange($feat)
    }
}

$grouped = $data | Sort-Object -Property Computername, Feature | Group-Object -Property Computername
foreach ($item in $grouped) {
    $fragments.Add("<H1>$($item.Name)</H1>")
    $TypeGroup = $item.Group | Group-Object -Property Type
    foreach ($t in $TypeGroup) {
        $fragments.Add("<H2>$($t.Name)</H2>")
        $nested = $t.group | Group-Object -Property InstallState
        foreach ($n in $nested) {
            $fragments.Add("<H3>$($n.Name)</H3>")
            $n.group | Sort-Object -property Feature |
            Select-Object -Property Feature, Name, Description |
            ConvertTo-Html -Fragment -As table | ForEach-Object { $fragments.Add($_) }
        }
    } #foreach $t
} #foreach item

#insert an attribute into the HTML
[xml]$meta = $info | Convertto-HTML -Fragment -as List
$class = $meta.CreateAttribute("class")
$meta.table.SetAttribute("class", "footer")
$footer = @"
<i>
$($meta.innerxml)
</i>
"@

$fragments.Add($footer)

ConvertTo-Html -Head $head -Title $reportTitle -Body $Fragments |
Out-File -FilePath $Path -Encoding ascii

