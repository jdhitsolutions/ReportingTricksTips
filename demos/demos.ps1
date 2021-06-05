Return "This is a demo script file."

#region Start with a working command

Get-WindowsFeature -ComputerName srv1 | Tee-Object -Variable feat

#endregion

#region Think objects, not text

#don't assume based on default output
$feat | Get-Member

#get a sample
$feat | Where-Object installed | Select-Object -First 1 -Property * | Tee-Object -Variable sample

$sample | Select-Object DisplayName, Installed, InstallState, Description

#you might need to add information

$sample |
Select-Object DisplayName, Installed, InstallState, Description,
@{Name = "Computername"; Expression = { "SRV1" } },
@{Name = "Report"; Expression = { Get-Date -Format g } }

$rpt = $feat |
Select-Object DisplayName, Installed, InstallState, Description,
@{Name = "Computername"; Expression = { "SRV1" } },
@{Name = "Report"; Expression = { Get-Date -Format g } }

$rpt

#you could easily export or convert this data.

#endregion

#region Format to Fit

$ft = $rpt | Format-Table -GroupBy Computername -Property Report,DisplayName,Installed,InstallState
$ft

$fl = $rpt | Format-List -GroupBy Computername -Property Report,DisplayName,Description,Installed,InstallState
$fl

#saving to files
$fl | out-file .\srv1-feat.txt
notepad .\srv1-feat.txt

#watch your width
#default is 80

$fl | out-file .\srv1-feat2.txt -Width 65
notepad .\srv1-feat2.txt

#Add-Border
#Install-Module PSScriptTools
help Add-Border
Add-Border -Text SRV1

#Create a formatted here-string
$title = "Windows Server Feature Report [$($rpt[0].Report)]"
$out = @"
$($title.PadLeft($title.length*1.5,' '))

$((Add-Border SRV1 | Out-String).trim())
"@

$rpt | Where-Object installed | Sort-Object Displayname | ForEach-Object {
    #build a string
    $f = @"

    $($_.DisplayName)
    $("-" * $_.Displayname.length)
    $($_.Description)

"@
    $out += $f
}

$out
$out | Out-File .\t.txt

notepad .\t.txt
#you could print or email this.

#using ANSI
#ANSI doesn't work in the PowerShell ISE
help Show-ANSISequence

#view sequences]
Get-PSReadLineOption

"$([char]27)[91mFoo$([char]27)[0m"

$feat |
Select-Object @{Name="Feature";Expression = {
if ($_.Installed) {
    "$([char]27)[92m$($_.DisplayName)$([char]27)[0m"
}
else {
    "$([char]27)[91m$($_.DisplayName)$([char]27)[0m"
}
}}, Description,
@{Name = "Computername"; Expression = { "SRV1" } },
@{Name = "Report"; Expression = { Get-Date -Format g } } |
Format-List

#formatting files
# help New-PSFormatXML
# $feat[0] | New-PSFormatXML -path .\feature.format.ps1xml -ViewName Report -Properties DisplayName,Description,Installed,InstallState -Wrap
# the read-only errors can be ignored

psedit .\feature.format.ps1xml
Update-Format .\feature.format.ps1xml
$feat[0..20] | Format-Table -view Report

#but this isn't as rich as I'd like

#endregion

#region custom objects and formatting
#create a rich object. It is better to have too much information.
#You can always pare down with property sets and custom format files.
psedit .\Get-FeatureInventory.ps1

. .\Get-FeatureInventory.ps1
$inv = Get-FeatureInventory srv1
$inv[0..10] | Select-Object *
$inv | Get-Member

# $inv[0] | New-PSFormatXML -Path .\featureinventory.format.ps1xml -ViewName default -GroupBy Computername -Properties Feature,Installed,Description -wrap

$inv | more

#using a custom view
$inv | Sort-Object type| Format-Table -view featuretype | more

#endregion

#region Creating HTML Reports

#read the help!

#I decided to use splatting for ConvertTo-HTML to make it easier to read
$conv = @{
    Title = "Server Feature Inventory"
    PreContent = "<H1>SRV1</H1>"
    Property = "Feature","Installed","InstallState","Description"
    PostContent = "<H5>Report by $($inv[0].ReportedBy) on $($inv[0].Report)</H5>"
    CssUri = ".\Alternating.css"
}

$inv | Convertto-html @conv | Out-File .\basic-inventory.html

Invoke-Item .\basic-inventory.html

#create a wrapper
psedit .\InventoryReport.ps1
.\InventoryReport.ps1 -Computername srv2
Invoke-Item .\InventoryReport.html

#some other examples

psedit .\ServerHealth.ps1
.\ServerHealth.ps1 -Computername srv1 -Path .\srv1-health.html

Invoke-Item .\srv1-health.html

psedit .\New-HVHealthReport.ps1
Invoke-Item .\hv.html

#for other reporting ideas take a look at https://github.com/jdhitsolutions/ADReportingTools

Get-Command -module ADReportingTools
Get-ADSummary
#this should be run in the console
Show-DomainTree
Get-ADbranch "OU=IT,DC=Company,DC=pri"

#endregion

