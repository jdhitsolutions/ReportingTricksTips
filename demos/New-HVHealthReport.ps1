#requires -version 5.1

# https://gist.github.com/jdhitsolutions/d6ec76a00525f18d87ca27d104ea00bd

<#
.SYNOPSIS
Create an HTML Hyper-V health report.
.DESCRIPTION
This command will create an HTML-based Hyper-V health report. It is designed to report on Hyper-V 3.0 or later servers or even Client Hyper-V on Windows 10. This script will retrieve data using PowerShell remoting from the Hyper-V Host. It is assumed you will run this from your desktop and specify a remote Hyper-V host. You do not need any Hyper-V tools installed locally to run this script.

The report only shows virtual machine information for any virtual machine that is not powered off. If you include performance counters, you will only get data on counters with a value other than 0.

Data from resource metering will only be available for running virtual machines with resource metering enabled.

If you don't specify a file name, the command will create a file in your Documents folder called HyperV-Health.htm. Be aware that the collapsible region feature may not work in all web browsers.
.PARAMETER Computername
The name of the Hyper-V server. You must have rights to administer the server.
.PARAMETER Credential
Specify an alternate administrator credential for the remote Hyper-V server
.PARAMETER RecentCreated
The number of days to check for recently created virtual machines.
.PARAMETER Hours
The number of hours to check for recent event log entries. The default is 24.
.PARAMETER LastUsed
The number of days to check for last used virtual machines. The default is 30.
.PARAMETER Performance
Specify if you do want Hyper-V performance counters in the report.
.PARAMETER Path
The filename and path for the HTML report file.
.PARAMETER Metering
Specify if you do want to include virtual machine resource metering in the report.
.PARAMETER NoEventLog
Skip gathering event log information.
.PARAMETER Logo
Specify the path to a graphic file to embed in the report. This should be a thumbnail size graphic.
.EXAMPLE
PS C:\Scripts> .\New-HVHealthReport.ps1 -computer HV01

Create a report for server HV01 with default values. The report will be saved locally in the documents folder as HyperV-Health.htm
.EXAMPLE
PS C:\Scripts> .\New-HVHealthReport.ps1 -computer HV01 -performance -metering -logo .\company.png -path c:\reports\HV01-Health.html

Create a report for server HV01 with default values including performance and resource meter data. The report will be saved locally in the C:\Reports folder.

.LINK
Get-VM
Get-VMHost
Get-VHD
Measure-VM
Get-CimInstance
Get-Counter
Get-Eventlog
.INPUTS
This command does not accept pipelined input.
.OUTPUTS
an HTML file
.NOTES
Last Updated : 4 December 2019
Version      : 4.0.1

Learn more about PowerShell:
http://jdhitsolutions.com/blog/essential-powershell-resources/

****************************************************************
* DO NOT USE IN A PRODUCTION ENVIRONMENT UNTIL YOU HAVE TESTED *
* THOROUGHLY IN A LAB ENVIRONMENT. USE AT YOUR OWN RISK.  IF   *
* YOU DO NOT UNDERSTAND WHAT THIS SCRIPT DOES OR HOW IT WORKS, *
* DO NOT USE IT OUTSIDE OF A SECURE, TEST SETTING.             *
****************************************************************
#>

[cmdletbinding()]

Param(
    [Parameter(Position = 0, HelpMessage = "The name of the Hyper-V server. You must have rights to administer the server. If necessary, specify an alternate credential.")]
    [ValidateNotNullorEmpty()]
    [Alias("CN")]
    [String]$Computername = [environment]::machinename,
    [PSCredential]$Credential,
    [Parameter(HelpMessage = "The path and filename for the HTML report.")]
    [ValidateNotNullorEmpty()]
    [ValidateScript( {
            if (Test-Path (Split-Path $_)) {
                $True
            }
            else {
                Throw "Can't validate part of the path $_"
            }
        })]
    [String]$Path = (Join-Path -path ([environment]::GetFolderPath("mydocuments")) -child "HyperV-Health.htm"),
    [Parameter(HelpMessage = "The number of days to check for recently created virtual machines.")]
    [ValidateScript({ $_ -ge 0 })]
    [int]$RecentCreated = 30,
    [Parameter(HelpMessage = "The number of days to check for last used virtual machines.")]
    [ValidateScript({ $_ -ge 0 })]
    [int]$LastUsed = 30,
    [Parameter(HelpMessage = "The number of hours to check for recent event log entries.")]
    [ValidateScript({ $_ -ge 0 })]
    [int]$Hours = 24,
    [Parameter(HelpMessage = "Specify if you want performance counters in the report.")]
    [switch]$Performance,
    [Parameter(HelpMessage = "Specify if you want resource metering in the report. This assumes you have enabled resource metering for the virtual machines.")]
    [switch]$Metering,
    [Parameter(HelpMessage = "Don't get any event log information. If you use this parameter, the Hours parameter will be ignored.")]
    [switch]$NoEventLog,
    [Parameter(HelpMessage = "Specify the path to a graphic file to use as a logo at the top of the report. A smaller graphic works best.")]
    [ValidateScript( { Test-Path $_ })]
    [ValidateNotNullOrEmpty()]
    [string]$Logo
)


$reportversion = "4.0.1"

#region initialize

<#
NOTE: All of the Hyper-V commands include the module name to avoid
any naming conflicts with cmdlets from VMware or System Center.
#>

#region define a scriptblock to run on the Hyper-V host to gather data

$datascriptblock = {
    Param(
        [string]$Computername = $env:computername,
        [int]$RecentCreated,
        [int]$LastUsed,
        [int]$Hours,
        [bool]$Performance,
        [bool]$Metering,
        [bool]$NoEventLog
    )
#region private helper functions

    function Get-VMLastUse {

        [cmdletbinding()]
        Param (
            [Parameter(Position = 0,
                HelpMessage = "Enter a Hyper-V virtual machine name",
                ValueFromPipeline, ValueFromPipelinebyPropertyName)]
            [ValidateNotNullorEmpty()]
            [alias("vm")]
            [object]$Name = "*"
        )

        Begin {
            Write-Verbose -Message "Starting $($MyInvocation.Mycommand)"

            #define a hashtable of parameters to splat to Get-VM
            $vmParams = @{
                ErrorAction = "Stop"
            }
        } #begin

        Process {
            if ($name -is [string]) {
                Write-Verbose -Message "Getting virtual machine(s)"
                $vmParams.Add("Name", $name)
                Try {
                    $vms = Hyper-V\Get-VM @vmParams
                }
                Catch {
                    Write-Warning "Failed to find a VM or VMs with a name like $name"
                    #bail out
                    Return
                }
            }
            elseif ($name -is [Microsoft.HyperV.PowerShell.VirtualMachine] ) {
                #otherwise we'll assume $Name is a virtual machine object
                Write-Verbose "Found one or more virtual machines matching the name"
                $vms = $name
            }
            else {
                #invalid object type
                Write-Error "The input object was invalid."
                #bail out
                return
            }
            foreach ($vm in $vms) {

                #if VM is on a remote machine using PowerShell remoting to get the information
                Write-Verbose "Processing $($vm.name)"
                $sb = {
                    param([string]$Path, [string]$vmname)
                    Try {
                        $diskfile = Get-Item -Path $Path -ErrorAction Stop
                        $diskFile | Select-Object @{Name = "LastUseTime"; Expression = { $diskFile.LastWriteTime } },
                        @{Name = "LastUseAge"; Expression = { (Get-Date) - $diskFile.LastWriteTime } }
                    }
                    Catch {
                        Write-Warning "$($vmname): Could not find $path."
                    }
                } #end scriptblock

                #get first drive file
                $diskpath = $vm.HardDrives[0].Path

                #only proceed if a hard drive path was found
                if ($diskpath) {
                    $icmParam = @{
                        ScriptBlock  = $sb
                        ArgumentList = @($diskpath, $vm.name)
                    }
                    Write-Verbose "Getting details for $(($icmParam.ArgumentList)[0])"
                    if ($vmParams.computername) {
                        $icmParam.Add("Session", $tmpSession)
                    }

                    $details = Invoke-Command @icmParam
                    #write a custom object to the pipeline
                    $objHash = [ordered]@{
                        VMName       = $vm.name
                        CreationTime = $vm.CreationTime
                        LastUseTime  = $details.LastUseTime
                        LastUseAge   = $details.LastUseAge
                    }

                    #if VM is running set the LastUseAge to 0:00:00
                    if ($vm.state -eq 'running') {
                        $objHash.LastUseAge = New-TimeSpan -hours 0
                    }

                    #write the object to the pipeline
                    New-Object -TypeName PSObject -Property $objHash

                } #if $diskpath
                Else {
                    Write-Warning "$($vm.name): No hard drives defined."
                }
            }#foreach
        } #process

        End {
            Write-Verbose -Message "Ending $($MyInvocation.Mycommand)"
        } #end

    } #end function

    function _getVMHost {
        [cmdletbinding()]
        Param([string]$Computername = $env:computername)

        Hyper-V\Get-VMHost -ComputerName $computername |
        Select-Object -property @{Name = "Name"; Expression = { $_.name.toUpper() } },
        @{Name = "Domain"; Expression = { $_.FullyQualifiedDomainName } },
        @{Name = "MemGB"; Expression = { $_.MemoryCapacity / 1GB -as [int] } },
        @{Name = "Max Migrations"; Expression = { $_.MaximumStorageMigrations } },
        @{Name = "Numa Spanning"; Expression = { $_.NumaSpanningEnabled } },
        @{Name = "IoV"; Expression = { $_.IoVSupport } },
        @{Name = "VHD Path"; Expression = { $_.VirtualHardDiskPath } },
        @{Name = "VM Path"; Expression = { $_.VirtualMachinePath } }
    }

    function _insertToggle {
        [cmdletbinding()]
        Param([string]$Text, [object[]]$Data, [string]$Heading = "H2", [switch]$NoConvert)

        $out = @()
        $div = $Text.Replace(" ", "_")
        $out += "<a href='javascript:toggleDiv(""$div"");' title='click to collapse or expand this section'><$Heading>$Text</$Heading></a><div id=""$div"">"
        if ($NoConvert) {
            $out += $Data
        }
        else {
            $out += $Data | ConvertTo-Html -Fragment
        }
        $out += "</div>"
        $out
    }

    Function _getVols {
        [cmdletbinding()]
        Param([string]$Computername = $env:computername)

        (Get-Volume -CimSession $computername).Where( { $_.drivetype -eq 'fixed' }) | Sort-Object -property DriveLetter | Select-Object -Property @{Name = "Drive"; Expression = {
                if ($_.DriveLetter) { $_.driveletter } else { "none" }
            }
        }, Path, HealthStatus,
        @{Name = "SizeGB"; Expression = { [math]::Round(($_.Size / 1gb), 2) } },
        @{Name = "FreeGB"; Expression = { [math]::Round(($_.SizeRemaining / 1gb), 4) } },
        @{Name = "PercentFree"; Expression = { [math]::Round((($_.SizeRemaining / $_.Size) * 100), 2) } }
    }
#endregion

    #parameters for Write-Progress
    $progParam = @{
        Activity        = "Hyper-V Health Report: $($computername.ToUpper())"
        Status          = "initializing"
        PercentComplete = 0
    }

    Write-Progress @progParam

    #initialize a variable for HTML fragments
    $fragments = @()
    $fragments += "<a href='javascript:toggleAll();' title='Click to toggle all sections'>+/-</a>"

#region get server information

    $progParam.Status = "Getting VM Host"
    $progParam.currentOperation = $computername
    Write-Progress @progParam

    $vmhost = _getVMHost
    $fragments += _insertToggle -Text "VM Host" -Data $vmhost

    $progParam.Status = "Getting Server information"
    $progParam.currentOperation = "Operating System"
    Write-Progress @progParam

    #some of these properties will be used for memory reporting later in the script
    $cimParams = @{
        ClassName    = 'Win32_OperatingSystem'
        ComputerName = $computername
        Property     = @('Caption', 'LastBootUptime', 'FreePhysicalMemory', 'FreeVirtualMemory', 'MaxProcessMemorySize', 'TotalVirtualMemorySize', 'TotalVisibleMemorySize')
    }
    $os = Get-CimInstance @cimParams

    $osdetail = $os | Select-Object -property @{Name = "OS"; Expression = { $_.caption } },
    LastBootUptime, @{Name = "Uptime"; Expression = { (Get-Date) - $_.LastBootUpTime } }

    $fragments += _insertToggle -Text "Operating System" -Data $osdetail

    $progparam.PercentComplete = 5
    $progParam.currentOperation = "Computer System"
    Write-Progress @progParam

    $cimParams.ClassName = 'Win32_ComputerSystem'
    $cimParams.Property = 'Manufacturer', 'Model', 'TotalPhysicalMemory', 'NumberOfLogicalProcessors', 'NumberofProcessors'
    $cs = Get-CimInstance @cimparams | Select-Object -property Manufacturer, Model, @{Name = "TotalMemoryGB"; Expression = { [int]($_.TotalPhysicalMemory / 1GB) } },
    NumberOfProcessors, NumberOfLogicalProcessors

    $fragments += _insertToggle -Text "Computer System" -Data $cs

#endregion

#region memory

    $mem = $os |
    Select-Object @{Name = "FreeGB"; Expression = { [math]::Round(($_.FreePhysicalMemory / 1MB), 2) } },
    @{Name = "TotalGB"; Expression = { [math]::Round(($_.TotalVisibleMemorySize / 1MB), 2) } },
    @{Name = "Percent Free"; Expression = { [math]::Round(($_.FreePhysicalMemory / $_.TotalVisibleMemorySize) * 100, 2) } },
    @{Name = "FreeVirtualGB"; Expression = { [math]::Round(($_.FreeVirtualMemory / 1MB), 2) } },
    @{Name = "TotalVirtualGB"; Expression = { [math]::Round(($_.TotalVirtualMemorySize / 1MB), 2) } }

    [xml]$html = $mem | ConvertTo-Html -fragment

    #check each row, skipping the TH header row
    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
        $class = $html.CreateAttribute("class")
        #check the value of the percent free MB column and assign a class to the row
        if (($html.table.tr[$i].td[2] -as [double]) -le 10) {
            $class.value = "memalert"
            [void]$html.table.tr[$i].ChildNodes[2].Attributes.Append($class)
        }
        elseif (($html.table.tr[$i].td[2] -as [double]) -le 20) {
            $class.value = "memwarn"
            [void]$html.table.tr[$i].ChildNodes[2].Attributes.Append($class)
        }
    }

    $fragments += _insertToggle -Text "Memory" -Data $html.innerXML -NoConvert

#endregion

#region network adapters
    $progParam.currentOperation = "Network Adapters"
    $progparam.PercentComplete = 10
    Write-Progress @progParam

    $netstats = Get-NetAdapterStatistics -CimSession $computername | Select-Object -property Name,
    @{Name = "RcvdUnicastMB"; Expression = { [math]::Round(($_.ReceivedUnicastBytes / 1MB), 2) } },
    @{Name = "SentUnicastMB"; Expression = { [math]::Round(($_.SentUnicastBytes / 1MB), 2) } },
    ReceivedUnicastPackets, SentUnicastPackets,
    ReceivedDiscardedPackets, OutboundDiscardedPackets

    $fragments += _insertToggle -Text "Network Adapters" -Data $netstats

#endregion

#region check disk space

    $progParam.Status = "Getting Server Details"
    $progParam.currentOperation = "checking volumes"
    $progparam.PercentComplete = 15
    Write-Progress @progParam

    $vols = _getVols
    [xml]$html = $vols | ConvertTo-Html -Fragment

    <#
        I don't know why, but I can't add attributes to two different nodes
        at the same time so we have to go through all the volumes once to
        look at health and then a second time to look at percent free space.
    #>

    #check each row, skipping the TH header row
    #add alert class if volume is not healthy
    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
        $class = $html.CreateAttribute("class")

        if ($html.table.tr[$i].td[2] -ne "Healthy") {
            $class.value = "alert"
            [void]$html.table.tr[$i].ChildNodes[2].Attributes.Append($class)
        }
        else {
            $class.value = "green"
            [void]$html.table.tr[$i].ChildNodes[2].Attributes.Append($class)
        }

    }
    #go through rows again and add class depending on % free space
    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
        $class = $html.CreateAttribute("class")

        if (($html.table.tr[$i].td[-1] -as [double]) -le 10) {
            $class.value = "memalert"
            [void]$html.table.tr[$i].ChildNodes[5].Attributes.Append($class)
        }
        elseif (($html.table.tr[$i].td[-1] -as [double]) -le 20) {
            $class.value = "memwarn"
            [void]$html.table.tr[$i].ChildNodes[5].Attributes.Append($class)
        }
    } #for

    $fragments += _insertToggle -Text "Volumes" -Data $html.InnerXml -NoConvert
#endregion

#region check services

    $progParam.currentOperation = "Checking Hyper-V Services"
    $progparam.PercentComplete = 20
    Write-Progress @progParam

    $cimParams.ClassName = "Win32_Service"
    $cimParams.Property = 'Name', 'Displayname', 'StartMode', 'State', 'Startname'
    $cimParams.filter = "name like 'vmi%' or name ='vmms'"

    $services = Get-CimInstance @cimParams | Select-Object $cimParams.Property

    [xml]$html = $services | ConvertTo-Html -Fragment
    #find stopped services and add Alert style
    for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
        $class = $html.CreateAttribute("class")
        #check the value of the State column and assign a class to the row
        if ($html.table.tr[$i].td[3] -eq 'running') {
            $class.value = "green"
            [void]$html.table.tr[$i].Attributes.Append($class)
        }
    }

    #add the revised html to the fragment
    $fragments += _insertToggle -Text "Services" -Data $html.InnerXml -NoConvert

#endregion

#region enum VM
    $progParam.Status = "Getting Virtual Machine information"
    $progParam.currentOperation = "Enumerating VMs"
    $progparam.PercentComplete = 25
    Write-Progress @progParam

    Try {
        #get all VMs that are not turned off
        $allVMs = Hyper-V\Get-VM -ErrorAction Stop
        $runningVMs = $allVMS | Where-Object State -ne 'off'
        $vmGroup = $runningVMs | Sort-Object -property State, Name | Group-Object -Property State | Sort-Object -property Count

        #define a set of properties to display for each VM

        #format memory values as MB
        $vmProps = "Name", "Uptime", "Status", "CPUUsage",
        @{Name = "MemAssignedMB"; Expression = { $_.MemoryAssigned / 1MB } },
        @{Name = "MemDemandMB"; Expression = { $_.MemoryDemand / 1MB } },
        "MemoryStatus",
        @{Name = "MemStartupMB"; Expression = { $_.MemoryStartup / 1MB } },
        @{Name = "MemMinimumMB"; Expression = { $_.MemoryMinimum / 1MB } },
        @{Name = "MemMaximumMB"; Expression = { $_.MemoryMaximum / 1MB } },
        "DynamicMemoryEnabled"

        $vmData = @()
        foreach ($item in $vmGroup) {

            [xml]$html = $item.Group | Select-Object $vmProps | ConvertTo-Html -Fragment

            $caption = $html.CreateElement("caption")
            [void]$html.table.AppendChild($caption)
            $html.table.caption = $item.Name

            for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
                $class = $html.CreateAttribute("class")
                #check the value of the MemoryStatus column and assign a class to the row
                if ($html.table.tr[$i].td[6] -eq "Low") {
                    $class.value = "memalert"
                    [void]$html.table.tr[$i].ChildNodes[6].Attributes.Append($class)
                }
                elseif ($html.table.tr[$i].td[6] -eq "Warning") {
                    $class.value = "memwarn"
                    [void]$html.table.tr[$i].ChildNodes[6].Attributes.Append($class)
                }
            } #for

            $vmdata += $html.InnerXml
        } #foreach

    } #try
    Catch {
        $vmdata += "<p style='color:red;'>No virtual machines detected</p>"
    }

#region created in the last 30 days
    $progParam.currentOperation = "Virtual Machines Created in last $RecentCreated Days"
    $progparam.PercentComplete = 28
    Write-Progress @progParam

    if ($allVMs) {
        $recent = ($allVMS).where( { $_.CreationTime -ge (Get-Date).AddDays(-$RecentCreated) }) |
        Select-Object -property Name, CreationTime, Notes
        if ($recent) {
            [xml]$html = $recent | ConvertTo-Html -Fragment
            $caption = $html.CreateElement("caption")
            [void]$html.table.AppendChild($caption)
            $html.table.caption = "Created in last $RecentCreated days"
            $vmdata += $html.InnerXml
        }
        else {
            $vmdata += "<table><caption>Created in last $RecentCreated days</caption><tr><td style='color:green'>No virtual machines created recently</td></tr></table>"
        }
    }
    else {
        $vmdata += "<p style='color:red;'>No virtual machines detected</p>"
    }

#endregion

#region last use
    $progParam.currentOperation = "Virtual Machines not used within last $LastUsed Days"
    $progparam.PercentComplete = 30
    Write-Progress @progParam

    $last = New-TimeSpan -Days $LastUsed
    $data = (Get-VMLastUse).where( { $_.lastuseage -gt $last }) | Sort-Object LastUseAge

    if ($data) {
        [xml]$html = $data | ConvertTo-Html -Fragment
        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "Not used in last $lastused days"
        $vmData += $html.InnerXml
    }
    else {
        $vmData += "<table><caption>Not used in last $lastused days</caption><tr><td style='color:green'>No unused virtual machines detected for the last $lastused days.</td></tr></table>"
    }
#endregion

#region Integrated Services Version
    $progParam.currentOperation = "Integrated Services Version"
    $progparam.PercentComplete = 35
    Write-Progress @progParam

    if ($runningVMs) {
        $isv = $runningVMs | Sort-Object -property IntegrationServicesVersion |
        Select-Object -property Name, IntegrationServicesVersion, @{Name = "Current"; Expression = {
                $test = (Hyper-V\Get-VMIntegrationService -VMName $_.Name ).Where( { $_.OperationalStatus -contains 'ProtocolMismatch' })

                if ($test.count -gt 0) {
                    $False
                }
                else {
                    $True
                }
            }
        }

        [xml]$html = $isv | ConvertTo-Html -Fragment
        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "Integration Services Version"

        1..($html.table.tr.count - 1) | ForEach-Object {
            #enumerate each TD
            $td = $html.table.tr[$_]

            #create a new class attribute
            $class = $html.CreateAttribute("class")

            if ($td.childnodes.item(2)."#text" -eq 'False') {
                $class.value = "alert"
            } #if critical

            #append the class
            [void]$td.childnodes.item(2).attributes.append($class)
        } #foreach

        $vmdata += $html.InnerXml
    }
    else {
        $vmdata += "<p style='color:red;'>No running virtual machines detected with integration services</p>"
    }
#endregion

#region Snapshots
    $progParam.currentOperation = "VM Snapshots"
    $progparam.PercentComplete = 38
    Write-Progress @progParam

    $sb = {
        Hyper-V\Get-VMSnapshot -VMName * | Select-Object -property VMName, Name,
        CreationTime, @{Name = "Age"; Expression = { (Get-Date) - $_.CreationTime } },
        SnapshotType,
        @{Name = "SizeMB"; Expression = {
                ($_.HardDrives | Get-Item | Measure-Object -Property length -sum).sum / 1MB
            }
        }
    }

    $snap = Invoke-Command -ScriptBlock $sb | Select-Object -property * -ExcludeProperty RunspaceID

    if ($snap) {
        [xml]$html = $snap | Select-Object -property * -exclude PS* | ConvertTo-Html -Fragment
        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "VM Snapshots"
        $vmdata += $html.InnerXml
    }
    else {
        $vmdata += "<p style='color:red;'>No snapshots detected</p>"
    }

#endregion

#endregion

#region VHD Utilization
    $progParam.currentOperation = "Analyzing Virtual Disks"
    $progparam.PercentComplete = 40
    Write-Progress @progParam

    $vmdata += "<h3>Virtual Disk Detail</h3>"

    if ($runningVMs) {
        $progParam.Status = "Getting Virtual Disk Detail"
        foreach ($vm in $runningVMs) {
            $progParam.currentOperation = $vm.name
            Write-Progress @progParam
            #get VHD details
            $vhdDetail = foreach ($drive in $vm.harddrives) {
                Try {
                    $detail = Hyper-V\Get-VHD -path $drive.path -ErrorAction Stop
                    $vhdHash = [ordered]@{
                        ControllerType     = $drive.ControllerType
                        ControllerNumber   = $drive.ControllerNumber
                        ControllerLocation = $drive.ControllerLocation
                        VHDFormat          = $detail.VHDFormat
                        VHDType            = $detail.VHDType
                        FileSizeMB         = [math]::Round(($detail.FileSize / 1MB), 2)
                        SizeMB             = [math]::Round(($detail.Size / 1MB), 2)
                        MinSizeMB          = [math]::Round(($detail.MinimumSize / 1MB), 2)
                        FragPercent        = $detail.FragmentationPercentage
                        Path               = $drive.path
                    }
                    New-Object -TypeName PSObject -Property $vhdhash
                } #try
                Catch {
                    $vmdata += "<p style='color:red'>$($_.Exception.Message)</p>"
                }
            } #foreach drive
            if ($vhdDetail) {
                [xml]$html = $vhdDetail | ConvertTo-Html -Fragment
                $caption = $html.CreateElement("caption")
                [void]$html.table.AppendChild($caption)
                $html.table.caption = $vm.Name
                $vmdata += $html.InnerXml
            }
        } #foreach vm
    }
    else {
        $vmdata += "<p style='color:red;'>No running virtual machines - no virtual disk files found</p>"
    }

#endregion

#region replication

    $progParam.currentOperation = "Analyzing Virtual Disks"
    $progparam.PercentComplete = 41
    Write-Progress @progParam

    $repl = Hyper-V\Get-VMReplication

    if ($repl) {
        [xml]$html = $repl |
        Select-Object -property Name, State, Health, Mode, PrimaryServer, ReplicaServer, LastReplicationTime | ConvertTo-Html -Fragment

        1..($html.table.tr.count - 1) | ForEach-Object {
            #enumerate each TD
            $td = $html.table.tr[$_]

            #create a new class attribute
            $class = $html.CreateAttribute("class")

            if ($td.childnodes.item(2)."#text" -eq 'Critical') {
                $class.value = "alert"
            } #if critical

            #append the class
            [void]$td.childnodes.item(2).attributes.append($class)

        } #foreach

        $vmdata += $html.Innerxml
    }
    else {
        $vmdata += "<p style='color:red;'>No VM replication configured.</p>"
    }

#endregion

#region Resource Metering
    if ($Metering) {
        $progParam.currentOperation = "Gathering Resource Metering Data"
        $progparam.PercentComplete = 43
        Write-Progress @progParam

#region Resource Pool
        $vmdata += "<h3>Resource Pool Metering</h3>"
        #turn off error handling. There might be some resource pool data for some
        #types
        $data = Hyper-V\Measure-VMResourcePool -name * -computer $computername -ErrorAction SilentlyContinue | Select-Object -property ResourcePoolname, AvgCPU, AvgRam, MinRam, MaxRam, TotalDisk,
        @{Name         = "NetworkInbound(M)";
            Expression = { ($_.NetworkMeteredTrafficReport | Where-Object direction -Eq 'inbound' | Measure-Object -property TotalTraffic -sum).Sum }
        }, MeteringDuration

        if ($data) {
            $vmdata += $data | ConvertTo-Html -Fragment
        }
        else {
            $vmdata += "<p style='color:red;'>No VM Resource Pool data found</p>"
        }
#endregion

#region VM metering

        $vmdata += "<h3>VM Resource Metering</h3>"

        if ($runningVMs) {
            $data = ($runningVMs).where( { $_.ResourceMeteringEnabled }) |
            ForEach-Object {
                Hyper-V\Measure-VM -name $_.vmname |
                Select-Object -property VMName, AvgCPU, AvgRAM, MinRam, MaxRam, TotalDisk,
                @{Name         = "NetworkInbound(M)";
                    Expression = { ($_.NetworkMeteredTrafficReport |
                            Where-Object direction -Eq 'inbound' | Measure-Object -property TotalTraffic -sum).Sum
                    }
                },
                @{Name         = "NetworkOutbound(M)";
                    Expression = { ($_.NetworkMeteredTrafficReport |
                            Where-Object direction -Eq 'outbound' | Measure-Object -property TotalTraffic -sum).Sum
                    }
                }, MeteringDuration
            } #foreach
            $vmdata += $data | ConvertTo-Html -Fragment
        }
        else {
            $vmdata += "<p style='color:red;'>No virtual machines detected</p>"
        }
    }

    #add Virtual Machines data
    $fragments += _insertToggle -Text "Virtual Machines" -Data $vmdata -NoConvert
#endregion
#endregion resource metering

#region check for recent event log errors and warnings
    if (-NOT $NoEventLog) {

        $progParam.currentOperation = "Checking System Event Log"
        $progparam.PercentComplete = 60
        Write-Progress @progParam

        #hashtable of parameters for Get-Eventlog
        $logParam = @{
            Computername = $Computername
            LogName      = "System"
            EntryType    = "Error", "Warning"
            After        = (Get-Date).AddHours(-$Hours)
        }
        $sysLog = Get-EventLog @logparam

        <#
            only get errors and warnings from these sources
            vmicheartbeat
            vmickvpexchange
            vmicrdv
            vmicshutdown
            vmictimesync
            vmicvss
        #>
        $progParam.currentOperation = "Checking Application Event log"
        $progparam.PercentComplete = 65
        Write-Progress @progParam

        $logParam.logName = "Application"

        $appLog = Get-EventLog @logparam -Source vmic*

        $LogData = @()
        $LogData += "<h3>System</h3>"

        if ($syslog) {
            $syslog | Group-Object -Property Source |
            Sort-Object -property Count -Descending | ForEach-Object {

                [xml]$html = $_.Group | Sort-Object -property TimeWritten -Descending |
                Select-Object -property TimeWritten, EntryType, InstanceID, Message |
                ConvertTo-Html -Fragment

                $caption = $html.CreateElement("caption")
                [void]$html.table.AppendChild($caption)
                $html.table.caption = $_.Name

                #find errors and add Alert style
                for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
                    $class = $html.CreateAttribute("class")
                    #check the value of the entry type column and assign a class to the row
                    if ($html.table.tr[$i].td[1] -eq 'error') {
                        $class.value = "alert"
                        [void]$html.table.tr[$i].Attributes.Append($class)
                    }
                } #for
                #add the revised html to the fragment
                $LogData += $html.InnerXml
            } #foreach
        } #if System entries
        else {
            $LogData += "<table></caption><tr><td style='color:green'>No relevant system errors or warnings found.</td></tr></table>"
        }
        $LogData += "<h3>Application</h3>"
        if ($applog) {
            $applog | Group-Object -Property Source |
            Sort-Object -property Count -Descending | ForEach-Object {
                $LogData += "<h4>$($_.Name)</h4>"
                [xml]$html = $_.Group | Sort-Object -property TimeWritten -Descending |
                Select-Object -property TimeWritten, EntryType, InstanceID, Message |
                ConvertTo-Html -Fragment

                $caption = $html.CreateElement("caption")
                [void]$html.table.AppendChild($caption)
                $html.table.caption = $_.Name

                #find errors and add Alert style
                for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
                    $class = $html.CreateAttribute("class")
                    #check the value of the entry type column and assign a class to the row
                    if ($html.table.tr[$i].td[1] -eq 'error') {
                        $class.value = "alert"
                        [void]$html.table.tr[$i].Attributes.Append($class)
                    }
                } #for
                #add the revised html to the fragment
                $LogData += $html.InnerXml
            } #foreach
        } #if
        else {
            $LogData += "<table></caption><tr><td style='color:green'>No relevant application errors or warnings found.</td></tr></table>"
        }

        #check operational logs
        $progParam.currentOperation = "Checking operational event logs"
        $progparam.PercentComplete = 68
        Write-Progress @progParam

        $LogData += "<h3>Operational Logs</h3>"

        #define a hash table of parameters to splat to Get-WinEvent
        $paramHash = @{
            ErrorAction   = "Stop"
            ErrorVariable = "MyErr"
            Computername  = $Computername
        }

        $start = (Get-Date).AddHours(-$hours)

        #construct a hash table for the -FilterHashTable parameter in Get-WinEvent
        $filter = @{
            Logname   = "Microsoft-Windows-Hyper-V*"
            Level     = 2, 3
            StartTime = $start
        }

        #add it to the parameter hash table
        $paramHash.Add("FilterHashTable", $filter)

        #search logs for errors and warnings
        Try {
            #add a property for each entry that translates the SID into
            #the account name
            #hash table of parameters for Get-WSManInstance
            $newHash = @{
                ResourceURI   = "wmicimv2/win32_SID"
                SelectorSet   = $null
                Computername  = $Computername
                ErrorAction   = "Stop"
                ErrorVariable = "myErr"
            }

            #Any remote server must have the firewall exception enabled for remote event log management
            $oplogs = Get-WinEvent @paramHash |
            Add-Member -MemberType ScriptProperty -Name Username -Value {
                Try {
                    #resolve the SID
                    $newHash.SelectorSet = @{SID = "$($this.userID)" }
                    $resolved = Get-WSManInstance @script:newhash
                }
                Catch {
                    Write-Warning $myerr.ErrorRecord
                }
                if ($resolved.accountname) {
                    #write the resolved name to the pipeline
                    "$($Resolved.ReferencedDomainName)\$($Resolved.Accountname)"
                }
                else {
                    #re-use the SID
                    $this.userID
                }
            } -PassThru

        }
        Catch {
            Write-Warning $MyErr.errorRecord
        }

        if ($oplogs) {
            $oplogs | Group-Object -Property Logname |
            Sort-Object -property Count -Descending | ForEach-Object {
                [xml]$html = $_.Group | Sort-Object -property TimeCreated -Descending |
                Select-Object -property TimeCreated, @{Name = "EntryType"; Expression = { $_.levelDisplayname } },
                ID, Username, Message | ConvertTo-Html -Fragment

                $caption = $html.CreateElement("caption")
                [void]$html.table.AppendChild($caption)
                $html.table.caption = $_.Name

                #find errors and add Alert style
                for ($i = 1; $i -le $html.table.tr.count - 1; $i++) {
                    $class = $html.CreateAttribute("class")
                    #check the value of the entry type column and assign a class to the row
                    if ($html.table.tr[$i].td[1] -eq 'error') {
                        $class.value = "alert"
                        [void]$html.table.tr[$i].Attributes.Append($class)
                    }
                } #for
                #add the revised html to the fragment
                $LogData += $html.InnerXml
            } #foreach

        }
        else {
            $LogData += "<table></caption><tr><td style='color:green'>No relevant application errors or warnings found.</td></tr></table>"
        }

        $fragments += _insertToggle -Text "Event Logs" -Data $LogData -NoConvert
    }
    else {
        Write-Verbose "Skipping event log queries"
    }
#endregion

#region get performance data
    if ($Performance) {
        $progParam.status = "Gathering Performance Data"
        $progparam.PercentComplete = 70
        $progParam.currentOperation = "..System"
        Write-Progress @progParam

        $PerfData = @()
        #system
        $ctrs = "\System\Processes", "\System\Threads", "\System\Processor Queue Length"
        $sysCounters = Get-Counter -counter $ctrs

        [xml]$html = ($sysCounters).CounterSamples |
        Select-Object -property Path, @{Name = "Value"; Expression = { $_.CookedValue } } | ConvertTo-Html -Fragment

        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "System"
        $PerfData += $html.InnerXml

        #memory
        $progParam.currentOperation = "..Memory"
        $progparam.PercentComplete = 72
        Write-Progress @progParam

        $ctrs = "\Memory\Page Faults/sec",
        "\Memory\% Committed Bytes In Use",
        "\Memory\Available MBytes"
        $memCounters = Get-Counter -counter $ctrs

        [xml]$html = ($memCounters).CounterSamples |
        Select-Object -property Path, @{Name = "Value"; Expression = { $_.CookedValue } } |
        ConvertTo-Html -Fragment

        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "Memory"
        $PerfData += $html.InnerXml

        #cpu
        $progParam.currentOperation = "..Processor"
        $progparam.PercentComplete = 75
        Write-Progress @progParam

        $ctrs = "\Processor(*)\% Processor Time"
        $procCounters = Get-Counter -counter $ctrs

        [xml]$html = ($procCounters).CounterSamples |
        Select-Object Path, @{Name = "Value"; Expression = { $_.CookedValue } } |
        ConvertTo-Html -Fragment

        $caption = $html.CreateElement("caption")
        [void]$html.table.AppendChild($caption)
        $html.table.caption = "Processor"
        $PerfData += $html.InnerXml

        #physicaldisk
        $progParam.currentOperation = "..PhysicalDisk"
        $progparam.PercentComplete = 77
        Write-Progress @progParam

        $ctrs = "\PhysicalDisk(*)\Current Disk Queue Length",
        "\PhysicalDisk(*)\Avg. Disk Queue Length",
        "\PhysicalDisk(*)\Avg. Disk Read Queue Length",
        "\PhysicalDisk(*)\Avg. Disk Write Queue Length",
        "\PhysicalDisk(*)\% Disk Time",
        "\PhysicalDisk(*)\% Disk Read Time",
        "\PhysicalDisk(*)\% Disk Write Time"

        Try {
            $diskCounters = Get-Counter -counter $ctrs  -ErrorAction Stop
            $data = ($diskCounters).CounterSamples | Where-Object CookedValue -gt 0
        }
        Catch {
            $PerfData += "<table><caption>$($counterset.CounterSetName)</caption><tr><td style='color:red'>$($_.Exception.Message)</td></tr></table>"
        }

        if ($data) {
            #non zero data found
            [xml]$html = $data |
            Select-Object -property Path, @{Name = "Value"; Expression = { $_.CookedValue } } |
            ConvertTo-Html -Fragment

            $caption = $html.CreateElement("caption")
            [void]$html.table.AppendChild($caption)
            $html.table.caption = "Physical Disk"
            $PerfData += $html.InnerXml
        }
        else {
            $PerfData += "<table><caption>$($counterset.CounterSetName)</caption><tr><td style='color:green'>No non-zero values for this counter set.</td></tr></table>"
        }
        #Hyper-V Perf counters
        $progParam.status = "Getting Hyper-V Performance Counters"
        $progparam.PercentComplete = 80
        Write-Progress @progParam

        $hvCounters = Get-Counter -ListSet Hyper-V* -ErrorAction SilentlyContinue

        if ($hvCounters) {
            $data = foreach ($counterset in $hvcounters) {
                $progParam.currentOperation = $counterset.countersetname
                Write-Progress @progParam

                #create reports for any counter with a value greater than 0
                try {
                    $data = (Get-Counter -Counter $counterset.counter  -ErrorAction Stop).CounterSamples |
                    Where-Object CookedValue -gt 0 |
                    Sort-Object -property Path | Select-Object -property Path, @{Name = "Value"; Expression = { $_.CookedValue } }
                    if ($data) {
                        [xml]$html = $data | ConvertTo-Html -Fragment
                        $caption = $html.CreateElement("caption")
                        [void]$html.table.AppendChild($caption)
                        $html.table.caption = $counterset.CounterSetName
                        $PerfData += $html.InnerXml
                    }
                    else {
                        $PerfData += "<table><caption>$($counterset.CounterSetName)</caption><tr><td style='color:green'>No non-zero values for this counter set.</td></tr></table>"
                    }
                } #try
                Catch {
                    $PerfData += "<table><caption>$($counterset.CounterSetName)</caption><tr><td style='color:red'>$($_.Exception.Message)</td></tr></table>"
                }
            }
        } #if hvcounters
        else {
            Write-Verbose "Could not find any Hyper-V performance counters."
            $PerfData += "<p style='color:red;'>No Hyper-V performance counters detected</p>"
        }

        $fragments += _insertToggle -Text "Performance Data" -Data $PerfData -NoConvert

    } #if not $NoPerformance

#endregion

    #write fragments as the result of this scriptblock
    $fragments

    $progParam.status = "Creating HTML Report"
    $progParam.currentOperation = "$using:Path"
    $progParam.percentcomplete = 95
    Write-Progress @progParam

} #close datascriptblock

#endregion

#region MAIN code
#run the scriptblock against the remote Hyper-V host
$icmParams = @{
    ScriptBlock  = $datascriptblock
    ComputerName = $Computername
    ArgumentList = @($computername, $RecentCreated, $LastUsed, $Hours, $Performance, $Metering, $NoEventLog)
}

if ($Credential) {
    $icmParams.Add("Credential", $Credential)
}
$html = Invoke-Command @icmParams

#endregion

#region create the local HTML report

$title = "$($Computername.ToUpper()) Hyper-V Health Report"
$head = @"
<Title>$($Title)</Title>
<style>
h2
{
width:95%;
background-color:#7BA7C7;
font-family:Tahoma;
font-size:12pt;
}
caption
{
background-color:#A9A9F5;
text-align:left;
font-weight:bold;
}
body
{
background-color:#FFFFFF;
font-family:Tahoma;
font-size:9pt;
}
td, th
{
border:1px solid black;
border-collapse:collapse;
}
th
{
color:black;
background-color:#F2F5A9;
}
table, tr, td, th
{
padding: 3px;
margin: 0px;
border-spacing:0;
}
table
{
width:95%;
margin-left:5px;
margin-bottom:20px;
}
tr:nth-child(odd) {background-color: lightgray}
.alert {color:red}
.green {color:green}
.memalert {background-color: red}
.memwarn {background-color: yellow}
a:link { color: black ; text-decoration: underline}
a:visited { color: black ; text-decoration: underline}
a:hover {color:yellow}
.footer {font-size:8pt;width:25%;}
.footer tr:nth-child(odd) {background-color: white}
.footer td,tr {
    border-collapse:collapse;
    padding:0px;
    border:none;
    }
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
$(if ($Logo) {
    #need to use different parameters based on PowerShell version
    if ($PSVersionTable.PSVersion.Major -gt 5) {
         $ImageBits = [Convert]::ToBase64String((Get-Content $Logo -asbyteStream))
    }
    else {
         $ImageBits = [Convert]::ToBase64String((Get-Content $Logo -Encoding Byte))
    }
    $ImageFile = Get-Item (Convert-Path $Logo)
    $ImageType = $ImageFile.Extension.Substring(1)
    $ImageHead = "<Img src='data:image/$ImageType;base64,$($ImageBits)' Alt='$($ImageFile.Name)' style='float:left' width='120' height='120' hspace=10>"
    $Imagehead
    })
<br><br>
<H1>$Title</H1>
<br><br><br>
"@

#HTML to display at the end of the report with metadata about where this report was generated
[xml]$meta = [pscustomobject]@{
    Date    = Get-Date
    Author  = "$env:USERDOMAIN\$env:username"
    Script  = $($myinvocation.mycommand).path
    Version = $reportVersion
    Source  = $($Env:COMPUTERNAME)
} | ConvertTo-Html -Fragment -as List

$class = $meta.CreateAttribute("class")
$meta.table.SetAttribute("class", "footer")
$footer = @"
<i>
$($meta.innerxml)
</i>
"@

$paramHash = @{
    Head        = $head
    Body        = $html
    Postcontent = $footer
}

ConvertTo-Html @paramHash | Out-File -FilePath $path -encoding ASCII

Write-Host "Report complete. Please see $(Resolve-Path $path)" -ForegroundColor Green

#endregion

<#
Change Log

v4.0.1
* code clean up and region restructuring

v4.0
* Revised to run over PowerShell remoting on the Hyper-V Host.
  This removes the requirement to have Hyper-V tools installed locally.
  It also allows this script to be run from PowerShell 7.
* Removed service pack values from operating system data
* Refactored code to use private functions, parenthetical expressions, and the Where() method
* Removed embedded logo graphic and added a parameter to specify a logo file
* Added a Credential parameter
* Added code to convert graphic files according to PowerShell version
* Updated the metadata footer and CSS style.

v3.1
* Modified CIM commands to only query specific properties.
* Expanded aliases

v3.0
* Changed LastUse to LastUseTime
* if no service pack set value none
* add replication information
* added parameter to skip event logs
* format vm memory as MB
* flag VMs needing integration services update
* better error handling for missing performance counters with Windows Server 2016

#>