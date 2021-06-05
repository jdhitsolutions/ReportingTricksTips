#requires -version 5.1
#requires -module ServerManager


Function Get-FeatureInventory {
    [cmdletbinding()]
    [outputType("FeatureInventoryItem")]
    Param(
        [Parameter(Mandatory, Position = 0, HelpMessage = "Enter the name of a Windows server", ValueFromPipeline)]
        [string]$Computername,
        [Parameter(HelpMessage = "Enter an alternate credential for the remote computer")]
        [PSCredential]$Credential
    )
    Begin {
        Write-Verbose "[$((Get-Date).TimeofDay) BEGIN  ] Starting $($myinvocation.mycommand)"
        $ReportDate = Get-Date -Format g
    } #begin

    Process {
        Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Getting features from $($Computername.ToUpper())"
        Try {
            $features = Get-WindowsFeature @PSBoundParameters -ErrorAction Stop
        }
        Catch {
            Write-Warning "Failed to inventory features from $($Computername.toUpper()). $($_.exception.message)."
        }

        if ($features) {
            Write-Verbose "[$((Get-Date).TimeofDay) PROCESS] Discovered $($features.count) features."
            #create a custom object
            foreach ($feature in $features) {
                [PSCustomObject]@{
                    PSTypeName   = "FeatureInventoryItem"
                    Feature      = $feature.DisplayName
                    Name         = $Feature.Name
                    Installed    = $feature.installed
                    InstallState = $feature.InstallState
                    Description  = $feature.description
                    Version      = "{0}.{1}" -f $feature.AdditionalInfo.MajorVersion, $feature.AdditionalInfo.MinorVersion
                    Type         = $feature.FeatureType
                    Computername = $Computername.toUpper()
                    Report       = $ReportDate
                }
            }
        }

    } #process

    End {
        Write-Verbose "[$((Get-Date).TimeofDay) END    ] Ending $($myinvocation.mycommand)"
    } #end

} #close Get-FeatureInventory

#You could extend the type
$splat = @{
 Typename = 'FeatureInventoryItem'
 MemberType = 'NoteProperty'
 MemberName = 'Source'
 Value = $env:COMPUTERNAME
 Force = $True
}
Update-TypeData @splat

$splat.MemberName = 'ReportedBy'
$splat.Value = "$env:USERDOMAIN\$env:username"
Update-TypeData @splat

#load a custom format file
Update-FormatData $PSScriptRoot\FeatureInventory.format.ps1xml