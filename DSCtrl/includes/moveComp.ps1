<#

.SYNOPSIS
Moves a computer. 
.DESCRIPTION
Moves an exsiting computer to a new Unit/Department (OU/Security Group). It can also remove any software collections it may be a part of.
.INPUTS
None. You cannot pipe objects to MoveComp.ps1.
.OUTPUTS
None. MoveComp.ps1 does not generate any output.
.EXAMPLE
C:\PS> .\MoveComp.ps1
.EXAMPLE
C:\PS> .\Update-Month.ps1 
#>

# get path of the include script
$includes = "$PSScriptRoot\includes"

# check if the global include script is where it is expected
if (Test-Path("$includes\global.ps1")) {
    # add the includes the popup scripts
    . "$includes\global.ps1"
} else {
    # Not there error out
    $wshell = New-Object -ComObject Wscript.Shell
    $wshell.Popup("Unable to find global include file",0,"Error!",0x0)
    exit
}

# get all the OUs
$ouListingFull = Get-ADOrganizationalUnit -Filter 'Name -like "*"' -SearchBase $ouPath -SearchScope 1

# get rid of the operational OUs so that we only have Unit OUs
$ouListing = $ouListingFull | Where-Object {$_.Name -ne 'M.UTOSPA.Test'} | Where-Object {$_.Name -ne 'M.UTOSPA.Groups'} | Where-Object {$_.Name -ne 'M.UTOSPA.Staging'}

$ouSelected = ListSelect $ouListing 'name' 'Please select the Unit:'
write-host "-$ouSelected-"

if ($ouSelected -ne $null) {

    $ouPath = "OU=" + $ouSelected + ',' + $ouPath
    $ouPath
    $ouGroupListing = Get-ADOrganizationalUnit -Filter 'Name -like "*Groups"' -SearchBase $ouPath -SearchScope Subtree

    $ouSelected = ListSelect $ouGroupListing 'name' 'Please select an Department:'

    $ouPath = "OU=" + $ouSelected + ',' + $ouPath
    $ouPath

    Get-ADGroup -SearchBase $ouPath -filter {GroupCategory -eq "Security"}


    # TODO: need to handle empty folders

} else {  
    PopupBox 'Nothing selected. Exiting' 'Exiting!'
}