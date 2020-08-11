<#
.SYNOPSIS
Utility for UTO Deskside (DS) to manage computer and other AD objects.
.DESCRIPTION
Provide various utilites for the UTO Deskside to quickly and accurately 
provide support to our customers.
#>

# setup required assembly. used mostly for inputboxes
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

$WORKING_DIR = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
$includes = @("global","DSForm","Invoke-Runas",'ShareReport','PrinterReport','GPOMigration','CreateGroup')

# itterate through include files and add them and 
ForEach ($include in $includes) {
            
    # check if the global include script is where it is expected
    if (Test-Path("${WORKING_DIR}\includes\${include}.ps1")) {
        # add the includes the popup scripts
        . "${WORKING_DIR}\includes\${include}.ps1"
        write-host "$include"
    } else {
        # Not there error out
        $wshell = New-Object -ComObject Wscript.Shell
        $wshell.Popup("Unable to find $include.ps1 include file",0,"Error!",0x0)
        exit
    }
}

# run the form
Start-Transcript -Path (Join-Path $Global:logdir "main-$((get-date).ToFileTime())") -IncludeInvocationHeader
$Global:SyncHash.window.ShowDialog() | out-null
Stop-Transcript