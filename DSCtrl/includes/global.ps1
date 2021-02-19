<#
.SYNOPSIS
Defines global functions & variables.
.DESCRIPTION
Defines global functions & variables for the DS Scripts
#>

# script version number
$globalVersion = "1.3.1"

# are we debug testing?
$globalDebug = $false
$global:debug = $globalDebug

# Configure the path for log files
$Global:logdir = Join-Path $env:localappdata '\ASU\DSCtrl\transcripts'
if ( (Test-Path -Path $Global:logdir) -ne $true) {
    New-Item -Path $Global:logdir -ItemType 'Directory' -Force
}

# Configure the path for backup files
$Global:backupdir = Join-Path $env:localappdata '\ASU\DSCtrl\backups'
if ( (Test-Path -Path $Global:backupdir) -ne $true) {
    New-Item -Path $Global:backupdir -ItemType 'Directory' -Force
}

# Configure settings file
$Global:settingsPath = Join-Path $env:localappdata '\ASU\DSCtrl\settings.json'
$Global:appSettings = [hashtable]::Synchronized(@{})
$Global:appSettings.add('path',$Global:settingsPath)
$Global:appSettings.add('settings',@{})
if (Test-Path -Path $Global:settingsPath) {
    $settingsdata = $null
    try {
        $settingsdata = Get-Content -Path $Global:settingsPath
        $settingsdata = ConvertFrom-Json -InputObject "${settingsdata}"
    } catch {
        $settingsdata = $null
    }

    if ($settingsdata -Is [PSObject]) {
        $settingsdata.psobject.properties | Foreach-Object {
            $Global:appSettings.settings.add($_.name,$_.value)
        }
    }
}
$Global:appSettings | Add-Member -Type ScriptMethod -name 'get' -Value {
    param($propName)
    $return = $null
    if ($this.settings.keys -icontains $propName) {
        $return = $this.settings."${propName}"
    }
    return $return
}
$Global:appSettings | Add-Member -Type ScriptMethod -name 'set' -Value {
    param($propName,$propValue)
    if ($this.settings.keys -contains $propName) {
        $this.settings."${propName}" = $propValue
    } else {
        $this.settings.add($propName,$propValue)
    }

    # Save the modified collection
    $this.settings | ConvertTo-Json | Set-Content -Path $this.path -Encoding 'utf8'
    return $null
}

# set the title of the main window
if ($globalDebug) {
    $globalTitle = "DSCTRL v$globalVersion !!!DEBUG MODE!!!"
} else {
    $globalTitle = "DSCTRL v$globalVersion"
}

# default OU path
$globalOUPath = 'OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu'
$global:OUPath = $globalOUPath

# Default AD server
$global:ADserver = ($ENV:LOGONSERVER -Replace '\\\\','')
if ( ($global:ADserver) -and ($global:ADserver.length -gt 6) ) {
    $global:ADserver = "${global:ADserver}.asurite.ad.asu.edu"
} else {
    $global:ADserver = 'asurite6.asurite.ad.asu.edu++'
}

# include popups scripts
. "$PSScriptRoot\popup.ps1"

function GlobalCheckModules {
    <#
    .SYNOPSIS
    Function checks to make sure required modules are installed
    .DESCRIPTION
    GlobalAddIncludes takes a list of include files to open. If it can not it will error out and exit.
    .PARAMETER modules
    The list of include modules to run the script
    .OUTPUTS
    None
    .EXAMPLE
    GlobalCheckModules $moduleList
    #>
  
    param ($modules)

    # itterate through modules files and add them and 
    ForEach ($module in $modules) {
        # test if the module is there
        if (-Not (Get-Module -ListAvailable -Name $module)) {
            # module not installed
            # error and exit
            $wshell = New-Object -ComObject Wscript.Shell
            $wshell.Popup("Required Module not installed '$module'. Exiting.",0,"Error!",0x0)
            exit
        }

    }
}


function GlobalLAPSApp {
    <#
    .SYNOPSIS
    Function runs provided application as UTOSPA with LAPS password
    .DESCRIPTION
    GlobalLAPSApp will attempt to run an application with as the remote computer's LAPS username and password.
    .PARAMETER computerName
    The name of the remote computer.
    .PARAMETER binaryPath
    The path of the binary to run
    .PARAMETER binaryArg
    The arguments for the binary
    .PARAMETER domainReq
    If True set the Domin switch on Invoke-Runas
    .OUTPUTS 
    None
    .EXAMPLE
    GlobalLAPSApp $computer $application $arguments $true
    #>
  
    param ([string]$computerName,[string]$binaryPath,[string]$binaryArg,[bool]$domainReq)

    # try to see if the computer exsits
    try{
        $Global:SyncHash.print("Checking if computer exists...", $false)
        Get-ADComputer -Identity $computerName -ErrorAction Stop
        $Global:SyncHash.print("yes", $false)
    } catch{
        # there was a problem getting the computer object. set error message
        $Global:SyncHash.print("&amp;#10;Error - The computer $computerName does not exsist. $($_.Exception.Message)", $false)
        return
    }
    
    $Global:SyncHash.print("&#10;Looking up password ...", $false)
    Write-Host ((Get-AdmPwdPassword -ComputerName $computerName).Password.Length)
    # is there a password set?    
    if (((Get-AdmPwdPassword -ComputerName $computerName).Password.Length) -gt 0) {
        # hopefully this is a valid password set
        # do we need to run with domain?
        $Global:SyncHash.print("Trying to run application ...", $false)
        if ($domainReq) {
            $domain = $computerName
        } else {
            $domain = '.'
        }
        if ($globalDebug) {
            $Global:SyncHash.print(" DEBUG: Running - Invoke-Runas -User 'UTOSPA' -Password (Get-AdmPwdPassword -ComputerName $computerName).Password -Domain $domain -Binary $binaryPath -Args $binaryArg -LogonType 0x2", $false)
        }
        Invoke-Runas -User "UTOSPA" -Computer $computerName -Domain $domain -Binary $binaryPath -Args $binaryArg -LogonType 0x2
    } else {
        # no password set error message
        if ($globalDebug) {
            $Global:SyncHash.print("Debug! trying to run App with provided cred.", $false)
            
            $printerQuestionText = 'Please enter UTOSPA password'

            # prompt for password
            $passwordQuestion = @(,($printerQuestionText,''))
            $passwordEntered = InputPopupBox $passwordQuestion
            
            # did they cancel?
            if (! ($passwordEntered.Status -like 'Cancel*') ) {

                # run application with default password
                if ($domainReq) {
                    Invoke-Runas -User "UTOSPA" -Computer ($printerEntered.Get_Item($printerQuestionText)) -Domain $computerName -Binary $binaryPath -Args $binaryArg -LogonType 0x2
                } else {
                    # no switch needed
                    Invoke-Runas -User "UTOSPA" -Computer ($printerEntered.Get_Item($printerQuestionText)) -Binary $binaryPath -Args $binaryArg -LogonType 0x2
                }
            }  
        } else {
            $Global:SyncHash.print("Error $computerName exisits in AD but could not find a password", $false)
        }
    }
}

function GlobalVarValidate {
    <#
    .SYNOPSIS
    Function checks to make sure that it matches it's expression
    .DESCRIPTION
    GlobalVarValidate checks to make sure that the variable provided matches the regular expresssion.
    .PARAMETER varChecking
    The string to check
    .PARAMETER regEx
    The regular expression to check variable to check
    .OUTPUTS
    True or False
    .EXAMPLE
     GlobalVarValidate $var '^([\w\.])+$'
    #>
  
    param ([string]$varChecking, [string]$regex)

    # test if the module is there
    if (($varChecking -imatch $regex)) {
        return $true
    } else {
        return $false
    }

}


# function from http://stackoverflow.com/questions/9999963/powershell-test-admin-rights-within-powershell-script
function GlobalIsAdmin {
    try {
        $identity = [Security.Principal.WindowsIdentity]::GetCurrent()
        $principal = New-Object Security.Principal.WindowsPrincipal -ArgumentList $identity
        return $principal.IsInRole( [Security.Principal.WindowsBuiltInRole]::Administrator )
    } catch {
        throw "Failed to determine if the current user has elevated privileges. The error was: '{0}'." -f $_
    }

    <#
        .SYNOPSIS
            Checks if the current Powershell instance is running with elevated privileges or not.
        .EXAMPLE
            PS C:\> Test-IsAdmin
        .OUTPUTS
            System.Boolean
                True if the current Powershell is elevated, false if not.
        .LINK
            http://stackoverflow.com/questions/9999963/powershell-test-admin-rights-within-powershell-script
    #>
}