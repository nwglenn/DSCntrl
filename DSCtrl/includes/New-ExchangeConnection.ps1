<#
.SYNOPSIS
Imports Exchange functions
.DESCRIPTION
Connects to the ASU on-prem Exchange servers via a PSRemote session and (by default) imports that session so the commands are available locally
.PARAMETER Credential
The name of the group to be interpreted
.PARAMETER UseImplicitCredentials
IMPLEMENTED BUT BROKEN! - DO NOT USE!
Attempts to use the local Kerberos token to authenticate, rather than require an explicit credentials object. This currently does not work, because the ASU Exchange servers are behind a load balancer that breaks the Kerberos chain of trust.
.PARAMETER NoImport
If set, skips importng the session once the connection has been established. This is handy if you plan on sending the PSSession to a child process, or want to invoke the commands through the session itself.
.OUTPUTS
On success, an object containing a reference the the PSSession
in all other cases, returns null
.EXAMPLE
ConvertFrom-GroupName -GroupName 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'
.NOTES
This script was written to integrate with the Exchange environment at Arizona State University (ASU). It is not intended for general use outside of the university.
#>
function New-ExchangeConnection {
    [CmdletBinding(DefaultParameterSetName = 'withCreds')]
    param (
        [Parameter(ParameterSetName = 'withCreds',
            Mandatory = $True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = 'Credentials object for connecting to Exchange')]
        [Alias('Creds')]
        [PSCredential]$Credential,

        [Parameter(ParameterSetName = 'noCreds',
            Mandatory = $True,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = 'Credentials object for connecting to Exchange')]
        [Alias('Implicit')]
        [switch]$UseImplicitCredentials,

        [Parameter(Mandatory = $False,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = 'Credentials object for connecting to Exchange')]
        [Alias('NI')]
        [switch]$NoImport
    )
    $session = $null
    $args = @{
        ConfigurationName = 'Microsoft.Exchange'
        ConnectionUri     = 'https://mail.asu.edu/Powershell/'
        Authentication    = 'Basic'
        AllowRedirection  = $True
    }

    if ($UseImplicitCredentials) {
        $args.Authentication = 'NegotiateWithImplicitCredential'
    }
    else {
        $args.add('Credential', $Credential)
        $args.Authentication = 'Basic'
    }

    try {
        $Session = New-PSSession @args
        if (!$NoImport) {
            Import-PSSession $Session | Out-Null
        }
    } catch {
        Write-Error $_
        $session = $null
    }
    return $session
}