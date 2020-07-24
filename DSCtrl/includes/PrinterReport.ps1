<#
.SYNOPSIS
Creates a report of the permissions of printers
.DESCRIPTION
Based on the print server and printer prefix that is provided a HTML report is created that 
will display the permission and the driver type on the printers.
.INPUTS
None. You cannot pipe objects to PrinterReport.ps1.
.OUTPUTS
None. PrinterReport.ps1 does not generate any output.
.EXAMPLE
PrinterReport
#>

#. "$PSScriptRoot\popup.ps1"

function PrinterReport {
    
    # get the path to search
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    # set the types of security groups
    $printServers = @([pscustomobject]@{name="ASUPRINT1"},
        [pscustomobject]@{name="DPCPRINT1"},
        [pscustomobject]@{name="POLYPRINT1"},
        [pscustomobject]@{name="WESTPRINT1"})
    
    # get the print server the printer(s) are on
    $printServersSelected = PopupListSelect $printServers 'name' 'Select the print server:'
    
    # did they cancel?
    if ($printServersSelected -eq $null) {
        PopupBox 'Canceled. Exiting.' "Information" "ok"
        return
    }
    
    $printerQuestionText = "Prefix of printers or provide list seperated by commas."

    Do {
        # get the printer name
        $printerQuestion = @(,($printerQuestionText,''))
        $printerEntered = InputPopupBox $printerQuestion
                        
        # did they cancel?
        if ($printerEntered.Status -like 'Cancel*') {
            PopupBox 'Canceled. Exiting.' "Information" "ok"
            return
        }
        
    } Until ( ($printerEntered.Get_Item($printerQuestionText) -ne $null) -and ($printerEntered.Get_Item($printerQuestionText) -ne '') )

    $printerList = $printerEntered.Get_Item($printerQuestionText) -split ','

    # start the html
    $html = "<html><style>div {padding:20px;}; li {font-weight:bold};</style><body><h1>Printer Permissions Report - $($printerEntered.Get_Item($printerQuestionText))</h1>"

    # display save as dialog to save the HTML file
    $saveDialog = New-Object windows.forms.savefiledialog
    $saveDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
    $saveDialog.title = "Save Printer Report"
    $saveDialog.filter = "HTML|*.html|All Files|*.*"
    $result = $saveDialog.ShowDialog()
    if($result -ne "OK") {
        # canceled so stop
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }

    $continue =  PopupBox 'Please be aware that this may take some time depending on the number of printers.' 'Continue?' 'oc' 

    if ($continue -eq 'Cancel') {
        # canceled so stop
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }
<#
    # create a variable to hold errors while scanning printers.
    $printerErrors = @()

    # traverse the printer Path.
    $printerFolders = Get-ChildItem -Recurse $printerPath -ErrorAction SilentlyContinue -ErrorVariable +printerErrors | where-object {($_.PsIsContainer)} | Sort-Object FullName

    # go through all the errors
    foreach ($printerError in $printerErrors) {

        $html += "<div style='background-color:#ffcccc;border-left:6px solid red;'><h2>Unable to access:</h2>"

        # make sure that it's not an error about the root
        if ($printerError.TargetObject -ne $printerPath) {
            # display any errors
            $html += "<li>$($printerError.CategoryInfo.Category) - $($printerError.TargetObject)</li>"
        }
    }

    $html += "</div>"
    $html += "<div style='background-color:#ffffcc;border-left:6px solid #ffeb3b;'><b>$printerPath Permissions</b><ul>"

    # get the permission for the printer Path
    $printerAcl = Get-Acl $printerPath 

    # set the name of the current path
    $currentPath = $printerPath

    # display all the permission for the root of the path
    foreach($printerACLAccess in $printerAcl.Access)
    {
        
        $html += "<li>$($printerACLAccess.IdentityReference) $($printerACLAccess.AccessControlType): $($printerACLAccess.FileSystemRights)</li>"
    }

    $html += "</ul></div>"
#>
    
    # go through all the printers entered
    foreach ($printerWild in $printerList) {
        
        # find all printers that start with prefix
        $printersFound = Get-Printer -ComputerName "$printServersSelected.asurite.ad.asu.edu" -Name "$($printerWild.Trim())*" -Full | Sort-Object Name

        $lastPrinter = ''

        foreach ($printer in $printersFound) {

            if ($lastPrinter -ne $printer.Name ) {
                
                # set new last printer
                $lastPrinter = $printer.Name

                # set the snipit for div
                $html += "<div style='background-color:#ffffcc;margin-top:10px;border-left:6px solid #ffeb3b'><h2>$($printer.Name)</h2><ul>"

                #check and see if the driver is package aware
                $printerDriver = Get-PrinterDriver -ComputerName "$printServersSelected.asurite.ad.asu.edu" -Name "$($printer.DriverName)"
                
                # get the ip addresss
                $printerIP = $($printer.PortName) -replace '(^IP_|_[0-9]*$)', ''

                # ping test the printer
                $printerPing = Test-Connection $printerIP -Quiet

                # display some printer details
                $html += "<li><b>Driver:</b> $($printer.DriverName)</li>"
                $html += "<li><b>Driver Package Aware:</b> $($printerDriver.IsPackageAware[0])"
                if (! $($printerDriver.IsPackageAware[0]) ) {
                    $html += ' <span style="color:red;font-size:150%;">!</span>'
                }
                
                $html += "</li><li><b>Location:</b> $($printer.Location)</li>"
                $html += "</li><li><b>Comments:</b> $($printer.Comment)</li>"
                $html += "</li><li><b>IP:</b> $printerIP</li>"
                $html += "<li><b>IP Ping:</b> $printerPing"
                
                if (! $printerPing ) {
                    $html += ' <span style="color:red;font-size:150%;">!</span>'
                }
                
                $html += "<li><b>Queue Status:</b> $($printer.PrinterStatus)"
                
                if ($($printer.PrinterStatus) -ne 'Normal' ) {
                    $html += ' <span style="color:red;font-size:150%;">!</span>'
                }

                $html += "</li><li><b>Permissions:</b> </li><ul><ul>"

                # SDDL work is based off of http://poshcode.org/3921        
                # get the permissions for the printer
                $printerSDDL = $printer.permissionSDDL

                # get the RawSecurityDescriptor
                $secDescriptor = [Int].Assembly.GetTypes() | Where-Object { 
                    $_.FullName -eq 'System.Security.AccessControl.RawSecurityDescriptor' 
                }

                # get some more context out of the SDDL
                $sddl = [Activator]::CreateInstance($secDescriptor, [Object[]] @($printerSDDL))

                # last account
                $lastAccount = ''
                # run through the permissions
                foreach($printerSDDL in $sddl.DiscretionaryAcl) {

                    # get the account
                    $account = $($printerSDDL.SecurityIdentifier).Translate([Security.Principal.NTAccount]).Value

                    # we only care about Builtin & ASURITE
                    #if ( ($account.StartsWith('BUILTIN\') ) -or ($account.StartsWith('ASURITE\')) ) {
                        # is this the same as the last account
                        if ($lastAccount -ne $account) {
                            # nope. put in the the account name
                            $html += "</ul><li>$account -</li><ul>"
                        }
                        # display permission
                        foreach ($dAcl in $printerSDDL) {
                            $html += "<li>$($printerSDDL.AceType) - "
                            switch ($printerSDDL.AccessMask) {
                                '131072' { 
                                    $html += 'Read'
                                }
                                '262144' { 
                                    $html += 'Change'
                                }
                                '983088' { 
                                    $html += 'Manage Document'
                                }
                                '983052' { 
                                    $html += 'Manage Printer'
                                }
                                '131080' { 
                                    $html += 'Print and Read'
                                }
                                '268435456' { 
                                    $html += 'Full Control'
                                }
                                Default {
                                    $html += 'Unknown'
                                }
                            }
                            $html += '</li>'
                        }
                    #}
                    # set the new last account
                    $lastAccount = $account
                }
                $html += '</ul></ul></div>'
            }
        }
        #$html += "</div><div></div>"
    }
    $html += "</body></html>"

    # write out HTML file
    $html | Out-File $saveDialog.filename

    # open HTML file
    Invoke-Item $saveDialog.filename

}

