<#
.SYNOPSIS
Creates a report of the permissions on a share
.DESCRIPTION
Based on the share that is provided a HTML report is created that 
will display the permission on a share
#>

function ShareReport {
    
    # get the path to search
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')

    $uncPath = [Microsoft.VisualBasic.Interaction]::InputBox("Enter a UNC Path", "Path")

    # make sure the path enters exsists
    if (! $uncPath) {
        $Global:SyncHash.print("Nothing was entered. Exiting", $false)
        return
    } elseif (! (Test-Path $uncPath)) {
        # nope. error and exit.
        $Global:SyncHash.print("${uncPath} does not exist. Exiting.", $false)
        return
    }

    # start the html
    $html = "<html><style>div {padding:20px;}; li {font-weight:bold};</style><body><h1>Share Permissions Report - $uncPath</h1>"

    # display save as dialog to save the HTML file
    $saveDialog = New-Object windows.forms.savefiledialog
    $saveDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
    $saveDialog.title = "Save Share Report"
    $saveDialog.filter = "HTML|*.html|All Files|*.*"
    $result = $saveDialog.ShowDialog()
    if($result -ne "OK") {
        # canceled so stop
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }

    $continue =  PopupBox 'Please be aware that this may take some time depending on the folder structure .' 'Continue?' 'oc' 

    if ($continue -eq 'Cancel') {
        # canceled so stop
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }

    # create a variable to hold errors while travering the directory structure.
    $uncErrors = @()

    # traverse the UNC Path.
    $uncFolders = Get-ChildItem -Recurse $uncPath -ErrorAction SilentlyContinue -ErrorVariable +uncErrors | where-object {($_.PsIsContainer)} | Sort-Object FullName

    # go through all the errors
    foreach ($uncError in $uncErrors) {

        $html += "<div style='background-color:#ffcccc;border-left:6px solid red;'><h2>Unable to access:</h2>"

        # make sure that it's not an error about the root
        if ($uncError.TargetObject -ne $uncPath) {
            # display any errors
            $html += "<li>$($uncError.CategoryInfo.Category) - $($uncError.TargetObject)</li>"
        }
        $html += "</div>"
    }

    $html += "<div style='background-color:#ffffcc;border-left:6px solid #ffeb3b;'><b>$uncPath Permissions</b><ul>"

    # get the permission for the UNC Path
    $uncAcl = Get-Acl $uncPath 

    # set the name of the current path
    $currentPath = $uncPath

    # display all the permission for the root of the path
    foreach($uncACLAccess in $uncAcl.Access)
    {
        
        $html += "<li>$($uncACLAccess.IdentityReference) $($uncACLAccess.AccessControlType): $($uncACLAccess.FileSystemRights)</li>"
    }

    $html += "</ul></div>"

    # go through all the folders
    foreach ($uncFolder in $uncFolders) {
        # create the spacing before each line - this is based on the number of backslashes
        $spaceMulit = (([regex]::Matches($uncFolder.FullName, "\\" )).count) - 3
        $spaceHeading = ( 15 * $spaceMulit ) + 20
        $perHtml = "<div style='background-color:#ffffcc;border-left:6px solid #ffeb3b;margin-left:$spaceHeading'>"

        # default the heading to false, so that it will run at least once
        $heading = $false

        # get the permissions for the folder
        $uncAcl = Get-Acl $uncFolder.FullName

        # run through the permissions
        foreach($uncACLAccess in $uncAcl.Access) {
            # we only care about not inherited  permissions
            if (! $uncACLAccess.IsInherited) {
                # diplay header if it is not already displayed
                if (! $heading) {
                    # place a gap between the last permission if this is not a sub folder
                    if ($uncFolder.FullName -notlike "$currentPath*") {
                        $html += "<div style='margin-left:$spaceHeading'>...</div>"
                    }
                    $html += "$perHtml <b>$($uncFolder.FullName) Permissions</b><ul>"
                    $heading = $true
                    $currentPath = $uncFolder.FullName
                }
                # display the access
                $html += "<li>$($uncACLAccess.IdentityReference) $($uncACLAccess.AccessControlType): $($uncACLAccess.FileSystemRights)</li>"
            }
        }
        if ($heading) {
            $html += "</ul></div>"
        }
    }
    $html += "</body></html>"

    # write out HTML file
    $html | Out-File $saveDialog.filename

    # open HTML file
    Invoke-Item $saveDialog.filename

}