<#
.SYNOPSIS
Duplicates links that are on an OU that is being migrated to UTOSPA 
.DESCRIPTION
This function will retreive all the linked GPOs on an OU and place them on 
a staging OU.
#>

function GPOMigration {

    # get the staging OU
    $Global:SyncHash.print("Starting GPO migration process...", $false)
    $stagingOUs = Get-ADOrganizationalUnit -SearchBase 'OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu' -SearchScope Subtree -Filter * | Where-Object { $_.DistinguishedName -notlike 'OU=M.UTOSPA_Staging,*' } |Sort-Object Name

    # can the  user see staging?
    if (($stagingOUs | Measure-Object).count -eq 0) {
        # they can't see into the OU so exit
        $Global:SyncHash.print("You do not have access to the staging OU. Please contact DSO if you expect to have access.", $false)
        return
    }

    # prompt user for the staging area the OU is moving to
    $stagingOU = PopupListSelect $stagingOUs 'Name' 'What staging area do you want to apply the GPOs:' 9

    if ($stagingOU -eq $null) {
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }

    # get the GPOs for the selected staging OU
    $stagingGPOs = Get-GPInheritance -Target "OU=M.UTOSPA_$stagingOU,OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu"
    
    # get the number of GPOs currently on the select staging OU.
    $stagingGPOCount = ($stagingGPOs.GpoLinks | Measure-Object).Count

    # is there anything to clean up from the selected staging OU?
    if ( $stagingGPOCount -ne 0 ) {
        # yup. 

        # confirm that they want to remove the GPOs
        $remStagingGPOs = PopupBox "Are you sure you want to remove the $stagingGPOCount GPOs linked to '$stagingOU'? `r`nThis may take a few moments depending on the number of GPOs." 'Confirmation!' 'yn'

        if ($remStagingGPOs -eq 'no') {
            $Global:SyncHash.print("Canceled. Exiting.", $false)
            return
        }

        # remove the GPO links.
        foreach ($stagingGPO in $stagingGPOs.GpoLinks) {
            try {
                Remove-GPLink -Guid $stagingGPO.GpoId -Target "OU=M.UTOSPA_${stagingOU},OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu" -ErrorAction Stop
            }
            catch {
                $Global:SyncHash.print("There was a problem removing GPO link from '${stagingOU}'. You should manually check the OU. - $error[0].Exception.Message", $false)
                return
            }
        }
    }
    # get the OU where the GPOs are linked
    $oldOU = Choose-ADOrganizationalUnit -HideNewOUFeature

    # make sure something was selected
    if (! $oldOU) {
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }
    
    # get the GPOs linked to the old OU.
    $gpoLinks = Get-GPInheritance -Target $oldOU.DistinguishedName
    
    # get the count of links
    $gpoLinksCount = ($gpoLinks.InheritedGpoLinks | Measure-Object).count

    # see if there are any GPOs linked
    if ( $gpoLinksCount -eq 0) {
        # no GPOs so exit
        $Global:SyncHash.print("No GPOs found. Exiting.", $false)
        return
    }

    # confirm that they want to complete the action
    $confirm = PopupBox "Are you sure you want to add $gpoLinksCount GPOs from '$($oldOU.Name)' to '$stagingOU'?`r`nThis may take a few moments depending on the number of GPOs." "Confirmation" "yn"

    # leave if they don't confirm
    if ($confirm -eq 'no') {
        $Global:SyncHash.print("Canceled. Exiting.", $false)
        return
    }

    # link SCCM and MBAM
    try {
        New-GPLink -Name 'M.UTOSPA.SCCM_Client' -Target "OU=M.UTOSPA_${stagingOU},OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu" -Order 1 -Enforced 'Yes' -LinkEnabled 'Yes'
        New-GPLink -Name 'M.UTOSPA.MBAM' -Target "OU=M.UTOSPA_${stagingOU},OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu" -Order 2 -Enforced 'Yes' -LinkEnabled 'Yes'
    } catch {
        $Global:SyncHash.print("There was a problem adding the GPO link for SCCM/MBAM. You should manually check the OU. - $error[0].Exception.Message", $false)
        return
    }

    # sort the links by order
    $inheritedGPOLinks = $gpoLinks.InheritedGpoLinks | Sort-Object Order

    # start up the link number
    $gpoOrderNumber = 2

    # Enable progress bar
    $syncHash.setProgressValue(0)
    $syncHash.setProgressMax(($inheritedGPOLinks | Measure-Object).count)
    $syncHash.progressVisible($true)

    # link each GPO to the staging OU
    foreach ($gpo in $inheritedGPOLinks) {

        # increment the order number
        $syncHash.setProgressValue($gpoOrderNumber)
        $gpoOrderNumber += 1

        # check if the GPO is enforced
        if ($gpo.Enforced) {
            $isEnforced = 'Yes'
        } else {
            $isEnforced = 'No'
        }
        
        # link GPO to staging OU
        try {
            New-GPLink -Guid $gpo.GpoID -Target "OU=M.UTOSPA_${stagingOU},OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu" -Order $gpoOrderNumber -Enforced $isEnforced -LinkEnabled 'Yes'
            $Global:SyncHash.print("$($gpo.DisplayName) linked to '${stagingOU}'", $false)
        } catch {
            $Global:SyncHash.print("There was a problem adding the GPO link '$($gpo.DisplayName)'. You should manually check the OU. - $error[0].Exception.Message", $false)
            return
        }
    }
    $syncHash.progressVisible($false)
    $Global:SyncHash.print("GPOs have been link to '${stagingOU}'", $false)
}
