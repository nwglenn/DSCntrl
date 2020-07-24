<#
.SYNOPSIS
Intended to be called as a DSCtrl child thread. Certain variables and functions will be missing if called in any other fashion; results may be unpredictable
#>
function Invoke-ExchangeOffboarding {

    # This function requires that the ADServer global variable be set.
    # Typically, this should happen when the global.ps1 file is loaded
    # on creation of a new DSCtrl thread. If this has not happened, we
    # default to ASURITE6 so we don't pass a null string later in the function
    if ( ($global:ADserver -isnot [String]) -or ($global:ADserver.length -lt 12) ) {
        $global:ADserver = 'asurite6.asurite.ad.asu.edu'
    }

    # Double-check that AD module is running
    $syncHash.print("[ OK! ] Loading AD module...")
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    # Connect to Exchange, if module not already loaded
    $syncHash.print("[ OK! ] Connecting to Microsoft Exchange...")
    if ($global:exSession -eq $null -or $global:exSession.State -eq 'Closed') {

        # If running within DSCtrl, call host method to get creds
        # Otherwise, use native function
        $creds= $null
        if ($syncHash.host) {
            $Creds = $syncHash.host.ui.PromptForCredential('Windows PowerShell credential request', 'Provide ASURITE credentials for accessing Microsoft Exchange',"ASURITE\$([Environment]::UserName)", "")
        } else {
            $Creds = Get-Credential -Message "Provide ASURITE credentials for accessing Microsoft Exchange"
        }

        # Attempt connection
        $global:exSession = New-ExchangeConnection -Credential $Creds -NoImport

        # Verify that connection was successful
        if (!$global:exSession) {
            Write-Host "[ERR] Connection to Microsoft Exchange failed"
            return 840
        }
    }

    # Define the form
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:WpfApplication2"
Title="Exchange DL membership report" Height="600" MinHeight="480" Width="800" MinWidth="800">
    <DockPanel>
        <StackPanel Orientation="Vertical" DockPanel.Dock="Top" Margin="16" Height="40">
            <StackPanel Orientation="Horizontal" VerticalAlignment="Center">
                <Label Margin="4" Height="24" VerticalAlignment="Center">ASURITE user name</Label>
                <TextBox Name="txtUserName" Margin="4" Height="24" VerticalAlignment="Center" Width="200"></TextBox>
                <Button Name="btnSearch" Margin="4" Height="24" Width="64" VerticalAlignment="Center">Search</Button>
                <Label Name="lblProgress" Visibility="Collapsed" Margin="16,0,8,0">Progress:</Label>
                <ProgressBar Name="progress" Visibility="Collapsed" Margin="2,0,8,0" Width="270" />
            </StackPanel>
        </StackPanel>
        <DockPanel DockPanel.Dock="Bottom">
            <GroupBox Header="DL membership" Margin="8" DockPanel.Dock="Left" Width="380">
                <DockPanel>
                    <StackPanel Orientation="Horizontal" DockPanel.Dock="Bottom" Height="42" HorizontalAlignment="Right">
                        <Button Name="btnRemove" Margin="8" Height="24" Width="180" VerticalAlignment="Center">Remove selected groups</Button>
                    </StackPanel>
                    <ListBox Margin="8,6,8,6" Name="lstMembership" VerticalAlignment="Stretch" DockPanel.Dock="Top">
                        <ListBox.ItemTemplate>
                            <HierarchicalDataTemplate>
                                <CheckBox Content="{Binding Name}" IsChecked="{Binding IsChecked}" IsEnabled="{Binding IsEnabled}"/>
                            </HierarchicalDataTemplate>
                        </ListBox.ItemTemplate>
                    </ListBox>
                </DockPanel>
            </GroupBox>
            <GroupBox Header="DL ownership" Margin="8" DockPanel.Dock="Right">
                <DockPanel>
                    <StackPanel Orientation="Horizontal" DockPanel.Dock="Bottom" Height="42" HorizontalAlignment="Right">
                        <Button Name="btnCopy" Margin="8" Height="24" Width="180" VerticalAlignment="Center">Copy ownership to clipboard</Button>
                    </StackPanel>
                    <ListBox Margin="8,6,8,6" Name="lstOwnership" VerticalAlignment="Stretch" DockPanel.Dock="Top">
                    </ListBox>
                </DockPanel>
            </GroupBox>
        </DockPanel>
    </DockPanel>
</Window>
"@

    # Generate paths for images
    $inputXML = $inputXML -replace '!imgpath!', "${IncludePath}\img"

    # Load XML; prep for deserialization
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms
    [xml]$XAML = $inputXML

    # Deserialize the form into an object
    $form = $null
    try {
        $Form = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xaml))
    } catch [System.Management.Automation.MethodInvocationException] {
        Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
        write-host $error[0].Exception.Message -ForegroundColor Red
        if ($error[0].Exception.Message -like "*button*") {
            write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"
        }
    } catch {
        #if it broke some other way :D
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }

    # Load form components into a hash table
    # This must be stored in global scope so event calls can reach it
    $Global:offboardForm = [hashtable]::Synchronized(@{})
    $Global:offboardForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:offboardForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:offboardForm.add('memberList', (New-Object -TypeName System.collections.ArrayList))
    $Global:offboardForm.add('OwnerList', (New-Object -TypeName System.collections.ArrayList))
    $Global:offboardForm.add('printProxy', [ref]$syncHash)
    $Global:offboardForm.lstMembership.itemsSource = $Global:offboardForm.memberList
    $Global:offboardForm.lstOwnership.itemsSource = $Global:offboardForm.OwnerList

    #
    # Functions for managing the progress indicator
    #
    $Global:offboardForm | Add-Member -Type ScriptMethod -name 'progressVisible' -Value {
        param($isVisible)
        if ($isVisible) {
            $this.Window.Dispatcher.invoke(
                [action]{
                    $this.progress.Visibility = "Visible"
                    $this.lblProgress.Visibility = "Visible"
                }
            )
        } else {
            $this.Window.Dispatcher.invoke(
                [action]{
                    $this.progress.Visibility = "Collapsed"
                    $this.lblProgress.Visibility = "Collapsed"
                }
            )
        }
    }
    $Global:offboardForm | Add-Member -Type ScriptMethod -name 'setProgressMax' -Value {
        param($maxValue)
        $this.Window.Dispatcher.invoke(
            [action]{ $this.progress.Maximum = "${maxValue}"}
        )
    }
    $Global:offboardForm | Add-Member -Type ScriptMethod -name 'setProgressValue' -Value {
        param($progressValue)
        $this.Window.Dispatcher.invoke(
            [action]{ $this.progress.Value = "${progressValue}"}
        )
    }

    # Create runspace for long-running tasks
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash", $Global:offboardForm)
    $Runspace.SessionStateProxy.SetVariable("includePath", $includePath)
    $Runspace.SessionStateProxy.SetVariable("adServer", $global:ADserver)
    $Runspace.SessionStateProxy.SetVariable("exSession", $global:exSession)

    $Global:offboardForm.btnSearch.add_Click({

        # Prep form for new query
        $syncHash.memberList.clear()
        $syncHash.OwnerList.clear()
        $syncHash.lstMembership.items.refresh()
        $syncHash.lstOwnership.items.refresh()

        $code = {
            . "${includePath}\global.ps1"
            . "${includePath}\popup.ps1"
            Start-Transcript -Path (Join-Path $Global:logdir "getExDLs-$((get-date).ToFileTime())") -IncludeInvocationHeader
            $syncHash.printProxy.value.print("[ OK! ] Starting DL search")

            $global:ADserver = $syncHash.adServer
            if ( ($global:ADserver -isnot [String]) -or ($global:ADserver.length -lt 12) ) {
                $global:ADserver = 'asurite6.asurite.ad.asu.edu'
            }

            # Reset progress indicator
            $syncHash.setProgressValue(0)
            $syncHash.setProgressMax(1)
            $syncHash.progressVisible($true)

            # Do we need to import exSession?
            # It may have been imported by a previous invocation
            if ([Bool](Get-Command Get-DistributionGroup -Erroraction SilentlyContinue) -ne $true) {
                Import-PSSession $exSession
            }

            # Pull user name from form and update controlls
            # While we are there, clear the form of any old data
            $script:uName = $null
            $syncHash.window.Dispatcher.invoke(
                [action]{
                    $syncHash.btnSearch.isEnabled = $false
                    $script:uName = $syncHash.txtUserName.Text

                    $syncHash.memberList.clear()
                    $syncHash.ownerList.clear()
                    $syncHash.lstMembership.items.refresh()
                    $syncHash.lstOwnership.items.refresh()
                }
            )
            Write-Host $uName
        
            # Get user information that we can pass around for querties
            Try {
                $targetUser = Get-ADUser -Identity $uName -Properties 'DistinguishedName', 'MemberOf' -Server $global:ADserver
            } Catch {
                $syncHash.printProxy.value.print("[ ERR ] Could not find AD user record for ${uName}, aborting operation")
                PopupBox "No user was found with ASURITE '${uName}'. Please check spelling and try again. "
                $syncHash.progressVisible($false)
                $syncHash.window.Dispatcher.invoke(
                    [action]{
                        $syncHash.btnSearch.isEnabled = $true
                        $syncHash.btnRemove.isEnabled = $true
                    }
                )
                exit
            }
            
            # Build list of groups that calling user is a member of;
            # This will be used to determine if we have access to manage a specific DL
            $myAccessToken = $null
            $myDN = Get-ADUser -Identity $ENV:USERNAME -Server $global:ADserver -Properties 'distinguishedname' | Select-Object -ExpandProperty 'distinguishedname'
            if ($myDN) {
                $myAccessToken = Get-ADUser -SearchScope 'Base' -SearchBase "${myDN}" -Server $global:ADserver -LDAPFilter '(objectClass=user)' -Properties @("distinguishedname", "samaccountname", "useraccountcontrol", "objectsid", "sidhistory", "primarygroupid", "lastlogontimestamp", "memberof", "tokenGroups") | Select-Object -ExpandProperty 'tokenGroups'
            }

            # Create filtered list to save time
            $dlList = $targetUser.memberOf | Where-Object -FilterScript {$_ -iMatch '^CN=DL\.'} | Sort-Object
            $dlProgress = 0

            # Start tracking progress, so folks know we are doing things #tempo :-)
            $syncHash.setProgressValue($dlProgress)
            $syncHash.setProgressMax($dlList.count)
            $syncHash.progressVisible($true)

            # Construct a list of 'DL.' prefixed groups
            $dlList | Foreach-Object {
                $newItem = [PSCustomObject]@{
                    name      = ""
                    isChecked = $false
                    isEnabled = $true
                }

                $someGroup = $null
                try {
                    $error.clear()
                    $someGroup = Get-DistributionGroup -Identity $_
                } catch {
                    $someGroup = $null
                }

                if ($someGroup -and ($error.count -eq 0) ) {
                    $newItem.name = $someGroup.DisplayName

                    # Determine if we have access to manage this DL
                    foreach ($owner in $someGroup.ManagedBy) {
                        #if ($owner -like "*${ENV:USERNAME}") {
                        if ($owner -like "*${ENV:USERNAME}") {
                            $newItem.isEnabled = $true
                        }
                    }

                    $syncHash.memberList.add($newItem)
                } else {
                    $syncHash.printProxy.value.print("[ ERR ] Could not find: $($_), it will be omitted from the report")
                }

                $dlProgress++
                $syncHash.setProgressValue($dlProgress)
            }

            # Populate ownership list
            Get-Recipient -Filter "ManagedBy -eq '$($targetUser.distinguishedname)'" | Select-Object -ExpandProperty 'name' | Sort-Object | ForEach-Object {
                $syncHash.ownerList.add($_)
            }

            # If no owned groups were added, add a 'NONE' to the list
            if ($syncHash.ownerList.count -eq 0) {
                $syncHash.ownerList.add('None')
            }

            $syncHash.progressVisible($false)
            $syncHash.window.Dispatcher.invoke(
                [action]{
                    $syncHash.lstMembership.items.refresh()
                    $syncHash.lstOwnership.items.refresh()
                    $syncHash.btnSearch.isEnabled = $true
                }
            )
            Stop-Transcript
        }
        $PSinstance = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })

    $Global:offboardForm.btnRemove.add_Click({
        $code = {
            . "${includePath}\global.ps1"
            . "${includePath}\popup.ps1"
            Start-Transcript -Path (Join-Path $Global:logdir "RemoveDL-Member-$((get-date).ToFileTime())") -IncludeInvocationHeader
            $syncHash.printProxy.value.print("[ OK! ] Starting list removal task")

            # Do we need to import exSession?
            # It may have been imported by a previous invocation
            if ([Bool](Get-Command Get-DistributionGroup -Erroraction SilentlyContinue) -ne $true) {
                Import-PSSession $exSession
            }

            # Reset progress bar
            $syncHash.setProgressValue(0)
            $syncHash.setProgressMax(1)
            $syncHash.progressVisible($true)

            # get user name from form
            $script:uName = $null
            $script:userDN = $null
            $syncHash.window.Dispatcher.invoke(
                [action]{
                    $syncHash.btnSearch.isEnabled = $false
                    $syncHash.btnRemove.isEnabled = $false
                    $script:uName = $syncHash.txtUserName.Text
                }
            )

            # Resolve DN for Exchange query
            if ($script:uName) {
                Try {
                    $script:userDN = Get-ADUser -Identity $script:uName -Server "ASURITE.ad.asu.edu" -Properties 'DistinguishedName'  | Select-Object -ExpandProperty 'DistinguishedName'
                } Catch {
                    $script:userDN = $null
                }
            }

            # Catch instances where resolving DN failed
            if ($script:userDN -isnot [String]) {
                PopupBox 'There was a problem resolving the Distinguished Name (DN) of the user account you attempted to remove. Please verify that the account name was entered correctly and try again. For additional information, check the application logs in the Help menu'
                $syncHash.window.Dispatcher.invoke(
                    [action]{
                        $syncHash.btnSearch.isEnabled = $true
                        $syncHash.btnRemove.isEnabled = $true
                    }
                )
                exit
            }

            # Build list of group names where;
            # the element is enabled (user is owner)
            # checkbox is checked
            $dlList = $syncHash.memberList | Where-Object -FilterScript {
                $_.isEnabled -and $_.isChecked
            }

            # Abort if no groups were selected
            if ($dlList.count -eq 0) {
                PopupBox 'No groups selected. Use the checkboxes to select at least one group.'
                $syncHash.window.Dispatcher.invoke(
                    [action]{
                        $syncHash.btnSearch.isEnabled = $true
                        $syncHash.btnRemove.isEnabled = $true
                    }
                )
                exit
            }
            $confirmMessage = @"
You are about to remove '${uName}'
from the following Distribution Lists

$($dlList.name | Out-String)

Do you wish to continue?
"@
            $confirm = PopupBox $confirmMessage "Confirm removal" "yn"
            if ($confirm -ne 'yes') {
                $syncHash.printProxy.value.print("[ WRN ] Group removal aborted")
                $syncHash.window.Dispatcher.invoke(
                    [action]{
                        $syncHash.btnSearch.isEnabled = $true
                        $syncHash.btnRemove.isEnabled = $true
                    }
                )
                exit
            }
            $syncHash.setProgressMax($dlList.count)
            $myProgress = 0
            # Attempt to remove the user from each DL
            $dlList | ForEach-Object {
                try {
                    $error.clear()
                    Remove-DistributionGroupMember -Identity $_.name -member $script:userDN -Confirm:$false
                } catch {
                    
                    # Ensure that the error stack contains at least one error,
                    # to trigger subsequent error checking login
                    if ($error.count -lt 1) {
                        Write-Error $_
                    }
                }

                # If things didn't break, assume success and temove group from master list
                if ($error.count -eq 0) {
                    $syncHash.printProxy.value.print("[ OK! ] removed ${uName} from $($_.name)")
                    $syncHash.memberList.remove($_)
                } else {
                    $syncHash.printProxy.value.print("[ ERR ] Could not remove ${uName} from $($_.name)")
                }

                $myProgress++
                $syncHash.setProgressValue($myProgress)
            }

            # log completion
            $syncHash.printProxy.value.print("[ OK! ] completed removal evaluation")

            # Clean up
            $syncHash.window.Dispatcher.invoke(
                [action]{
                    $syncHash.btnSearch.isEnabled = $true
                    $syncHash.btnRemove.isEnabled = $true
                    $syncHash.lstMembership.items.refresh()
                }
            )
            $syncHash.progressVisible($false)
            Stop-Transcript
        }
        $PSinstance = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })

    $Global:offboardForm.btnCopy.add_Click({
        Set-Clipboard -Value "$($Global:offboardForm.ownerList | Out-String)"
    })

    # Form visibility changed
    $Global:offboardForm.window.add_IsVisibleChanged( {
        if ($Global:offboardForm.window.isVisible -eq $true) {
            $Global:offboardForm.window.topmost = $true
            $Global:offboardForm.window.topmost = $false
            $Global:offboardForm.window.focus()
        }
    })

    $global:offboardForm.window.showdialog()
    $Global:offboardForm.memberList
}
