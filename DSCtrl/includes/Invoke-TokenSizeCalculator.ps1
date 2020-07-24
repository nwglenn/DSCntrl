<#
.SYNOPSIS
Displays a form to estimate the size of a users token
.DESCRIPTION
Performs analytics on the groups in a users TokenGroups dynamic property to estimate the size of the users token
.NOTES
Based on a script written by Jeremy Saunders
http://www.jhouseconsulting.com/2013/12/20/script-to-create-a-kerberos-token-size-report-1041
#>
function Invoke-TokenSizeCalculator {

    # This function requires that the ADServer global variable be set.
    # Typically, this should happen when the global.ps1 file is loaded
    # on creation of a new DSCtrl thread. If this has not happened, we
    # default to ASURITE6 so we don't pass a null string later in the function
    if ( ($global:ADserver -isnot [String]) -or ($global:ADserver.length -lt 12) ) {
        $global:ADserver = 'asurite6.asurite.ad.asu.edu'
    }

    # Double-check that AD module is running
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    # Define the form
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
Title="Token Calculator" Height="600" MinHeight="480" Width="1000" MinWidth="1000">
    <DockPanel Margin="16">
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Height="32">
            <TextBlock Margin="8">NOTE: Values shown are not exact; they represent estimates based on available data</TextBlock>
        </StackPanel>
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Left" DockPanel.Dock="Top">
            <Label>ASURITE user name:</Label>
            <TextBox Name="txtUserName" Width="256" Margin="8,0,8,0"></TextBox>
            <Button Name="btnCalculate" Width="80">Calculate!</Button>
        </StackPanel>
        <GroupBox Header="Token count" VerticalAlignment="Top" Height="80" DockPanel.Dock="Top">
            <StackPanel Orientation="Horizontal" Margin="8">
                <TextBlock Name="itemCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Right" Margin="0,0,4,0">0</TextBlock>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal">
                        <Canvas Name="itemCountDisplay" Background="Black" Height="16" Width="1" HorizontalAlignment="Left"></Canvas>
                        <Canvas Name="errorItemsDisplay" Background="Red" Height="16" Width="0" HorizontalAlignment="Left"></Canvas>
                    </StackPanel>
                    <Image Height="32" Width="800" Source="!imgpath!\itemBar-halfscale.png" HorizontalAlignment="Left"></Image>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        <GroupBox Header="Estimated token size" VerticalAlignment="Top" Height="80" DockPanel.Dock="Top">
            <StackPanel Orientation="Horizontal" Margin="8">
                <TextBlock Name="sizeCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Right" Margin="0,0,4,0">0</TextBlock>
                <StackPanel Orientation="Vertical">
                    <StackPanel Orientation="Horizontal">
                        <Canvas Name="sizeDisplay" Background="Black" Height="16" Width="1" HorizontalAlignment="Left"></Canvas>
                        <Canvas Name="errorDisplay" Background="Red" Height="16" Width="0" HorizontalAlignment="Left"></Canvas>
                    </StackPanel>
                    <Image Height="32" Width="800" Source="!imgpath!\tokenBar-halfscale.png" HorizontalAlignment="Left"></Image>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        <StackPanel Orientation="Horizontal" DockPanel.Dock="Top">
            <GroupBox Header="Domain-Local" Width="450" Height="120">
                <StackPanel Orientation="Horizontal">
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="localCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Local</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="domainGlobalCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Global</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="DomainUniversalCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Universal</TextBlock>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Foreign" Width="300" Height="120">
                <StackPanel Orientation="Horizontal">
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="foreignUniversalCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Global</TextBlock>
                    </StackPanel>
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="ForeignGlobalCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Universal</TextBlock>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Unknown" Width="140" Height="120">
                <StackPanel Orientation="Horizontal">
                    <StackPanel Orientation="Vertical">
                        <TextBlock Name="errorCount" FontWeight="Bold" FontSize="48" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2" Foreground="#FFF10505">0</TextBlock>
                        <TextBlock FontWeight="Bold" Width="128" VerticalAlignment="Center" TextAlignment="Center" Margin="2">Unresolvable SID</TextBlock>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
        </StackPanel>
        <GroupBox Header="Principal groups" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" DockPanel.Dock="Bottom">
            <DataGrid Name="groupGrid" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Group name" IsReadOnly="True" Binding="{Binding name}" Width="512" />
                    <DataGridTextColumn Header="children" IsReadOnly="True" Binding ="{Binding children}" Width="64" />
                    </DataGrid.Columns>
                </DataGrid>
        </GroupBox>
    </DockPanel>
</Window>
"@

    # Generate paths for images
    $inputXML = $inputXML -replace '!imgpath!',"${IncludePath}\img"

    # Load XML; prep for deserialization
    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
    [xml]$XAML = $inputXML

    # Deserialize the form into an object
    $form = $null
    try{
        $Form = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xaml))
    } catch [System.Management.Automation.MethodInvocationException] {
        Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
        write-host $error[0].Exception.Message -ForegroundColor Red
        if ($error[0].Exception.Message -like "*button*"){
            write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"
        }
    } catch {
        #if it broke some other way :D
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }

    # Load form components into a hash table
    # This must be stored in global scope so event calls can reach it
    $Global:tokenForm = [hashtable]::Synchronized(@{})
    $Global:tokenForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:tokenForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:tokenForm.add('okButtonClicked',$false)
    $Global:tokenForm.add('groupList',(New-Object -TypeName System.collections.ArrayList))
    $Global:tokenForm.groupGrid.itemsSource = $Global:tokenForm.groupList

    # Create runspace for long-running tasks
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$Global:tokenForm)
    $Runspace.SessionStateProxy.SetVariable("includePath",$includePath)

    $Global:tokenForm.btnCalculate.add_Click({
        
        # Validate the text that was entered
        $uName = $Global:tokenForm.txtUserName.Text
        
        # Disable the components of the form so we don't spawn multiple threads
        $Global:tokenForm.txtUserName.isEnabled = $false
        $Global:tokenForm.btnCalculate.isEnabled = $false
        $Global:tokenForm.groupList.clear()
        $Global:tokenForm.groupGrid.items.refresh()

        # Spawn the thread and start calculating
        $code = {
            . "${includePath}\global.ps1"
            Start-Transcript -Path (Join-Path $Global:logdir "calctoken-worker-$((get-date).ToFileTime())") -IncludeInvocationHeader
            $script:AccountName = $null
            $script:localCount = 0
            $script:domainGlobalCount = 0
            $script:ForeignGlobalCount = 0
            $script:domainUniversalCount = 0
            $script:foreignUniversalCount = 0
            $script:SIDHistoryCount = 0
            $script:errorCount = 0
            $script:groupcache = @{}

            $syncHash.Window.Dispatcher.invoke(
                [action]{ $script:AccountName = $syncHash.txtUserName.Text}
            )

            # Only update the form every few seconds - this improves performance by reducing the number of times we have to wait for the WPF dispatcher
            $script:formLastUpdated = (get-date).ToFileTime()
            $script:minimumWaitTime = 3 # interval time in seconds
            function updateFormData {
                $script:TokenSize = 1200 + (40 * ($script:localCount + $script:ForeignGlobalCount + $script:foreignUniversalCount + $script:SIDHistoryCount)) + (8 * ($script:domainGlobalCount + $script:domainUniversalCount))
                $syncHash.Window.Dispatcher.invoke(
                [action]{

                    $syncHash.itemCount.text = ($script:localCount + $script:domainGlobalCount + $script:domainUniversalCount + $script:ForeignGlobalCount + $script:foreignUniversalCount)
                    $syncHash.itemCountDisplay.width = ($script:localCount + $script:domainGlobalCount + $script:domainUniversalCount + $script:ForeignGlobalCount + $script:foreignUniversalCount) / 2
                    $syncHash.errorItemsDisplay.width = $script:errorCount / 2

                    $syncHash.sizeCount.text = $script:TokenSize
                    $syncHash.sizeDisplay.width = $script:TokenSize / 20
                    $syncHash.errorDisplay.width = ($script:errorCount *8) / 20
                    
                    $syncHash.localCount.text = $script:localCount
                    $syncHash.domainGlobalCount.text = $script:domainGlobalCount
                    $syncHash.domainUniversalCount.text = $script:domainUniversalCount
                    $syncHash.ForeignGlobalCount.text = $script:ForeignGlobalCount
                    $syncHash.foreignUniversalCount.text = $script:foreignUniversalCount
                    $syncHash.errorCount.text = $script:errorCount
                })
                $script:TokenSize = 0
                $script:formLastUpdated = (get-date).ToFileTime()
            }

            # Start by getting the target users information
            $user = $null
            try {
                $user = Get-ADUser -Identity $script:AccountName -server 'asurite.ad.asu.edu' -Properties 'distinguishedname' | Select-Object -ExpandProperty 'distinguishedname'
                $user = Get-ADUser -SearchScope 'Base' -SearchBase "${user}" -LDAPFilter '(objectClass=user)' -Properties @("distinguishedname","samaccountname","useraccountcontrol","objectsid","sidhistory","primarygroupid","lastlogontimestamp","memberof","tokenGroups")
            }
            catch {
                $user = $null
                $syncHash.Window.Dispatcher.invoke(
                    [action]{
                        $syncHash.txtUserName.isEnabled = $True
                        $syncHash.btnCalculate.isEnabled = $True
                    }
                )
                #TODO: Provide additional feedback to user about what went wrong
                exit 9
            }

            # Determine users primary group by combining the domain SID and group ID properties
            $primaryGroup = $null
            try {
                $primaryGroup = Get-ADGroup -Identity "$($user.objectsid.AccountDomainSid)-$($user.primarygroupid)"
            }
            catch {
                $primaryGroup = $null
                $syncHash.Window.Dispatcher.invoke(
                    [action]{
                        $syncHash.txtUserName.isEnabled = $True
                        $syncHash.btnCalculate.isEnabled = $True
                    }
                )
                #TODO: Provide additional feedback to user about what went wrong
                exit 9
            }

            # Populate the groups table
            # Start by adding the users primary group, then add memberOf groups
            $synchash.groupList.add(
                [PSCustomObject]@{
                    name = $primaryGroup.name
                    children = 0
                }
            )
            $user.memberof | Sort-Object | ForEach-Object {
                $displayName = ([RegEx]'(?i)CN=(?<name>[^,]+)').Match($_)
                if (
                    $displayName.success -and
                    $displayName.groups.item('name').value -ne $null
                ) {
                    $displayName = $displayName.groups.item('name').value
                } else {
                    $displayName = $_
                }
                $synchash.groupList.add(
                    [PSCustomObject]@{
                        name = $displayName
                        children = 0
                    }
                )
            }
            $syncHash.Window.Dispatcher.invoke(
                [action]{
                    $synchash.groupGrid.items.refresh()
                }
            )

            #TODO: Skipping code that deals with SID history: source file line 420

            # This is where the math starts...
            $user.tokengroups | ForEach-Object {
                $someGroup = $null
                try {
                    $someGroup = Get-ADGroup -Identity $_ -Server 'asurite.ad.asu.edu' -Properties @('DistinguishedName','SID','GroupCategory','GroupScope','SIDHistory','memberof')
                }
                catch {
                    $someGroup = $null
                }

                if ($someGroup -ne $null) {

                    # Cache for later computation
                    $script:groupcache.add($someGroup.name,$someGroup)
                    
                    switch ($someGroup.groupScope) {
                        'DomainLocal' {
                            $script:localCount++
                        }
                        'Global' {
                            if ($user.SID.AccountDomainSid -ne $_.accountdomainsid) {
                                $script:domainGlobalCount++
                            } else {
                                $script:ForeignGlobalCount++
                            }
                        }
                        'Universal' {
                            if ($user.SID.AccountDomainSid -ne $_.accountdomainsid) {
                                $script:domainUniversalCount++
                            } else {
                                $script:foreignUniversalCount++
                            }
                        }
                        Default {}
                    }

                    #TODO: fix this - it results in values higher than expected
                    #$script:SIDHistoryCount += $someGroup.SIDHistory.count

                    #TODO: If the account is trusted for delegation, we need to multiply this value by 2 before reporting it
                    
                } else {
                    if ($user.SID.AccountDomainSid -ne $_.accountdomainsid) {
                        $script:foreignUniversalCount++
                    } else {
                        $script:errorCount++
                    }
                }

                # Only update the form if it has been awhile since we posted new data
                $lastUpdateSeconds = ((get-date).ToFileTime() - $script:formLastUpdated) / 10000000
                if ( $lastUpdateSeconds -gt $script:minimumWaitTime ) {
                    updateFormData
                }
            }

            # Run the update one more time to ensure we have the final data posted
            updateFormData

            # Calculate child counts using $script:groupcache
            function countChildren {
                param($childName)
                $returnCount = 0
                $aGroup = $script:groupcache."${childname}"

                # Abort if object was not found in cache
                if ($aGroup -eq $null) {return $returnCount}

                # Base case = return if no child groups
                if (!$aGroup.memberof -or $aGroup.memberof.count -eq 0) {
                    return $returnCount
                }

                # Call ourselves again for each child group
                $aGroup.memberof | ForEach-Object {
                    $displayName = ([RegEx]'(?i)CN=(?<name>[^,]+)').Match($_).groups.item('name').value
                    $returnCount += countChildren $displayName
                    $returnCount++
                }
                return $returnCount
            }
            for ($i = 0; $i -lt $synchash.groupList.Count; $i++) {
                $synchash.groupList[$i].children = countChildren $synchash.groupList[$i].name
            }
            $syncHash.Window.Dispatcher.invoke(
                [action]{
                    $synchash.groupGrid.items.refresh()
                }
            )

            # Unlock the form
            $syncHash.Window.Dispatcher.invoke(
            [action]{
                $syncHash.txtUserName.isEnabled = $True
                $syncHash.btnCalculate.isEnabled = $True
            })
            Stop-Transcript
        }
        $PSinstance = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })

    # Form visibility changed
    $Global:tokenForm.window.add_IsVisibleChanged({
        if ($Global:tokenForm.window.isVisible -eq $true) {
            $Global:tokenForm.window.topmost = $true
            $Global:tokenForm.window.topmost = $false
            $Global:tokenForm.window.focus()
        }
    })

    $global:tokenForm.window.showdialog()
}
