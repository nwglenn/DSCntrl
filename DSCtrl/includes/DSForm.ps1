# this script is a heavly modified version that came
# from https://github.com/1RedOne/FoxDeploy-GUI-Series-Post4-Resources/blob/master/FancyUI.ps1

if ($includePath -eq $null) {
    $script:includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}

$inputXML = @"
<Window x:Class="Deskside_Control_Panel.Tabs"
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:Deskside_Control_Panel"
mc:Ignorable="d"
Title="$globalTitle" Height="660" Width="1100">
<DockPanel>
<Menu DockPanel.Dock="Top">
    <MenuItem Header="_File">
        <MenuItem Header="Copy log" Name="CopyLogMenuItem" />
        <MenuItem Header="_Clear log" Name="ClearMenuItem" />
        <Separator />
        <MenuItem Header="_Exit" Name="exitMenuItem" />
    </MenuItem>
    <MenuItem Header="_Computer">
        <MenuItem Header="Verify" Name="Verify">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Test if computer is online</TextBlock>
                        <TextBlock TextWrapping="Wrap">Performs a ping test using the provided computer name and returns the result in the log window</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Get Security Groups" Name="CompSecGroup">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock TextWrapping="Wrap">Returns a list of security groups the computer is a member of</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Get password" Name="LAPS">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Get LAPS password</TextBlock>
                        <TextBlock TextWrapping="Wrap">For machines managed by the LAPS client, this function obtains the most recently configured password for the provided computer name</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Go to asset" Name="goToAsset">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Go to Service-Now asset record</TextBlock>
                        <TextBlock TextWrapping="Wrap">Uses your default browser to search the Service-Now asset database for the machine name you specified</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Start Bomgar session" Name="BOMGAR">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Remote support using Bomgar</TextBlock>
                        <TextBlock TextWrapping="Wrap">Uses the computer name provided to attempt a remote support session using Bomgar.</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
    </MenuItem>
    <MenuItem Header="_User">
        <MenuItem Header="Exchange DL offboarding" Name="ExDLRemoveMember">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Remove a user from managed DLs</TextBlock>
                        <TextBlock TextWrapping="Wrap">Remove a user account from any Exchange Distribution Lists (DLs) that you have access to manage.</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Token Calculator" Name="CalcToken">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Estimate the token size for a user account</TextBlock>
                        <TextBlock TextWrapping="Wrap">This tool examines the group membership information for a provided user account in order to estimate the size of their token object</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
    </MenuItem>
    <MenuItem Header="_Groups">
        <MenuItem Header="_UTOSPA Security Groups">
            <MenuItem Header="Create" Name="CreateSecGrp">
                <MenuItem.ToolTip>
                    <ToolTip>
                        <StackPanel Width="256">
                            <TextBlock FontWeight="Bold">Create new security group</TextBlock>
                            <TextBlock TextWrapping="Wrap">Creates a new security group that conforms to the UTOSPA naming convention</TextBlock>
                        </StackPanel>
                    </ToolTip>
                </MenuItem.ToolTip>
            </MenuItem>
            <MenuItem Header="Bulk import" Name="ImportSecGrp">
                <MenuItem.ToolTip>
                    <ToolTip>
                        <StackPanel Width="256">
                            <TextBlock FontWeight="Bold">Import security groups from a migration worksheet</TextBlock>
                            <TextBlock TextWrapping="Wrap">This tool will walk you through creating a collection of security groups from a migration worksheet.</TextBlock>
                        </StackPanel>
                    </ToolTip>
                </MenuItem.ToolTip>
            </MenuItem>
            <MenuItem Header="Bulk rename" Name="RenameSecGrp">
                <MenuItem.ToolTip>
                    <ToolTip>
                        <StackPanel Width="256">
                            <TextBlock FontWeight="Bold">Rename security groups</TextBlock>
                            <TextBlock TextWrapping="Wrap">Use this tool when you need to rename all security groups that belong to a department or business unit within the nested security group standard.</TextBlock>
                        </StackPanel>
                    </ToolTip>
                </MenuItem.ToolTip>
            </MenuItem>
        </MenuItem>
    </MenuItem>
    <MenuItem Header="_OU">
        <MenuItem Header="Create Saved Queries" Name="SQXML">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Create Saved Queries</TextBlock>
                        <TextBlock TextWrapping="Wrap">Generate XML files to import as Saved Queries into ADUC.</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Migrate GPO" Name="MIGGPO">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Migrate GPO</TextBlock>
                        <TextBlock TextWrapping="Wrap">Copy GPO links from an existing OU to one of the migration staging areas</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Clean computers in OU" Name="UPCOMP">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Clean Computers in OU</TextBlock>
                        <TextBlock TextWrapping="Wrap">Standardizes the description field of computer objects and checks for stale/unused computers</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <Separator />
        <MenuItem Header="Prepare OU for Deletion" Name="CLEANOU">
            <MenuItem.Icon>
                <Image Source="$includePath/img/MS-stdlib-warning-128x128.png" />
            </MenuItem.Icon>
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Prepare OU for Deletion</TextBlock>
                        <TextBlock TextWrapping="Wrap">Prepares an existing OU for retirement by offering to remove unused GPO objects, computer accounts and security groups</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
    </MenuItem>
    <MenuItem Header="_Resources">
        <MenuItem Header="File Share Report" Name="SHRRPT">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">File Share Report</TextBlock>
                        <TextBlock TextWrapping="Wrap">Collects information about a network share and generates an HTML report</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
        <MenuItem Header="Printer Report" Name="PRTRPT">
            <MenuItem.ToolTip>
                <ToolTip>
                    <StackPanel Width="256">
                        <TextBlock FontWeight="Bold">Printer Report</TextBlock>
                        <TextBlock TextWrapping="Wrap">Collects information on specified printers and generates a report in HTML format</TextBlock>
                    </StackPanel>
                </ToolTip>
            </MenuItem.ToolTip>
        </MenuItem>
    </MenuItem>
    <MenuItem Header="_Help">
        <MenuItem Header="Open AD Standard KB article" Name="kbADSTD" />
        <MenuItem Header="Open DSCtrl KB article" Name="kbDSCTRL" />
        <Separator />
        <MenuItem Header="Create Tech 2 Tech ticket" Name="techToTech" />
        <Separator />
        <MenuItem Header="Check for updates" Name="hlpCheckUpdate" />
        <MenuItem Header="View application logs" Name="openLogs" />
        <MenuItem Header="About" Name="hlpAbout" />
    </MenuItem>
    <MenuItem Header="_Test">
        <MenuItem Header="Add to Software Group" Name="AddToGroup" />
        <MenuItem Header="Add to Printer Group" Name="AddToPrinter" />
        <MenuItem Header="Move and Disable Computers" Name="MoveAndDisable" />
        <MenuItem Header="Update Dell BIOS" Name="UpdateBIOS" />
        <MenuItem Header="Get Printer IP" Name="GetPrintIP" />
        <MenuItem Header="Ping" Name="Ping" />
        <MenuItem Header="PSRemote" Name="PSRemote" />
        <MenuItem Header="Generate GPO Report" Name="GPOReport" />
        <MenuItem Header="Canned Phrases" Name="CannedPhrases" />
        <MenuItem Header="Num Groups" Name="NumGroups" />
    </MenuItem>
</Menu>
<ToolBarTray DockPanel.Dock="Top">
    <ToolBar HorizontalAlignment="Stretch" VerticalAlignment="Top" HorizontalContentAlignment="Stretch">
        <Label Content="Computer Name:" Height="24" Margin="4" />
        <TextBox AcceptsReturn="false" Name="compNameText" Text="${ENV:ComputerName}" Width="128" />
        <Label Name="lblProgress" Content="Task progress:" Visibility="Collapsed" Height="24" Margin="4" />
        <ProgressBar Name="progress" Minimum="0" Maximum="100" Value="1" Visibility="Collapsed" Height="24" Margin="4" Width="300"></ProgressBar>
        <Button Content="Get Computer" Height="31" VerticalAlignment="Top" Width="89" RenderTransformOrigin="0.52,0.548" Name="GetComputer"/>
        <CheckBox Content="w/Properties" Height="32" VerticalAlignment="Top" Name="WithProperties"/>
        <ComboBox Name="PropertiesDropbox" Height="32" Margin="0" VerticalAlignment="Top" Width="120" StaysOpenOnEdit="True" Text="Properties">
            <ComboBoxItem>ComboBox Item #1</ComboBoxItem>
            <ComboBoxItem>ComboBox Item #2</ComboBoxItem>
            <ComboBoxItem>ComboBox Item #3</ComboBoxItem>
        </ComboBox>
    </ToolBar>
</ToolBarTray>
<ScrollViewer Name="outputScroller">
    <TextBlock TextWrapping="Wrap" Name="outputText" FontFamily="Courier New" />
</ScrollViewer>
</DockPanel>
</Window>
"@

$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
[void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
[xml]$XAML = $inputXML

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

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

# Storing this data in a Thread-Safe hash so we can pass it aroud to threads using Run Spaces (SOON...)

$Global:SyncHash = [hashtable]::Synchronized(@{})
$Global:SyncHash.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:SyncHash.add($_.Name, $Form.FindName($_.Name))
}

# Add a function that can be called to write log messages
$Global:SyncHash | Add-Member -Type ScriptMethod -name 'print' -Value {
    param($aMessage,$clearBeforePrinting)
    $nowTime = (get-date -Format 'T').toString()
    $Global:stringToPrint = "[ ${nowTime} ]   "

    # Beautify the string that the caller passed in
    $offsetString = ' ' * $Global:stringToPrint.length
    $Global:stringToPrint += $aMessage -replace "`n","`n${offsetString}"

    if ($clearBeforePrinting) {
        $this.Window.Dispatcher.invoke(
            [action]{ $this.outputText.Text = ""}
        )
    } 
    
    else {
        $this.Window.Dispatcher.invoke(
            [action]{ $this.outputText.Text += "${Global:stringToPrint}`n"}
        )
    }

    $this.Window.Dispatcher.invoke(
        [action]{ $this.outputScroller.ScrollToEnd()}
    )
    
    Write-Host "${Global:stringToPrint}"
}

# Add functions for manipulating the progress bar
$Global:SyncHash | Add-Member -Type ScriptMethod -name 'progressVisible' -Value {
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
$Global:SyncHash | Add-Member -Type ScriptMethod -name 'setProgressMax' -Value {
    param($maxValue)
    $this.Window.Dispatcher.invoke(
        [action]{ $this.progress.Maximum = "${maxValue}"}
    )
}
$Global:SyncHash | Add-Member -Type ScriptMethod -name 'setProgressValue' -Value {
    param($progressValue)
    $this.Window.Dispatcher.invoke(
        [action]{ $this.progress.Value = "${progressValue}"}
    )
}

#===========================================================================
# Additional SyncHash properties
#===========================================================================
$Global:SyncHash.Host = $host   # Provides access to host UI from children

#===========================================================================
# Create a Run Space environment that we use to spawn tasks inside of threads
#===========================================================================
# Inspired by: https://smsagent.wordpress.com/2015/09/07/powershell-tip-utilizing-runspaces-for-responsive-wpf-gui-applications/
$Runspace = [runspacefactory]::CreateRunspace()
$Runspace.ApartmentState = "STA"
$Runspace.ThreadOptions = "ReuseThread"
$Runspace.Open()
$Runspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
$Runspace.SessionStateProxy.SetVariable("includePath",$includePath)

#===========================================================================
# Use this space to add code to the various form elements in your GUI
#===========================================================================
$Global:SyncHash.Verify.Add_Click({

    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "verify-$((get-date).ToFileTime())") -IncludeInvocationHeader
        # Let the customer know we are about to try pinging their host
        $Global:SyncHash.print("Attempting to ping...", $false)

        # Pull the computer name from the form and store it locally
        # These calls are syncronized; so we want to access the form as little as possible
        $global:compToPing = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToPing = $Global:SyncHash.compNameText.text}
        )

        # Is the computer name valid?
        if ( (GlobalVarValidate $global:compToPing '^([\w\.-])+$')  ) {
            if (Test-Connection -Cn $global:compToPing -BufferSize 16 -Count 1 -ea 0 -quiet) {
                # yup
                # set the display with more verbose info
                $global:pingResult = Test-Connection $global:compToPing | Format-Table Address, ResponseTime | Out-string -Width 128
                $Global:SyncHash.print($global:pingResult, $false)
            } else {
                # nope there was a bad test. display message
                $Global:SyncHash.print("$($global:compToPing) did not respond", $false)
            }
        } else {
            # not a valid computer name.
            $Global:SyncHash.print("$($global:compToPing) is not a valid computer name.`n", $false)
        }
        Stop-Transcript
    }
    $PSinstance = [powershell]::Create().AddScript($Code)
    $PSinstance.Runspace = $Runspace
    $PSinstance.BeginInvoke()

})


$Global:SyncHash.CompSecGroup.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "Get-CompSecGroups-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\Invoke-ViewComputerMembership.ps1"

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        # is the computer name valid?
        if ( (GlobalVarValidate $global:compToTest '^([\w\.-])+$')  ) {
            $Global:SyncHash.print("Querying groups...", $false)
            Invoke-ViewComputerMembership -ComputerName $global:compToTest
        } else {
            # invalid computer name.
            $Global:SyncHash.print("${global:compToTest} is not a valid computer name.`n", $false)
        }
        Stop-Transcript
    }
    $PSexit = [powershell]::Create().AddScript($Code)
    $PSexit.Runspace = $Runspace
    $myJob = $PSexit.BeginInvoke()
})

$global:SyncHash.goToAsset.Add_Click({
    $global:compToTest = ''
    $syncHash.Window.Dispatcher.invoke(
            [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
    )
    (New-Object -Com Shell.Application).Open("https://asu.service-now.com/cmdb_ci_computer_list.do?sysparm_query=name=${global:compToTest}
    ")
})

$Global:SyncHash.LAPS.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "LAPS-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\Invoke-ClipBoardWindow.ps1"

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        # is the computer name valid?
        if ( (GlobalVarValidate $global:compToTest '^([\w\.-])+$')  ) {
            $Global:SyncHash.print("Does the computer exisist...", $false)

            # try and see if the computer exisits
            try{
                Get-ADComputer $global:compToTest -ErrorAction Stop
            }
            catch{
                # computer does not exisit in the AD
                $Global:SyncHash.print("The computer $($global:compToTest) does not exsist. $($_.Exception.Message)", $false)
            }
            # must exist.
            $Global:SyncHash.print("Yes `nLooking up password ...`n", $false)

            # grab the password
            $compObject = Get-ADComputer -Identity $global:compToTest -Property 'Name','ms-Mcs-AdmPwd','ms-Mcs-AdmPwdExpirationTime'

            # does the password exsist
            if ($compObject.'ms-Mcs-AdmPwd'.length -gt 0) {
                # looks like there is one
                # popup the password
                Invoke-ClipBoardWindow -Value $compObject.'ms-Mcs-AdmPwd' -Title $global:compToTest -Timer 120
            } else {
                # dispaly message that there is no password parameters.
                $Global:SyncHash.print("${global:compToTest} exisits in AD but could not find a password`n", $false)
            }
        } else {
            # invalid computer name.
            $Global:SyncHash.print("${global:compToTest} is not a valid computer name.`n", $false)
        }
        Stop-Transcript
    }
    $PSexit = [powershell]::Create().AddScript($Code)
    $PSexit.Runspace = $Runspace
    $myJob = $PSexit.BeginInvoke()
    })

$Global:SyncHash.BOMGAR.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "creategroup-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\Invoke-BomgarJump.ps1"

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        # is the computer name valid?
        if ( (GlobalVarValidate $global:compToTest '^([\w\.-])+$')  ) {
            $invokeResult = Invoke-BomgarJump -ComputerName $global:compToTest

            if ($invokeResult -ne $null) {
                $nestConfirm = PopupBox "There was a problem calling the Bomgar Representative Console, would you like to launch the web console instead?" "Bomgar" "yn"
                if ($nestConfirm -ne 'yes') {
                    exit 5
                }

                # launch web console
                (New-Object -Com Shell.Application).Open("https://bomgar.asu.edu/login/console")
            }
        } else {
            PopupBox "Could not read computer name from form. Please check that the computer name is entered and try again."
        }
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.CreateSecGrp.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "creategroup-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\CreateGroup.ps1"
        CreateGroup
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.ImportSecGrp.Add_Click({
    $code = {
        Stop-Transcript
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "bulkGroup-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\CreateGroup.ps1"
        ImportGroup
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.RenameSecGrp.Add_Click({
    $code = {
        Stop-Transcript
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "renameGroup-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\Rename-DepartmentGroups.ps1"
        Rename-DepartmentGroups
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.CalcToken.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "calcToken-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\Invoke-TokenSizeCalculator.ps1"
        Invoke-TokenSizeCalculator
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.ExDLRemoveMember.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "ExDLoffboard-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\popup.ps1"
        . "${includePath}\New-ExchangeConnection.ps1"
        . "${includePath}\Invoke-ExchangeOffboarding.ps1"
        Invoke-ExchangeOffboarding

        # A little clean-up after running for subsequent threads
        Get-PSSession | Remove-PSSession
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.SQXML.Add_Click({

    # get all Unit Group OUs in M.UTOSPA
    $ouListing = Get-ADOrganizationalUnit -Filter 'Name -like "*.Groups" -and Name -ne "M.UTOSPA.Groups"' -SearchBase $globalOUPath -SearchScope 2 | Sort-Object

    # prompt and have user select the Unit they would like to add the group to.
    $ouSelected = PopupListSelect $ouListing 'name' 'Please select the Unit:' 9 7

    # did they cancel?
    if ($ouSelected -eq $null) {
        PopupBox 'Canceled. Exiting.' "Information" "ok"
        return
    }

    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    $FolderBrowser.Description = 'Select a folder to save the xml files'
    [void]$FolderBrowser.ShowDialog()

    if ($FolderBrowser.SelectedPath -eq '') {
        PopupBox 'Canceled. Exiting.' "Information" "ok"
        return
    }

    # create two sub-folders
    mkdir "$($FolderBrowser.SelectedPath)\Security_Groups"
    mkdir "$($FolderBrowser.SelectedPath)\Computers"

    # various security groups
    $sgType = @('GPO', 'PRT', 'SHR','CMP','USR')

    # create xml file for the various security groups
    foreach ($sg in $sgType) {

        $sqxml = "<QUERY><NAME>M.UTOSPA.$ouSelected.Groups.$sg</NAME><DESCRIPTION>$ouSelected - $sg Security Groups</DESCRIPTION><DN>OU=M.UTOSPA.$ouSelected.Groups,OU=M.UTOSPA.$ouSelected,OU=M.UTOSPA,</DN><FILTERLASTLOGON>-1</FILTERLASTLOGON><LDAPQUERY>(&amp;(objectCategory=group)(objectClass=group)(cn=M.UTOSPA.UTO.Groups.$sg*))</LDAPQUERY><ONELEVEL>FALSE</ONELEVEL><COLUMNID>{6432329B-5691-4941-A544-0F518115B0F0}</COLUMNID></QUERY>"

        # write file
        $sqxml | Out-File "$($FolderBrowser.SelectedPath)\Security_Groups\$ouSelected_Group_$sg.xml"
    }

    # get computer security groups in the Unit OU
    $groupListing = Get-ADGroup -Filter "Name -like 'M.UTOSPA.$ouSelected.Groups.CMP.*'" -SearchBase "OU=M.UTOSPA.$ouSelected.Groups,OU=M.UTOSPA.$ouSelected,OU=M.UTOSPA,DC=ASURITE,DC=AD,DC=ASU,DC=EDU"

    # go through all the computer groups and create xml files
    foreach ($group in $groupListing) {
        $sqxml = "<QUERY><NAME>$($($group.Name).Replace('Groups.CMP.','')) Computers</NAME><DESCRIPTION>Computers in Group - $($group.Name)</DESCRIPTION><DN></DN><FILTERLASTLOGON>-1</FILTERLASTLOGON><LDAPQUERY>(&amp;(objectcategory=computer)(memberof=CN=$($group.Name),OU=M.UTOSPA.$ouSelected.Groups,OU=M.UTOSPA.$ouSelected,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu))</LDAPQUERY><ONELEVEL>FALSE</ONELEVEL><COLUMNID>{64E007CE-5A56-48BC-AE00-5B71C0B94F4C}</COLUMNID></QUERY>"

        # write file
        $sqxml | Out-File "$($FolderBrowser.SelectedPath)\Computers\$($group.Name)_Computer.xml"
    }

})

# run the Share Report function
$Global:SyncHash.SHRRPT.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "shareReport-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\Get-FSTree.ps1"
        Get-FSTree
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

# run the Printer Report function
$Global:SyncHash.PRTRPT.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "printer-report-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\PrinterReport.ps1"
        PrinterReport
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

# run GPO Migration
$Global:SyncHash.MIGGPO.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "miggpo-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\GPOMigration.ps1"
        GpoMigration
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.PropertiesDropbox.add_DropDownClosed({
    $syncHash.Window.Dispatcher.invoke([action]{$global:selectedProperty = $Global:SyncHash.PropertiesDropbox.text})
})

$Global:SyncHash.GetComputer.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "getcomputer-$((get-date).ToFileTime())") -IncludeInvocationHeader

        $syncHash.Window.Dispatcher.invoke([action]{$global:compToTest = $Global:SyncHash.compNameText.text})
        $syncHash.Window.Dispatcher.invoke([action]{$global:withProps = $Global:SyncHash.WithProperties.ischecked})

        $outputPath = ("{0}\outputs" -f (Split-path ${includePath} -parent))

        $global:SyncHash.print("Attempting to get $global:compToTest...", $false)
        $global:SyncHash.print($global:selectedProperty, $false)

        if($global:withProps) {
            try{
                $computer = Get-ADComputer $global:compToTest -Properties $global:properties | Out-string -Width 128
                $Global:SyncHash.print(($computer), $false)
                $computer | out-file -FilePath "$outputpath\Get-ADComputer.txt"
                $Global:SyncHash.print("Output saved to $outputpath\Get-ADComputer.txt", $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't find the computer, please try again."), $false)
                $Global:syncHash.print($_, $false)
            }
        }

        else {
            try{
                $computer = Get-ADComputer $global:compToTest | Out-string -Width 128
                $Global:SyncHash.print(($computer), $false)
                $computer | out-file -FilePath "$outputpath\Get-ADComputer.txt"
                $Global:SyncHash.print("Output saved to $outputpath\Get-ADComputer.txt", $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't find the computer, please try again."), $false)
                $Global:syncHash.print($_, $false)
            }
        }
        $Global:SyncHash.print("-------------------", $false)
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.AddToGroup.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "addtogroup-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\AddingToSoftWareDynamic.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.AddToPrinter.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "addtoprinter-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\addingtoprinters.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.MoveAndDisable.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "moveanddisable-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\MoveAndDisable.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.UpdateBIOS.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "updatebios-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\UpdatingDellBIOS.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.GetPrintIP.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "getprintip-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\GettingPrinterIP.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.Ping.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "ping-$((get-date).ToFileTime())") -IncludeInvocationHeader

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        $global:SyncHash.print("Attempting to ping $global:compToTest...", $false)

        $pingInfo = Test-NetConnection $global:compToTest

        if($pingInfo.PingSucceeded) {
            $global:SyncHash.print("Ping result: success", $false)
            $global:SyncHash.print(("IP address: {0}" -f $pingInfo.RemoteAddress.IPAddressToString), $false)
            $global:SyncHash.print(("MS: {0}" -f($pingInfo.PingReplyDetails.RoundTripTime), $false))

            $FQDN = nslookup $pingInfo.RemoteAddress.IPAddressToString

            foreach($line in $FQDN) {
                if($line.StartsWith("Name:")) {
                    $global:SyncHash.print(("FQDN: {0}" -f ($line -split " ")[4]), $false)
                }
            }
        }
        else {
            $global:SyncHash.print("Ping not successful", $false)
        }

        $global:SyncHash.print("----------")

    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})


$Global:SyncHash.PSRemote.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "ping-$((get-date).ToFileTime())") -IncludeInvocationHeader

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        $global:SyncHash.print("Attempting to start a remote powershell session with $global:compToTest...", $false)

        start-process powershell.exe -argument "-noexit -nologo -noprofile -command Enter-PSSession $global:compToTest"

        $global:SyncHash.print("----------")

    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.GPOReport.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "ping-$((get-date).ToFileTime())") -IncludeInvocationHeader

        $outputPath = ("{0}\outputs" -f (Split-path ${includePath} -parent))

        $global:compToTest = ''
        $syncHash.Window.Dispatcher.invoke(
                [action]{ $global:compToTest = $Global:SyncHash.compNameText.text}
        )

        $global:SyncHash.print("Generating GPO report for $global:compToTest...", $false)

        try {
            Get-GPResultantSetOfPolicy -computer $global:compToTest -ReportType html -Path ("{0}\GPOResults.html" -f $outputPath)
            $global:SyncHash.print(("Report generated successfully. Report saved to {0}\GPOResults.html" -f $outputPath), $false)
        }

        catch {
            $global:SyncHash.print($_, $false)
        }
       
        
        Start-Process ("{0}\GPOResults.html" -f $outputPath)

        $global:SyncHash.print("----------")

    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.CannedPhrases.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "cannedphrases-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\CannedResponses.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

# $Global:SyncHash.EditProperties.Add_Click({
#     $code = {
#         . "${includePath}\global.ps1"
#         Start-Transcript -Path (Join-Path $Global:logdir "cannedphrases-$((get-date).ToFileTime())") -IncludeInvocationHeader
#         . "${includePath}\Get-ADComputerProperties.ps1"
#     }
#     $PStask = [powershell]::Create().AddScript($Code)
#     $PStask.Runspace = $Runspace
#     $myJob = $PStask.BeginInvoke()
# })

# $Global:SyncHash.RBC.Add_Click({
#     $code = {
#         . "${includePath}\global.ps1"
#         Start-Transcript -Path (Join-Path $Global:logdir "cannedphrases-$((get-date).ToFileTime())") -IncludeInvocationHeader
#         . "${includePath}\rbcForm.ps1"
#     }
#     $PStask = [powershell]::Create().AddScript($Code)
#     $PStask.Runspace = $Runspace
#     $myJob = $PStask.BeginInvoke()
# })

# $Global:SyncHash.TestMenuItem.Add_Click({
#     $code = {
#         . "${includePath}\global.ps1"
#         Start-Transcript -Path (Join-Path $Global:logdir "cannedphrases-$((get-date).ToFileTime())") -IncludeInvocationHeader
#         $global:SyncHash.print("Hello World!", $false) 
#         . "${includePath}\insertfilehere.ps1"
#     }
#     $PStask = [powershell]::Create().AddScript($Code)
#     $PStask.Runspace = $Runspace
#     $myJob = $PStask.BeginInvoke()
# })

$Global:SyncHash.NumGroups.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "cannedphrases-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\NumGroups.ps1"
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

# Clean GPO, Computers and security groups from a selected OU
$Global:SyncHash.CLEANOU.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "cleanou-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\GPOMigration.ps1"

        # Warn the user that proceeding might result in things being deleted
        $nestConfirm = PopupBox "This tool will walk you through identifying and deleting assets from an OU that is being retired. Are you sure you want to continue? Doing so may result in objects being deleted." "Prepare OU for Deletion" "yn"
        if ($nestConfirm -ne 'yes') {
            exit 5
        }

        # Get the OU that the user would like to clean
        $cleanOU = Choose-ADOrganizationalUnit -HideNewOUFeature
        if ( $cleanOU -isnot [PSCustomObject]) {
            $Global:SyncHash.print("Canceled. Exiting.", $false)
            exit
        }
        $Global:SyncHash.print("Starting cleanup tool. Bacjups will be saved to:`n${Global:backupdir}", $false)

        # Create a list of GPO that are linked here
        $Global:SyncHash.print("Obtaining list of linked GPO", $false)
        $GPOList = Get-GPInheritance -Target $cleanOU.DistinguishedName | Select-Object -ExpandProperty 'GPOLinks'

        # Determine which policies are only linked here
        $singletons = New-Object -TypeName 'System.Collections.ArrayList'
        $GPOList | ForEach-Object {
            $reviewObject = Get-GPO -Guid $_.GpoID
            $Global:SyncHash.print("Calculating links for $($_.DisplayName)...", $false)
            if ($reviewObject -is [Microsoft.GroupPolicy.Gpo] ) {

                # Attempt to obtain a report for this object
                $xmlStuff = ''
                try {
                    $xmlStuff = Get-GPOReport -Guid $reviewObject.id -ReportType 'XML'
                    $xmlStuff = [XML]$xmlStuff
                } catch {
                    # Should we ignore, or tell the user?
                    # I'm going to ignore for now
                }

                if (($xmlStuff -is [XML]) -and
                    ($xmlStuff.GPO.LinksTo) ) {
                    $linkCount = $xmlStuff.GPO.LinksTo | Measure-Object | Select-Object -ExpandProperty 'count'

                    if ($linkCount -lt 2) {
                        $singletons.add($reviewObject)
                    }
                }
            }
        }

        # Generate a string containing a list of policies that will be unlinked
        $policyListString = New-Object -TypeName System.Text.StringBuilder
        if ($GPOList -is [Object[]]) {
            $GPOList | ForEach-Object {
                $policyListString.Append($_.DisplayName)
                $policyListString.Append("`n")
            }
        } elseif ($GPOList -is [Microsoft.GroupPolicy.GpoLink]) {
            $policyListString.Append($GPOList.DisplayName)
            $policyListString.Append("`n")
        }

        $nestConfirm = PopupBox @"
The following polices will be unlinked from $($cleanOU.name)

$($policyListString.toString())

Select Yes to continue, or No to abort OU cleaning
"@ "Unlink policies" "yn"
        if ($nestConfirm -ne 'yes') {
            exit 6
        }

        # Unlink policies
        $GPOList | ForEach-Object {
            ########
            ######## DANGER ZONE! Disabled during testing
            ########
            $Global:SyncHash.print("Unlink $($_.GpoId) which has name of $($_.DisplayName) from $($cleanOU.DistinguishedName)", $false)
            Remove-GPLink -Guid $_.GpoId -Target $cleanOU.DistinguishedName
        }

        # Delete unused policies
        if ( ($singletons -is [System.Collections.ArrayList]) -and ($singletons.count -gt 0)  ) {

            # Create a friendly string with policy names
            $deleteList = New-Object -TypeName System.Text.StringBuilder
            $singletons | ForEach-Object {
                $deleteList.Append($_.DisplayName)
                $deleteList.Append("`n")
            }

            # Display warning before deleting
            $nestConfirm = PopupBox @"
The following policies are no longer linked anywhere in this domain and can be deleted

$($policyListString.toString())

Select Yes to DELETE these policy objects, or No to abort OU cleaning
"@ "Purge unused policies" "yn"
            if ($nestConfirm -ne 'yes') {
                exit 6
            }

            # If we have not exited at this point: let's delete everything!
            $singletons | ForEach-Object {

                ########
                ######## DANGER ZONE! Disabled during testing
                ######## TODO: Replace simulation with real thing
                ########
                $Global:SyncHash.print("DELETE: $($_.GpoId) which has name of $($_.DisplayName)", $false)
                Remove-GPO -Guid $_.Id
            }
        }

        # Search for computers in this OU
        $Global:SyncHash.print("Searching for computers...", $false)
        $compsToDelete = Get-ADComputer -SearchBase $cleanOU.DistinguishedName -SearchScope 'onelevel' -Filter '*' -Properties 'name','description'

        # If computers were found, we need to delete them. Ask the user if it is ok to delete these computers
        if ( ($compsToDelete -is [Object[]]) -and ($compsToDelete.count -gt 0) -and ($compsToDelete[0] -is [Microsoft.ActiveDirectory.Management.ADComputer]) ) {

        # Let the user know we found some comuters and need to delete them
        $nestConfirm = PopupBox @"
While examining OU $($cleanOU.name)
$($compsToDelete.count.toString()) computer objects were found

Would you like to delete these computer objects?
"@ "Delete computer objects" "yn"
            if ($nestConfirm -ne 'yes') {
                exit 7
            }

            # Delete them
            $compsToDelete | ForEach-Object {

                ########
                ######## DANGER ZONE! Disabled during testing
                ######## TODO: Replace simulation with real thing
                ########
                $Global:SyncHash.print("DELETE: $($_.Name)", $false)
                Remove-ADComputer -ComputerName $_.Name
            }
        }

        # Search for security groups
        $Global:SyncHash.print("Searching for groups...", $false)
        $groupsToDelete = Get-ADGroup -SearchBase $cleanOU.DistinguishedName -SearchScope 'onelevel' -Filter '*' -Properties 'name','description'

        # Filter out groups created by MSS
        $groupsToDelete = $groupsToDelete | Where-Object -FilterScript {
            $_.name -inotmatch '(OUadmins|OUoperators|OUpermit|OUusers)$'
        }

        # If groups were found, we need to delete them. Ask the user if it is ok to delete these groups
        if ( ($groupsToDelete -is [Object[]]) -and ($groupsToDelete.count -gt 0) -and ($groupsToDelete[0] -is [Microsoft.ActiveDirectory.Management.ADGroup]) ) {

            # Let the user know we found some groups and need to delete them
            $nestConfirm = PopupBox @"
While examining OU $($cleanOU.name)
$($compsToDelete.count.toString()) group objects were found

Would you like to delete these group objects?
"@ "Delete group objects" "yn"
            if ($nestConfirm -ne 'yes') {
                exit 7
            }

            # Delete them
            $groupsToDelete | ForEach-Object {

                # Back up data before deletion
                $Global:SyncHash.print("BACKUP: group $($_.Name)", $false)
                $backupLocation = Join-Path $Global:backupdir "$($_.Name).xml"
                $groupDetails = Get-ADGroup -Identity $_ -Properties "*"
                $groupDetails | Export-Clixml -Path $backupLocation

                ########
                ######## DANGER ZONE! Disabled during testing
                ######## TODO: Replace simulation with real thing
                ########
                if (Test-Path =Path $backupLocation) {
                    $Global:SyncHash.print("DELETE: group $($_.Name)", $false)
                    Remove-ADGroup -Identity $_
                } else {
                    $Global:SyncHash.print("SKIPPING: Backup file not created, the following group WILL NOT be deleted;  $($_.Name)", $false)
                }
                
            }
        }
        $Global:SyncHash.print("Cleaning completed for $($cleanOU.name)", $false)
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

# run Update Computer Description
$Global:SyncHash.UPCOMP.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "clean-computers-$((get-date).ToFileTime())") -IncludeInvocationHeader
        . "${includePath}\Select-ADComputer.ps1"
        . "${includePath}\GPOMigration.ps1"
        $syncHash.print('Starting computer cleanup...')

        $Global:SyncHash.print("Starting cleanup tool. Backups will be saved to:`n${Global:backupdir}", $false)
        
        # Create array to hold computers which are not in ServiceNow
        $NotInSN = @()

        #$DefaultQuestion = 'Do you want to run with default settings? - Non-ServiceNow Check - 12 mo. inactivity - Search all child OUs'
        $DefaultQuestion = "Do you want to run with default settings?`n- Non-ServiceNow Check`n- 12 mo. inactivity`n- Search all child OUs"

        # Ask the user if they want to find systems not in ServiceNow
        if ((PopupBox $DefaultQuestion 'Settings' 'yn') -eq 'Yes') {
            $SNCandidates = $true
            $syncHash.print('Check for non-ServiceNow candidates.')
            $Months = '12'
            $syncHash.print('Search for systems that have not communicated with domain for 12 months.')
            $OUSearch = 'Subtree'
            $syncHash.print('Search for systems in all child OUs.')

        } else {

            # Ask the user if they want to find systems not in ServiceNow
            if ((PopupBox 'Do you want search all child OUs?' 'Search Subtree' 'yn') -eq 'Yes') {
                $OUSearch = 'Base'
                $syncHash.print('Search for systems in all child OUs.')
            } else {
                $OUSearch = 'OneLevel'
                $syncHash.print('Only search selected OU.')
            }
        
            # Ask the user if they want to find systems not in ServiceNow
            if ((PopupBox 'Do you want to find candidates to disable that are not in ServiceNow?' 'ServiceNow Test' 'yn') -eq 'Yes') {
                $SNCandidates = $true
                $syncHash.print('Check for non-ServiceNow candidates.')
            } else {
                $SNCandidates = $false
                $syncHash.print('Do not check for non-ServiceNow candidates.')
            }
        
            # Ask the user what is the month threshold 
            $Months = -1
            do {
                if ( ([int]$Months -le 5) -and ([int]$Months -ne -1) ) {
                    $popResult = ''
                    $popResult = PopupBox 'Months must be 6 or greater.' 'Information' 'ok'
                }   
                Add-Type -AssemblyName Microsoft.VisualBasic
                $Months = [Microsoft.VisualBasic.Interaction]::InputBox('Number of months without communication to AD? (min 6 month)', 'Month Threshold', "12")
            } while ( ([int]$Months -le 5) -or ($Months -eq '') )
            # Make sure it is all digits
            if (! ($Months -match '\d+')) {
                if ([int]$Months -ne -1) {
                    $popResult = ""
                    $popResult = PopupBox 'Invalid number of months. Exiting.' 'Error!' 'ok'
                }
                $syncHash.print('Aborting cleanup process: operation cancelled')
                exit 1
            } else {
                $syncHash.print("Search for systems that have not communicated with domain for $Months months.")        
            }

        }

        # Let the user know that we need to log into Service-Now
        $popResult = ""
        $popResult = PopupBox 'In order to query Service-Now for computer information, you will need to log into your ASURITE account. Once you click OK, an Internet Explorer window will appear, and require you to log into ASU SSO.' "Information" "oc"

        if ($popResult -ne 'OK') {
            $syncHash.print('Aborting cleanup process: operation cancelled')
            exit 1
        }

        # Launch IE
        $ie = $null
        try {
            $ie = new-object -ComObject "InternetExplorer.Application"
        }
        catch {
            $syncHash.print('Error starting Internet Explorer; cancelling process.')
            Exit 91
        }

        if ($ie -eq $null) {
            $syncHash.print('Error starting Internet Explorer; No error was thrown, but container variable is null.')
            Exit 92
        }

        # Start navigating while we set up IE
        $ie.Navigate('https://asu.service-now.com')

        # Configure IE to render JSON files
        if ( !(Test-Path -Path 'HKCR:') ) {
            New-PSDrive -Name 'HKCR' -PSProvider 'Registry' -Root 'HKEY_CLASSES_ROOT'
        }
        if ( !(Test-Path -Path "HKCR:\MIME\Database\Content Type\application$([char]0x2F)json") ) {
            $popResult = ""
            $popResult = PopupBox @"
Internet Explorer is currently not configured to allow rendering of JSON documents. In order to properly query Serivce Now, we need to set the following registry key:

HKCR:\MIME\Database\Content Type\application/json

Yhis action requires elevation. Please click OK and authorize the elevation to set this registry key. Not setting this registry key will cause queries to Service-Now to fail

"@ "Information" "oc"

            if ($popResult -ne 'OK') {
                $syncHash.print('Aborting cleanup process: operation cancelled')
                exit 1
            }

            try {
                $setupCode = '-STA -command New-PSDrive -Name "HKCR" -PSProvider "Registry" -Root "HKEY_CLASSES_ROOT"; $key = (get-item "HKCR:\").OpenSubKey("MIME\Database\Content Type\", $true); $key.CreateSubKey("application/json"); $key.Close(); New-ItemProperty -Path "HKCR:\MIME\Database\Content Type\application/json" -Name "CLSID" -Value "{25336920-03F9-11cf-8FD0-00AA00686F13}" -PropertyType "STRING" -Force; New-ItemProperty -Path "HKCR:\MIME\Database\Content Type\application/json" -Name "Encoding" -Value  ([Byte[]](08,00,00,00)) -PropertyType "BINARY" -Force' -Replace '"',"'"
                $setupProc = Start-Process -FilePath 'powershell.exe' -ArgumentList $setupCode -verb 'RunAs' -PassThru
                $setupProc.WaitForExit()
            }
            catch {
                $syncHash.print('Aborting cleanup process: failed to configure Internet Explorer for JSON rendering')
                exit 98
            }
        }

        # Customize the window
        $ie.silent = $true
        $ie.width = 1024
        $ie.height = 800
        $ie.addressbar = $false
        $ie.Toolbar = 0
        $ie.Left = 64
        $ie.Top = 64
        $ie.Visible = $true

        # All done with HKCR: Collect Garbage and disconnect
        [GC]::Collect()
        Remove-PSDrive -Name 'HKCR'

        # Wait for user to log in
        $syncHash.print('Waiting for Service-Now login...')
        Start-Sleep -Seconds 16
        while (($ie.LocationURL -eq $null) -or ($ie.LocationURL -inotmatch '^https://asu.service-now.com/*') ) {
            Start-Sleep -Seconds 2
        }

        # Make sure user didn't close IE before logging in
        if (($ie -eq $null) -or ($ie.locationURL -eq $null)) {
            $syncHash.print('Error communicating with Internet Explorer; window may have been closed. cancelling operation')
            Exit 92
        }

        # Hide IE
        $syncHash.print('Login detected!')
        $ie.Visible = $false

        # get the OU
        $cleanOU = Choose-ADOrganizationalUnit -HideNewOUFeature
        if ( $cleanOU -isnot [PSCustomObject]) {
            $Global:SyncHash.print("Canceled. Exiting.", $false)
            exit
        }

        # Set up parameters for computers query
        $searchParams = @{
            filter = '*'
            SearchBase = 'OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu'
            SearchScope = $OUSearch
            Properties = '*'

        }
        
        # Override default quey location if selected OU appears valid
        # Abort if not valid; this indicates that selection failed
        if ($cleanOU.DistinguishedName -Match '^(OU=).+') {
            $searchParams.SearchBase = $cleanOU.DistinguishedName
        } else {
            $Global:SyncHash.print("ERROR! Selected OU does not appear to be valid '$($cleanOU.DistinguishedName)'", $false)
            exit
        }

        # get all the computers in the selected OU
        $compsToClean = Get-ADComputer @searchParams
        $syncHash.print("Found $(($compsToClean | Measure-Object).count) computers")

        # Configure the progress bar
        $syncHash.setProgressValue(0)
        $syncHash.setProgressMax(($compsToClean | Measure-Object).count)
        $syncHash.progressVisible($true)

        # Update description fields
        $updateSuccesses = 0
        $updateFailures = 0
        $syncHash.print('Starting Service-Now query: first query may take up to 1 minute')
        $compsToClean | Foreach-Object {
            Start-Sleep -Seconds 3
            # Wait for IE to become available
            while ($ie.busy -ne $false) {
                Start-Sleep 5
            }
            $syncHash.print("$($_.Name) : Starting Service-Now query")
            $maxQueryTime = 120
            $queryURL = "https://asu.service-now.com/cmdb_ci_computer_list.do?sysparm_query=name=$($_.Name)&JSONv2"
            $queryResult = ''
            $queryError = $null
            $sys_id = $null
            $newDescription = ''

            $ie.navigate($queryURL)
            While ($ie.busy -and ($maxQueryTime -gt 0)) {
                Start-Sleep -Seconds 1
                $maxQueryTime = $maxQueryTime - 1
            }

            if (($ie.LocationURL -ne $queryURL) -or $ie.busy) {
                $queryError = 'Initial query failed'
            } else {
                if ($ie.Document.body.innerText) {
                    $queryResult = $ie.Document.body.innerText
                } else {
                    $queryError = "Document had no body"
                }
            }

            # Attempt to read the resulting JSON object
            if (($queryResult -Is [string]) -and !$queryError) {
                try {
                    $queryResult = ConvertFrom-Json -InputObject $queryResult
                } catch {
                    $queryError = "query returned invalid JSON"
                    $queryResult = $null
                }
            } elseif (!$queryError) {
                $queryError = 'Query returned no data'
                $queryResult = $null
            } else {
                $queryResult = $null
            }

            # Get sys_id
            if ( ($queryResult -is [PSCustomObject]) -and ($queryResult.records) -and (($queryResult.records | Measure-Object).count -gt 0) ) {
                $sys_id = $queryResult.records[0].sys_id
            } elseif (!$queryError) {
                $queryError = 'Unable to read sys_id from resulting payload'
            }
            
            # Add to array of computers not in ServiceNow
            if ($queryError) {
                $NotInSN += ,$($_.Name)
            }
            
            # Stop processing this computer if the ExtensionAttribute8 is already populated
            if (($_.ExtensionAttribute8 -is [String]) -and ($_.ExtensionAttribute8.length -gt 0)) {
                $queryError = "skipped because EA8 is not empty"
            }

            # Perform additional query to get inventory information
            $queryResult = $null
            if ($sys_id -and !$queryError) {
                $queryURL = "https://asu.service-now.com/alm_asset_list.do?sysparm_query=ci=${sys_id}&JSONv2"
                $ie.navigate($queryURL)
                While ($ie.busy -and ($maxQueryTime -gt 0)) {
                    Start-Sleep -Seconds 1
                    $maxQueryTime = $maxQueryTime - 1
                }
                if (($ie.LocationURL -ne $queryURL) -or $ie.busy) {
                    $queryError = "detail query failed"
                } else {
                    if ($ie.Document.body.innerText) {
                        $queryResult = $ie.Document.body.innerText
                    } else {
                        $queryError = "detail query bad data"
                    }
                }

                if ($queryResult -Is [string]) {
                    try {
                        $queryResult = ConvertFrom-Json -InputObject $queryResult
                    } catch {
                        $queryError = "query returned invalid JSON"
                        $queryResult = $null
                    }
                } else {
                    $queryResult = $null
                }
            }

            # generate description
            if ( ($queryResult -is [PSCustomObject]) -and ($queryResult.records) -and (($queryResult.records | Measure-Object).count -gt 0) ) {
                $newDescription = "$($queryResult.records[0].serial_number) \ $($queryResult.records[0].display_name)"
                $syncHash.print("$($_.Name) : New description is: ${newDescription}")
            }

            # Apply changes
            if ( ($newDescription -is [String]) -and ($newDescription.length -gt 8) -and !$queryError) {
                $oldDescription = $_.Description
                Set-ADComputer -Identity $_ -Description $newDescription
                Set-ADComputer -Identity $_ -Add @{extensionAttribute8=$oldDescription;description=$newDescription}
                $syncHash.print("$($_.Name) : changes applied")
                $updateSuccesses++
            } else {
                $syncHash.print("$($_.Name) : Skipping - ${queryError}")
                $updateFailures++
            }
            $syncHash.setProgressValue(($updateSuccesses + $updateFailures))
        }
        $syncHash.print("Finished updating computer descriptions: ${updateSuccesses} successes and ${updateFailures} skipped")
        if ($ie.LocationURL -ne $null) {
            $ie.Quit()
            Remove-Variable -Name 'ie'
        }
        $syncHash.progressVisible($false)

        # Perform some analysis to determine if any of these machines are unused
        $MonthsAgo = (Get-Date).AddMonths(-($Months))
        $compsToDisable = $compsToClean | Where-Object -FilterScript {

            # Perform a score based filter
            # If two tests match: Object passes through filter
            $objectScore = 0

            # Test if computer isn't in ServicNow
            if ($NotInSN -match $($_.Name) -and $SNCandidates){
                # Automatic candidate to disable
                $objectScore = 3
            } 

            # Test if last logon occurred more than entered timeframe
            if ($_.LastLogonTimestamp -and $_.LastLogonTimestamp -lt $MonthsAgo.ToFileTime()) {
                $objectScore++
            } elseif ($_.LastLogonTimestamp -eq $null) {
                $objectScore++
            }

            # Test for LAPS password with reset date older than entered timeframe
            if ($_.'ms-Mcs-AdmPwdExpirationTime' -and
                $_.'ms-Mcs-AdmPwdExpirationTime' -lt $MonthsAgo.ToFileTime() -and
                $_.'ms-Mcs-AdmPwdExpirationTime' -gt 0){
                    $objectScore++
            } elseif ($_.'ms-Mcs-AdmPwdExpirationTime' -eq $null) {
                $objectScore++
            } elseif (!$_.'ms-Mcs-AdmPwdExpirationTime') {
                $objectScore++
            }

            # Check if computer account password has been changed in the entered timeframe
            # MS says 30 days is default (REF=https://blogs.technet.microsoft.com/askds/2009/02/15/machine-account-password-process-2/)
            if ($_.pwdLastSet -lt $MonthsAgo.ToFileTime()) {
                $objectScore++
            }

            # Check OS name - Blank means probably never joined
            if ([String]::IsNullOrEmpty($_.OperatingSystem) ) {
                $objectScore++
            }

            # Exclude objects that are already disabled
            if ($_.Enabled -eq $false) {
                $objectScore = 0
            }

            # If all match; then the object should be disabled
            if ($objectScore -gt 2) {
                $true
            } else {
                $false
            }
        }

        # If some number of computers were selected to disabled
        # We need to give the user the option to select which ones to disable
        if ( ($compsToDisable | Measure-Object).count -gt 0) {
            Try {
                $compsToDisable = Select-ADComputer -ADComputers $compsToDisable -SNMissing $NotInSN -invertSelection
            } Catch {

                # An error was likely thrown if user cancelled; or something broke
                # Clear the list of computers so we don't accidentally disable something
                $compsToDisable = $null
                $syncHash.print("Operation cancelled: ${_}")
                Exit 9
            }

            # Disable any computers that were not selected
            $compsToDisable | Foreach-Object {
                $syncHash.print("Disabling computer $($_.name)")
                Set-ADComputer -Identity $_ -Add @{ExtensionAttribute7="Disabled on: $(Get-Date), lastLogon=$([datetime]::FromFileTime($_.LastLogon)), LAPS=$([datetime]::FromFileTime($_.'ms-Mcs-AdmPwdExpirationTime')), lastPassword=$([datetime]::FromFileTime($_.'pwdLastSet')), appVer=$($globalVersion)"}
                Disable-ADAccount -Identity $_
            }
        }
        $syncHash.print("Computer cleanup completed")
        Stop-Transcript
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$global:SyncHash.CopyLogMenuItem.Add_Click({
    Set-Clipboard -Value $Global:SyncHash.outputText.text
})

$global:SyncHash.ClearMenuItem.Add_Click({
    $Global:SyncHash.print("", $true)
})

$global:SyncHash.exitMenuItem.Add_Click({
    $Global:SyncHash.print("Exiting...", $false)
    $SyncHash.Window.Close()
    exit
})

$global:SyncHash.kbADSTD.Add_Click({
    (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0014939")
})

$global:SyncHash.kbDSCTRL.Add_Click({
    (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015713")
})

$global:SyncHash.techToTech.Add_Click({
    (New-Object -Com Shell.Application).Open("https://asu.service-now.com/sp?id=sc_cat_item&sys_id=99c6e66a13b02a4094ef7e776144b040")
})

$global:SyncHash.openLogs.Add_Click({
    (New-Object -Com Shell.Application).Open("${Global:logDir}")
})

# Checks for updates to the DSCtrl app
function checkForUpdates {
    $code = {
        . "${includePath}\global.ps1"
        . "${includePath}\Invoke-AskToUpdate.ps1"
        Start-Transcript -Path (Join-Path $Global:logdir "check-update-$((get-date).ToFileTime())") -IncludeInvocationHeader

        # Download XML file containing information about updates
        $syncHash.print('Checking for updates...')
        $updateFile = 'https://bitbucket.org/!api/2.0/snippets/utodso/A6geLg/HEAD/files/dsupdate.xml'
        $updateData = $null
        try {
            $updateData = Invoke-WebRequest -UseBasicParsing -Uri $updateFile
        }
        catch {
            $syncHash.print('Failed to contact update server, no updates found')
            $updateData = $null
        }

        if (($updateData -isnot [Microsoft.PowerShell.Commands.BasicHtmlWebResponseObject]) -or ($updateData.StatusCode -ne 200)) {
            $syncHash.print('Update check aborted')
            exit -1
        }

        # Parse downloaded data into an XML object
        try {
            $updateData = [XML]$updateData
        }
        catch {
            $updateData = $null
        }

        if (($updateData -isnot [XML]) -or ($updateData.data -eq $null)) {
            $syncHash.print('Update data validation error: aborting.')
            exit -1
        }

        # Test that we have write access to the DSCtrl folder
        $writeTest = Join-Path $includePath '\touch.test'
        try {
            New-Item -Path $writeTest -ItemType 'File'
        } catch {
            $syncHash.print('An update is available, but you do not have write access to the DSCtrl folder. Restart DSCtrl with elevated privilages to install this update.')
            exit -1
        } Finally {
            if (Test-Path -Path $writeTest) {
                Remove-Item -Path $writeTest -Force
            }
        }

        # Calculate version differance
        $versionparts = $globalversion.split('.')
        $localVersion = ([Int]$versionparts[0] * 10000) + ([Int]$versionparts[1] * 100) + ([Int]$versionparts[2])
        $remoteVersion = ([Int]$updateData.data.stable.version.major * 10000) + ([Int]$updateData.data.stable.version.minor * 100) + ([Int]$updateData.data.stable.version.patch)

        if ($remoteVersion -le $localVersion) {
            $syncHash.print('DSCtrl is up to date.')
            exit 0
        }

        $updateConfirm = Invoke-AskToUpdate -HTML @"
$($updateData.data.stable.html.'#cdata-section')
"@
        if ($updateConfirm -ne $true) {
            $syncHash.print("Update process cancelled. To update later, select the 'Check for Updates' option in the Help menu")
            exit 0
        }

        # An updated version appears to be available! Lets try downloading it
        $dlDestination = Join-Path $env:temp '\dsctrl-update.zip'
        $dlClient = New-Object System.Net.WebClient

        if (Test-Path -Path $dlDestination) {
            Remove-Item -Path $dlDestination -Force
        }

        try {
            $syncHash.print('Downloading update...')
            $dlClient.DownloadFile($updateData.data.stable.link,$dlDestination)
        }
        catch {
            $syncHash.print('Download error: update was not installed')
            exit -6
        }

        # Check the hash value
        $newHash = Get-FileHash -Path $dlDestination -Algorithm SHA256
        if ($newHash.Hash -ne $updateData.data.stable.SHA256) {
            $syncHash.print('Security error when checking signature of update file: cancelling installation')
            exit -7
        }

        # unblock the new file
        Unblock-File -Path $dlDestination

        # Extract the file; replace all files
        Expand-Archive -Path $dlDestination -DestinationPath (Split-Path -Path $includePath -Parent) -Force
        Remove-Item -Path $dlDestination -Force
        $syncHash.print('Update installed, restart the application to complete update process.')
        exit 1
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
}
$global:SyncHash.hlpCheckUpdate.Add_Click({
    checkForUpdates
})

$global:SyncHash.hlpAbout.Add_Click({
    $code = {
        . "${includePath}\global.ps1"
        $aboutMsg = @"
--------------------------------
|            DSCTrl            |
|    Deskside Control Panel    |
--------------------------------
Version: ${globalVersion}
Arizona State University

For support, use the 'Tech 2 Tech' option in the Help menu,
and submit your ticket to 'UTO Deskside Operations Management'
--------------------------------
Conncected to: ${global:adserver}
"@
        $syncHash.print($aboutMsg)
    }
    $PStask = [powershell]::Create().AddScript($Code)
    $PStask.Runspace = $Runspace
    $myJob = $PStask.BeginInvoke()
})

$Global:SyncHash.window.Add_Closing({
         Write-Output $null >> "$Home\.dsrmt\_close_"
})

# Skip update checking if running in Debug mode
if (!$global:debug) {
    checkForUpdates
}