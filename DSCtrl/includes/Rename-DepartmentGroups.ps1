<#
.SYNOPSIS
Self-contained function for renaming UTOSPA security groups
.DESCRIPTION
Graphical utility that walks a caller through the process of renaming security groups in bulk
.INPUTS
None. You cannot pipe objects to CreateGroup.ps1.
.OUTPUTS
None.
#>
function Rename-DepartmentGroups {

    # Print start of operation to main form
    $syncHash.print("Starting bulk renaming utility...")

    # Double-check that AD module is running
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    # Define the fancy new monolithic form
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:WpfApplication2"
Title="Bulk rename" Height="600" MinHeight="480" Width="888" MinWidth="570">
    <DockPanel Margin="8">
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Height="32">
            <Button Name="btnOk" Width="64" Height="24" TabIndex="100" Margin="4" IsDefault="True">OK</Button>
            <Button Name="btnCancel" Width="64" Height="24" TabIndex="101" Margin="4">Cancel</Button>
        </StackPanel>
        <GroupBox Header="Information" HorizontalAlignment="Stretch" DockPanel.Dock="Top">
            <StackPanel Orientation="Vertical" Margin="8">
                <StackPanel Orientation="Horizontal" Margin="2">
                    <Label Width="160">Unit:</Label>
                    <ComboBox Name="cmbUnitName" Width="256" TabIndex="2"></ComboBox>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="2">
                    <Label Width="160">Parent Group:</Label>
                    <ComboBox Name="cmbParentGroup" Width="256" TabIndex="3"></ComboBox>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="2">
                    <Label Width="160">Change:</Label>
                    <TextBox Name="txtGroupChange" Width="256" TabIndex="4"></TextBox>
                </StackPanel>
            </StackPanel>
        </GroupBox>
        <GroupBox Header="Group inheritance tree" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" DockPanel.Dock="Bottom">
            <DataGrid Name="groupGrid" Margin="8" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False" VerticalAlignment="Stretch" MinHeight="200" IsReadOnly="True">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Current name" Width="350" IsReadOnly="True" Binding="{Binding oldName}" />
                    <DataGridTextColumn Header="New name" Width="350" IsReadOnly="True" Binding="{Binding newName}" />
                    <DataGridTextColumn Header="Length" Width="50" IsReadOnly="True" Binding="{Binding length}" />
                </DataGrid.Columns>
            </DataGrid>
        </GroupBox>
    </DockPanel>
</Window>
"@
    
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
    $Global:renamingForm = [hashtable]::Synchronized(@{})
    $Global:renamingForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:renamingForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:renamingForm.add('okButtonClicked',$false)
    $Global:renamingForm.add('ADCache',$null)
    $Global:renamingForm.add('groupList',(New-Object -TypeName System.collections.ArrayList))
    $Global:renamingForm.groupGrid.itemsSource = $Global:renamingForm.groupList

    #
    # Form functins
    #

    # updateParentList - Clears and reloads the list of available parent security groups based on group type and unit selections
    $Global:renamingForm | Add-Member -Type ScriptMethod -name 'updateParentList' -Value {
        $selectedUnit = $this.cmbUnitName.SelectedValue.ToString()

        # Also abort if selected unit name is empty
        if ( ($selectedUnit -eq $null) -or ($selectedUnit -eq '') ) {
            return
        }

        # Generate new list of items
        $basePath = "OU=M.UTOSPA.${selectedUnit}.Groups,OU=M.UTOSPA.${selectedUnit},$globalOUPath"
        $groupList = $null
        $groupList = Get-ADGroup -SearchBase $basePath -filter {GroupCategory -eq "Security" -and name -like "*.Groups.CMP*" -and name -notlike "*.Groups.OU*" -and name -notlike "*.T" -and name -notlike "*.D" -and name -notlike "*.P" -and name -notlike "*.W" -and name -notlike "*.H"} | Sort-Object -Property 'Name'

        # Reset parent group selection
        if ($groupList -ne $null) {
            $this.cmbParentGroup.SelectedIndex = -1
            $this.cmbParentGroup.Items.Clear()
            $groupList | ForEach-Object {
                $this.cmbParentGroup.Items.Add($_.Name) | Out-Null
            }
            $this.cmbParentGroup.SelectedIndex = 0
        } else {
            $this.cmbParentGroup.SelectedIndex = -1
        }
    }

    # updateGroupList - Update list of groups and transformations. The function will reload from the local ADCache object, unless that object is nullified before calling. In that case, it will pull fresh data from AD to build a new cache.
    $Global:renamingForm | Add-Member -Type ScriptMethod -name 'updateGroupList' -Value {
        $selectedUnit = $this.cmbUnitName.SelectedValue
        $selectedParent = $this.cmbParentGroup.SelectedValue
        $selectedReplace = $this.txtGroupChange.text
        
        if ( ($selectedUnit -eq $null) -or ($selectedParent -eq $null) ) {
            return
        }

        # Generate new list of items
        $basePath = "OU=M.UTOSPA.${selectedUnit}.Groups,OU=M.UTOSPA.${selectedUnit},$globalOUPath"
        
        if ($Global:renamingForm.ADCache -eq $null) {
            $Global:renamingForm.ADCache = Get-ADGroup -SearchBase $basePath -filter {GroupCategory -eq "Security" -and name -like "*.Groups.*" -and name -notlike "*.Groups.OU*"} | Sort-Object -Property 'Name'
        }
        # Flush collection so we can re-populate it
        $Global:renamingForm.groupList.clear()

        # Generate the portion of the names being replaced
        $replaceString = $null
        $nameResult = ([Regex]'(?i)(CMP\.)(?<team>.+)').Match($selectedParent)
        if ($nameResult.Success) {
            $scopeFilter = $nameResult.Groups.Item('team').Value
            [Regex]$replaceString = $scopeFilter -Replace '^(.+)\.',''

            $Global:renamingForm.ADCache | Where-Object {$_.Name -like "*${scopeFilter}*"} | ForEach-Object {
                $gData = [PSCustomObject]@{
                    oldName=''
                    newName=''
                    length=0
                }

                $gData.oldName = $_.Name
                if ( ($selectedReplace -is [String]) -and ($selectedReplace.length -gt 0) ) {
                    $gData.newName = $replaceString.replace($_.Name, $selectedReplace, 1)
                    $gData.length = "$($gData.newName)".length
                }
                $Global:renamingForm.groupList.add($gdata) | Out-Null
            }
        }
        $global:renamingForm.groupGrid.items.refresh()
    }

    #
    # Event listeners
    #

    # Update list of parent groups when a unit is selected
    $Global:renamingForm.cmbUnitName.add_SelectionChanged({
        $Global:renamingForm.updateParentList()
    })

    $Global:renamingForm.cmbParentGroup.add_SelectionChanged({
        $Global:renamingForm.updateGroupList()
    })

    $Global:renamingForm.btnOk.add_Click({
        $errorMessage = $null

        # Test that there are items in the list
        if ($Global:renamingForm.groupList.count -lt 1) {
            $errorMessage = 'The list of groups to rename is empty. Change your selected options and try again'
        }

        # Test that provided replacement does not contain any strange characters
        if ($Global:renamingForm.txtGroupChange.text -inotmatch '^[a-z0-9\-_]+$') {
            $errorMessage = 'The provided replacement is not valid. Please only use characters, numbers dashes and underscores'
        }

        # Test that none of the new names are longer than 64 characters
        if ($errorMessage -eq $null) {
            $longNames = $null
            $Global:renamingForm.groupList | ForEach-Object {
                if ($_.newName.length -gt 64) {
                    $longNames = "${longNames}`n  $($_.newName)"
                }
            }
            if ($longNames -is [String]) {
                $errorMessage = "The following groups would contain more than 64 characters when renamed. Shorted these names manually, or modify the replacement string to be the same length or shorter than the original.`n${longNames}"
            }
        }

        # Close if no errors
        if ($errorMessage -eq $null) {
            $Global:renamingForm.okButtonClicked = $true
            $Global:renamingForm.window.close()
        } else {
            $Global:renamingForm.okButtonClicked = $false
            PopupBox $errorMessage
        }
    })

    $Global:renamingForm.btnCancel.add_Click({
        $Global:renamingForm.okButtonClicked = $false
        $Global:renamingForm.window.close()
    })

    $Global:renamingForm.txtGroupChange.add_TextChanged({

        # Nullify local cache, as it is no longer relevant
        $Global:renamingForm.ADCache = $null

        # Update list with new data
        # TODO: there is a faster way to do this that requires a little refactoring
        $Global:renamingForm.updateGroupList()
    })

    # Form visibility changed
    $Global:renamingForm.window.add_IsVisibleChanged({
        if ($Global:renamingForm.window.isVisible -eq $true) {
            $Global:renamingForm.window.topmost = $true
            $Global:renamingForm.window.topmost = $false
            $Global:renamingForm.window.focus()
        }
    })

    #
    # Form setup
    #

    # Populate unit list
    $unitFilter = [Regex]"(\w+)(.Groups)$"
    Get-ADOrganizationalUnit -Filter 'Name -like "*.Groups" -and Name -ne "M.UTOSPA.Groups"' -SearchBase $globalOUPath -SearchScope 2 | Sort-Object | Foreach-Object {
        $Global:renamingForm.cmbUnitName.Items.Add($unitFilter.Match($_.Name).Groups[1].Value) | Out-Null
    }
    if ($Global:renamingForm.cmbUnitName.Items.count -gt 0) {$Global:renamingForm.cmbUnitName.SelectedIndex = 0}

    # If the application settings store contains a unit property, select that unit by default
    $supGroup = $Global:appSettings.get('unit')
    if ($supGroup -is [String]) {
        :selectType for ($i = 0; $i -lt $Global:renamingForm.cmbUnitName.Items.Count; $i++) {
            if ($Global:renamingForm.cmbUnitName.Items[$i] -imatch $supGroup) {
                $Global:renamingForm.cmbUnitName.SelectedIndex = $i
                break selectType
            }
        }
    }

    $Global:renamingForm.window.showdialog()

    if ($Global:renamingForm.okButtonClicked -eq $true) {
        $Global:renamingForm.groupList | ft
        $Global:renamingForm.groupList | Foreach-Object {
            $syncHash.print("Rename: $($_.oldName) to $($_.newName)")
            try {
                Get-ADGroup -Identity $_.oldName | Select-Object -First 1 | Rename-ADObject -NewName $_.newName -PassThru | Set-ADGroup -SamAccountName $_.newName
            } catch {
                $syncHash.print("ERROR: $($_.message)")
            }
        }
    }
    $syncHash.print("Bulk rename utility exited")
}
