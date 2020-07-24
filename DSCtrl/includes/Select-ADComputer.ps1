function Select-ADComputer {
    <#
        .SYNOPSIS
        Provides a UI for selecting computer objects from a list
        .DESCRIPTION
        When this function is provided a collection of ADComputer objects, it displays that collection in a Windows form; allowing the end user to select from the list using checkboxes.

        When the Accept/OK button is clicked, the function filters the list based on which checkboxes were selected and returs the filtered list

        NOTE: While the default behavior is to return the list of SELECTED objects, you can specify the -Invert parameter to return a list of unselected objects

        In the event that the form is closed by any other means than the Accept/OK button, this function will throw an error and halt execuion
        .EXAMPLE
        Get-ADComputer -SearchBase (Get-ADRootDSE).defaultNamingContext -Filter "name -like 'UTO*'" | Select-ADComputer -Missing $NotInSN 
        .EXAMPLE
        Select-ADComputer -ADComputers $labComputers -Missing $NotInSN -Invert | Format-List name
        .PARAMETER ADComputers
        A collection (array) of ADComputer objects. This collection will be presented to the user, who will have an opportunity to select which members they would like to filter.
        .PARAMETER SNMissing
        A collection (array) of computer names. These computers are not in the ServiceNow instance.
        .PARAMETER invertSelection
        The default behavior of this function is to return the items the user selected from the collection. Specifying this switch changes that; and will result in the function returning the unselected items.
        .NOTES
        When calling this function, be aware that it will throw an exception and halt execution if the user closes the window or uses the cancel button; resulting in Null being returned in the data pipeline.
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='List of computers to display')]
        [Alias('Computer')]
        [Microsoft.ActiveDirectory.Management.ADComputer[]]$ADComputers,

        [Parameter(Mandatory=$True,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='List of computers not in ServiceNow')]
        [Alias('Missing')]
        [string[]]$SNMissing,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='If set, returns unselected rather than selected items')]
        [Alias('Invert')]
        [Switch]$invertSelection
    )

    # Form defenition
    $inputXML = @"
<Window x:Class="WpfPlayground.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfPlayground"
        mc:Ignorable="d"
        Title="Select computers" Height="600" Width="800">
    <DockPanel>
        <ToolBarTray DockPanel.Dock="Top">
            <ToolBar HorizontalAlignment="Stretch" VerticalAlignment="Top" HorizontalContentAlignment="Stretch">
                <Label Content="Select any computers you would like to exclude" />
                <Button Name="accept" Content="OK" Height="26" VerticalAlignment="Top" Width="75"/>
                <Button Name="cancel" Content="Cancel" Height="26" VerticalAlignment="Top" Width="75"/>
                <Button Name="reset" Content="Reset" Height="26" VerticalAlignment="Top" Width="75"/>
                <Button Name="save" Content="Save" Height="26" VerticalAlignment="Top" Width="75"/>
            </ToolBar>
        </ToolBarTray>
        <DataGrid Name="dataGrid" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False">
            <DataGrid.Columns>
                <DataGridCheckBoxColumn CanUserResize="False" Header="Exclude?" Binding ="{Binding exclude}" />
                <DataGridTextColumn Header="Computer name" IsReadOnly="True" Binding ="{Binding name}" />
                <DataGridTextColumn Header="Description" IsReadOnly="True" Binding ="{Binding description}" />
                <DataGridTextColumn Header="Last logon" IsReadOnly="True" Binding ="{Binding LastLogonDate}" />
                <DataGridTextColumn Header="Last password change" IsReadOnly="True" Binding ="{Binding pwdLastSet}" />
                <DataGridTextColumn Header="LAPS Expiration" IsReadOnly="True" Binding ="{Binding laps}" />
                <DataGridTextColumn Header="In ServiceNow" IsReadOnly="True" Binding ="{Binding ServiceNow}" />
            </DataGrid.Columns>
        </DataGrid>
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
    $Global:selectCompForm = [hashtable]::Synchronized(@{})
    $Global:selectCompForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:selectCompForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:selectCompForm.add('acceptButtonClicked',$false)

    # Load provided computer objects into the DataGrid
    $dataGridItems = New-Object -TypeName System.collections.ArrayList
    $ADComputers | ForEach-Object {
        if ($SNMissing -match $_.Name){
            $CompMissingSN = "FALSE"
        } else {
            $CompMissingSN = "TRUE"
        }

        $dataGridItems.add(
            [PSCustomObject]@{
                exclude=$false
                name = $_.name
                description = $_.description
                LastLogonDate = $_.LastLogonDate
                pwdLastSet = [DateTime]::FromFileTime($_.pwdLastSet)
                laps = [DateTime]::FromFileTime($_.'ms-Mcs-AdmPwdExpirationTime')
                ServiceNow = $CompMissingSN
                
            }
        ) | Out-Null
    }

    # Bind the data grid view to the new array of PSObjects
    $global:selectCompForm.dataGrid.itemssource = $dataGridItems

    # Add a listened to short-circuit the default functionality of the DataGrid
    $global:selectCompForm.dataGrid.add_BeginningEdit({

        # Normally, we would want to check which column is being edited, but as this form only has one editable column; I'm skipping that

        # Cancel the event so we can make our changes manually
        # This solves the problem of the last modified cell not being captured when the OK button is clicked
        $_.cancel = $true

        # Update the boolean propert manually by inverting it
        $_.Row.Item.exclude = !$_.Row.Item.exclude

        # Refresh the list... This is a very dirty way of doing this
        $global:selectCompForm.datagrid.items.refresh()
    })
    $global:selectCompForm.cancel.add_Click({
        $Global:selectCompForm.acceptButtonClicked = $false
        $global:selectCompForm.Window.close()
    })
    $global:selectCompForm.accept.add_Click({
        $Global:selectCompForm.acceptButtonClicked = $true
        $global:selectCompForm.Window.close()
    })
    $global:selectCompForm.save.add_Click({
        $save = New-Object -TypeName System.Windows.Forms.SaveFileDialog
        if ($save.ShowDialog() -eq 'OK') {
            $save.Filter = "CSV Files (*.csv)|*.csv"
            $dataGridItems | Export-Csv -Path  $save.Filename -NoTypeInformation # -Encoding UTF8 -Delimiter '|'
        }
    })
    $global:selectCompForm.reset.add_Click({
        $dataGridItems | Foreach-Object {
            $_.exclude = $false
        }
        $global:selectCompForm.datagrid.items.refresh()
    })

    # Display the form; this will block execution until it is closed
    $global:selectCompForm.Window.showDialog() | Out-Null

    # Was the form cancelled?
    if ($Global:selectCompForm.acceptButtonClicked -ne $true) {
        throw "Operation was cancelled"
        return $null
    }

    # Build a hash table relationship between computer names and exclusion status
    $exclusionTable = New-Object -TypeName System.Collections.Hashtable
    $dataGridItems | ForEach-Object {
        $exclusionTable.add($_.Name,$_.exclude)
    }

    # Dispose of global vars that are no longer needed
    $global:selectCompForm = $null

    # Filter the computer list based on the new hash table
    return $ADComputers | Where-Object -FilterScript {
        $test = $_.name
        if ($invertSelection) {
            !($exclusionTable."${test}")
        } else {
            $exclusionTable."${test}"
        }
    }
}
