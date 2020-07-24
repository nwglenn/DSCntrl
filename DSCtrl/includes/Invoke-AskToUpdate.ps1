function Invoke-AskToUpdate {
    <#
        .SYNOPSIS
        Dsplays the passed HTML data and asks the user if they want to update
        .EXAMPLE
        Invoke-AskToUpdate -HTML $someHTML
        .PARAMETER HTML
        A string containing the valid HTML to be displayed
        .NOTES
        Returns True if the update button is clicked, False in al other cass
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='String to display')]
        [Alias('String')]
        [String]$HTML
    )

    # Form defenition
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:WpfApplication2"
Title="Release notes" Height="600" MinHeight="480" Width="800" MinWidth="570">
    <DockPanel Margin="8">
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Height="32">
            <Button x:Name="btnOk" Width="84" Height="24" TabIndex="100" Margin="4" IsDefault="True">Update Now</Button>
            <Button Name="btnCancel" Width="84" Height="24" TabIndex="101" Margin="4">Later</Button>
        </StackPanel>
        <GroupBox Header="A new update is available!" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" DockPanel.Dock="Bottom">
            <WebBrowser  Name="htmlView" Margin="0,0,0,0"/>
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
    $Global:updatesForm = [hashtable]::Synchronized(@{})
    $Global:updatesForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:updatesForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:updatesForm.add('response',$false)

    # Form visibility changed
    $Global:updatesForm.window.add_IsVisibleChanged({
        if ($Global:updatesForm.window.isVisible -eq $true) {
            $Global:updatesForm.window.topmost = $true
            $Global:updatesForm.window.topmost = $false
            $Global:updatesForm.window.focus()
        }
    })

    $Global:updatesForm.btnOk.Add_Click({
        $Global:updatesForm.response = $true
        $Global:updatesForm.window.close() | Out-Null
    })

    $Global:updatesForm.btnCancel.Add_Click({
        $Global:updatesForm.response = $false
        $Global:updatesForm.window.close() | Out-Null
    })

    $global:updatesForm.htmlView.NavigateToString($HTML)
    $global:updatesForm.Window.showDialog() | Out-Null
    Write-Output $Global:updatesForm.response
}
