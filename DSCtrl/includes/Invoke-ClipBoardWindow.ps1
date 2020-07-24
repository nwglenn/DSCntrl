function Invoke-ClipBoardWindow {
    <#
        .SYNOPSIS
        Displays the provided string in a window and allows the user to copy that string to the clipboard
        .EXAMPLE
        Invoke-ClipBoardWindow -Value "Hello, World!" -Timer 120
        .PARAMETER Value
        A string to be displayed and optionally copied to the clipboard
        .PARAMETER Timer
        An optional time in seconds to limit the amount of time the data is available on the clipboard before it is cleared
        .PARAMETER Title
        An optional string to be displayed as the title of the window
        .NOTES
        Timer value is not used if the user does not click the 'copy to clipboard' button
    #>
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='String to display')]
        [Alias('string')]
        [string]$value,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='An optional timer value in seconds')]
        [Alias('Seconds')]
        [int32]$timer = 60,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='If set, returns unselected rather than selected items')]
        [Alias('WindowTitle')]
        [String]$title
    )

    # Format the provided string with bolded numbers
    $strBuild = New-Object -TypeName System.Text.StringBuilder
    for ($i = 0; $i -lt $value.length; $i++) {
        if ($value[$i] -match "\d") {
            $strBuild.append("<Bold>") | Out-Null
            $strBuild.append($value[$i]) | Out-Null
            $strBuild.append("</Bold>") | Out-Null
        } else {
            $strBuild.append($value[$i]) | Out-Null
        }
    }
    $strBuild = $strBuild.ToString()

    # Form defenition
    $inputXML = @"
<Window x:Class="WpfPlayground.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:WpfPlayground"
        mc:Ignorable="d"
        Title="${title}" Height="280" Width="340">
    <StackPanel Orientation="Vertical" Margin="20">
        <Label Name="heading">LAPS Password (numbers are in bold)</Label>
        <RichTextBox Name="richTextBox" MinHeight="80" Padding="20" IsEnabled="False">
            <FlowDocument>
                <Paragraph TextAlignment="Center">
                    ${strBuild}
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <ProgressBar Name="statusBar" Height="20" Minimum="0" Maximum="${timer}" Value="${timer}"/>
        <Label Name="statusText" Height="40">   </Label>
        <StackPanel Orientation="Horizontal" VerticalAlignment="Center" HorizontalAlignment="center" Height="40">
            <Button Name="copy" Content="Copy to clipboard" Margin="8" />
            <Button Name="close" Content="Close" Margin="8" />
        </StackPanel>
    </StackPanel>
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
    $Global:stringDisplayForm = [hashtable]::Synchronized(@{})
    $Global:stringDisplayForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:stringDisplayForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:stringDisplayForm.add('timer',$timer)
    $Global:stringDisplayForm.add('admvalue',$value)

    # Add function to refresh progress indicator
    $Global:stringDisplayForm | Add-Member -Type ScriptMethod -name 'setProgress' -Value {
        param($newValue)
        $this.Window.Dispatcher.invoke(
            [action]{$this.statusBar.value = $newValue}
        )
        $this.Window.Dispatcher.invoke(
            [action]{$this.statusText.Content = "Password copied, will be cleared in $($newValue) seconds"}
        )
    }

    # Function to reset the form
    $Global:stringDisplayForm | Add-Member -Type ScriptMethod -name 'reset' -Value {
        $this.Window.Dispatcher.invoke(
            [action]{$this.statusBar.value = $this.timer}
        )
        $this.Window.Dispatcher.invoke(
            [action]{$this.statusText.Content = ""}
        )
    }

    # Form visibility changed
    $Global:stringDisplayForm.window.add_IsVisibleChanged({
        if ($Global:stringDisplayForm.window.isVisible -eq $true) {
            $Global:stringDisplayForm.window.topmost = $true
            $Global:stringDisplayForm.window.topmost = $false
            $Global:stringDisplayForm.window.focus()
        }
    })

    # Thread pool for countdown
    $counterRunspace = [runspacefactory]::CreateRunspace()
    $counterRunspace.ApartmentState = "STA"
    $counterRunspace.ThreadOptions = "ReuseThread"
    $counterRunspace.Open()
    $counterRunspace.SessionStateProxy.SetVariable("syncHash",$Global:stringDisplayForm)

    $Global:stringDisplayForm.copy.Add_Click({
        $taskCode = {
            Set-Clipboard -Value $Global:SyncHash.admvalue
            $timeRemain = $Global:SyncHash.timer
            :whileCopied while ($timeRemain -gt 0) {
                if ((Get-Clipboard -Format 'text') -eq $Global:SyncHash.admvalue ) {
                    $Global:SyncHash.setProgress($timeRemain)
                    Start-Sleep -Seconds 1
                    $timeRemain--
                } else {
                    break whileCopied
                }
            }
            $Global:SyncHash.reset()

            # Clear clipboard if it still contains the password
            if ((Get-Clipboard -Format 'text') -eq $Global:SyncHash.admvalue ) {
                Set-Clipboard -Value ""
            }
        }
        $psTimer = [powershell]::Create().AddScript($taskCode)
        $psTimer.Runspace = $counterRunspace
        $psTimer.BeginInvoke() | Out-Null
    })

    $Global:stringDisplayForm.close.Add_Click({
        $Global:stringDisplayForm.window.close()
    })

    $global:stringDisplayForm.Window.showDialog() | Out-Null

}
