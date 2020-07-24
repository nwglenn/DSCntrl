[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Get Computer" Height="140.166" Width="157.833" WindowStyle="ToolWindow" Topmost="True" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,-8,-18">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="160"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Computer:" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="26" Width="69"/>
        <TextBox Name="Computer" HorizontalAlignment="Left" Height="23" Margin="10,26,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Content="Ok" Name="btnOk" HorizontalAlignment="Left" Margin="10,74,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <Button Content="Cancel" Name="btnCancel" HorizontalAlignment="Left" Margin="75,74,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <CheckBox Content="Include Properties" Name="Properties" HorizontalAlignment="Left" Margin="10,54,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>

'@

#I added this comment.

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

# $xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

$form.Topmost = $false

$btnOk = $Form.FindName("btnOk")
$btnCancel = $Form.FindName('btnCancel')
$inputBox = $Form.FindName("Computer")
$properties = $Form.FindName("Properties")

$btnOk.add_click({
    if($properties.IsChecked){
        $inputText = $inputBox.text
        try{
            $computer = Get-ADComputer $inputText -properties * |  Out-string -Width 128
            $Global:SyncHash.print(($computer), $false)
        }
        catch{
            $Global:SyncHash.print(("Couldn't find the computer, please try again."), $false)
        }

    }
    else{
        $inputText = $inputBox.text
        try{
            $computer = Get-ADComputer $inputText |  Out-string -Width 128
            $Global:SyncHash.print(($computer), $false)
        }
        catch{
            $Global:SyncHash.print(("Couldn't find the computer, please try again."), $false)
        }

    }
})

$btnCancel.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

