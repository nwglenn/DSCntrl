[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Away Tools" Height="234" Width="262">
        <Grid HorizontalAlignment="Left" Height="207" VerticalAlignment="Top" Width="262" Margin="0,0,-8,-4">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <CheckBox Name="Enabled" Content="Enabled" HorizontalAlignment="Left" Margin="22,52,0,0" VerticalAlignment="Top" Height="15" Width="71"/>
        <CheckBox Name="LastLogon" Content="Last Logon" HorizontalAlignment="Left" Margin="131,52,0,0" VerticalAlignment="Top" Height="15" Width="86"/>
        <CheckBox Name="IPAddress" Content="IPv4 Address" HorizontalAlignment="Left" Margin="22,94,0,0" VerticalAlignment="Top" Height="15" Width="87"/>
        <CheckBox Name="LogonCount" Content="Logon Count" HorizontalAlignment="Left" Margin="131,94,0,0" VerticalAlignment="Top" Height="15" Width="92"/>
        <Button Name="SetBtn" Content="Set" HorizontalAlignment="Left" Margin="22,145,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Content="Cancel" HorizontalAlignment="Left" Margin="142,145,0,0" VerticalAlignment="Top" Width="75"/>
        <Label Content="Choose Which Properties to Query:" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top"/>
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

$xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

$SetBtn.add_click({
    $global:properties = @()
    if($enabled.ischecked) {
        $global:properties += "Enabled"
    }
    if($lastlogon.ischecked) {
        $global:properties += "LastLogon"
    }
    if($IPAddress.ischecked) {
        $global:properties += "ipv4address"
    }
    if($LogonCount.ischecked) {
        $global:properties += "LogonCount"
    }
    $syncHash.Window.Dispatcher.invoke(
})



$CloseBtn.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

