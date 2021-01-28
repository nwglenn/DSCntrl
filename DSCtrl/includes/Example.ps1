[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="DSCtrl" Height="167" Width="262">
        <Grid HorizontalAlignment="Left" Height="193" VerticalAlignment="Top" Width="262" Margin="0,0,-8,0">
            <Label Content="Computer:" HorizontalAlignment="Left" Margin="25,29,0,0" VerticalAlignment="Top"/>
            <TextBox Name="MyTextBox" HorizontalAlignment="Left" Height="23" Margin="91,33,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
            <Button Name="MyButton" Content="Okay" HorizontalAlignment="Left" Margin="75,80,0,0" VerticalAlignment="Top" Width="75"/>
        </Grid>
</Window>
'@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

$MyButton.add_click({
    $global:SyncHash.print($MyTextBox.text, $false)
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null