[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="RBC" Height="144.834" Width="219.667">
    <Grid HorizontalAlignment="Left" Height="224" VerticalAlignment="Top" Width="412" Margin="0,0,0,-1">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Label Content="ASURITE:" HorizontalAlignment="Left" Margin="15,16,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="ASURITE" HorizontalAlignment="Left" Height="23" Margin="78,19,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="109"/>
        <Button Name="RunBtn" Content="Run" HorizontalAlignment="Left" Margin="15,57,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <CheckBox Name="Confirm" Content="Confirm" HorizontalAlignment="Left" Margin="113,60,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
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

if ($includePath -eq $null) {
    $includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}

$RunBtn.add_click({
    if($Confirm.ischecked()) {
        . "${includePath}\rbc.ps1" -action "del" -userID $ASURITE.text -dept "UTO"
    }
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

