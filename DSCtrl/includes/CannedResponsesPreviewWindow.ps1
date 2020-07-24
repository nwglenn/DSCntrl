[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Away Tools" Height="451" Width="702">
    <Grid HorizontalAlignment="Left" Height="422" VerticalAlignment="Top" Width="702" Margin="0,0,-8,-2">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Button Name="CopyBtn" Content="Copy" HorizontalAlignment="Left" Margin="165,392,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
        <Button Name="CloseBtn" Content="Close" HorizontalAlignment="Left" Margin="263,392,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
        <RichTextBox Name="TextBlock" HorizontalAlignment="Left" Height="362" Grid.RowSpan="2" VerticalAlignment="Top" Width="668" Margin="10,10,0,0">
            <FlowDocument>
                <Paragraph>
                    <Run Name="TextBlockRun" Text="RichTextBox"/>
                </Paragraph>
            </FlowDocument>
        </RichTextBox>
        <Button Name="SaveBtn" Content="Save" HorizontalAlignment="Left" Margin="449,392,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
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

$settingsPath = ("{0}\settings" -f (Split-path ${includePath} -parent))

$Salutation.text = $toPerson

$TextBlockRun.text = Get-Clipboard

$CopyBtn.add_click({
    $TextBlock.SelectAll()
    Set-Clipboard -value $TextBlock.Selection.Text
})

$CloseBtn.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

