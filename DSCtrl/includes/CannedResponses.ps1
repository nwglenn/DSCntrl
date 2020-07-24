[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Away Tools" Height="273" Width="324">
        <Grid HorizontalAlignment="Left" Height="239" VerticalAlignment="Top" Width="314">
            <Grid.ColumnDefinitions>
                <ColumnDefinition/>
                <ColumnDefinition Width="0*"/>
            </Grid.ColumnDefinitions>
            <Grid.RowDefinitions>
                <RowDefinition Height="0*"/>
                <RowDefinition/>
            </Grid.RowDefinitions>
            <Label Content="To:" HorizontalAlignment="Left" Margin="41,27,0,0" Grid.RowSpan="2" VerticalAlignment="Top" RenderTransformOrigin="1.48,0.654" Height="26" Width="25"/>
            <TextBox Name="toField" HorizontalAlignment="Left" Height="23" Margin="66,31,0,0" Grid.RowSpan="2" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
            <Label Content="From:" HorizontalAlignment="Left" Margin="26,82,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Height="26" Width="40"/>
            <TextBox Name="fromField" HorizontalAlignment="Left" Height="23" Margin="66,86,0,0" Grid.RowSpan="2" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
            <Button Name="CopyBtn" Content="Copy" HorizontalAlignment="Left" Margin="26,185,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
            <Button Name="CloseBtn" Content="Close" HorizontalAlignment="Left" Margin="229,185,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
            <ComboBox Name="Template" HorizontalAlignment="Left" Margin="66,135,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="120">
                <ComboBoxItem>Closing</ComboBoxItem>
                <ComboBoxItem>Appointment</ComboBoxItem>
            </ComboBox>
            <Label Content="Template:" HorizontalAlignment="Left" Margin="5,131,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
            <Button Name="PreviewBtn" Content="Preview" HorizontalAlignment="Left" Margin="131,185,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
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

if ($includePath -eq $null) {
    $includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}

$settingsPath = ("{0}\settings" -f (Split-path ${includePath} -parent))


$Template.add_DropDownCLosed({
    $script:selectedTemplate = $Template.text
})

$CopyBtn.add_click({
    $fromPerson = $fromField.text
    $toPerson = $toField.text
    Set-Clipboard -value ("Hi {0},`n`nI'm glad that we could get your issue resolved today! If you have any questions or concerns please feel free to contact us again.`n`nThanks,`n`n{1}`nArizona State University`nUTO Poly Deskside Support" -f $toPerson, $fromPerson)
})

$PreviewBtn.add_click({
    $fromPerson = $fromField.text
    $toPerson = $toField.text

    switch($script:selectedTemplate) {
        "Closing" {
            $salutation = "Hi {0}," -f $toPerson
            $body = Get-Content "$settingsPath\Closing.txt"
            $sigName = $fromPerson
            $signature = Get-Content "$settingsPath\Signature.txt"
            $message = ($salutation, "", ($body -join "`n"), "", $sigName, ($signature -join "`n")) -join "`n"
            Set-Clipboard -value $message
        }
        "Appointment" {
            Set-ClipBoard -value ("Hi {0},`n`nWe have received your ticket and have decided that we need to meet with you in order to address this issue, what would be the best time for you to be able to meet with us?`n`nNOTE: Because of the Covid-19 health precautions, we reserve the right to refuse service to any person not wearing the proper PPE during their appointment.`n`nThanks,`n`n`{1}`nArizona State University`nUTO Poly Deskside Support" -f $toPerson, $fromPerson)
        }
    }

    . "${includePath}\CannedResponsesPreviewWindow.ps1"
})

$CloseBtn.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

