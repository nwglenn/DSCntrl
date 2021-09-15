[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="TestPopout" Height="211.875" Width="220.624">
    <Grid Margin="0,0,2,-1">
        <Button Name="OkBtn" Content="Ok" HorizontalAlignment="Left" Margin="19,127,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Content="Cancel" HorizontalAlignment="Left" Margin="116,127,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBox Name="TestText" HorizontalAlignment="Left" Height="23" Margin="19,34,0,0" TextWrapping="Wrap" Text="Put text here!" VerticalAlignment="Top" Width="120"/>
        <CheckBox Name="TestCheck" Content="Check me!" HorizontalAlignment="Left" Margin="19,81,0,0" VerticalAlignment="Top"/>

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

$Global:testForm = [hashtable]::Synchronized(@{})
$Global:testForm.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:testForm.add($_.Name, $Form.FindName($_.Name))
}

$Global:testForm.window.add_IsVisibleChanged({
    if ($Global:testForm.window.isVisible -eq $true) {
        $Global:testForm.window.topmost = $true
        $Global:testForm.window.topmost = $false
        $Global:testForm.window.focus()
    }
})

$Global:testForm.OkBtn.add_click({
    $Global:SyncHash.print(("The text within the text box was: {0}" -f $testForm.TestText.text))
    $Global:SyncHash.print(("The checkbox was set to: {0}" -f $testForm.TestCheck.ischecked))
    $Global:SyncHash.print("-------------------", $false)
    $form.close()
})

$Global:testForm.CancelBtn.add_click({
    $form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null