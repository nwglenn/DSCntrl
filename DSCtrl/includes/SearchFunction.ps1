
if ($includePath -eq $null) {
    $script:includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="SearchWindow" Height="277.5" Width="349.063">
    <Grid>
        <Button Name="SearchBtn" Content="Search" HorizontalAlignment="Left" Margin="191,22,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="RunBtn" Content="Run" HorizontalAlignment="Left" Margin="68,190,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Content="Cancel" HorizontalAlignment="Left" Margin="191,190,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBox Name="SearchBox" HorizontalAlignment="Left" Height="23" Margin="53,21,0,0" TextWrapping="Wrap" Text="Search" VerticalAlignment="Top" Width="120"/>
        <ListBox Name="ListBox" HorizontalAlignment="Left" Height="100" Margin="53,66,0,0" VerticalAlignment="Top" Width="238"/>
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

$Global:SearchFunction = [hashtable]::Synchronized(@{})
$Global:SearchFunction.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:SearchFunction.add($_.Name, $Form.FindName($_.Name))
}

$Global:SearchFunction.window.add_IsVisibleChanged({
    if ($Global:SearchFunction.window.isVisible -eq $true) {
        $Global:SearchFunction.window.topmost = $true
        $Global:SearchFunction.window.topmost = $false
        $Global:SearchFunction.window.focus()
    }
})

$searchHash = @{
    RemoteComputerInfo = ("remote", "computer", "information", "active directory", "ad")
    OUReport = ("ou", "organizational unit", "report", "information", "active directory", "ad")
    PrinterIP = ("ip", "printer", "mame")
    CountSecGroup = ("count", "security", "groups", "number", "tokens")
}

$Global:SearchFunction.SearchBtn.add_Click({
    $Global:SearchFunction.ListBox.items.clear()
    foreach ($item in $searchHash.GetEnumerator()) {
        $userInput = $Global:SearchFunction.SearchBox.Text
        $lowered = $userInput.ToLower()
        if ($item.value.contains($lowered)) {
            $Global:SearchFunction.ListBox.AddChild($item.Name)
        }
    }
})

$Global:SearchFunction.RunBtn.add_Click({
    if($Global:SearchFunction.ListBox.SelectedItem -eq "RemoteComputerInfo") {
        . "${includePath}\Get-ADComputerPopup.ps1"
    }

    elseif($Global:SearchFunction.ListBox.SelectedItem -eq "OUReport") {
        . "${includePath}\OUReporting.ps1"
    }

    elseif($Global:SearchFunction.ListBox.SelectedItem -eq "PrinterIP") {
        . "${includePath}\GettingPrinterIP.ps1"
    }

    elseif($Global:SearchFunction.ListBox.SelectedItem -eq "CountSecGroup") {
        . "${includePath}\NumGroups.ps1"
    }
})

$Global:SearchFunction.CancelBtn.add_Click({
    $form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null



