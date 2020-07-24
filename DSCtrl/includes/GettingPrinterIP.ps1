[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Find Printer IP" Height="200.666" Width="156.333" WindowStyle="ToolWindow" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen" Topmost="True">
    <Grid Margin="0,0,-8,13">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="160"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Printer Name:" HorizontalAlignment="Left" Margin="15,0,0,0" VerticalAlignment="Top" Height="26" Width="95"/>
        <TextBox Name="Printer" HorizontalAlignment="Left" Height="23" Margin="10,26,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Content="Ok" Name="btnOk" HorizontalAlignment="Left" Margin="42,128,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <Label Content="IP:" HorizontalAlignment="Left" Margin="15,76,0,0" VerticalAlignment="Top"/>
        <Button Content="Get IP" Name="btnGetIP" HorizontalAlignment="Left" Margin="42,54,0,0" VerticalAlignment="Top" Width="55"/>
        <TextBox Name="IP" HorizontalAlignment="Left" Height="23" Margin="10,97,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
    </Grid>
</Window>
'@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}



$xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}
<#$btnOk = $Form.FindName("btnOk")
$printer = $Form.FindName("Printer")
$btnGetIP = $Form.FindName("GetIP")
$IP = $Form.FindName("IP") #>

$btnGetIP.add_click({
    $printerName = ("*" + $printer.Text + "*")
    $printerIP = get-printer -ComputerName polyprint1 | Sort-Object | Select-Object name, portname | where {$_.name -like $printerName}
    if(! ($null -eq $printerIP)){
        $Global:SyncHash.print(("The IP of " + $printerIP.name + " is: " + $printerip.portname + "."), $false)
        $IP.text = $printerIP.portname
    }
    else{
        $Global:SyncHash.print(($printer.Text + " does not coorelate with any printer on Polyprint1, please try again."), $false)
        $IP.text = "Printer not found"
    }

})

$btnOk.add_click({
    $form.close()
})


# $IP.focus = $true
#$form.activate()

$Form.ShowDialog() | out-null
#$form.focus()

$form.Topmost = $false