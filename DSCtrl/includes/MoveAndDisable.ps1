[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Move and Disable" Height="140.166" Width="157.833" WindowStyle="ToolWindow" Topmost="True" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,-8,-18">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="160"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Computer:" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="26" Width="69"/>
        <TextBox Name="Computer" HorizontalAlignment="Left" Height="23" Margin="10,26,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Content="Ok" Name="btnOk" HorizontalAlignment="Left" Margin="10,64,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <Button Content="Cancel" Name="btnCancel" HorizontalAlignment="Left" Margin="75,64,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
    </Grid>
</Window>

'@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

$form.Topmost = $false

$btnOk = $Form.FindName("btnOk")
$btnCancel = $Form.FindName("btnCancel")
$Computer = $Form.FindName("Computer")

$btnOk.add_click({
    $computerText = $computer.Text
    $computerInfo = Get-ADComputer $computerText # get the information about the current computer
    if($computerInfo.enabled -eq $true){
        try{ # if it's enabled, attempt to disable it
            Disable-ADAccount -identity $computerInfo.distinguishedname
            $Global:SyncHash.print(("Successfully disabled $computerText"), $false)
        }
        catch{
            $Global:SyncHash.print(("Unable to disable $computerText"), $false)
        }

        try{ # attempt to move it to the disabled computer OU
            Move-ADObject -identity $computerInfo.distinguishedname -TargetPath "OU=P.POLY.CLS.SocialSciences.Computers,OU=P.POLY.CLS.SocialSciences,OU=P.POLY.CLS,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $Global:SyncHash.print(("Moved $computerText to OU=P.POLY.CLS.SocialSciences.Computers,OU=P.POLY.CLS.SocialSciences,OU=P.POLY.CLS,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"), $false)
        }
        catch{
            $Global:SyncHash.print(("Unable to move $computerText"), $false)
        }
    }
})

$btnCancel.add_click({
    $form.close()
})



$Form.ShowDialog() | out-null