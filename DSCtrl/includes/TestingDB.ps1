[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="Database" Height="254.562" Width="265.625">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <Label Content="INSERT" HorizontalAlignment="Left" Margin="10,10,0,0" VerticalAlignment="Top" Grid.Column="1" Height="26" Width="50"/>
        <Label Grid.ColumnSpan="2" Content="Username:" HorizontalAlignment="Left" Margin="27,41,0,0" VerticalAlignment="Top"/>
        <TextBox Name="InputUserText" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="23" Margin="101,45,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Name="CloseBtn" Grid.ColumnSpan="2" Content="Close" HorizontalAlignment="Left" Margin="92,166,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="InsertBtn" Grid.ColumnSpan="2" Content="Insert" HorizontalAlignment="Left" Margin="146,106,0,0" VerticalAlignment="Top" Width="75"/>
        <Label Grid.ColumnSpan="2" Content="SELECT" HorizontalAlignment="Left" Margin="12,100,0,0" VerticalAlignment="Top"/>
        <Button Name="GetUsersBtn" Grid.ColumnSpan="2" Content="Retrieve Users" HorizontalAlignment="Left" Margin="23,131,0,0" VerticalAlignment="Top" Width="85"/>
        <Label Grid.ColumnSpan="2" Content="User ID:" HorizontalAlignment="Left" Margin="41,69,0,0" VerticalAlignment="Top"/>
        <TextBox Name="UserIDText" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="23" Margin="101,73,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
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

Import-Module PSSQLite

$dbPath = ("{0}\database" -f (Split-path ${includePath} -parent))
$Database = "$dbpath\Database.sqlite"

$Global:TestingDB = [hashtable]::Synchronized(@{})
$Global:TestingDB.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:TestingDB.add($_.Name, $Form.FindName($_.Name))
}

$Global:TestingDB.window.add_IsVisibleChanged({
    if ($Global:TestingDB.window.isVisible -eq $true) {
        $Global:TestingDB.window.topmost = $true
        $Global:TestingDB.window.topmost = $false
        $Global:TestingDB.window.focus()
    }
})

$Global:TestingDB.InsertBtn.add_click({
    $insertQuery = ("INSERT INTO USERS (USERID, Name) VALUES ({0}, '{1}')" -f $Global:TestingDB.UserIDText.text, $Global:TestingDB.InputUserText.text)
    try{
        Invoke-SqliteQuery -DataSource $Database -query $insertQuery
        $global:SyncHash.print(("Successfully added {0} to the database." -f $Global:TestingDB.InputUserText.text))
    }
    catch{
        $global:SyncHash.print(("Could not add {0} to the database." -f $Global:TestingDB.InputUserText.text))
    }
})


$Global:TestingDB.GetUsersBtn.add_click({
    $query = ("SELECT * FROM USERS")
    $userInfo = Invoke-SqliteQuery -DataSource $Database -query $query
    $global:SyncHash.print($userInfo)
})

$Global:TestingDB.CloseBtn.add_click({
    $form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null