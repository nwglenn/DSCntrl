<#
    Title: get-lastlogondate-csv
    Version: 1.0
    Author: Jake Hartman
    Description: Takes a csv of computer names and exports a csv of AD lastlogondates for each.
    Input: Full path name to csv of computer names (i.e. C:\Users\jahartm3\computers.csv)
    Outout: LastLogonReport.csv
#>

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="Last Logon" Height="148.167" Width="235.667">
    <Grid HorizontalAlignment="Left" Height="115" VerticalAlignment="Top" Width="226">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Label Content="CSV Path:" HorizontalAlignment="Left" Margin="22,19,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="CSVPath" HorizontalAlignment="Left" Height="23" Margin="82,23,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Name="CloseBtn" Content="Close" HorizontalAlignment="Left" Margin="121,70,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <Button Name="RunBtn" Content="Run" HorizontalAlignment="Left" Margin="22,70,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
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

$RunBtn.add_click({
    
    $outputPath = ("{0}\outputs" -f (Split-path ${includePath} -parent))


    try {
        $ComputerNames = Import-CSV $csvPath.text -Header 'hostname'
        $MyFileName = "LastLogonReport.csv"
        $LastLogonReport = Join-Path $outputPath $MyFileName
    
        #Remove-Item $LastLogonReport -ErrorAction SilentlyContinue
    
        $STR = "Computer Name, Last Logon"
    
        Add-Content $LastLogonReport $STR
    
        Foreach ($computer in $ComputerNames) {
            
            $computerLastLogon = Get-ADComputer -Identity $computer.hostname -Properties * | select-object Name, LastLogonDate 
            
            $computerName = $computerLastLogon.Name
            $computerLastLogon = $computerLastLogon.LastLogonDate
    
            $STRNew = "$ComputerName , $computerLastLogon"
    
            Add-Content $LastLogonReport $STRNew
        }
        $Global:SyncHash.print(("Successfully created the report at {0}." -f $LastLogonReport))
    }

    catch {
        $Global:SyncHash.print("Couldn't find the CSV path, ensure you've entered a valid complete Windows path to the CSV file.")
    }

})

$CloseBtn.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

