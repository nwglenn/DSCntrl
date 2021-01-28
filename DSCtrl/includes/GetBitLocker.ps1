[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Get BitLocker" Height="247" Width="257">
        <Grid HorizontalAlignment="Left" Height="211" VerticalAlignment="Top" Width="257" Margin="0,0,-8,0">
            <Grid.ColumnDefinitions>
                <ColumnDefinition Width="125*"/>
                <ColumnDefinition Width="181*"/>
            </Grid.ColumnDefinitions>
            <Label Content="Computer:" HorizontalAlignment="Left" Margin="25,29,0,0" VerticalAlignment="Top"/>
            <TextBox Name="ComputerName" HorizontalAlignment="Left" Height="23" Margin="91,33,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="2"/>
            <Button Name="OkayBtn" Content="Okay" HorizontalAlignment="Left" Margin="35,163,0,0" VerticalAlignment="Top" Width="75"/>
            <Label Content="OR" HorizontalAlignment="Left" Margin="10,59,0,0" VerticalAlignment="Top" Grid.Column="1"/>
            <Label Content="BitLocker ID:" HorizontalAlignment="Left" Margin="15,86,0,0" VerticalAlignment="Top"/>
            <TextBox Name="BitLockerID" HorizontalAlignment="Left" Height="23" Margin="91,87,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="2"/>
            <Label Content="Search Base:" HorizontalAlignment="Left" Margin="15,117,0,0" VerticalAlignment="Top"/>
            <TextBox Name="SearchBase" HorizontalAlignment="Left" Height="23" Margin="91,120,0,0" TextWrapping="NoWrap" Text="DC=asurite,DC=ad,DC=asu,DC=edu" VerticalAlignment="Top" Width="120" Grid.ColumnSpan="2"/>
            <Button Name="CancelBtn" Content="Cancel" HorizontalAlignment="Left" Margin="31,163,0,0" VerticalAlignment="Top" Width="75" Grid.Column="1"/>
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

$OkayBtn.add_click({
    # Grab all the text from the text boxes for easier access
    $ComputerNameText = $ComputerName.text
    $SearchBaseText = $SearchBase.text
    $BitLockerIDText = $BitLockerID.text

    # If there is a computer name listed, just query that and nothing else.
    if($ComputerNameText.length -gt 1) {
        # Get the computer information from AD
        $objComputer = Get-ADComputer $ComputerNameText
        # Retrieve the recovery information from the AD object, making sure the password is included
        $Bitlocker_Object = Get-ADObject -Filter {objectclass -eq 'msFVE-RecoveryInformation'} -SearchBase $objComputer.DistinguishedName -Properties 'msFVE-RecoveryPassword'
        # The object may have multiple BitLocker entries, sort the list and grab the last one's (the most recent) recovery password
        $LatestPassword = ($Bitlocker_Object | select name, 'msFVE-RecoveryPassword' | Sort-Object)[-1].'msFVE-RecoveryPassword'
        #Print the information to the user
        $Global:SyncHash.print("The latest BitLocker recovery password for $ComputerNameText is:`n`n$LatestPassword")
        $Global:SyncHash.print("----------")
    }

    # If both the ID and Search Base fields are filled out search for the ID within the search base
    elseif($BitLockerIDText.length -gt 1 -and $SearchBaseText.length -gt 1) {
        # Create the wildcarded string to pinpoint the correct recovery ID
        $BitLockerIDSearch = "*{$BitlockerIDText-*"
        # Warn the user this may take a while
        $Global:SyncHash.print("Querying $SearchBaseText for the specified BitLocker ID, this could take a minute or two...")
        # Query AD searching for the ID within the msfve-recoveryinformation objects from the search base provided
        $Bitlocker_Object = (Get-ADObject -Filter {objectclass -eq 'msfve-recoveryinformation' -and name -like $BitLockerIDSearch} -SearchBase $SearchBaseText -Properties "msfve-recoverypassword")
        # Grab the password specifically
        $BitLockerPassword = $Bitlocker_Object.'msfve-recoverypassword'
        # Print the information to the user
        $Global:SyncHash.print("BitLocker Recovery Password for ${BitLockerIDText}:`n`n$BitLockerPassword")
        $Global:SyncHash.print("----------")
        
    }

    else {
        $Global:SyncHash.print("Please input either a valid computer name or both an ID and a Search Base (an OU's DistinguishedName) and try again.")
        $Global:SyncHash.print("----------")
    }
})

$CancelBtn.add_click({
    $Form.close()
})

$Form.ShowDialog() | out-null