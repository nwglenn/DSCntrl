# Window to assist with adding a user to a software group. The techician will provide the distinguished name of the OU that contains the software groups for their area
# Then they will click the update button, the Group dropdown will then be generated which the technician can select a group and type in a computer name, then click add

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Software Group" Height="273" Width="248.666" WindowStyle="ToolWindow" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen" Topmost="True">
    <Grid Margin="0,0,-54,-18">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="231"/>
            <ColumnDefinition Width="4"/>
        </Grid.ColumnDefinitions>
        <Label Content="Computer:" HorizontalAlignment="Left" Margin="10,113,0,0" VerticalAlignment="Top" Height="30" Width="80" FontSize="15" Padding="2,5,5,5"/>
        <TextBox Name="Computer" HorizontalAlignment="Left" Height="25" Margin="10,143,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" FontSize="15"/>
        <Label Content="Software Group:" HorizontalAlignment="Left" Margin="10,58,0,0" VerticalAlignment="Top" Height="30" Width="124" FontSize="15" Padding="2,5,5,5"/>
        <Button Content="Add" Name="btnAdd" HorizontalAlignment="Left" Margin="30,189,0,0" VerticalAlignment="Top" Width="60" Height="25" FontSize="15"/>
        <Button Content="Cancel" Name="btnCancel" HorizontalAlignment="Left" Margin="136,189,0,0" VerticalAlignment="Top" Width="60" Height="25" FontSize="15"/>
        <ComboBox Name="GroupMenu" HorizontalAlignment="Left" Margin="10,88,0,0" VerticalAlignment="Top" Width="212" Height="25" FontSize="13">
            <ComboBoxItem>Adobe Pro</ComboBoxItem>
            <ComboBoxItem>MatLab</ComboBoxItem>
        </ComboBox>
        <Label Content="OU Distinguished Name:" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="30" Width="200" FontSize="15" Padding="2,5,5,5"/>
        <TextBox Name="OU" HorizontalAlignment="Left" Height="25" Margin="10,29,0,0" VerticalAlignment="Top" Width="212" FontSize="15" MaxLines="1" TextWrapping="Wrap"/>
        <Button Content="Update" Name="btnUpdate" HorizontalAlignment="Left" Margin="158,58,0,0" VerticalAlignment="Top" Width="38" Height="19" FontSize="10"/>
    </Grid>
</Window>

'@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
<#try{#>$Form=[Windows.Markup.XamlReader]::Load( $reader )#}
#catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

$form.Topmost = $false

# Find the elements of the box
$btnAdd = $Form.FindName("btnAdd")
$btnCancel = $Form.FindName("btnCancel")
$computer = $Form.FindName("Computer")
$groupMenu = $Form.FindName("GroupMenu")
$OU = $Form.FindName("OU")
$btnUpdate = $Form.FindName("btnUpdate")

if(Test-Path ${includepath}\settings\DefaultSoftwareOU.txt){
    $OU.Text = (Get-Content ${includepath}\settings\DefaultSoftwareOU.txt)
}

$OUText = $OU.Text

if($OUText -ne ""){
    try{
        $groups = Get-ADGroup -SearchBase $OUText -filter * | Sort-Object
    }
    catch{
        $Global:SyncHash.print(("The current OU couldn't be retrieved, please enter the OU distinguished name again and press the update button."), $false)
    }
    $groupNames = $groups.name
    $groupMenu.items.Clear()

    foreach($group in $groupNames){
        $groupMenu.items.add($group)
    }
}


$btnUpdate.add_click({
    $OUText = $OU.Text
    try{
        $groups = Get-ADGroup -SearchBase $OUText -filter * | Sort-Object
    }
    catch{
        $Global:SyncHash.print(("The OU couldn't be retrieved, please enter the OU distinguished name again and press the update button again."), $false)
        return
    }

    $groupNames = $groups.name
    $groupMenu.items.Clear()

    foreach($group in $groupNames){
        $groupMenu.items.add($group)
    }

    $OUText | Out-File ${includePath}\Settings\DefaultSoftwareOU.txt
})

# Attempt to add the computer to the selected group when the user clicks the "Add" button
$btnAdd.add_click({
    $computerText = $computer.Text # get the text that the user entered into the computer text box
    $currentItem = $groupMenu.SelectionBoxItem # get the current item that is selected in the dropdown

    try{ # check to see if the computer entered is valid, get its distinguished name
        $computerInfo = Get-ADComputer $computerText -ErrorAction Stop
        $distinguishedName = $computerInfo.distinguishedName
    }
    catch{ # if it doesn't exist, print that it's not valid and exit
        $Global:SyncHash.print(("$computerText is not a valid computer, please try again."), $false)
        return
    }

    try{
        $currentGroup = Get-ADGroup $currentItem
    }
    catch{
        $Global:SyncHash.print(("Could not find the group: $currentItem"), $false)
    }

    try{
        Add-ADGroupMember -Identity $currentGroup -Members "$distinguishedName" -ErrorAction Stop
        $Global:SyncHash.print(("Added $computerText to $currentGroup"), $false)
    }
    catch{
        $Global:SyncHash.print(("Couldn't add $computerText to $currentGroup"), $false)
    }
})

# if the user clicks the cancel button, close the window
$btnCancel.add_click({
    $form.close()
})



# show the form
$Form.ShowDialog() | out-null

