[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Software Group" Height="260.499" Width="253.166" WindowStyle="ToolWindow" Topmost="True" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
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
        <Label Content="Software OU:" Name="OU" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="30" Width="105" FontSize="15" Padding="2,5,5,5"/>
        <TextBox HorizontalAlignment="Left" Height="22" Margin="10,29,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="212" FontSize="13" TextAlignment="Center" HorizontalScrollBarVisibility="Auto"/>
        <Button Content="Update" HorizontalAlignment="Left" Margin="158,56,0,0" VerticalAlignment="Top" Width="38" Height="19" FontSize="10"/>
    </Grid>
</Window>

'@

<#            <ComboBoxItem> Adobe Pro </ComboBoxItem>
            <ComboBoxItem> Adobe Pro 11 </ComboBoxItem>
            <ComboBoxItem> Adobe CC </ComboBoxItem>
            <ComboBoxItem> ArcGIS </ComboBoxItem>
            <ComboBoxItem> Citrix </ComboBoxItem>
            <ComboBoxItem> ERDAS </ComboBoxItem>
            <ComboBoxItem> GrassGIS </ComboBoxItem>
            <ComboBoxItem> Java - BLOCK </ComboBoxItem>
            <ComboBoxItem> Labview </ComboBoxItem>
            <ComboBoxItem> Lockdown Browser </ComboBoxItem>
            <ComboBoxItem> Logger Pro </ComboBoxItem>
            <ComboBoxItem> Maple </ComboBoxItem>
            <ComboBoxItem> Mathematica </ComboBoxItem>
            <ComboBoxItem> MatLab </ComboBoxItem>
            <ComboBoxItem> Maya </ComboBoxItem>
            <ComboBoxItem> Minitab </ComboBoxItem>
            <ComboBoxItem> Office 2016 64bit </ComboBoxItem>
            <ComboBoxItem> R </ComboBoxItem>
            <ComboBoxItem> SAS </ComboBoxItem>
            <ComboBoxItem> SPSS </ComboBoxItem>
            <ComboBoxItem> SPSSAMOS </ComboBoxItem>
            <ComboBoxItem> Turning Point </ComboBoxItem>
            <ComboBoxItem> Visio </ComboBoxItem>#>

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
<#try{#>$Form=[Windows.Markup.XamlReader]::Load( $reader )#}
#catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}


# Find the elements of the box
$btnAdd = $Form.FindName("btnAdd")
$btnCancel = $Form.FindName("btnCancel")
$computer = $Form.FindName("Computer")
$groupItem = $Form.FindName("GroupMenu")

# Attempt to add the computer to the selected group when the user clicks the "Add" button
$btnAdd.add_click({
    $computerText = $computer.Text # get the text that the user entered into the computer text box
    $currentItem = $groupItem.SelectionBoxItem # get the current item that is selected in the dropdown

    try{ # check to see if the computer entered is valid, get its distinguished name
        $computerInfo = Get-ADComputer $computerText -ErrorAction Stop
        $distinguishedName = $computerInfo.distinguishedName
    }
    catch{ # if it doesn't exist, print that it's not valid and exit
        $Global:SyncHash.print(("$computerText is not a valid computer, please try again."), $false)
        return
    }

    switch($currentItem){ # List of security groups for the UTO Poly OU
        "Adobe Pro" { # if the current item is Adobe Pro, attempt to add the computer to the adobe pro security group
            $OU = "P.POLY.Groups.Software.AdobeAcrobatPro"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName" -ErrorAction Stop 
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{ # if it gets an error, print that it couldn't be completed
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "MatLab" { # the rest of the cases are the same as the Adobe Pro entry, just the names are edited
            $OU = "P.POLY.Groups.Software.MatLab"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Adobe Pro 11" {
            $OU = "P.POLY.Groups.Software.AdobeAcrobatPro11"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Adobe CC" {
            $OU = "P.POLY.Groups.Software.AdobeCC"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "ArcGIS" {
            $OU = "P.POLY.Groups.Software.ArcGIS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Citrix" {
            $OU = "P.POLY.Groups.Software.Citrix"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "ERDAS" {
            $OU = "P.POLY.Groups.Software.ERDAS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "GrassGIS" {
            $OU = "P.POLY.Groups.Software.GrassGIS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Java - BLOCK" {
            $OU = "P.POLY.Groups.Software.Java-BLOCK"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Labview" {
            $OU = "P.POLY.Groups.Software.Labview"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Lockdown Browser" {
            $OU = "P.POLY.Groups.Software.LockdownBrowser"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Logger Pro" {
            $OU = "P.POLY.Groups.Software.LoggerPro"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Maple" {
            $OU = "P.POLY.Groups.Software.Maple"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Mathematica" {
            $OU = "P.POLY.Groups.Software.Mathematica"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Maya" {
            $OU = "P.POLY.Groups.Software.Maya"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Minitab" {
            $OU = "P.POLY.Groups.Software.Minitab"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Office 2016 64bit" {
            $OU = "P.POLY.Groups.Software.Office2016-64bit"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "R" {
            $OU = "P.POLY.Groups.Software.R"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "SAS" {
            $OU = "P.POLY.Groups.Software.SAS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "SPSS" {
            $OU = "P.POLY.Groups.Software.SPSS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "SPSSAMOS" {
            $OU = "P.POLY.Groups.Software.SPSSAMOS"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Turning Point" {
            $OU = "P.POLY.Groups.Software.TurningPoint"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        "Visio" {
            $OU = "P.POLY.Groups.Software.Visio"
            try{
                Add-ADGroupMember -Identity $OU -Members "$distinguishedName"
                $Global:SyncHash.print(("Added $computerText to $OU."), $false)
            }
            catch{
                $Global:SyncHash.print(("Couldn't add $computerText to $OU."), $false)
            }
            Break
        }
        default {
            $Global:SyncHash.print(("Please select a software group from the dropdown."), $false)
        }
    }
})

# if the user clicks the cancel button, close the window
$btnCancel.add_click({
    $form.close()
})

# show the form
$Form.ShowDialog() | out-null