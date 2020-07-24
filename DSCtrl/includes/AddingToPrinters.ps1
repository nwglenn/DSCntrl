[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCntrl_WPF"
        Title="Software Group" Height="238.166" Width="157.833" WindowStyle="ToolWindow" Topmost="True" ResizeMode="CanMinimize" WindowStartupLocation="CenterScreen">
    <Grid Margin="0,0,-8,-18">
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="160"/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="ASURITE:" HorizontalAlignment="Left" Margin="10,0,0,0" VerticalAlignment="Top" Height="26" Width="69"/>
        <TextBox Name="ASURITE" HorizontalAlignment="Left" Height="23" Margin="10,26,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="Department:" HorizontalAlignment="Left" Margin="10,54,0,0" VerticalAlignment="Top" Height="26" Width="96"/>
        <Button Content="Add" Name="btnAdd" HorizontalAlignment="Left" Margin="10,171,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <Button Content="Cancel" Name="btnCancel" HorizontalAlignment="Left" Margin="75,171,0,0" VerticalAlignment="Top" Width="55" Height="20"/>
        <ComboBox Name="DepartmentMenu" HorizontalAlignment="Left" Margin="10,80,0,0" VerticalAlignment="Top" Width="120" Height="22" SelectedIndex="-1">
            <ComboBoxItem>CLS</ComboBoxItem>
            <ComboBoxItem>EOSS</ComboBoxItem>
            <ComboBoxItem>ES</ComboBoxItem>
            <ComboBoxItem>MLFTC</ComboBoxItem>
            <ComboBoxItem>Provost</ComboBoxItem>
            <ComboBoxItem>WPC</ComboBoxItem>
        </ComboBox>
        <ComboBox Name="PrinterMenu" HorizontalAlignment="Left" Margin="10,133,0,0" VerticalAlignment="Top" Width="120"/>
        <Label Content="Printer:" HorizontalAlignment="Left" Margin="10,107,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
'@

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

$form.Topmost = $false


$btnCancel = $Form.FindName("btnCancel")
$deptMenu = $Form.FindName("DepartmentMenu")
$printMenu = $Form.FindName("PrinterMenu")
$btnAdd = $Form.FindName("btnAdd")
$ASURITE = $Form.FindName("ASURITE")


$currentDept = $deptMenu.selectionboxitem

$deptMenu.add_DropDownClosed({

    $currentDept = $deptMenu.Text

    switch($currentDept){
        "CLS"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.CLS,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
        "EOSS"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.EOSS,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
        "ES"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.ES,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
        "MLFTC"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.MLFTC,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
        "Provost"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.Provost,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
        "WPC"{
            $printMenu.items.Clear()
            $script:OU = "OU=P.POLY.WPC,OU=P.POLY,DC=asurite,DC=ad,DC=asu,DC=edu"
            $printers = Get-ADGroup -filter {name -like "*polyprint1*"} -SearchBase $script:OU | Sort-Object | select name
            $printerList = @()
            foreach($printer in $printers){
                $printerSplit = $printer -split("-")
                $printerSplitAgain = $printerSplit -split("}")
                $printerList += $printerSplitAgain[1]
            }
            $sortedList = $printerList | Sort-Object
            foreach($printer in $sortedList){
                $printMenu.items.Add($printer)
            }
            break
        }
    }

})



$btnAdd.add_click({
    $ASURITEText = $ASURITE.text
    $currentPrinter = $printMenu.selectionboxitem
    $fullName = ("*polyprint1-" + $currentPrinter)

    try{
        Get-ADUser $ASURITEText -ErrorAction Stop
    }
    Catch{
        $Global:SyncHash.print(("$ASURITEText is not a valid ASURITE. Please try again."), $false)
        return
    }

    try{
        $printerName = Get-ADGroup -filter {name -like $fullName} -searchbase $script:OU | select name -ErrorAction Stop -ErrorVariable myError
    }
    catch{
        $Global:SyncHash.print(("Unable to find the printer name. Please try again."), $false)
    }
  
    try{
        Add-ADGroupMember -Identity $printerName.name -Members "$ASURITEText" -ErrorAction Stop -ErrorVariable myError
        $Global:SyncHash.print(("Added $asuritetext to " + $printername.name), $false)
        Write-Host ("Added $ASURITEText to " + $printername.name)
    }
    Catch{
        $Global:SyncHash.print(("Unable to add $ASURITEText to " + $printerName), $false)
    }
})

$btnCancel.add_click({
    $form.close()
})

$Form.ShowDialog() | out-null