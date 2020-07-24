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
$computer = $Form.FindName("Computer")


$btnOk.add_click({
    $Global:SyncHash.print(("Updating BIOS... This may take a few minutes. DSControl may be unresponsive until the script has finished. If it has been more than a few minutes, please check the log file to see if something went wrong."), $false)

    $startTime = Get-Date

    # Enter in the name of the computer
    $computerText = $computer.Text
    
    # Start a session with the computer
    try{
        $Session = New-PSSession -ComputerName $computerText -ErrorAction Stop
    }
    catch{
        $Global:SyncHash.print(("Unable to create a session with $computerText"), $false)
        return
    }

    
    # Check to see if the files are there, if they are it deletes and re-adds them
    $PathCheck = Invoke-Command -Session $Session -ScriptBlock {
        $PathCheck = $true
        if((Test-Path -Path "${programfiles(x86)}\Dell\Command Configure\X86_64\cctk.exe") -and (Test-Path -Path "${programfiles(x86)}\Dell\CommandUpdate\dcu-cli.exe")){
            $PathCheck = $true
        }
        elseif ((Test-Path -Path "${env:ProgramFiles(x86)}\Dell") -and (!((Test-Path -Path "${programfiles(x86)}\Dell\Command Configure\X86_64\cctk.exe") -and (Test-Path -Path "${programfiles(x86)}\Dell\CommandUpdate\dcu-cli.exe")))) {
            $PathCheck = $false
            Remove-Item -Path "${env:ProgramFiles(x86)}\Dell\" -recurse
        }
        else{
            $PathCheck = $false
        }
        Return $PathCheck
    }
    
    # If it isn't there, it just does a straight copy of the files
    if($PathCheck -eq $false){
        Copy-Item "${env:ProgramFiles(x86)}\Dell\" -Destination "${env:ProgramFiles(x86)}\Dell\" -ToSession $session -Recurse -ErrorAction SilentlyContinue
    }
    
    # Run commands on the remote computer
    Invoke-Command -Session $Session -ScriptBlock {
        Suspend-BitLocker -MountPoint "c:" # suspend bitlocker

        # Run the cctk program to clear the BIOS password
        Set-Location "${env:ProgramFiles(x86)}\Dell\Command Configure\X86_64\"
        .\cctk.exe --setuppwd=  --valsetuppwd=UT0P0LY
    
        # Run the Dell Command Update - Command Line Interface with our policy file specifying to only perform a BIOS update
        Set-Location "${env:ProgramFiles(x86)}\Dell\CommandUpdate\"
        .\dcu-cli.exe /silent /policy "${env:ProgramFiles(x86)}\Dell\CommandUpdate\MyPolicy.xml" /log C:\logs
    
        # Run the cctk program again to re-set the BIOS password
        Set-Location "${env:ProgramFiles(x86)}\Dell\Command Configure\X86_64\"
        .\cctk.exe --setuppwd=UT0P0LY
    }
    
    $endTime = Get-Date
    
    $totalTime = $endTime - $startTime
    
    $props = [ordered]@{
        Computer = $computerText
        Time = $totalTime
    }
    
    $obj = New-Object psobject -Property $props
    
    $obj | Export-Csv -Path "C:\users\nwglenn\documents\BIOSUpdateAttempts.csv" -Append -NoTypeInformation

    $Global:SyncHash.print(("The command has finished running. Please check the log file to see if there was anything that failed."), $false)
})

$btnCancel.add_click({
    $Form.close()
})

$Form.ShowDialog() | out-null