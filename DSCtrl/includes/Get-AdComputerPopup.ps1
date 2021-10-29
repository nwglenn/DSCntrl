[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="GetComputer" Height="437.062" Width="347.479">
    <Grid>
        <Button Name="RunBtn" Content="Run" HorizontalAlignment="Left" Margin="30,314,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Content="Cancel" HorizontalAlignment="Left" Margin="219,314,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBox Name="CompNameText" HorizontalAlignment="Left" Height="23" Margin="153,21,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="Computer Name:" HorizontalAlignment="Left" Margin="38,18,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="ipChk" Content="Ipv4 Address" HorizontalAlignment="Left" Margin="17,116,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="EnabledChk" Content="Enabled" HorizontalAlignment="Left" Margin="110,116,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="DNChk" Content="DitinguishedName" HorizontalAlignment="Left" Margin="200,116,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="DescriptionChk" Content="Description" HorizontalAlignment="Left" Margin="17,154,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="LogonTimeChk" Content="Last Logon" HorizontalAlignment="Left" Margin="110,154,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="OSChk" Content="Operating System" HorizontalAlignment="Left" Margin="200,154,0,0" VerticalAlignment="Top"/>
        <Separator HorizontalAlignment="Left" Height="17" Margin="17,174,0,0" VerticalAlignment="Top" Width="297"/>
        <Label Content="Additional Information (Requires remote permissions)" HorizontalAlignment="Left" Margin="23,180,0,0" VerticalAlignment="Top" Height="25"/>
        <CheckBox Name="BIOSChk" Content="BIOS" HorizontalAlignment="Left" Margin="38,220,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="UpTimeChk" Content="Up Time" HorizontalAlignment="Left" Margin="124,220,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="BitlockerChk" Content="Bitlocker" HorizontalAlignment="Left" Margin="219,220,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="LogonUserChk" Content="Last Logged in User" HorizontalAlignment="Left" Margin="48,265,0,0" VerticalAlignment="Top"/>
        <CheckBox Name="HardwareChk" Content="Hardware" HorizontalAlignment="Left" Margin="200,265,0,0" VerticalAlignment="Top"/>
        <Button Name="CSVBtn" Content="Save As CSV" HorizontalAlignment="Left" Margin="124,314,0,0" VerticalAlignment="Top" Width="75"/>
        <TextBox Name="OUTxt" HorizontalAlignment="Left" Height="23" Margin="153,70,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="OU:" HorizontalAlignment="Left" Margin="109,66,0,0" VerticalAlignment="Top"/>
        <Label Content="OR" HorizontalAlignment="Left" Margin="200,44,0,0" VerticalAlignment="Top"/>
        <RadioButton Name="CompRadio" Content="&#xD;&#xA;" HorizontalAlignment="Left" Margin="286,26,0,0" VerticalAlignment="Top" Height="16" Width="16" RenderTransformOrigin="0.188,0.314" IsChecked="True"/>
        <RadioButton Name="OURadio" Content="" HorizontalAlignment="Left" Margin="286,75,0,0" VerticalAlignment="Top" Width="14"/>
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

$Global:ADCompPopup = [hashtable]::Synchronized(@{})
$Global:ADCompPopup.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:ADCompPopup.add($_.Name, $Form.FindName($_.Name))
}

$Global:ADCompPopup.window.add_IsVisibleChanged({
    if ($Global:ADCompPopup.window.isVisible -eq $true) {
        $Global:ADCompPopup.window.topmost = $true
        $Global:ADCompPopup.window.topmost = $false
        $Global:ADCompPopup.window.focus()
    }
})

$Global:ADCompPopup.CompNameText.text = $Global:compToTest

$Global:ADCompPopup.RunBtn.add_click({

    if($Global:ADCompPopup.OURadio.ischecked) {
        $computers = Get-ADComputer -SearchBase $Global:ADCompPopup.OUTxt.text -filter * | select name
        $computerlist = @($computers.name)
    }
    else {
        $computerlist = @($Global:ADCompPopup.CompNameText.text)
    }

    $Global:syncHash.print("Gathering computer information...")
    foreach($computer in $computerList) {
        $propertiesSelected = @("name")
        if($Global:ADCompPopup.ipChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "ipv4address"
        }
    
        if($Global:ADCompPopup.EnabledChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "enabled"
        }
    
        if($Global:ADCompPopup.DNChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "distinguishedname"
        }
    
        if($Global:ADCompPopup.DescriptionChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "description"
        }
    
        if($Global:ADCompPopup.LogonTimeChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "LastLogonDate"
        }
    
        if($Global:ADCompPopup.OSChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "OperatingSystem"
        }
    
        $computerInfo = Get-ADComputer $computer -Properties $propertiesSelected | select $propertiesSelected
        if(!$computerInfo.ipv4address -and $Global:ADCompPopup.ipChk.ischecked) {
            $computerInfo.ipv4address = "None Found"
        }
    
        if(!$computerInfo) {
            $Global:SyncHash.print(("Could not find {0}, please try again." -f $computer), $false)
        }
    
        foreach($property in $computerInfo.psobject.properties) {
            $global:SyncHash.print(("{0}: {1}`n" -f $property.Name, $property.value), $false)
        }
    
        if($Global:ADCompPopup.BIOSChk.ischecked -or $Global:ADCompPopup.UpTimeChk.ischecked -or $Global:ADCompPopup.BitlockerChk.ischecked -or $Global:ADCompPopup.LogonUserChk.ischecked -or $Global:ADCompPopup.HardwareChk.ischecked) {
            $remoteAble = $false
            
            try {
                $Global:SyncHash.print(("Attempting to remote into {0} for additonal`ninformation, this may take a few seconds...`n" -f $computer))
                New-PSSession $computer -ErrorAction stop
                $remoteAble = $true
            }
            catch {
                $Global:SyncHash.print(("Could not remote into {0} for additonal information." -f $computer))
            }
    
            if($Global:ADCompPopup.BIOSChk.IsChecked -and $remoteAble) {
                $biosInfo = Invoke-Command -ComputerName $computer -Command {Get-WmiObject win32_bios | select SMBIOSBIOSVersion, Manufacturer}
                $Global:SyncHash.print("**BIOS Information**")
                $Global:SyncHash.print(("`tVersion: {0}" -f $biosInfo.SMBIOSBIOSVersion))
                $Global:SyncHash.print(("`tManufacturer: {0}`n" -f $biosInfo.Manufacturer))
            }
    
            if($Global:ADCompPopup.UpTimeChk.IsChecked -and $remoteAble) {
                $bootTime = Invoke-Command -ComputerName $computer -Command {(Get-CimInstance win32_operatingsystem).lastbootuptime}
                $upTime = (Get-Date) - $bootTime
                $Global:SyncHash.print("**Up Time Information**")
                $Global:SyncHash.print(("`tDays:`t`t{0}" -f $upTime.days))
                $Global:SyncHash.print(("`tHours:`t`t{0}" -f $upTime.hours))
                $Global:SyncHash.print(("`tMinutes:`t{0}`n" -f $upTime.minutes))
            }
    
            if($Global:ADCompPopup.BitlockerChk.IsChecked -and $remoteAble) {
                $bitlockerInfo = Invoke-Command -ComputerName $computer -Command {Get-BitLockerVolume}
                $Global:SyncHash.print("**Bitlocker Status**")
                $Global:SyncHash.print(("`tDrive Letter: {0}" -f $bitlockerInfo.mountpoint))
                $Global:SyncHash.print(("`tCapacity: {0}" -f $bitlockerInfo.capacitygb))
                $Global:SyncHash.print(("`tVolume Status: {0}`n" -f $bitlockerInfo.volumestatus))
            }
    
            if($Global:ADCompPopup.LogonUserChk.IsChecked -and $remoteAble) {
                $lastLoggedOnUser = Invoke-Command -ComputerName $computer -Command {
                    $Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
                    Get-ItemProperty -Path $Path | select LastLoggedOnDisplayName, LastLoggedOnUser
                }
                $Global:SyncHash.print("**Last Logged on User**")
                $Global:SyncHash.print(("Display Name: {0}" -f $lastLoggedOnUser.LastLoggedOnDisplayName))
                $Global:SyncHash.print(("Username: {0}`n" -f $lastLoggedOnUser.LastLoggedOnUser))
            }
    
            if($Global:ADCompPopup.HardwareChk.IsChecked -and $remoteAble) {
                $totalMemory = 0
                $computerModel = Invoke-Command -ComputerName $computer -command {Get-WmiObject Win32_computersystem}
                $videoCard = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_videocontroller}
                $memory = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_physicalmemory}
                $processor = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_processor}
                $disk = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_logicaldisk}
    
                $Global:syncHash.print("**Hardware Information**")
                $Global:SyncHash.print("--Model--")
                $Global:SyncHash.print(("`tManufacturer: {0}" -f $computerModel.Manufacturer))
                $Global:SyncHash.print(("`tModel: {0}" -f $computerModel.model))
                $Global:SyncHash.print("--Video Card--")
                foreach($card in $videoCard) {
                    $global:SyncHash.print(("`tName: {0}" -f $card.Name), $false)
                }
                $Global:SyncHash.print("--Memory--")
                foreach($stick in $memory) {
                    $stickSize = [math]::round($stick.capacity / 1GB)
                    $global:SyncHash.print(("`tSize: {0}GB" -f $stickSize))
                    $totalMemory = $totalMemory + $stickSize
                }
                $global:SyncHash.print(("`tTotal: {0}GB" -f $totalMemory), $false)
                $Global:SyncHash.print("--Processor--")
                $Global:SyncHash.print(("`tName: {0}" -f $processor.name))
                $Global:SyncHash.print("--Hard Drive--")
                foreach($drive in $disk) {
                    if($drive.drivetype -eq 3) {
                        $diskSize = [math]::round($drive.size / 1GB)
                        $remainningSpace = [math]::round($drive.FreeSpace / 1GB)
                        $Global:SyncHash.print(("`tDrive Letter: {0}" -f $drive.DeviceID))
                        $Global:SyncHash.print(("`tSize: {0}GB" -f $diskSize))
                        $Global:SyncHash.print(("`tFree Space: {0}GB`n" -f $remainningSpace))
    
                    } 
                }
            }
        }
    }
    $Global:SyncHash.print("`n-------------------`n", $false)
})

$Global:ADCompPopup.CSVBtn.add_click({

    $fileName = ("GetComputerInfo " + (get-date).tofiletime())

    if($Global:ADCompPopup.OURadio.ischecked) {
        $computers = Get-ADComputer -SearchBase $Global:ADCompPopup.OUTxt.text -filter * | select name
        $computerlist = @($computers.name)
    }
    else {
        $computerlist = @($Global:ADCompPopup.CompNameText.text)
    }

    $Global:syncHash.print("Gathering computer information...")
    foreach($computer in $computerList) {
        $props = [ordered]@{
            Name = $computer
        }
        $propertiesSelected = @("name")
        if($Global:ADCompPopup.ipChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "ipv4address"
        }
    
        if($Global:ADCompPopup.EnabledChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "enabled"
        }
    
        if($Global:ADCompPopup.DNChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "distinguishedname"
        }
    
        if($Global:ADCompPopup.DescriptionChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "description"
        }
    
        if($Global:ADCompPopup.LogonTimeChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "LastLogonDate"
        }
    
        if($Global:ADCompPopup.OSChk.ischecked) {
            $propertiesSelected = $propertiesSelected + "OperatingSystem"
        }
    
        $computerInfo = Get-ADComputer $computer -Properties $propertiesSelected | select $propertiesSelected
        if(!$computerInfo.ipv4address -and $Global:ADCompPopup.ipChk.ischecked) {
            $computerInfo.ipv4address = "None Found"
        }
    
        if(!$computerInfo) {
            $Global:SyncHash.print(("Could not find {0}, please try again." -f $computer), $false)
        }
    
        foreach($property in $computerInfo.psobject.properties) {
            $props[$property.name] = $property.value
        }
    
        if($Global:ADCompPopup.BIOSChk.ischecked -or $Global:ADCompPopup.UpTimeChk.ischecked -or $Global:ADCompPopup.BitlockerChk.ischecked -or $Global:ADCompPopup.LogonUserChk.ischecked -or $Global:ADCompPopup.HardwareChk.ischecked) {
            $remoteAble = $false
            
            try {
                $Global:SyncHash.print(("Attempting to remote into {0} for additonal`ninformation, this may take a few seconds...`n" -f $computer))
                New-PSSession $computer -ErrorAction stop
                $remoteAble = $true
            }
            catch {
                $Global:SyncHash.print(("Could not remote into {0} for additonal information." -f $computer))
            }
    
            if($Global:ADCompPopup.BIOSChk.IsChecked -and $remoteAble) {
                $biosInfo = Invoke-Command -ComputerName $computer -Command {Get-WmiObject win32_bios | select SMBIOSBIOSVersion, Manufacturer}
                $Global:SyncHash.print("**BIOS Information**")
                $Global:SyncHash.print(("`tVersion: {0}" -f $biosInfo.SMBIOSBIOSVersion))
                $Global:SyncHash.print(("`tManufacturer: {0}`n" -f $biosInfo.Manufacturer))
            }
    
            if($Global:ADCompPopup.UpTimeChk.IsChecked -and $remoteAble) {
                $bootTime = Invoke-Command -ComputerName $computer -Command {(Get-CimInstance win32_operatingsystem).lastbootuptime}
                $upTime = (Get-Date) - $bootTime
                $Global:SyncHash.print("**Up Time Information**")
                $Global:SyncHash.print(("`tDays:`t`t{0}" -f $upTime.days))
                $Global:SyncHash.print(("`tHours:`t`t{0}" -f $upTime.hours))
                $Global:SyncHash.print(("`tMinutes:`t{0}`n" -f $upTime.minutes))
            }
    
            if($Global:ADCompPopup.BitlockerChk.IsChecked -and $remoteAble) {
                $bitlockerInfo = Invoke-Command -ComputerName $computer -Command {Get-BitLockerVolume}
                $Global:SyncHash.print("**Bitlocker Status**")
                $Global:SyncHash.print(("`tDrive Letter: {0}" -f $bitlockerInfo.mountpoint))
                $Global:SyncHash.print(("`tCapacity: {0}" -f $bitlockerInfo.capacitygb))
                $Global:SyncHash.print(("`tVolume Status: {0}`n" -f $bitlockerInfo.volumestatus))
            }
    
            if($Global:ADCompPopup.LogonUserChk.IsChecked -and $remoteAble) {
                $lastLoggedOnUser = Invoke-Command -ComputerName $computer -Command {
                    $Path = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\Authentication\LogonUI"
                    Get-ItemProperty -Path $Path | select LastLoggedOnDisplayName, LastLoggedOnUser
                }
                $Global:SyncHash.print("**Last Logged on User**")
                $Global:SyncHash.print(("Display Name: {0}" -f $lastLoggedOnUser.LastLoggedOnDisplayName))
                $Global:SyncHash.print(("Username: {0}`n" -f $lastLoggedOnUser.LastLoggedOnUser))
            }
    
            if($Global:ADCompPopup.HardwareChk.IsChecked -and $remoteAble) {
                $totalMemory = 0
                $computerModel = Invoke-Command -ComputerName $computer -command {Get-WmiObject Win32_computersystem}
                $videoCard = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_videocontroller}
                $memory = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_physicalmemory}
                $processor = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_processor}
                $disk = Invoke-command -ComputerName $computer -command {Get-WmiObject Win32_logicaldisk}
    
                $Global:syncHash.print("**Hardware Information**")
                $Global:SyncHash.print("--Model--")
                $Global:SyncHash.print(("`tManufacturer: {0}" -f $computerModel.Manufacturer))
                $Global:SyncHash.print(("`tModel: {0}" -f $computerModel.model))
                $Global:SyncHash.print("--Video Card--")
                foreach($card in $videoCard) {
                    $global:SyncHash.print(("`tName: {0}" -f $card.Name), $false)
                }
                $Global:SyncHash.print("--Memory--")
                foreach($stick in $memory) {
                    $stickSize = [math]::round($stick.capacity / 1GB)
                    $global:SyncHash.print(("`tSize: {0}GB" -f $stickSize))
                    $totalMemory = $totalMemory + $stickSize
                }
                $global:SyncHash.print(("`tTotal: {0}GB" -f $totalMemory), $false)
                $Global:SyncHash.print("--Processor--")
                $Global:SyncHash.print(("`tName: {0}" -f $processor.name))
                $Global:SyncHash.print("--Hard Drive--")
                foreach($drive in $disk) {
                    if($drive.drivetype -eq 3) {
                        $diskSize = [math]::round($drive.size / 1GB)
                        $remainningSpace = [math]::round($drive.FreeSpace / 1GB)
                        $Global:SyncHash.print(("`tDrive Letter: {0}" -f $drive.DeviceID))
                        $Global:SyncHash.print(("`tSize: {0}GB" -f $diskSize))
                        $Global:SyncHash.print(("`tFree Space: {0}GB`n" -f $remainningSpace))
    
                    } 
                }
            }
    
        }
        $obj = New-Object psobject -Property $props
        $outputPath = ("{0}\outputs" -f (Split-path ${includePath} -parent))
        $obj | Export-Csv -Path $outputPath\$fileName.csv -NoTypeInformation -Append
    }
    $Global:SyncHash.print("File saved at: $outputPath\$fileName.csv")
    $Global:SyncHash.print("`n-------------------`n", $false)
})


$Global:ADCompPopup.CancelBtn.add_click({
    $form.close()
})


#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null