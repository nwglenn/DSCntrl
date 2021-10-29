if ($includePath -eq $null) {
    $script:includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}

[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="OU Report" Height="452.75" Width="643.133">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="OU Distinguished Name:" HorizontalAlignment="Left" Margin="3,10,0,0" VerticalAlignment="Top" Height="26" Width="140" Grid.Column="1"/>
        <TextBox Name="OUDNText" HorizontalAlignment="Left" Margin="10,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="252" Height="32" Grid.Column="1"/>
        <Label Content="Last Logon Cutoff Date:" HorizontalAlignment="Left" Margin="3,77,0,0" VerticalAlignment="Top" Height="26" Width="134" Grid.Column="1"/>
        <DatePicker Name="DatePicker" Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="151,108,0,0" VerticalAlignment="Top"/>
        <ComboBox Name="DateCombo" Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="142,81,0,0" VerticalAlignment="Top" Width="120">
            <ComboBoxItem Content="1 Month"/>
            <ComboBoxItem Content="3 Months"/>
            <ComboBoxItem Content="6 Months"/>
            <ComboBoxItem Content="1 Year"/>
        </ComboBox>
        <Button Name="RunBtn" Grid.ColumnSpan="2" Content="Run" HorizontalAlignment="Left" Margin="42,164,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Grid.ColumnSpan="2" Content="Cancel" HorizontalAlignment="Left" Margin="151,164,0,0" VerticalAlignment="Top" Width="74"/>
        <TreeView Name="ResultsTree" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="196" Margin="296,41,0,0" VerticalAlignment="Top" Width="323"/>
        <Label Grid.ColumnSpan="2" Content="Results:" HorizontalAlignment="Left" Margin="296,10,0,0" VerticalAlignment="Top"/>
        <Separator Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="20" Margin="62,202,0,0" VerticalAlignment="Top" Width="426" RenderTransformOrigin="0.5,0.5">
            <Separator.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="90.1"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Separator.RenderTransform>
        </Separator>
        <Label Grid.ColumnSpan="2" Content="Total:" HorizontalAlignment="Left" Margin="296,242,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Grid.ColumnSpan="2" Content="Computers:" HorizontalAlignment="Left" Margin="296,268,0,0" VerticalAlignment="Top" FontSize="10" RenderTransformOrigin="0.48,0.576"/>
        <Label Name="ComputersTotal"  Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,268,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Grid.ColumnSpan="2" Content="Groups:" HorizontalAlignment="Left" Margin="296,321,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="GroupsTotal" Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,321,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Grid.ColumnSpan="2" Content="Inactive Computers:" HorizontalAlignment="Left" Margin="296,285,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="InactiveTotal" Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,285,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Grid.ColumnSpan="2" Content="Empty Groups:" HorizontalAlignment="Left" Margin="296,339,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="EmptyGroupTotal" Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,339,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Grid.ColumnSpan="2" Content="Users Added By Name:" HorizontalAlignment="Left" Margin="296,358,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="UserByNameTotal" Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,358,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Grid.ColumnSpan="2" Content="Past Last Login:" HorizontalAlignment="Left" Margin="296,304,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="PastLogin" Grid.ColumnSpan="2" Content="0" HorizontalAlignment="Left" Margin="409,304,0,0" VerticalAlignment="Top" FontSize="10"/>
        <Label Name="StatusText" Grid.ColumnSpan="2" Content="" HorizontalAlignment="Left" Margin="107,204,0,0" VerticalAlignment="Top"/>
        <Label Grid.ColumnSpan="2" Content="Legend:" HorizontalAlignment="Left" Margin="452,242,0,0" VerticalAlignment="Top" FontWeight="Bold"/>
        <Label Grid.ColumnSpan="2" Content="" HorizontalAlignment="Left" Margin="553,278,0,0" VerticalAlignment="Top" Background="Red" Width="36" Height="8"/>
        <Label Grid.ColumnSpan="2" Content="Disabled:" HorizontalAlignment="Left" Margin="452,268,0,0" VerticalAlignment="Top" FontSize="10" RenderTransformOrigin="0.48,0.576"/>
        <Label Grid.ColumnSpan="2" Content="Last Logon Beyond:" HorizontalAlignment="Left" Margin="452,285,0,0" VerticalAlignment="Top" FontSize="10" RenderTransformOrigin="0.48,0.576"/>
        <Label Grid.ColumnSpan="2" Content="" HorizontalAlignment="Left" Margin="553,294,0,0" VerticalAlignment="Top" Background="Yellow" Width="36" Height="8"/>
        <Button Name="MoveBtn" Grid.ColumnSpan="2" Content="Move Selected" HorizontalAlignment="Left" Margin="436,374,0,0" VerticalAlignment="Top" Width="84"/>
        <Label Grid.ColumnSpan="2" Content="Good Computer:" HorizontalAlignment="Left" Margin="452,304,0,0" VerticalAlignment="Top" FontSize="10" RenderTransformOrigin="0.48,0.576"/>
        <Label Grid.ColumnSpan="2" Content="" HorizontalAlignment="Left" Margin="553,312,0,0" VerticalAlignment="Top" Background="LightGreen" Width="36" Height="8"/>
        <Label Grid.ColumnSpan="2" Content="Move To:" HorizontalAlignment="Left" Margin="436,336,0,0" VerticalAlignment="Top"/>
        <TextBox Name="MoveToBox" Grid.ColumnSpan="2" HorizontalAlignment="Left" Height="23" Margin="499,339,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Name="UpdateBtn" Grid.ColumnSpan="2" Content="Update Location" HorizontalAlignment="Left" Margin="525,374,0,0" VerticalAlignment="Top" Width="94"/>
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

$Global:OUReporting = [hashtable]::Synchronized(@{})
$Global:OUReporting.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:OUReporting.add($_.Name, $Form.FindName($_.Name))
}

$Global:OUReporting.window.add_IsVisibleChanged({
    if ($Global:OUReporting.window.isVisible -eq $true) {
        $Global:OUReporting.window.topmost = $true
        $Global:OUReporting.window.topmost = $false
        $Global:OUReporting.window.focus()
    }
})

function Add-ChildNodes($parentNode, $currentItems) {
    foreach($item in $currentItems) {
        $treeItem = [Windows.Controls.TreeViewItem]::new()
        $treeItem.header = $item.name
        try {
            if($item.ObjectClass -eq "organizationalUnit") {
                $childItems = Get-ChildItem ("AD:\{0}" -f $item.distinguishedName)
                Add-ChildNodes -parentNode $treeItem -currentItems $childItems
            }
            elseif($item.ObjectClass -eq "Group") {
                $Global:NumGroups = $Global:NumGroups + 1
                $members = Get-ADGroupMember $item.Name

                if($members.count -eq 0) {
                    $Global:NumEmpty = $Global:NumEmpty + 1
                }

                foreach($member in $members) {
                    if($member.objectClass -eq "user") {
                        $Global:NumNamedUsers = $Global:NumNamedUsers + 1
                    }
                }
            }
            elseif($item.objectClass -eq "Computer") {
                
                $Global:NumComputers = $Global:NumComputers + 1
                $currentComputer = Get-ADComputer $item.name -properties LastLogonDate, Enabled
                #Write-Host $currentComputer.enabled
                if($currentComputer.enabled -eq $false) {
                    $Global:NumInactive = $Global:NumInactive + 1
                    $treeItem.Background = "Red"
                }
                else {
                    $currentDate = Get-Date
                    if(($currentDate - $currentComputer.LastLogonDate).days -gt $Global:pastDate) {
                        $Global:CompsPastLogin = $Global:CompsPastLogin + 1
                        $treeItem.Background = "Yellow"
                    }
                    else{
                        $treeItem.Background = "LightGreen"
                    }
                }
            }
        }
        catch {
            Write-Host "Error encountered."
            Write-Host $_
        }
        $parentNode.Items.Add($treeItem) | Out-Null
    }
}

$Global:OUReporting.window.add_IsVisibleChanged({
    if ($Global:OUReporting.window.isVisible -eq $true) {
        $Global:OUReporting.window.topmost = $true
        $Global:OUReporting.window.topmost = $false
        $Global:OUReporting.window.focus()
    }
})

$Global:OUReporting.DatePicker.add_SelectedDateChanged({
    $Global:pastDate = ((Get-Date) - $Global:OUReporting.DatePicker.selectedDate).days
})

$Global:OUReporting.DateCombo.add_SelectionChanged({
    $today = Get-Date
    if($Global:OUReporting.DateCombo.SelectedValue.Content -eq "1 Year") {
        $Global:pastDate = 365
        
    }
    elseif ($Global:OUReporting.DateCombo.SelectedValue.Content -eq "6 Months") {
        $Global:pastDate = 182
    }

    elseif ($Global:OUReporting.DateCombo.SelectedValue.Content -eq "3 Months") {
        $Global:pastDate = 91
    }

    elseif ($Global:OUReporting.DateCombo.SelectedValue.Content -eq "1 Month") {
        $Global:pastDate = 30
    }
})

$Global:OUReporting.RunBtn.add_Click({
    $Global:OUReporting.ResultsTree.Items.Clear()
    $Global:NumComputers = 0
    $Global:NumGroups = 0
    $Global:NumInactive = 0
    $Global:NumEmpty = 0
    $Global:NumNamedUsers = 0
    $Global:CompsPastLogin = 0

    $Global:rootOU = Get-ADOrganizationalUnit -SearchBase $Global:OUReporting.OUDNText.text -filter * -searchscope 0
    $rootItem = [Windows.Controls.TreeViewItem]::new()
    $rootItem.header = $Global:rootOu.Name
    $Global:OUReporting.ResultsTree.Items.Add($rootItem) | Out-Null
    Add-ChildNodes -parentNode $rootItem -currentItems (Get-ChildItem ("AD:\{0}" -f $Global:OUReporting.OUDNText.text))

    $Global:OUReporting.GroupsTotal.Content = $Global:NumGroups
    $Global:OUReporting.UserByNameTotal.Content = $Global:NumNamedUsers
    $Global:OUReporting.ComputersTotal.Content = $Global:NumComputers
    $Global:OUReporting.InactiveTotal.Content = $Global:NumInactive
    $Global:OUReporting.PastLogin.Content = $Global:CompsPastLogin
    $Global:OUReporting.EmptyGroupTotal.Content = $Global:NumEmpty

})

$Global:OUReporting.ResultsTree.add_SelectedItemChanged({
    $Global:GroupDN = Get-ADGroup -SearchBase $Global:rootOU.distinguishedName -filter ('Name -eq "{0}"' -f $Global:OUReporting.ResultsTree.SelectedItem.Header)
})

$Global:OUReporting.MoveBtn.add_Click({
    $computerInfo = Get-ADComputer -SearchBase $Global:rootOU.DistinguishedName -filter ('Name -eq "{0}"' -f $Global:OUReporting.ResultsTree.SelectedItem.Header)
    Move-ADObject -identity $computerInfo.distinguishedName -TargetPath $Global:OUReporting.MoveToBox.text
})

$Global:OUReporting.UpdateBtn.add_Click({
    $Global:OUReporting.MoveToBox.text = $Global:GroupDN.distinguishedname
})

$Global:OUReporting.CancelBtn.add_Click({
    $form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null