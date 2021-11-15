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
        Title="Share Drive Reporting" Height="561" Width="744">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition Width="0*"/>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Label Content="Share Drive File Path:" HorizontalAlignment="Left" Margin="3,10,0,0" VerticalAlignment="Top" Height="26" Width="140" Grid.Column="1"/>
        <TextBox Name="PathText" HorizontalAlignment="Left" Margin="10,41,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="252" Height="44" Grid.Column="1"/>
        <Button Name="RunBtn" Grid.ColumnSpan="2" Content="Run" HorizontalAlignment="Left" Margin="282,63,0,0" VerticalAlignment="Top" Width="75"/>
        <Button Name="CancelBtn" Grid.ColumnSpan="2" Content="Cancel" HorizontalAlignment="Left" Margin="363,63,0,0" VerticalAlignment="Top" Width="74"/>
        <TreeView Name="ResultsTree" HorizontalAlignment="Center" Height="196" Margin="0,135,0,0" VerticalAlignment="Top" Width="724" Grid.Column="1"/>
        <Label Grid.ColumnSpan="2" Content="Results:" HorizontalAlignment="Left" Margin="3,109,0,0" VerticalAlignment="Top"/>
        <Label Name="PathLabel" Content="Path: " Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="10,351,0,0" VerticalAlignment="Top"/>
        <Label Name="OwnerLabel" Content="Owner: " Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="10,377,0,0" VerticalAlignment="Top"/>
        <Label Name="AccessesLabel" Content="Accesses:" Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="10,403,0,0" VerticalAlignment="Top"/>
        <Label Name="AccessesContentLabel" Content="" Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="28,429,0,0" VerticalAlignment="Top"/>
        <TextBlock Name="AccessesBlock" Grid.ColumnSpan="2" HorizontalAlignment="Left" Margin="23,429,0,0" Text="" TextWrapping="Wrap" VerticalAlignment="Top">
            <TextBlock.Inlines>
            </TextBlock.Inlines>
        </TextBlock>
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

$Global:ShareReporting = [hashtable]::Synchronized(@{})
$Global:ShareReporting.Window = $Form
$xaml.SelectNodes("//*[@Name]") | ForEach-Object {
    $Global:ShareReporting.add($_.Name, $Form.FindName($_.Name))
}

$Global:ShareReporting.window.add_IsVisibleChanged({
    if ($Global:ShareReporting.window.isVisible -eq $true) {
        $Global:ShareReporting.window.topmost = $true
        $Global:ShareReporting.window.topmost = $false
        $Global:ShareReporting.window.focus()
    }
})

function Build-FullPath($treeItem, $currentPath) {
    $currentPath = ("{0}\{1}" -f $treeItem.header, $currentPath)
    if($treeItem.parent) {
        Build-FullPath($treeItem.parent, $currentPath)
    }
    $currentPath
}

function Add-ChildNodes($parentNode, $currentItems) {
    foreach($item in $currentItems) {
        $ACL = Get-ACL $item

        $treeItem = [Windows.Controls.TreeViewItem]::new()
        $treeItem.header = $item.name

        $inheritanceCount = ($acl.access | Where-Object -Property 'IsInherited' -eq $true | Measure-Object).count
        if ($inheritanceCount -eq 0) {
            $treeItem.background = "LightYellow"
            Write-Host "Found not inherited"
            
            $parent = $parentNode
        
            while ($parent.header) {
                $parent.background = "LightYellow"
                $parent = $parent.parent
            }
        }

        $acl.access | Where-Object -Property 'IsInherited' -eq $false | ForEach-Object {
            "$($_.IdentityReference) $($_.AccessControlType): $($_.FileSystemRights)`n"
            if ($($_.IdentityReference) -eq 'CREATOR OWNER') {
                $errorStr += "Warning: Account CREATOR OWNER has explicit permission on folder.`n"
                $treeItem.background = "Red"
                Write-Host "Found CREATOR OWNER"
            
                $parent = $parentNode
            
                while ($parent.header) {
                    $parent.background = "LightRed"
                    $parent = $parent.parent
                }
            } elseif ($_.IdentityReference.AccountDomainSid) { 
                $errorStr += "Warning: User account ($($_.IdentityReference.value)) has explicit permission on folder.`n"
                $treeItem.background = "Red"
                Write-Host "Found user with explicit permissions"
            
                $parent = $parentNode
            
                while ($parent.header) {
                    $parent.background = "LightRed"
                    $parent = $parent.parent
                }
            } else {
                $userAccount = ($_.IdentityReference.value).Split('\')
                if ( (-NOT $userAccount[1].Contains('.')) -and ($userAccount[1].length -le 8)) {
                    if ($userAccount[0] -eq 'ASURITE') {
                        $domainServer = "asurite.ad.asu.edu"
                    } else {
                        $domainServer = "ad.asu.edu"
                    }
                    if ($(get-ADObject -Server $domainServer -Filter "Name -eq '$($userAccount[1])'").ObjectClass -eq 'user') {
                        $errorStr += "Warning: User account ($($userAccount[1])) has explicit permission on folder.`n"
                        $treeItem.background = "Red"
                        Write-Host "Found user with explicit permissions"
            
                        $parent = $parentNode
                    
                        while ($parent.header) {
                            $parent.background = "LightRed"
                            $parent = $parent.parent
                        }
                    }
                }
            }
        }
        
        try {
            if($item.PSIsContainer) {
                $childItems = Get-ChildItem $item
                Add-ChildNodes -parentNode $treeItem -currentItems $childItems
            }
        }
        catch {
            Write-Host "Error encountered."
            Write-Host $_
        }

        $parentNode.Items.Add($treeItem) | Out-Null

        $parent = $item.Parent
        $fullPath = $item.header
    
        while ($parent.header) {
            $fullPath = ("{0}\{1}" -f $parent.Header, $fullPath)
            $parent = $parent.parent
        }
    
    }
}

$Global:ShareReporting.window.add_IsVisibleChanged({
    if ($Global:ShareReporting.window.isVisible -eq $true) {
        $Global:ShareReporting.window.topmost = $true
        $Global:ShareReporting.window.topmost = $false
        $Global:ShareReporting.window.focus()
    }
})

$Global:ShareReporting.RunBtn.add_Click({
    $Global:ShareReporting.ResultsTree.Items.Clear()

    $pathValidation = Get-Childitem $Global:ShareReporting.PathText.text -ErrorAction SilentlyContinue
    $rootItem = [Windows.Controls.TreeViewItem]::new()
    $rootItem.header = $Global:ShareReporting.PathText.text
    if ($pathValidation -ne $null) {
        $Global:ShareReporting.ResultsTree.Items.Add($rootItem) | Out-Null
        Add-ChildNodes -parentNode $rootItem -currentItems $pathValidation
    }
    else {
        $GLobal:SyncHash.print(("{0} is not a valid share drive file path, please try again." -f $rootItem.Header))
    }
    
})

$Global:ShareReporting.ResultsTree.add_SelectedItemChanged({
    $currentItem = $Global:ShareReporting.ResultsTree.SelectedItem
    $parent = $currentItem.Parent
    $fullPath = $currentItem.header

    while ($parent.header) {
        $fullPath = ("{0}\{1}" -f $parent.Header, $fullPath)
        $parent = $parent.parent
    }

    $ACL = Get-ACL $fullPath
    $Global:ShareReporting.AccessesBlock.Inlines.Clear()
    foreach($access in $acl.access) {

        $accessInheritedRun = New-Object System.Windows.Documents.Run
        $accessInheritedRun.Name = "AccessesInheritedRun"
        $accessInheritedRun.FontStyle = "Italic"
        if($access.IsInherited) {
            $inheritedText = "Inherited"
            $accessInheritedRun.Background = "LightGreen"
        }
        else {
            $inheritedText = "Not Inherited"
            $accessInheritedRun.Background = "LightYellow"
        }
        $accessInheritedRun.Text = $inheritedText

        $accessLevelRun = New-Object System.Windows.Documents.Run
        $accessLevelRun.name = "AccessLevelRun"
        $accessLevelRun.FontWeight = "Bold"
        $accessLevelRun.Text = (" {0}: " -f $access.FilesystemRights)

        $accessIdentityRun = New-Object System.Windows.Documents.Run
        $accessIdentityRun.name = "AccessIdenityRun"
        $accessIdentityRun.text = $access.IdentityReference

        $lineBreak =  New-Object System.Windows.Documents.LineBreak

        $Global:ShareReporting.AccessesBlock.AddChild($accessInheritedRun)
        $Global:ShareReporting.AccessesBlock.AddChild($accessLevelRun)
        $Global:ShareReporting.AccessesBlock.AddChild($accessIdentityRun)
        $Global:ShareReporting.AccessesBlock.AddChild($lineBreak)
    }
    # $Global:ShareReporting.ResultsLabel.Content = "Path: `t{0}`nOwner: `t{1}`nInherited:Accesses:`n`t{2}" -f $fullPath, $acl.owner, $accessText
    $Global:ShareReporting.PathLabel.Content = "Path: $fullPath" 
    $Global:ShareReporting.OwnerLabel.Content = ("Owner: {0}" -f $acl.Owner)
    
    # $runOne = New-Object System.Windows.Documents.Run
    # $runOne.Name = "RunOne"
    # $runOne.Text = "Rune One Text"

    # $runTwo = New-Object System.Windows.Documents.Run
    # $runTwo.Name = "RunOne"
    # $runTwo.FontWeight = "Bold"
    # $runTwo.Text = "Rune Two Text"

    # $Global:ShareReporting.AccessesBlock.AddChild($runOne)
    # $Global:ShareReporting.AccessesBlock.AddChild($lineBreak)
    # $Global:ShareReporting.AccessesBlock.AddChild($runTwo)

})

$Global:ShareReporting.CancelBtn.add_Click({
    $form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null