[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:DSCtrl"
        Title="Building Tree Test" Height="516.167" Width="570.334">
    <Grid HorizontalAlignment="Left" Height="434" VerticalAlignment="Top" Width="560" Margin="0,0,0,-1">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <TreeView Name="Tree" HorizontalAlignment="Left" Height="416" Margin="10,10,-328,-178" Grid.RowSpan="2" VerticalAlignment="Top" Width="544" Grid.ColumnSpan="2"/>
        <Label Content="Currently Selected:" HorizontalAlignment="Left" Height="27" Margin="21,434,0,-27" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Label Name="CurrentSelection" Content="" HorizontalAlignment="Left" Height="27" Margin="130,434,0,-27" Grid.RowSpan="2" VerticalAlignment="Top"/>
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

$rootItem = [Windows.Controls.TreeViewItem]::new()
$rootItem.header = "C:"
$Tree.Items.Add($rootItem) | Out-Null

$rootChildren = Get-ChildItem "C:\"

function Add-ChildNodes($parentNode, $currentItems) {
    foreach($item in $currentItems) {
        $treeItem = [Windows.Controls.TreeViewItem]::new()
        $treeItem.header = $item.name
        $parentNode.Items.Add($treeItem) | Out-Null
        try {
            $itemProperties = Get-Item $item.FullName
            if($itemProperties -is [System.IO.DirectoryInfo]) {
                $childItems = Get-ChildItem $item.FullName
                Add-ChildNodes -parentNode $treeItem -currentItems $childItems
            }
        }
        catch {
            Write-Host "Error encountered."
        }
    }
}

Add-ChildNodes -parentNode $rootItem -currentItems $rootChildren

$Tree.add_GotFocus({
    $CurrentSelection.Content = $_.OriginalSource.Header
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null