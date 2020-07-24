<#
.SYNOPSIS
Display list of groups a computer is a member of
.DESCRIPTION
Generates and displays a tree view of all groups a computer is a member of
#>
function Invoke-ViewComputerMembership {
    [CmdletBinding(DefaultParameterSetName='none')]
    param (
        [Parameter(
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Computer name')]
        [Alias('Computer')]
        [String]$ComputerName,

        [Parameter(
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='The maximum depth to search')]
        [Alias('Type')]
        [Int]$MaxDepth
    )
    
    # Double-check that AD module is running
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:WpfApplication2"
Title="Computer object group report" Height="600" MinHeight="480" Width="800" MinWidth="570">
<TabControl Margin="8">
    <TabItem Header="Group tree">
        <DockPanel>
            <GroupBox Header="Information" HorizontalAlignment="Stretch" DockPanel.Dock="Top" Height="100">
                <StackPanel Orientation="Vertical" Margin="8">
                    <StackPanel Orientation="Horizontal">
                        <Label Width="256">Windows 10 update group:</Label>
                        <Label Name="lblUpdateGroup">Group 4</Label>
                    </StackPanel>
                    <StackPanel Orientation="Horizontal">
                        <Label Width="256">Update group inherited from:</Label>
                        <Label Name="lblUpdateGroupSource">None specified, global default used</Label>
                    </StackPanel>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="Group inheritance tree" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" DockPanel.Dock="Bottom">
                <ScrollViewer Name="listScroller" Margin="8">
                    <TreeView Name="groupTree"
                            VerticalAlignment="Stretch"
                            HorizontalAlignment="Stretch">
                        <TreeView.ItemContainerStyle>
                            <Style TargetType="{x:Type TreeViewItem}">
                                <!-- Style for the selected item -->
                                <Setter Property="BorderThickness" Value="1"/>
                                <Style.Triggers>
                                    <!-- Selected and has focus -->
                                    <Trigger Property="IsSelected" Value="True">
                                        <Setter Property="BorderBrush" Value="#7DA2CE"/>
                                    </Trigger>
                                    <MultiTrigger>
                                        <MultiTrigger.Conditions>
                                            <Condition Property="IsSelected" Value="True"/>
                                            <Condition Property="IsSelectionActive" Value="False"/>
                                        </MultiTrigger.Conditions>
                                        <Setter Property="BorderBrush" Value="#D9D9D9"/>
                                    </MultiTrigger>
                                </Style.Triggers>
                                <Style.Resources>
                                    <Style TargetType="Border">
                                        <Setter Property="CornerRadius" Value="2"/>
                                    </Style>
                                </Style.Resources>
                            </Style>
                        </TreeView.ItemContainerStyle>
                        <TreeView.Resources>
                            <LinearGradientBrush x:Key="{x:Static SystemColors.HighlightBrushKey}" EndPoint="0,1" StartPoint="0,0">
                                <GradientStop Color="#FFDCEBFC" Offset="0"/>
                                <GradientStop Color="#FFC1DBFC" Offset="1"/>
                            </LinearGradientBrush>
                            <LinearGradientBrush x:Key="{x:Static SystemColors.ControlBrushKey}" EndPoint="0,1" StartPoint="0,0">
                                <GradientStop Color="#FFF8F8F8" Offset="0"/>
                                <GradientStop Color="#FFE5E5E5" Offset="1"/>
                            </LinearGradientBrush>
                            <SolidColorBrush x:Key="{x:Static SystemColors.HighlightTextBrushKey}" Color="Black" />
                            <SolidColorBrush x:Key="{x:Static SystemColors.ControlTextBrushKey}" Color="Black" />
                        </TreeView.Resources>
                    </TreeView>
                </ScrollViewer>
            </GroupBox>
        </DockPanel>
    </TabItem>
    <TabItem Header="Software groups">
        <DataGrid Name="softwareGrid" Margin="8" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False" VerticalAlignment="Stretch" MinHeight="200" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Width="450" IsReadOnly="True" Binding="{Binding name}" />
                <DataGridTextColumn Header="Inherited From" Width="300" IsReadOnly="True" Binding="{Binding parent}" />
            </DataGrid.Columns>
        </DataGrid>
    </TabItem>
    <TabItem Header="Resource groups">
        <DataGrid Name="resourceGrid" Margin="8" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False" VerticalAlignment="Stretch" MinHeight="200" IsReadOnly="True">
            <DataGrid.Columns>
                <DataGridTextColumn Header="Name" Width="450" IsReadOnly="True" Binding="{Binding name}" />
                <DataGridTextColumn Header="Inherited From" Width="300" IsReadOnly="True" Binding="{Binding parent}" />
            </DataGrid.Columns>
        </DataGrid>
    </TabItem>
</TabControl>
</Window>
"@

    # Load XML; prep for deserialization
    $inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Add-Type -AssemblyName PresentationCore,PresentationFramework,WindowsBase,system.windows.forms
    [xml]$XAML = $inputXML

    # Deserialize the form into an object
    $form = $null
    try{
        $Form = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xaml))
    } catch [System.Management.Automation.MethodInvocationException] {
        Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
        write-host $error[0].Exception.Message -ForegroundColor Red
        if ($error[0].Exception.Message -like "*button*"){
            write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"
        }
    } catch {
        #if it broke some other way :D
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }

    # Load form components into a hash table
    # This must be stored in global scope so event calls can reach it
    $Global:groupTreeForm = [hashtable]::Synchronized(@{})
    $Global:groupTreeForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:groupTreeForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:groupTreeForm.add('okButtonClicked',$false)
    $Global:groupTreeForm.add('softwareList',(New-Object -TypeName System.collections.ArrayList))
    $Global:groupTreeForm.softwareGrid.itemsSource = $Global:groupTreeForm.softwareList
    $Global:groupTreeForm.add('resourceList',(New-Object -TypeName System.collections.ArrayList))
    $Global:groupTreeForm.resourceGrid.itemsSource = $Global:groupTreeForm.resourceList
    $rootItem = New-Object -Type 'System.Windows.Controls.TreeViewItem'
    $rootItem.tag = 'root'
    $rootItem.header = $ComputerName
    
    # Being window to fromnt when made visible
    $Global:groupTreeForm.window.add_IsVisibleChanged({
        if ($Global:groupTreeForm.window.isVisible -eq $true) {
            $Global:groupTreeForm.window.topmost = $true
            $Global:groupTreeForm.window.topmost = $false
            $Global:groupTreeForm.window.focus()
        }
    })

    $compObject = $null
    try {
        $compObject = Get-ADComputer -Identity $ComputerName -Properties 'name','memberOf'
    }
    catch {
        $compObject = $null
    }

    if ($compObject -eq $null) {
        return 9
    }

    function addGroup {
        param($Private:child,$Private:isComputer)
        $Private:aGroup = $null
        try {
            if ($Private:isComputer -eq $true) {
                $Private:aGroup = Get-ADComputer -Identity "${Private:child}" -Properties 'Name','MemberOf'
            } else {
                $Private:aGroup = Get-ADGroup -Identity $Private:child -Properties 'Name','MemberOf'
            }
        } catch {
            $Private:aGroup = $null
            return "ERR------${private:child}"
        }
        $Private:i = New-Object -Type 'System.Windows.Controls.TreeViewItem'
        $Private:i.header = $Private:child
        $Private:aGroup.memberOf | Sort-Object | Foreach-Object {
            $Private:gName = ([Regex]'CN=(?<name>[^,]+),*').match($_).groups.item('name').value
            Write-Host "  ${Private:child}>>${Private:gName}"
            $Private:subObject = addGroup $Private:gName $false
            if ($Private:subObject -is [System.Windows.Controls.TreeViewItem]) {
                $Private:i.items.add($Private:subObject) | Out-Null
            }

            # Also add to resource and software groups if matches name standard
            $gridItem = [PSCustomObject]@{
                name = $Private:gName
                parent = $Private:child
            }
            if ($Private:gName  -match '(?i)\.Software') {
                $Global:groupTreeForm.softwareList.add($gridItem) | Out-Null

                if ($Private:gName -match '(?i)Win10_Update_Latest_Release_') {
                    $setGroup = $false
                    $Private:upGroup = ([Regex]'(?i)Update_Latest_Release_(?<name>[\w]+)$').match($Private:gName).groups.item('name').value -Replace '_',' '

                    # Determine if the group listed on the form is later than this one
                    $currentGroup = $null
                    $currentGroup = $Global:groupTreeForm.lblUpdateGroup.Content
    
                    if (
                        $private:upgroup -imatch 'test' -and
                        $currentGroup -inotmatch 'text'
                    ) {
                        $setGroup = $true
                    } elseif (
                        $private:upgroup -imatch 'pilot' -and
                        $currentGroup -inotmatch 'text|pilot'
                    ) {
                        $setGroup = $true
                    } elseif (
                        $private:upgroup -imatch 'group 1' -and
                        $currentGroup -inotmatch 'text|pilot|group 1'
                    ) {
                        $setGroup = $true
                    } elseif (
                        $private:upgroup -imatch 'group 2' -and
                        $currentGroup -inotmatch 'text|pilot|group 1|group 2'
                    ) {
                        $setGroup = $true
                    } elseif (
                        $private:upgroup -imatch 'group 3' -and
                        $currentGroup -inotmatch 'text|pilot|group 1|group 2|group 3'
                    ) {
                        $setGroup = $true
                    } elseif (
                        $private:upgroup -imatch 'group 4' -and
                        $currentGroup -inotmatch 'text|pilot|group 1|group 2|group 3|group 4'
                    ) {
                        $setGroup = $true
                    }
    
                    if ($setGroup) {
                        $Global:groupTreeForm.lblUpdateGroup.Content = $private:upgroup
                        $Global:groupTreeForm.lblUpdateGroupSource.Content = $private:child
                    }
                }
            } elseif ($Private:gName -match '(?i)\.GPO|\.SHR|\.PRT|\.TaskSequence') {
                $Global:groupTreeForm.resourceList.add($gridItem) | Out-Null
            }
        }
        
        return $Private:i
    }

    $rootItem = addGroup $compObject.name $true
    $rootItem.ExpandSubTree()
    $Global:groupTreeForm.groupTree.items.add($rootItem)
    $Global:groupTreeForm.softwareGrid.items.refresh()
    $Global:groupTreeForm.resourceGrid.items.refresh()
    $Global:groupTreeForm.listScroller.ScrollToTop()
    $Global:groupTreeForm.window.showdialog()
}