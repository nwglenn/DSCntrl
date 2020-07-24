<#
.SYNOPSIS
Converts an M.UTOSPA security group name into its parts
.DESCRIPTION
In the M.UTOSPA naming standard, security groups are made up of multiple parts seperated by a period. This function takes a properly formatted security group name and returns a hash map with each of its respective parts
.PARAMETER GroupName
The name of the group to be interpreted
.OUTPUTS
In the case of an improperly formatted name, the function returns null or throws an error. If a properly formatted name is provided, a hash map is returned with the following parts;

    name: The full name of the group, as it was provided
    type: One of; Computer, User, GPO, Printer, Share
    parent: The name of the parent group for Computer and User groups
    unit: The unit to which this group belongs
    suffix: The portion of the name that follows the type descriptor
    resource: In the case of a security group that pertains to a resource, this
        field will contain a path to the defined resource. In all other cases, this field will be null
.EXAMPLE
ConvertFrom-GroupName -GroupName 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'
.NOTES
For more information o the naming standard, read: https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015381
#>
function ConvertFrom-GroupName {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Group name')]
        [Alias('Name')]
        [String]$GroupName
    )

    $partsFinder = [Regex]'(?i)^(M\.UTOSPA\.)(?<unit>[^\.]+)\.groups\.(?<type>[^\.]+)(\.(?<suffix>.+))?'
    $parentFinder = [Regex]'(?i)^(?<parent>M\.UTOSPA\..*)(?<child>\.[^\.]*)$'

    $gData = $null
    $parts = $partsFinder.Match($GroupName)

    $parts = $partsFinder.Match($GroupName)
    if ($parts.Success) {
        $gData = [PSCustomObject]@{
            name=''
            type=''
            parent=''
            unit=''
            suffix=''
            resource = $null
        }

        try {
            $gData.name = $parts.groups.Item(0).value
            $gData.type = $parts.groups.Item('type').value
            $gData.unit = $parts.groups.Item('unit').value
            $gData.suffix = $parts.groups.Item('suffix').value

            # Attempt to detect parent
            $parent = $parentFinder.match($GroupName)
            if ($parent.success) {
                $gData.parent = $parent.groups.Item('parent').value
                $gData.suffix = $parent.groups.Item('child').value
            }

            # Ensure that there are no leading or trailing periods
            if ($gData.suffix -imatch '^\.|\.$') {
                $gData.suffix = $gData.suffix -Replace '^\.|\.$',''
            }

        } catch {
            $gData = $null
        }

        # Translate type identifiers to type names
        # Also, clear parent names for groups that do not support parenting
        if ($gData.type -imatch '^SHR') {
            $gData.type = 'Share'
            $gData.parent = ''
        } elseif ($gData.type -imatch '^CMP') {
            $gData.type = 'Computer'
        } elseif ($gData.type -imatch '^PRT') {
            $gData.type = 'Printer'
        } elseif ($gData.type -imatch '^USR') {
            $gData.type = 'User'
        } elseif ($gData.type -imatch '^GPO') {
            $gData.type = 'GPO'
            $gData.parent = ''
        }

        # Assign printer, share and GPO groups to relative computer groups for prefix generation purposes
        $gData.parent = $gData.parent -Replace '(?i)\.(GPO|SHR-\w\w|PRT)','.CMP'

        # In the case of Printer or Share group, decode the resource name and provide it to the caller
        if ($gData.type -imatch 'SHARE|PRINTER') {
            $resourceFinder = [Regex]'(?i)^(?<server>[^_]*)_(?<path>[^\.~]*)(~\d)?$'
            $resource = $resourceFinder.match($gData.suffix)
            if ($resource.success) {
                $gData.resource = "\\$($resource.Groups.Item('server').value)\$($resource.Groups.Item('path').value)"
            }
        }

        return $gdata
    }
}

<#
.SYNOPSIS
Converts a set of name information into a M.UTOSPA group name
.DESCRIPTION
The inverse of the ConvertFrom-GroupName function, generates a M.UTOSPA compliant group name.
.PARAMETER GroupType
One of; Computer, User, Printer, Share, GPO
.PARAMETER GroupUnit
The unit name as it appears in the UTOSPA OU structure, typically the acronym for the respective unit
.PARAMETER GroupParent
When requesting a computer or user group, this name will be used to determine the nesting structure
.PARAMETER GroupSuffix
The portion of the name that follows either the parent name, or base group name.

NOTE: ignored for printer, share and GPO groups
.PARAMETER CampusInitial
A single letter that dissambiguates between sections of a group across campuses

NOTE: Only used for user and computer groups
.PARAMETER ResourceServer
When specifying a printer or share group, the name of the server the resource resides on. Ex; POLYPRINT1
.PARAMETER ResourcePath
The path to the resource on the server. Do not include the server name portion, only the ath below that point
.PARAMETER AccessType
One of; RW,RO,LO,DA
.PARAMETER SkipADQuery
Do not query AD when trying to generate names that are longer than 62 characters. This AD query is designed to prevent collisions with existing long names in AD,
When enabled, it will append a ~# to the end of names after the 62ng characted, where the number is the next available index in AD/
When this query is disabled by specifying this option, long names will always end in ~1, regardless of collisions
.OUTPUTS
A string containing the calculated group name
.EXAMPLE
ConvertFrom-GroupName -GroupName 'M.UTOSPA.UTO.Groups.CMP.ITCSS.UCC.Poly'
.NOTES
For more information o the naming standard, read: https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015381
#>
function ConvertTo-GroupName {
    [CmdletBinding(DefaultParameterSetName='standard')]
    param (
        [Parameter(Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Group name')]
        [Alias('Type')]
        [String]$GroupType,

        [Parameter(Mandatory=$True,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='One of; Computer, GPO, Printer, Share, User')]
        [Alias('Unit')]
        [String]$GroupUnit,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='Unit name as it appears in OU structure; USI, CLAS, UTO, CPW, Etc...')]
        [Alias('Parent')]
        [String]$GroupParent,

        [Parameter(ParameterSetName='standard',
        Mandatory=$true,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='The name of the parent group')]
        [Alias('Suffix')]
        [String]$GroupSuffix,

        [Parameter(Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='The name of the parent group')]
        [Alias('Campus')]
        [String]$CampusInitial,

        [Parameter(ParameterSetName='resource',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='The name of the resource server')]
        [Alias('Server')]
        [String]$ResourceServer,

        [Parameter(ParameterSetName='resource',
        Mandatory=$true,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='The name of the resource server')]
        [Alias('Path')]
        [String]$ResourcePath,

        [Parameter(ParameterSetName='resource',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='Either; RW,RO,LO,DA')]
        [Alias('SubType')]
        [String]$AccessType = 'RW',

        [Parameter(ParameterSetName='resource',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='If set, AD query will be skipped for resource groups')]
        [Alias('SkipAD')]
        [Boolean]$SkipADQuery = $false
    )

    # Normalize inputs
    if ($ResourcePath) {

        # Replaces backslashes with underscores
        $ResourcePath = $ResourcePath -Replace '[\\]','_'

        # remove invalid characters
        $ResourcePath = $ResourcePath -ireplace '[\[\];=+@]',''
    }
    $newPrefix = $null
    switch ($GroupType) {
        'COMPUTER' {
            if ($GroupParent.length -lt 8) {
                #https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015392
                $newPrefix = "M.UTOSPA.${GroupUnit}.Groups.CMP.${GroupSuffix}"
            } else {
                $newPrefix = "${GroupParent}.${GroupSuffix}"
            }
        }
        'GPO' {

            # Use either suffix OR resourcePath, whichever is specified
            $newPrefix = "M.UTOSPA.${GroupUnit}.Groups.GPO.${GroupSuffix}${ResourcePath}"            
        }
        'PRINTER' {
            $newPrefix = "M.UTOSPA.${GroupUnit}.Groups.PRT"
            if ($GroupParent.length -gt 8) {
                $getGroup = [Regex]"(CMP|USR)\.(.*)$"
                $rxResult = $getGroup.Match($GroupParent)
                if ($rxResult.Success) {
                    $newPrefix = "${newPrefix}.$($rxResult.Groups[2].Value)"
                }
            }

            # Add campus initial if provided
            if ($CampusInitial) {
                $newPrefix = "${newPrefix}.${CampusInitial}"
            }

            # generate suffix from provided ResourceServer and ResourcePath
            $newPrefix = "${newPrefix}.${ResourceServer}`_${ResourcePath}"

            # Truncate names longer than 62 characters
            if ($newPrefix.length -gt 63) {
                $newPrefix = "$($newPrefix.substring(0,62))~1"
            }
        }
        'SHARE' {
            $newPrefix = "M.UTOSPA.${GroupUnit}.Groups.SHR-${AccessType}"
            if ($GroupParent.length -gt 8) {
                $getGroup = [Regex]"(CMP|USR)\.(.*)$"
                $rxResult = $getGroup.Match($GroupParent)
                if ($rxResult.Success) {
                    $newPrefix = "${newPrefix}.$($rxResult.Groups[2].Value)"
                }
            }

            # Add campus initial if provided
            if ($CampusInitial) {
                $newPrefix = "${newPrefix}.${CampusInitial}"
            }

            # generate suffix from provided ResourceServer and ResourcePath
            $newPrefix = "${newPrefix}.${ResourceServer}`_${ResourcePath}"

            # Truncate names longer than 62 characters
            if ($newPrefix.length -gt 63) {
                $newPrefix = "$($newPrefix.substring(0,62))~1"
            }
        }
        'USER' {
            if ($GroupParent.length -lt 8) {
                #https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015392
                $newPrefix = "M.UTOSPA.${GroupUnit}.Groups.USR.${GroupSuffix}"
            } else {
                $newPrefix = "${GroupParent}.${GroupSuffix}" -Replace '\.CMP\.','.USR.'
            }
        }
        Default {
            $newPrefix = $null
        }
    }

    # Check AD for duplicate objects if the generated name contains a tilda
    # this is an indication that a long name was shortened,
    # Therefore, we need to check for collisions
    if (
        ($newPrefix -match '~1$') -and
        ($SkipADQuery -ne $true)
    ) {
        $queryargs = @{
            searchBase = "OU=M.UTOSPA.${GroupUnit},${global:OUPath}"
            filter = 'Name -like "M.UTOSPA*" '
            SearchScope = 'SubTree'
            property = 'name'
            Server = $global:ADserver
        }
        $queryResult = $null
        try {

            # If running a pester test, use mockable version of Get-ADGroup
            # this is a hack to get around an issue where mocking Get-ADGroup
            # directly seems to break constantly
            # TODO: Solve this really terrible issue
            # possible reference: https://blogs.technet.microsoft.com/heyscriptingguy/2015/12/17/testing-script-modules-with-pester/
            if ([Boolean](get-command invoke-testproxy -Erroraction SilentlyContinue)) {
                $queryResult = Invoke-TestProxy | Select-Object -ExpandProperty Name
            } else {
                $queryResult = Get-ADGroup @queryargs | Select-Object -ExpandProperty Name
            }
        }
        catch {
            $queryResult = $null
        }

        $groupexists = $queryResult -contains $newPrefix
        while ($groupexists) {
            
            # Increment number
            $currentNumber = [Int]($newPrefix.subString($newPrefix.length - 1))
            $newPrefix = $newPrefix -Replace "~${currentNumber}","~$($currentNumber + 1)"
            $groupexists = $queryResult -contains $newPrefix
        }
    }

    # Add campus initial if applicable
    if (
        ($CampusInitial.length -eq 1) -and
        ($newPrefix -Match '\.(CMP|USR)\.')
    ) {
        return "${newPrefix}.${CampusInitial}"
    } else {
        return $newPrefix
    }
}

<#
.SYNOPSIS
Creates a new Security Group in the M.UTOSPA OU.
.DESCRIPTION
Creates a new Security Group in the M.UTOSPA OU based on the UTO Deskside Standards.
.INPUTS
None. You cannot pipe objects to CreateGroup.ps1.
.OUTPUTS
None. CreateGroup.ps1 does not generate any output.
.EXAMPLE
C:\PS> .\CreateGroup.ps1
#>
function CreateGroup {
    [CmdletBinding(DefaultParameterSetName='none')]
    param (
        [Parameter(ParameterSetName='import',
        Mandatory=$True,
        ValueFromPipeline=$True,
        ValueFromPipelineByPropertyName=$True,
        HelpMessage='Group name')]
        [Alias('Name')]
        [String]$GroupName,

        [Parameter(ParameterSetName='import',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='An optional ticket number')]
        [Alias('Ticket')]
        [String]$GroupTicket,

        [Parameter(ParameterSetName='import',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='An optional ticket number')]
        [Alias('Department')]
        [String]$GroupDept,

        [Parameter(ParameterSetName='import',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='An optional support group name')]
        [Alias('sGroup')]
        [String]$SupportGroup,

        [Parameter(ParameterSetName='import',
        Mandatory=$False,
        ValueFromPipeline=$False,
        ValueFromPipelineByPropertyName=$False,
        HelpMessage='An optional update group name')]
        [Alias('uGroup')]
        [String]$UpdateGroup
    )

    # This function requires that the ADServer global variable be set.
    # Typically, this should happen when the global.ps1 file is loaded
    # on creation of a new DSCtrl thread. If this has not happened, we
    # default to ASURITE6 so we don't pass a null string later in the function
    if ( ($global:ADserver -isnot [String]) -or ($global:ADserver.length -lt 12) ) {
        $global:ADserver = 'asurite6.asurite.ad.asu.edu'
    }

    # Print start of operation to main form
    $syncHash.print("[ OK! ] Starting group creation workflow...")

    # Double-check that AD module is running
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    #region formLoad

    # Define the fancy new monolithic form
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:WpfApplication2"
Title="Create Group" Height="600" MinHeight="480" Width="620" MinWidth="620">
    <DockPanel>
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Height="32">
            <CheckBox Name="chkDebugCreation" Margin="4" Visibility="Collapsed">!!DEBUG MODE!! Create group in test OU rather than actial unit OU</CheckBox>
            <Button Name="btnOk" Width="64" Height="24" TabIndex="100" Margin="4" IsDefault="True">OK</Button>
            <Button Name="btnCancel" Width="64" Height="24" TabIndex="101" Margin="4">Cancel</Button>
        </StackPanel>
        <ScrollViewer DockPanel.Dock="Bottom" Margin="16" VerticalAlignment="Stretch" HorizontalAlignment="Stretch">
            <StackPanel Orientation="Vertical" Margin="16" VerticalAlignment="Stretch" HorizontalAlignment="Stretch">
                <StackPanel Name="grpBulkGroupName" Orientation="Horizontal" Margin="4" Visibility="Collapsed">
                    <Image Source="!imgpath!\MS-stdlib-warning-128x128.png" Width="64" Height="64" />
                    <StackPanel Orientation="Vertical" Margin="8">
                        <TextBlock TextWrapping="Wrap" Width="420">Before continuing, verify that the automatically selected options match the expected values for the following group:</TextBlock>
                        <TextBlock Name="lblBulkGroupName" FontWeight="Bold">ERROR: Name not provided!</TextBlock>
                    </StackPanel>
                </StackPanel>
                <StackPanel Name="grpGroupType" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Group Type:</Label>
                    <ComboBox Name="cmbGroupType" Width="256" TabIndex="1">
                        <ComboBoxItem>Computer</ComboBoxItem>
                        <ComboBoxItem>GPO</ComboBoxItem>
                        <ComboBoxItem>Printer</ComboBoxItem>
                        <ComboBoxItem>Share</ComboBoxItem>
                        <ComboBoxItem>User</ComboBoxItem>
                    </ComboBox>
                    <TextBlock Padding="4">
                        <Hyperlink Name="lnkOpenKB">
                            Go to KB
                        </Hyperlink>
                    </TextBlock>
                </StackPanel>
                <StackPanel Name="grpUnitName" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Unit:</Label>
                    <ComboBox Name="cmbUnitName" Width="256" TabIndex="2"></ComboBox>
                </StackPanel>
                <StackPanel Name="grpParentGroup" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Parent Group:</Label>
                    <ComboBox Name="cmbParentGroup" Width="256" TabIndex="3"></ComboBox>
                </StackPanel>
                <StackPanel Name="grpSupportGroup" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Support Group:</Label>
                    <ComboBox Name="cmbSupportGroup" Width="256" TabIndex="3"></ComboBox>
                </StackPanel>
                <StackPanel Name="grpUpdateGroup" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Update branch:</Label>
                    <ComboBox Name="cmbUpdateGroup" Width="256" TabIndex="3" SelectedIndex="0">
                        <ComboBoxItem>(None - inherit from parent)</ComboBoxItem>
                        <ComboBoxItem>Group_1</ComboBoxItem>
                        <ComboBoxItem>Group_2</ComboBoxItem>
                        <ComboBoxItem>Group_3</ComboBoxItem>
                        <ComboBoxItem>Group_4</ComboBoxItem>
                    </ComboBox>
                    <Label Width="64">(optional)</Label>
                </StackPanel>
                <StackPanel Name="grpCampusInitial" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Campus initial:</Label>
                    <ComboBox Name="cmbCampusInitial" Width="256" TabIndex="1" SelectedIndex="0">
                        <ComboBoxItem>(None)</ComboBoxItem>
                        <ComboBoxItem>D</ComboBoxItem>
                        <ComboBoxItem>H</ComboBoxItem>
                        <ComboBoxItem>P</ComboBoxItem>
                        <ComboBoxItem>T</ComboBoxItem>
                        <ComboBoxItem>W</ComboBoxItem>
                    </ComboBox>
                    <Label Width="64">(optional)</Label>
                </StackPanel>
                <StackPanel Name="grpPrintServer" Orientation="Horizontal" Margin="4" Visibility="Collapsed">
                    <Label Width="160">Print server:</Label>
                    <ComboBox Name="cmbPrintServer" Width="256">
                        <ComboBoxItem>ASUPRINT1</ComboBoxItem>
                        <ComboBoxItem>DPCPRINT1</ComboBoxItem>
                        <ComboBoxItem>POLYPRINT1</ComboBoxItem>
                        <ComboBoxItem>WESTPRINT1</ComboBoxItem>
                    </ComboBox>
                    <Label Name="lblPrintersLoading" Visibility="Collapsed">Loading printers...</Label>
                </StackPanel>
                <StackPanel Name="grpPrinterName" Orientation="Horizontal" Margin="4" Visibility="Collapsed">
                    <Label Width="160">Printer:</Label>
                    <ComboBox Name="cmbPrinterName" Width="256">
                    </ComboBox>
                </StackPanel>
                <StackPanel Name="grpShareName" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Path to share:</Label>
                    <TextBox Name="txtShareName" Width="256"></TextBox>
                </StackPanel>
                <StackPanel Name="grpGroupName" Orientation="Horizontal" Margin="4">
                    <Label Width="160">Name:</Label>
                    <TextBox Name="txtGroupName" Width="256"></TextBox>
                </StackPanel>
                <GroupBox Header="Generated name" Height="64" HorizontalAlignment="Stretch">
                    <TextBlock Name="lblGroupName" VerticalAlignment="Center" HorizontalAlignment="Stretch"></TextBlock>
                </GroupBox>
                <StackPanel Orientation="Horizontal" Margin="4">
                    <Label Name="lblDepartmentName" Width="160">Department:/section name:</Label>
                    <TextBox Name="txtDepartmentName" Width="256" TabIndex="6"></TextBox>
                    <Label Name="lblApproverOptional" Visibility="Collapsed" Width="64">(optional)</Label>
                </StackPanel>
                <StackPanel Orientation="Horizontal" Margin="4">
                    <Label Width="160">Ticket Number:</Label>
                    <TextBox Name="txtTicketNumber" Width="256" TabIndex="7"></TextBox>
                    <Label Width="64">(optional)</Label>
                </StackPanel>
                <GroupBox Name="grpShareOptions" Visibility="Collapsed" Header="Create groups with these types">
                    <StackPanel>
                        <CheckBox Name="chkAlsoCreateLO" Margin="4" >SHR-LO - List Only permissions</CheckBox>
                        <CheckBox Name="chkAlsoCreateRO" Margin="4" >SHR-RO - Read Only permissions</CheckBox>
                        <CheckBox Name="chkAlsoCreateRW" IsChecked="True" Margin="4" >SHR-RW - Read and write permissions</CheckBox>
                        <CheckBox Name="chkAlsoCreateDA" IsChecked="True" Margin="4" >SHR-DA - Deny All permissions</CheckBox>
                    </StackPanel>
                </GroupBox>
                <GroupBox Name="grpPrintInfo" Visibility="Collapsed" Header="Printer information">
                    <StackPanel Orientation="Vertical" Margin="4">
                        <StackPanel Orientation="Horizontal" Margin="4">
                            <Label Width="150">Make/Model:</Label>
                            <TextBox Name="txtPrinterMakeModel" Width="256" TabIndex="6"></TextBox>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="4">
                            <Label Width="150">IP:</Label>
                            <TextBox Name="txtPrinterIP" Width="256" TabIndex="6"></TextBox>
                        </StackPanel>
                        <StackPanel Orientation="Horizontal" Margin="4">
                            <Label Width="150">Location:</Label>
                            <TextBox Name="txtPrinterLocation" Width="256" TabIndex="6"></TextBox>
                        </StackPanel>
                    </StackPanel>
                </GroupBox>
                <StackPanel Name="grpCompOptions" Orientation="Vertical" Margin="4">
                    <CheckBox Name="chkNesting" Margin="4" TabIndex="8" IsChecked="True">Nest group within the parent group (recommended)</CheckBox>
                    <CheckBox Name="chkCreateCmpUsrPair" Margin="4" TabIndex="8" IsChecked="True">Create a set of Computer (CMP) and User (USR) groups</CheckBox>
                    <StackPanel Name="grpUseStaging" Orientation="Horizontal">
                        <CheckBox Name="chkUseStaging" Margin="4" TabIndex="9">Populate membership with computers in a staging area</CheckBox>
                        <ComboBox Name="cmbStagingOUNumber" TabIndex="10">
                            <ComboBoxItem>1</ComboBoxItem>
                            <ComboBoxItem>2</ComboBoxItem>
                            <ComboBoxItem>3</ComboBoxItem>
                            <ComboBoxItem>4</ComboBoxItem>
                            <ComboBoxItem>5</ComboBoxItem>
                        </ComboBox>
                    </StackPanel>
                </StackPanel>
            </StackPanel>
        </ScrollViewer>
    </DockPanel>
</Window>
"@

    # Provide path to image files
    $inputXML = $inputXML -replace '!imgpath!',"${IncludePath}\img"

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
    $Global:createGroupForm = [hashtable]::Synchronized(@{})
    $Global:createGroupForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:createGroupForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:createGroupForm.add('okButtonClicked',$false)
    $Global:createGroupForm.add('printerList',(New-Object -TypeName System.collections.ArrayList))
    $Global:createGroupForm.cmbPrinterName.itemsSource = $Global:createGroupForm.printerList
    if ($global:debug) {$syncHash.print("DEBUG: load form complete")}
    if ($global:debug) {$Global:createGroupForm.chkDebugCreation.Visibility = 'Visible'}

    #endregion

    #region formFunctions

    # updateParentList - Clears and reloads the list of available parent security groups based on group type and unit selections
    $Global:createGroupForm | Add-Member -Type ScriptMethod -name 'updateParentList' -Value {
        $selectedUnit = $this.cmbUnitName.SelectedValue.ToString()
        $selectedType = $this.cmbGroupType.SelectedValue.Content.ToUpper().ToString()

        # Abort if group type has not been selected
        if ( ($selectedType -eq $null) -or ($selectedType -eq '') ) {
            return
        }

        # Also abort if selected unit name is empty
        if ( ($selectedUnit -eq $null) -or ($selectedUnit -eq '') ) {
            return
        }

        # Generate new list of items
        $basePath = "OU=M.UTOSPA.${selectedUnit}.Groups,OU=M.UTOSPA.${selectedUnit},$globalOUPath"
        $groupList = $null
        switch ($selectedType) {
            'COMPUTER' {
                # get all the Computer Groups except for campus ones.
                $groupList = Get-ADGroup -SearchBase $basePath -filter {GroupCategory -eq "Security" -and name -like "*.Groups.CMP*" -and name -notlike "*.Groups.OU*" -and name -notlike "*.T" -and name -notlike "*.D" -and name -notlike "*.P" -and name -notlike "*.W" -and name -notlike "*.H"} -Server $global:ADserver | Sort-Object -Property 'Name'
            }
            'USER' {
                $groupList = Get-ADGroup -SearchBase $basePath -filter {GroupCategory -eq "Security" -and name -like "*.Groups.USR*" -and name -notlike "*.Groups.OU*" -and name -notlike "*.T" -and name -notlike "*.D" -and name -notlike "*.P" -and name -notlike "*.W" -and name -notlike "*.H"} -Server $global:ADserver | Sort-Object -Property 'Name'
            }
            Default {
                # get all the Computer Groups
                $groupList = Get-ADGroup -SearchBase $basePath -filter {GroupCategory -eq "Security" -and name -like "*.Groups.CMP*" -and name -notlike "*.Groups.OU*" } -Server $global:ADserver | Sort-Object -Property 'Name'
            }
        }

        # Reset parent group selection
        if ($groupList -ne $null) {
            $this.cmbParentGroup.SelectedIndex = -1
            $this.cmbParentGroup.Items.Clear()
            $groupList | ForEach-Object {
                $this.cmbParentGroup.Items.Add($_.Name)
            }
            $this.cmbParentGroup.SelectedIndex = 0
        } else {
            $this.cmbParentGroup.SelectedIndex = -1
        }
    }

    # Update the fields that are displayed based on group type selection
    $Global:createGroupForm | Add-Member -Type ScriptMethod -name 'updateVisibility' -Value {
        $selectedGroup = $Global:createGroupForm.cmbGroupType.SelectedValue.Content.ToString().ToUpper()
        switch ($selectedGroup) {
            'COMPUTER' {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Visible'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpGroupName.Visibility = 'Visible'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Visible'
                $Global:createGroupForm.grpUseStaging.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Department/Unit name:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Collapsed'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $True
            }
            'GPO' {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpGroupName.Visibility = 'Visible'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Department/Unit name:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Collapsed'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $False
                $Global:createGroupForm.txtDepartmentName.Text = ''
            }
            'PRINTER' {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Visible'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Visible'
                $Global:createGroupForm.grpShareName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpGroupName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Visible'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Authorized approvers:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Visible'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $True
            }
            'SHARE' {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareName.Visibility = 'Visible'
                $Global:createGroupForm.grpGroupName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Visible'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Authorized approvers:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Visible'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $True
            }
            'USER' {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Visible'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpGroupName.Visibility = 'Visible'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Visible'
                $Global:createGroupForm.grpUseStaging.Visibility = 'Collapsed'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Department/Unit name:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Collapsed'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $True
            }
            Default {
                $Global:createGroupForm.grpGroupType.Visibility = 'Visible'
                $Global:createGroupForm.grpUnitName.Visibility = 'Visible'
                $Global:createGroupForm.grpParentGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpSupportGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpUpdateGroup.Visibility = 'Visible'
                $Global:createGroupForm.grpCampusInitial.Visibility = 'Visible'
                $Global:createGroupForm.grpPrintServer.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrinterName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpShareName.Visibility = 'Collapsed'
                $Global:createGroupForm.grpGroupName.Visibility = 'Visible'
                $Global:createGroupForm.grpShareOptions.Visibility = 'Collapsed'
                $Global:createGroupForm.grpPrintInfo.Visibility = 'Collapsed'
                $Global:createGroupForm.grpCompOptions.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Visibility = 'Visible'
                $Global:createGroupForm.lblGroupName.Text = ''
                $Global:createGroupForm.lblDepartmentName.Content = 'Department/Unit name:'
                $Global:createGroupForm.lblApproverOptional.Visibility = 'Collapsed'
                $Global:createGroupForm.txtDepartmentName.isEnabled = $True
            }
        }
    }

    # Returns a hash table containing form state information
    $Global:createGroupForm | Add-Member -Type ScriptMethod -name 'getState' -Value {
        $groupParams = @{}
        
        # query group type
        if ($this.cmbGroupType.SelectedValue.Content -ne $null) {
            $groupParams.GroupType = $this.cmbGroupType.SelectedValue.Content.ToString().ToUpper()
        }

        # query unit name
        if ($this.cmbUnitName.SelectedValue -ne $null) {
            $groupParams.GroupUnit = $this.cmbUnitName.SelectedValue.ToString()
        }

        # query parent group name
        if ($this.cmbParentGroup.SelectedValue -ne $null) {
            $groupParams.GroupParent = $this.cmbParentGroup.SelectedValue.ToString()
        }

        # query campus initial
        $anInitial = $this.cmbCampusInitial.SelectedValue.Content
        if ( ($anInitial -ne $null) -and ($anInitial.length -eq 1) ) {
            $groupParams.CampusInitial = $anInitial.ToUpper()
        }

        # Collect resource server and path information
        switch ($groupParams.GroupType) {
            'PRINTER' {
                $groupParams.ResourceServer = $this.cmbPrintServer.SelectedValue.Content.ToString()
                $groupParams.ResourcePath = $this.cmbPrinterName.SelectedValue
            }
            'SHARE' {
                $path = $this.txtShareName.Text
                $nameSplitter = ([RegEx]'\\\\(?<server>[^\\]+)\\(?<path>.*)').Match($path)

                if ($nameSplitter.Success) {
                    $groupParams.ResourceServer = $nameSplitter.Groups.Item('server').Value
                    $groupParams.ResourcePath = $nameSplitter.Groups.Item('path').Value
                    $groupParams.AccessType = 'RW'
                }
            }
            Default {
                # query name/suffix portion of desired name
                $aSuffix = $this.txtGroupName.Text
                if ( ($aSuffix -ne $null) -and ($aSuffix.length -gt 0) ) {
                    $groupParams.GroupSuffix = $aSuffix
                }
            }
        }
        return $groupParams
    }

    # updatePrefix - Generates the security group name prefix based on the selected group type, unit, paren OU and campus initial
    $Global:createGroupForm | Add-Member -Type ScriptMethod -name 'updatePrefix' -Value {
        $myState = $this.getState()
        try {
            
            # Temporarily skip AD query for shares, because it makes the UI freeze
            if ($myState.groupType -imatch 'SHARE') {
                $myState.SkipADQuery = $true
            }
            $newName = ConvertTo-GroupName @myState
        } catch {
            $newName = ''
        }
        $this.lblGroupName.Text = "${newName}"
    }

    #endregion

    #region prepareForm

    # Create runspace for long-running tasks
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash",$Global:createGroupForm)
    $Runspace.SessionStateProxy.SetVariable("includePath",$includePath)

    # Populate unit list
    $unitFilter = [Regex]"(\w+)(.Groups)$"
    Get-ADOrganizationalUnit -Filter 'Name -like "*.Groups" -and Name -ne "M.UTOSPA.Groups"' -SearchBase $globalOUPath -SearchScope 2 -Server $global:ADserver | Sort-Object | Foreach-Object {
        $Global:createGroupForm.cmbUnitName.Items.Add($unitFilter.Match($_.Name).Groups[1].Value)
    }
    $Global:createGroupForm.cmbGroupType.Visibility = 'Visible'
    $Global:createGroupForm.cmbUnitName.Visibility = 'Visible'

    # Populate support groups
    $supportFilter = [Regex]'(?i)^(M\.UTOSPA_Groups\.SupportTeam\.)(?<team>[^\.]+)'
    Get-ADGroup -Filter 'Name -like "M.UTOSPA_Groups.SupportTeam*" ' -SearchBase $globalOUPath -SearchScope 'SubTree' -Server $global:ADserver | Sort-Object | Foreach-Object {
        $teamParts = $supportFilter.Match($_.Name)
        if ($teamParts.success) {
            $Global:createGroupForm.cmbSupportGroup.Items.Add($teamParts.Groups.item('team').value)
        }
    }

    # change support group if it exists in settings file
    $supGroup = $Global:appSettings.get('supportGroup')
    if ($supGroup -is [String]) {
        :selectType for ($i = 0; $i -lt $Global:createGroupForm.cmbSupportGroup.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbSupportGroup.Items[$i] -imatch $supGroup) {
                $Global:createGroupForm.cmbSupportGroup.SelectedIndex = $i
                break selectType
            }
        }
    }

    #endregion

    #region populatefields

    # If the caller provided us with arguments,
    # we need to populate the form with that data before subscribing listeners
    if (
        ($GroupName -is [String]) -and
        ($GroupName.length -gt 8)
    ) {

        # Convert from group name to parts object
        $groupParts = ConvertFrom-GroupName -GroupName $GroupName

        # Enable the bulk group warning and update the label (TextBlock)
        $Global:createGroupForm.grpBulkGroupName.Visibility = 'Visible'
        $Global:createGroupForm.lblBulkGroupName.Text = "${GroupName}"

        # Type
        :selectType for ($i = 0; $i -lt $Global:createGroupForm.cmbGroupType.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbGroupType.Items[$i].Content -imatch $groupParts.type) {
                $Global:createGroupForm.cmbGroupType.SelectedIndex = $i
                $Global:createGroupForm.updateVisibility()
                break selectType
            }
        }

        # Unit
        :selectUnit for ($i = 0; $i -lt $Global:createGroupForm.cmbUnitName.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbUnitName.Items[$i] -imatch $groupParts.unit) {
                $Global:createGroupForm.cmbUnitName.SelectedIndex = $i
                $Global:createGroupForm.updateParentList()
                $Global:createGroupForm.updatePrefix()
                break selectUnit
            }
        }

        # Parent
        :selectParent for ($i = 0; $i -lt $Global:createGroupForm.cmbParentGroup.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbParentGroup.Items[$i] -imatch $groupParts.parent) {
                $Global:createGroupForm.cmbParentGroup.SelectedIndex = $i
                $Global:createGroupForm.updatePrefix()
                break selectParent
            }
        }

        # Set resource path field for those types
        if ($groupParts.type -imatch 'SHARE|PRINTER') {
            $Global:createGroupForm.txtShareName.Text = $groupParts.resource
        } elseif ($groupParts.type -imatch 'GPO') {
            $Global:createGroupForm.txtShareName.Text = $groupParts.suffix
        }

        # Name/Suffix information
        $Global:createGroupForm.txtGroupName.Text = $groupParts.suffix

        # Department info
        if ($GroupDept -is [String]) {
            $Global:createGroupForm.txtDepartmentName.Text = $GroupDept
        }

        # Select support group
        :selectP for ($i = 0; $i -lt $Global:createGroupForm.cmbSupportGroup.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbSupportGroup.Items[$i] -imatch $SupportGroup) {
                $Global:createGroupForm.cmbSupportGroup.SelectedIndex = $i
                break selectP
            }
        }

        # Select update group
        :selectP for ($i = 0; $i -lt $Global:createGroupForm.cmbUpdateGroup.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbUpdateGroup.Items[$i] -imatch $UpdateGroup) {
                $Global:createGroupForm.cmbUpdateGroup.SelectedIndex = $i
                break selectP
            }
        }

        # Ticket info
        if ($groupTicket -is [String]) {
            $Global:createGroupForm.txtTicketNumber.Text = $GroupTicket
        }
    } else {

        # Set unit to last selected, if that entry exists in the settings file
        $savedUnit = $Global:appSettings.get('unit')
        if ($savedUnit -is [String]) {
            :selectUnit for ($i = 0; $i -lt $Global:createGroupForm.cmbUnitName.Items.Count; $i++) {
                if ($Global:createGroupForm.cmbUnitName.Items[$i] -imatch $savedUnit) {
                    $Global:createGroupForm.cmbUnitName.SelectedIndex = $i
                    $Global:createGroupForm.updateParentList()
                    $Global:createGroupForm.updatePrefix()
                    break selectUnit
                }
            }
        }
    }

    #endregion


    #region formEvents

    # Update dependent fields when group type is changed
    $Global:createGroupForm.cmbGroupType.add_SelectionChanged({
        $Global:createGroupForm.updateParentList()
        $Global:createGroupForm.updateVisibility()
        $Global:createGroupForm.updatePrefix()
    })

    # Open AD standards KB when link is clicked
    $Global:createGroupForm.lnkOpenKB.add_click({
        $selectedGroup = $Global:createGroupForm.cmbGroupType.SelectedValue.Content.ToString().ToUpper()
        switch ($selectedGroup) {
            'COMPUTER' {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015392")
            }
            'GPO' {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015382")
            }
            'PRINTER' {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015386")
            }
            'SHARE' {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015390")
            }
            'USER' {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015393")
            }
            Default {
                (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0014939")
            }
        }
    })

    # Update list of parent groups when a unit is selected
    $Global:createGroupForm.cmbUnitName.add_SelectionChanged({
        $Global:createGroupForm.updateParentList()
        $Global:createGroupForm.updatePrefix()
    })

    # Update group prefix when parent group is changed
    $Global:createGroupForm.cmbParentGroup.add_SelectionChanged({
        $Global:createGroupForm.updatePrefix()
    })

    # Update group prefix when campus initial is changed
    $Global:createGroupForm.cmbCampusInitial.add_SelectionChanged({
        $Global:createGroupForm.updatePrefix()
    })

    # Update list of printers when print server is selected
    $Global:createGroupForm.cmbPrintServer.add_SelectionChanged({

        # Cancel if ComboBox is disabled
        # this indicates that another task may be running against it
        if ($Global:createGroupForm.cmbPrinterName.isEnabled -eq $false) {
            return
        }

        # Disable the printer list until the query completes
        $Global:createGroupForm.cmbPrinterName.SelectedIndex = -1
        $Global:createGroupForm.cmbPrinterName.isEnabled = $false
        $Global:createGroupForm.cmbPrintServer.isEnabled = $false
        $Global:createGroupForm.lblPrintersLoading.Visibility = 'Visible'

        $code = {
            . "${includePath}\global.ps1"
            Start-Transcript -Path (Join-Path $Global:logdir "getPrinters-$((get-date).ToFileTime())") -IncludeInvocationHeader
            $script:compName = $null
            $syncHash.Window.Dispatcher.invoke(
                [action]{ $script:compName = $syncHash.cmbPrintServer.SelectedValue.Content.ToString() + '.ASURITE.AD.ASU.EDU'}
            )

            $script:printers = Get-Printer -ComputerName $compName | Sort-Object -Property 'Name'

            if ($script:printers -eq $null) {
                $syncHash.Window.Dispatcher.invoke(
                    [action]{
                        $syncHash.cmbPrintServer.isEnabled = $true
                        $syncHash.cmbPrinterName.isEnabled = $true
                        $syncHash.lblPrintersLoading.Visibility = 'Collapsed'
                    }
                )
                exit 9
            }

            # Place printer list in SyncHash for object reference safety
            $syncHash.printerList.clear()
            $script:printers | ForEach-Object {
                $syncHash.printerList.Add("$($_.Name)")
            }
            $syncHash.window.Dispatcher.invoke(
                [action]{$syncHash.cmbPrinterName.Items.Refresh()}
            )

            # Determine if a printer name exists in the main form resource field and attempt to resolve that printer in our list
            $script:resourcePath = $null
            $syncHash.window.Dispatcher.invoke(
                [action]{$script:resourcePath = $syncHash.txtShareName.Text}
            )
            if ($script:resourcePath -ne $null) {
                $script:serverName = ([RegEx]'(?i)^\\\\(?<server>[^\\]*)\\(?<path>.*)').Match($script:resourcePath)
                $script:serverName = $serverName.Groups.Item('path').value
                :selectPrinter for ($i = 0; $i -lt $syncHash.printerList.count; $i++) {
                    if ($syncHash.printerList[$i] -imatch $serverName) {
                        $global:printerIndex = $i
                        $syncHash.window.Dispatcher.invoke(
                           [action]{$syncHash.cmbPrinterName.SelectedIndex = $global:printerIndex}
                        )
                        break selectPrinter
                    }
                }
            }

            # Unlock the comboboxes
            $syncHash.window.Dispatcher.invoke(
                [action]{
                    $syncHash.cmbPrintServer.isEnabled = $true
                    $syncHash.cmbPrinterName.isEnabled = $true
                    $syncHash.lblPrintersLoading.Visibility = 'Collapsed'
                }
            )
            Stop-Transcript
        }
        $PSinstance = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })

    # Update printer information when printer selection changes
    $Global:createGroupForm.cmbPrinterName.add_SelectionChanged({

        # Cancel if ComboBox is disabled
        # or if index is still set to default value of -1
        # this indicates that another task may be running against it
        if (
            ($Global:createGroupForm.cmbPrinterName.isEnabled -eq $false) -or
            ($Global:createGroupForm.cmbPrinterName.selectedIndex -lt 0) -or
            ($Global:createGroupForm.cmbPrinterName.selectedValue -isnot [String])
        ) {
            return
        }

        # Refresh prefix data now, as it may be unsafe to call it from the
        # spawned thread later (race condition?)
        $Global:createGroupForm.updatePrefix()

        # Disable the text fields until this process is complete
        $Global:createGroupForm.txtPrinterMakeModel.isEnabled = $false
        $Global:createGroupForm.txtPrinterIP.isEnabled = $false
        $Global:createGroupForm.txtPrinterLocation.isEnabled = $false

        $code = {
            . "${includePath}\global.ps1"
            Start-Transcript -Path (Join-Path $Global:logdir "getPrinterInfo-$((get-date).ToFileTime())") -IncludeInvocationHeader

            $script:compName = $null
            $script:printerName = $null
            $syncHash.Window.Dispatcher.invoke(
                [action]{
                    $script:compName = $syncHash.cmbPrintServer.SelectedValue.Content.ToString() + '.ASURITE.AD.ASU.EDU'
                    $script:printerName = $syncHash.cmbPrinterName.SelectedValue.ToString()
                }
            )

            # Check that fields are not blank and attempt query
            $script:printerInfo = $null
            if ( ($script:compName -ne $null) -and ($script:printerName -ne $null)) {
                $script:printerInfo = Get-Printer -ComputerName $script:compName -Name $script:printerName
            }

            if ($script:printerInfo -eq $null) {
                $syncHash.Window.Dispatcher.invoke(
                    [action]{
                        $syncHash.txtPrinterMakeModel.isEnabled = $true
                        $syncHash.txtPrinterIP.isEnabled = $true
                        $syncHash.txtPrinterLocation.isEnabled = $true
                    }
                )
                eexit 404
            }

            # Populate form with found data and re-enable fields before exit
            $syncHash.Window.Dispatcher.invoke(
                [action]{
                    $syncHash.txtPrinterMakeModel.Text = $script:printerInfo.DriverName
                    $syncHash.txtPrinterIP.Text = $script:printerInfo.PortName -ireplace 'IP_',''
                    $syncHash.txtPrinterLocation.Text = $script:printerInfo.Location
                    $syncHash.txtPrinterMakeModel.isEnabled = $true
                    $syncHash.txtPrinterIP.isEnabled = $true
                    $syncHash.txtPrinterLocation.isEnabled = $true
                }
            )
            exit 9
            Stop-Transcript
        }
        $PSinstance = [powershell]::Create().AddScript($Code)
        $PSinstance.Runspace = $Runspace
        $PSinstance.BeginInvoke()
    })

    # Update group name recommendation when share name changes
    $Global:createGroupForm.txtShareName.add_TextChanged({
        $Global:createGroupForm.updatePrefix()
    })

    $Global:createGroupForm.txtGroupName.add_TextChanged({
        $Global:createGroupForm.updatePrefix()
    })

    $Global:createGroupForm.btnOk.Add_Click({

        # Perform validation before closing the form
        $errorString = $null

        # Pull in all the fields we need to check
        $formState = $Global:createGroupForm.getState()

        # Check that generated name looks correct
        if ($Global:createGroupForm.lblGroupName.Text.length -lt 8) {
            $errorString = 'The generated name is not valid. Check that all required fields have been selected and that the generated name looks correct.'
        }
        
        # Common validations across all types
        if (!$formState.GroupType -or ($formState.GroupType.length -lt 2)) {
            $errorString = 'Could not determine selected group type. Make sure a group type has been selected.'
        } elseif ($formState.GroupUnit -isnot [String]) {
            $errorString = 'Unit name was not valid, make sure a unit has been selected'
        } elseif ($formState.GroupUnit.length -lt 2) {
            $errorString = 'Unit name is too short'
        } elseif ( (ConvertTo-GroupName @formState).length -gt 64 ) {
            $errorString = 'The generated name is too long. Please modify the name to be less than 64 characters.'
        }

        # Only perform additional validation if no errors were found
        if ($errorString -eq $null) {
            switch ($formState.GroupType) {
                'COMPUTER' {

                    # If user selected to import from staging OU, make sure they selected a staging number
                    if ($Global:createGroupForm.chkUseStaging.isChecked -and ($Global:createGroupForm.cmbStagingOUNumber.SelectedIndex -lt 0) ) {
                        $errorString = 'The staging area option is selected, but no staging area was selected. Either disable the use of the staging area, or select a staging area to use.'
                    }

                    # A parent group is required
                    if ($formState.GroupParent -isnot [String] -or ($formState.GroupParent.length -lt 8)) {
                        $errorString = 'Parent group is not selected or does not appear to be valid'
                    }
                }
                'USER' {

                    # If user selected to import from staging OU, make sure they selected a staging number
                    if ($Global:createGroupForm.chkUseStaging.isChecked -and ($Global:createGroupForm.cmbStagingOUNumber.SelectedIndex -lt 0) ) {
                        $errorString = 'The staging area option is selected, but no staging area was selected. Either disable the use of the staging area, or select a staging area to use.'
                    }

                    # A parent group is required
                    if ($formState.GroupParent -isnot [String] -or ($formState.GroupParent.length -lt 8)) {
                        $errorString = 'Parent group is not selected or does not appear to be valid'
                    }
                }
                Default {
                    # No additional common validation needed
                }
            }
        }

        # Display error if it exists, otherwise, close the form
        if ($errorString) {
            PopupBox $errorString
        } else {
            $Global:createGroupForm.okButtonClicked = $true
            $Global:createGroupForm.window.close()
        }
    })

    # Cancel button clicked
    $Global:createGroupForm.btnCancel.add_Click({
        $Global:createGroupForm.okButtonClicked = $false
        $Global:createGroupForm.window.close()
    })

    # Form visibility changed
    $Global:createGroupForm.window.add_IsVisibleChanged({
        if ($Global:createGroupForm.window.isVisible -eq $true) {
            $Global:createGroupForm.window.topmost = $true
            $Global:createGroupForm.window.topmost = $false
            $Global:createGroupForm.window.focus()
            $Global:createGroupForm.updatePrefix()
        }
    })

    #endregion

    #region execStart

    # today's date
    $today = (get-date).ToShortDateString()

    # get the staging OU
    $stagingOUs = Get-ADOrganizationalUnit -SearchBase 'OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu' -SearchScope Subtree -Filter * -Server $global:ADserver | Where-Object { $_.DistinguishedName -notlike 'OU=M.UTOSPA_Staging,*' } |Sort-Object Name

    # stagingAccess states whether or not they have access to staging OU
    $stagingAccess = $false

    # can the user see staging area?
    if (($stagingOUs | Measure-Object).count -gt 0) {
        # yup.
        $stagingAccess = $true
    }

    if ($globalDebug) {$syncHash.print("DEBUG: Form load complete")}
    if ($Global:createGroupForm.cmbGroupType.SelectedIndex -lt 0) {
        $Global:createGroupForm.cmbGroupType.SelectedIndex = 0
    }

    # If the resource path field was populated by a prior step and
    # that field contains one of the printer names in the printers combobox,
    # Atempt to resolve the server
    $serverName = ([RegEx]'(?i)^\\\\(?<server>[^\\]*)').Match($Global:createGroupForm.txtShareName.Text)
    if ($serverName.success) {
        $serverName = $serverName.Groups.Item('server').value
    }
    if ($serverName -is [String]) {
        :selectServer for ($i = 0; $i -lt $Global:createGroupForm.cmbPrintServer.Items.Count; $i++) {
            if ($Global:createGroupForm.cmbPrintServer.Items[$i].Content -imatch $serverName) {
                $Global:createGroupForm.cmbPrintServer.SelectedIndex = $i
                $Global:createGroupForm.updatePrefix()
                break selectServer
            }
        }
    }
    
    $Global:createGroupForm.window.showdialog()

    # Determine if the form was closed using the OK button
    if ($Global:createGroupForm.okButtonClicked -eq $false) {
        return 999
    } else {
        $Global:createGroupForm.updatePrefix()
    }

    # Gather information to create groups
    $formState = $Global:createGroupForm.getState()
    $script:ouPath = "OU=M.UTOSPA.$($formState.groupUnit).Groups,OU=M.UTOSPA.$($formState.groupUnit),OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu"
    $listOfGroups = New-Object -TypeName 'System.collections.ArrayList'
    switch ($formState.GroupType) {
        'COMPUTER' {
            $compGroup = @{
                name=$Global:createGroupForm.lblGroupName.Text
                description=''
                notes=''
                path=$script:ouPath
                parent=$formState.GroupParent
            }

            # Generate the Description
            $compGroup.Description = "Computers of $($Global:createGroupForm.txtDepartmentName.Text)"

            # Generate the notes
            $compGroup.notes = @"
$($compGroup.Description)
Ticket Number: $($Global:createGroupForm.txtTicketNumber.Text)
Created: ${today}
Tech: ${env:username}
"@

            # Add the group to the list
            $listOfGroups.add($compGroup)

            # Add additional group if the option was selected
            if ($Global:createGroupForm.chkCreateCmpUsrPair.isChecked) {
                $compGroup2 = @{
                    name=''
                    description=''
                    notes=''
                    path=$script:ouPath
                }
                # Modify name
                $compGroup2.name = $compGroup.name -replace '.CMP.','.USR.'

                # modify description
                $compGroup2.description = $compGroup.description -replace 'Computers of','Users of'

                # modify notes
                $compGroup2.notes = $compGroup.notes -replace 'Computers of','Users of'

                # Modify parent name
                $compGroup2.parent = $compGroup.parent -replace '.CMP','.USR'

                $listOfGroups.add($compGroup2)
            }
        }
        'GPO' {
            $compGroup = @{
                name=$Global:createGroupForm.lblGroupName.Text
                description=''
                notes=''
                path=$script:ouPath
            }

            # Generate the Description
            $compGroup.Description = "GPO filter for $($Global:createGroupForm.lblGroupName.Text)"

            # Generate the notes
            $compGroup.notes = @"
Membership filters the application of policies in $($Global:createGroupForm.lblGroupName.Text)
Ticket Number: $($Global:createGroupForm.txtTicketNumber.Text)
Created: ${today}
Tech: ${env:username}
"@
            $listOfGroups.add($compGroup)
        }
        'PRINTER' {
            $compGroup = @{
                name=$Global:createGroupForm.lblGroupName.Text
                description=''
                notes=''
                path=$script:ouPath
            }

            # Generate the Description
            $compGroup.Description = "Access to Printer \\$($Global:createGroupForm.cmbPrintServer.SelectedValue.Content)\$($Global:createGroupForm.cmbPrinterName.SelectedValue)"
            $pmodel = $Global:createGroupForm.txtPrinterMakeModel.Text
            $pip = $Global:createGroupForm.txtPrinterIP.Text
            $ploc = $Global:createGroupForm.txtPrinterLocation.Text
            $compGroup.notes = @"
$($compGroup.Description)
make & model: ${pmodel}
Location: ${ploc}
IP: ${pip}
Authorized approver: $($Global:createGroupForm.txtDepartmentName.Text)
ticket: $($Global:createGroupForm.txtTicketNumber.Text)
Created: ${today}
Tech: ${env:username}
"@
            $listOfGroups.add($compGroup)
        }
        'SHARE' {

            # We need to create several groups depending on which ones the user selected

            # Start by generating a common template
            $nameBuilder = $Global:createGroupForm.lblGroupName.Text

            # Generate a template for the description
            $descTemplate = "[ACCESSTYPE] Access to Share $($Global:createGroupForm.txtShareName.Text)"
            $notesTemplate = @"
Membership [ACCESSTYPE] access to $($Global:createGroupForm.txtShareName.Text)
Authorized approver: $($Global:createGroupForm.txtDepartmentName.Text)
Ticket Number: $($Global:createGroupForm.txtTicketNumber.Text)
Created: ${today}
Tech: ${env:username}
"@
            if ($Global:createGroupForm.chkAlsoCreateRW.IsChecked) {
                $rwGroup = @{
                    name=$nameBuilder
                    description=($descTemplate -Replace '\[ACCESSTYPE\]','Read/Write')
                    notes=($notesTemplate -Replace '\[ACCESSTYPE\]','grants Read/Write')
                    path=$script:ouPath
                }

                $listOfGroups.add($rwGroup)
            }
            if ($Global:createGroupForm.chkAlsoCreateLO.IsChecked) {
                $loGroup = @{
                    name=($nameBuilder -Replace '-RW','-LO')
                    description=($descTemplate -Replace '\[ACCESSTYPE\]','List contents')
                    notes=($notesTemplate -Replace '\[ACCESSTYPE\]','grants list')
                    path=$script:ouPath
                }

                $listOfGroups.add($loGroup)
            }
            if ($Global:createGroupForm.chkAlsoCreateRO.IsChecked) {
                $reGroup = @{
                    name=($nameBuilder -Replace '-RW','-RO')
                    description=($descTemplate -Replace '\[ACCESSTYPE\]','Read')
                    notes=($notesTemplate -Replace '\[ACCESSTYPE\]','grants read')
                    path=$script:ouPath
                }

                $listOfGroups.add($reGroup)
            }
            if ($Global:createGroupForm.chkAlsoCreateDA.IsChecked) {
                $daGroup = @{
                    name=($nameBuilder -Replace '-RW','-DA')
                    description=($descTemplate -Replace '\[ACCESSTYPE\]','Deny all')
                    notes=($notesTemplate -Replace '\[ACCESSTYPE\]','denies all')
                    path=$script:ouPath
                }

                $listOfGroups.add($daGroup)
            }
        }
        'USER' {
            $compGroup = @{
                name=$Global:createGroupForm.lblGroupName.Text
                description=''
                notes=''
                path=$script:ouPath
                parent=($formState.GroupParent -Replace '\.(USR|CMP)\.','.USR.') -Replace '\.$',''
            }

            # Generate the Description
            $compGroup.Description = "Users of $($Global:createGroupForm.txtDepartmentName.Text)"

            # Generate the notes
            $compGroup.notes = @"
$($compGroup.Description)
Ticket Number: $($Global:createGroupForm.txtTicketNumber.Text)
Created: ${today}
Tech: ${env:username}
"@
            # Add the group to the list
            $listOfGroups.add($compGroup)

            # Add additional group if the option was selected
            if ($Global:createGroupForm.chkCreateCmpUsrPair.isChecked) {
                $compGroup2 = @{
                    name=''
                    description=''
                    notes=''
                    path=$script:ouPath
                }
                # Modify name
                $compGroup2.name = $compGroup.name -replace '.USR.','.CMP.'

                # modify description
                $compGroup2.description = $compGroup.description -replace 'Users of','Computers of'

                # modify notes
                $compGroup2.notes = $compGroup.notes -replace 'Users of','Computers of'

                # Modify parent
                $compGroup2.parent = $compGroup.parent -replace '.USR','.CMP'
                $listOfGroups.add($compGroup2)
            }
        }
        Default {
        }
    }

    # WE ARE ALMOST DONE!
    # We simply need to use the data in the ArrayList to create the respective groups!

    # Get confirmation that list of groups looks correct
    $ccMessage = @"
The following groups are about to be created. Click Yes to continue, or click No to abort the process.

$($listOfGroups.name | Out-String)
"@
    $createConfirm = PopupBox $ccMessage "Confirm creation" "yn"
    if ($createConfirm -ne 'yes') {
        $syncHash.print("[ WRN ] Group creation aborted")
        return
    }
    foreach ($newGroup in $listOfGroups) {

        # check that the group doens't already exist
        $gcError = $null
        if ($newGroup.name.length -gt 4) {
            $checkGroupExsist = Get-ADObject -SearchBase (Get-ADRootDSE).defaultNamingContext -SearchScope 'SubTree' -Filter {name -eq "$($newGroup.name)"} -Server $global:ADserver

            if ($checkGroupExsist -ne $null) {
                $gcError = "$($newGroup.name) already exists: skipping"
            }
        } else {
            $gcError = "Aboring group creation: generated name was not valid"
        }

        if (!$gcError -and $newGroup.name.Length -gt 64) {
            $gcError = "$($newGroup.name) Generated group name is too long: skipping"
        }

        # If simulation is selected, create in testing OU instead
        if ($global:debug -and ($Global:createGroupForm.chkDebugCreation.IsChecked)) {
            $newGroup.path = 'OU=M.UTOSPA_Test,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu'
        }

        # Create the group
        if (!$gcError) {
            $syncHash.print("[ OK! ] Creating new group: $($newGroup.name)")
            if ($global:debug -eq $true) {($newGroup | Format-List)}
            try {
                New-ADGroup -Name $newGroup.name -GroupCategory Security -GroupScope Global -Path $newGroup.path -Description $newGroup.description -OtherAttributes @{info=$newGroup.notes} -ErrorAction Stop -Server $global:ADserver
            } catch {
                $gcError = "$($newGroup.name) encountered an unexpected error when creating this group $($_.message)"
            }

            # Wait for AD to catch up with us
            if (!$gcError) {
                $chkTimer = [DateTime]::Now
                $chkResult = $False
                while ($chkResult -eq $False) {
                    $checkGroup = $null
                    try {
                        $checkGroup = Get-ADGroup -Identity $newGroup.name -Server $global:ADserver
                    }
                    catch {
                        $checkGroup = $null
                    }

                    if ($checkGroup -eq $null) {
                        $elapsedTime = [DateTime]::Now - $chkTimer
                        $syncHash.print("[ WRN ] Waiting for AD to sync changes... ($($elapsedTime.Minutes):$($elapsedTime.Seconds).$($elapsedTime.Milliseconds) elapsed)")
                        $chkResult = $False
                        Start-Sleep -Seconds 10
                    } else {
                        $chkResult = $True
                    }
                }
            }

            # check if nesting is wanted
            if (!$gcError -and ($Global:createGroupForm.chkNesting.IsChecked) -and ($Global:createGroupForm.cmbGroupType.SelectedValue.Content.ToUpper() -Match '^(USER|COMPUTER)$') ) {
                # try and nest the new group within its parent
                $syncHash.print("[ OK! ] Nesting $($newGroup.name) within $($newGroup.parent)")
                try {
                    Add-ADGroupMember -Identity $newGroup.parent -Members $newGroup.name -Server $global:ADserver
                } catch {
                    $gcError = "$($newGroup.name) encountered an unexpected error when nesting group to parent"
                }
            }

            # check if staging population is wanted
            if (($Global:createGroupForm.chkUseStaging.IsChecked) -and ($newGroup.name -imatch 'CMP') ) {

                # Get computers from staging area
                $stagingNumber = $Global:createGroupForm.cmbStagingOUNumber.SelectedValue.Content.ToString()
                try {
                    # add the staging computers to the new group.
                    $stagingComps = Get-ADComputer -SearchBase "OU=M.UTOSPA_Staging.${stagingNumber},OU=M.UTOSPA_Staging,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu" -SearchScope 'OneLevel' -Filter '*' -Properties 'Name' -Server $global:ADserver
                    $stagingComps | Format-Table 'Name'
                    Add-ADGroupMember -Identity $newGroup.name -Members $stagingComps -ErrorAction stop -Server $global:ADserver
                } catch {
                    $gcError = "$($newGroup.name) encountered an unexpected error when populating group with staging area"
                }
            }

            # Determine if we need to add the group to a support group
            if (
                ($newGroup.name -imatch 'CMP') -and
                ($Global:createGroupForm.cmbSupportGroup.SelectedIndex -gt -1)
            ) {
                $supSuffix = $Global:createGroupForm.cmbSupportGroup.SelectedItem
                $supGroupname = "M.UTOSPA_Groups.SupportTeam.${supSuffix}"

                $syncHash.print("[ OK! ] Adding $($newGroup.name) to suport group $($supGroupname)")
                try {
                    Add-ADGroupMember -Identity $supGroupname -Members $newGroup.name -Server $global:ADserver
                } catch {
                    $gcError = "$($newGroup.name) encountered an unexpected error when adding to support group"
                }
            }

            # Determine if we need to add to update group
            if (
                ($newGroup.name -imatch 'CMP') -and
                ($Global:createGroupForm.cmbUpdateGroup.SelectedIndex -gt 0)
            ) {
                $supSuffix = $Global:createGroupForm.cmbUpdateGroup.SelectedItem.Content
                $supGroupname = "M.UTOSPA_Groups.Software.REQ.Win10_Update_Latest_Release_${supSuffix}"

                $syncHash.print("[ OK! ] Adding $($newGroup.name) to update group $($supGroupname)")
                try {
                    Add-ADGroupMember -Identity $supGroupname -Members $newGroup.name -Server $global:ADserver
                } catch {
                    $gcError = "$($newGroup.name) encountered an unexpected error when adding to update group"
                }
            }

            # If group was a SHR-DA group, add it to the global Deny All group
            if ($newGroup.name -imatch 'SHR-DA') {
                $syncHash.print("[ OK! ] Adding $($newGroup.name) to global deny group")
                try {
                    Add-ADGroupMember -Identity 'CN=M.UTOSPA_Groups.Access.DenyALL,OU=M.UTOSPA_Groups.Access,OU=M.UTOSPA_Groups,OU=M.UTOSPA,DC=asurite,DC=ad,DC=asu,DC=edu' -Members $newGroup.name -Server $global:ADserver
                } catch {
                    $gcError = "$($newGroup.name) encountered an unexpected error when adding to update group"
                }
            }

            # Display any errors in log window
            if ($gcError) {
                $syncHash.print("[ERROR] ${gcError}")
            } else {
                $syncHash.print("[ OK! ] $($newGroup.name) created successfuly!")
            }
        } else {
            $syncHash.print("[ERROR] $($newGroup.name) encountered error $($newGroup.parent)")
        }

        # All done!

    }
    
    #endregion

    #region saveSettings

    # Save most recently used unit
    $Global:appSettings.set('unit',$Global:createGroupForm.cmbUnitName.SelectedValue.ToString())

    #endregion

    $syncHash.print('[ OK! ] Group creation workflow completed')
    return 0
}

function ImportGroup {

    # This function requires that the ADServer global variable be set.
    # Typically, this should happen when the global.ps1 file is loaded
    # on creation of a new DSCtrl thread. If this has not happened, we
    # default to ASURITE6 so we don't pass a null string later in the function
    if ( ($global:ADserver -isnot [String]) -or ($global:ADserver.length -lt 12) ) {
        $global:ADserver = 'asurite6.asurite.ad.asu.edu'
    }

    # Print start of operation to main form
    $syncHash.print("[ OK! ] Starting bulk import dialog...")

    # Double-check that AD module is running
    if ([boolean](Get-Module -Name 'ActiveDirectory') -eq $false) {
        Import-Module -Name 'ActiveDirectory'
    }

    # Define the form
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
xmlns:local="clr-namespace:Deskside_Control_Panel"
Title="Bulk Import" Height="480" MinHeight="300" Width="900" MinWidth="720">
    <DockPanel Margin="8">
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="4" DockPanel.Dock="Top" Height="32">
            <Label Margin="4" Width="100">
                Ticket Number:
            </Label>
            <TextBox Name="txtTicketNumber" Width="128" Margin="4" />
            <Label Margin="4" Width="180">
                Department name/description:
            </Label>
            <TextBox Name="txtDepartment" Width="256" Margin="4" />
        </StackPanel>
        <StackPanel Orientation="Horizontal" HorizontalAlignment="Left" VerticalAlignment="Center" Margin="4" DockPanel.Dock="Top" Height="32">
            <Label Margin="4" Width="100">
                Update Group::
            </Label>
            <ComboBox Name="cmbUpdateGroup" Width="128" Margin="4">
                <ComboBoxItem>(None - inherit from parent)</ComboBoxItem>
                <ComboBoxItem>Group_1</ComboBoxItem>
                <ComboBoxItem>Group_2</ComboBoxItem>
                <ComboBoxItem>Group_3</ComboBoxItem>
                <ComboBoxItem>Group_4</ComboBoxItem>
            </ComboBox>
            <Label Margin="4" Width="180">
                Supported by:
            </Label>
            <ComboBox Name="cmbSupportGroup" Width="256" Margin="4" />
        </StackPanel>
        <StackPanel Orientation="Vertical" HorizontalAlignment="Right" VerticalAlignment="Center" Width="170" Margin="4" DockPanel.Dock="Right">
            <Button Content="Import from clipboard" Name="btnImport" Height="26" Margin="4"></Button>
            <Separator></Separator>
            <Button Content="Clear" Name="btnClear" Height="26" Margin="4"></Button>
            <Separator></Separator>
            <Button Content="Continue" Name="btnContinue"  Height="26" Margin="4"></Button>
            <Button Content="Cancel" Name="btnCancel" Height="26" Margin="4"></Button>
            <TextBlock Padding="4" HorizontalAlignment="Center">
            <Hyperlink Name="lnkOpenKB">
                Veiw online documentation
            </Hyperlink>
            </TextBlock>
        </StackPanel>
        <DockPanel Margin="4" DockPanel.Dock="Left">
            <DataGrid Name="groupGrid" AutoGenerateColumns="False" SelectionMode="Single" CanUserAddRows="False" Margin="2">
                <DataGrid.Columns>
                    <DataGridTextColumn Header="Name" Width="312" IsReadOnly="True" Binding="{Binding name}" />
                    <DataGridTextColumn Header="Type" Width="64" IsReadOnly="True" Binding="{Binding type}" />
                    <DataGridTextColumn Header="Parent Group" Width="256" IsReadOnly="True" Binding="{Binding parent}" />
                </DataGrid.Columns>
            </DataGrid>
        </DockPanel>
    </DockPanel>
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
    $Global:bulkGroupForm = [hashtable]::Synchronized(@{})
    $Global:bulkGroupForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:bulkGroupForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:bulkGroupForm.add('okButtonClicked',$false)
    $Global:bulkGroupForm.add('groupList',(New-Object -TypeName System.collections.ArrayList))
    $Global:bulkGroupForm.groupGrid.itemsSource = $Global:bulkGroupForm.groupList

    # Populate support groups
    $supportFilter = [Regex]'(?i)^(M\.UTOSPA_Groups\.SupportTeam\.)(?<team>[^\.]+)'
    Get-ADGroup -Filter 'Name -like "M.UTOSPA_Groups.SupportTeam*" ' -SearchBase $globalOUPath -SearchScope 'SubTree' -Server $global:ADserver | Sort-Object | Foreach-Object {
        $teamParts = $supportFilter.Match($_.Name)
        if ($teamParts.success) {
            $Global:bulkGroupForm.cmbSupportGroup.Items.Add($teamParts.Groups.item('team').value)
        }
    }

    # Event listeners
    $Global:bulkGroupForm.btnImport.add_Click({
        $data = Get-Clipboard -Format 'Text'
        $script:ticketNumber = ''

        # Check that first line contains header
        if ( ($data[0] -iMatch 'Enter Computer|Enter All|Exportable') -ne $true) {
            PopupBox 'There was a problem parsing the copied data. A header appears to be missing. Make sure to select all cells before copying the spreadsheet from Google Docs'
            return
        }

        # Attempt to get ticket number from first row
        $ticketResult = ([RegEx]'(?i)(?<ticket>TASK\d+)').Match($data[0])
        if ($ticketResult.Success) {
            $script:ticketNumber = $ticketResult.Groups.Item('ticket').value
            $Global:bulkGroupForm.Window.Dispatcher.invoke(
                [action]{
                    $Global:bulkGroupForm.txtTicketNumber.Text = $script:ticketNumber
                }
            )
        }

        # Check for department name
        $deptResult = ([RegEx]'(?i)DN:(?<dept>[\w\ ]+)').Match($data[0])
        if ($deptResult.Success) {
            $script:ticketNumber = $deptResult.Groups.Item('dept').value
            $Global:bulkGroupForm.Window.Dispatcher.invoke(
                [action]{
                    $Global:bulkGroupForm.txtDepartment.Text = $script:ticketNumber
                }
            )
        }

        # Attempt to read support group
        $supportGroup = ([RegEx]'(?i)SG:(?<supportGroup>[\w\ ]+)').Match($data[0])
        if ($supportGroup.success) {
            $groupName = $supportGroup.Groups.Item('supportGroup').Value
            :selectP for ($i = 0; $i -lt $Global:bulkGroupForm.cmbSupportGroup.Items.Count; $i++) {
                if ($Global:bulkGroupForm.cmbSupportGroup.Items[$i] -imatch $groupName) {
                    $Global:bulkGroupForm.cmbSupportGroup.SelectedIndex = $i
                    break selectP
                }
            }
        }

        # Attempt to read the updategroup
        # If we see 'Inherited from Parent' default to that option
        $updateGroup = ([RegEx]'(?i)(?<updateGroup>Group_\d)').Match($data[0])
        if ($data[0].Contains('Inherited from')) {
            $Global:bulkGroupForm.cmbUpdateGroup.SelectedIndex = 0
        } elseif ($updateGroup.success) {
            $groupName = $updateGroup.Groups.Item('updateGroup').Value
            :selectP for ($i = 0; $i -lt $Global:bulkGroupForm.cmbUpdateGroup.Items.Count; $i++) {
                if ($Global:bulkGroupForm.cmbUpdateGroup.Items[$i] -imatch $groupName) {
                    $Global:bulkGroupForm.cmbUpdateGroup.SelectedIndex = $i
                    break selectP
                }
            }
        }

        # Determine which column contains the group names
        $headers = $data[1].trim().Split("`t")
        $targetColumn = -1
        :findHeader for ($i = 0; $i -lt $headers.Count; $i++) {
            if ($headers[$i] -iMatch '(security)?group name$') {
                $targetColumn = $i
                Break findHeader
            }
        }

        # If header for group name was not found, abort
        if ($targetColumn -lt 0) {
            PopupBox 'Could not find the column header for group names. Please make sure you select all cells in the spreadsheet when copying from Google Docs'
            return
        }

        #Process members of the column and add to the list
        for ($i = 2; $i -lt $data.Count; $i++) {

            # Determine if entry is an example
            $isExample = $data[$i] -imatch 'Example:'
    
            $row = $data[$i].Split("`t")
            $cell = $row[$targetColumn].trim()
        
            # Remove any trailing .'s that sometimes crop up out of spreadsheet
            $cell = $cell -Replace '\.$',''
        
            # Remove any instances of itfs1 FQDN: we want folks to be using the shortened name
            $cell = $cell -Replace 'itfs1\.asu\.edu','itfs1'

            $nameparts = $null
            if ( ($cell -is [String]) -and (!$isExample) -and ($cell.length -gt 8) ) {
                $nameParts = ConvertFrom-GroupName -GroupName $cell
            }

            if ($nameParts -is [PSCustomObject]) {

                # Only add group if it does not already exist or if array is empty
                if (
                    ($Global:bulkGroupForm.groupList.count -eq 0) -or
                    ($Global:bulkGroupForm.groupList.name -inotcontains $nameParts.name)
                ) {
                    $Global:bulkGroupForm.groupList.add($nameParts)
                }
            }
        }

        # Determine if we need to add any additional groups to satisfy nesting
        if ($Global:bulkGroupForm.groupList.count -gt 0) {
            $AdditionalGroups = New-Object -TypeName System.Collections.ArrayList
            $Global:bulkGroupForm.groupList | ForEach-Object {
                if ( ($_.parent -is [String]) -and ($_.parent.length -gt 8) ) {

                    # Add to list if parent is not in local queue
                    if (
                        ($Global:bulkGroupForm.groupList.name -inotcontains $_.parent) -and
                        ($AdditionalGroups -inotcontains $_.parent)
                    ) {
                        $AdditionalGroups.add($_.parent)
                    }
                }
            }

            #TODO: Determine if any of the additional groups already exist in AD
            # If they do, remove them from the list
            ##$AdditionalGroups | 

            # Add any remaining groups to the queue
            $AdditionalGroups | ForEach-Object {
                $groupExists = $false
                try {
                    $groupExists = Get-ADGroup -Identity $_ -Server $global:ADserver
                }
                catch {
                    $groupExists = $false
                }

                if ($groupExists -eq $false) {
                    $nameParts = ConvertFrom-GroupName -GroupName $_
                    if ($nameParts -is [PSCustomObject]) {
                        
                    # Only add group if it does not already exist or if array is empty
                    if (
                        ($Global:bulkGroupForm.groupList.count -eq 0) -or
                        ($Global:bulkGroupForm.groupList.name -inotcontains $nameParts.name)
                    ) {
                        $Global:bulkGroupForm.groupList.add($nameParts)
                    }
                }
                }
            }
        }

        # Bubble-seort the groups so that these are true
        # - Computer groups are created before all other groups
        # - User groups are created next
        # - any other groups are created after those
        #bubble sort 'borrowed' from https://gallery.technet.microsoft.com/scriptcenter/Bubble-Sort-for-PowerShell-c25964c2
        [bool]$sorted = $false
        $counter = 0
        for ($pass = 1; ($pass -lt $Global:bulkGroupForm.groupList.Count) -and -not $sorted; $pass++)
        {
            # Assume the array is sorted
            $sorted = $true
            for ($index = 0; $index -lt ($Global:bulkGroupForm.groupList.Count - $pass); $index++)
            {
                $counter++
                $nextIndex = $index + 1
                if (
                    (
                        ($Global:bulkGroupForm.groupList[$nextIndex].type -eq 'Computer') -and
                        ($Global:bulkGroupForm.groupList[$index].type -ne 'Computer')
                    ) -or
                    (
                        ($Global:bulkGroupForm.groupList[$nextIndex].type -eq 'User') -and
                        ($Global:bulkGroupForm.groupList[$index].type -notmatch 'Computer|User')
                    ) -or (
                        ($Global:bulkGroupForm.groupList[$nextIndex].type -eq $Global:bulkGroupForm.groupList[$index].type) -and
                        ($Global:bulkGroupForm.groupList[$nextIndex].name.length -lt $Global:bulkGroupForm.groupList[$index].name.length)
                    )
                ) {
                    # Swap items
                    $temp = $Global:bulkGroupForm.groupList[$index]
                    $Global:bulkGroupForm.groupList[$index] = $Global:bulkGroupForm.groupList[$nextIndex]
                    $Global:bulkGroupForm.groupList[$nextIndex] = $temp
                    $sorted = $false
                }
            }
        }

        # Check that none of the entries are null
        $invalidEntry = $false
        $Global:bulkGroupForm.groupList | ForEach-Object {
            if ($_ -eq $null) {
                $invalidEntry = $true
            }
        }
        if ($invalidEntry -ne $false) {
            PopupBox "There was an error parsing one or more of the security groups from the migration worksheet. The index of the error was: $($invalidEntry)"
            return
        }

        # Finish up by redrawing the data displayed in the grid
        $global:bulkGroupForm.groupGrid.items.refresh()
    })

    $Global:bulkGroupForm.btnClear.add_Click({
        $Global:bulkGroupForm.groupList.clear()
        $global:bulkGroupForm.groupGrid.items.refresh()
    })

    $Global:bulkGroupForm.btnContinue.add_Click({
        if ($Global:bulkGroupForm.groupList.count -gt 0) {
            $global:bulkGroupForm.okButtonClicked = $true
            $global:bulkGroupForm.window.close()
        } else {
            PopupBox 'There are no groups in the import list. Please import some groups, or use the Cancel button to abort.'
        }
    })

    $Global:bulkGroupForm.btnCancel.add_Click({
        $global:bulkGroupForm.okButtonClicked = $false
        $global:bulkGroupForm.window.close()
    })

    $Global:bulkGroupForm.lnkOpenKB.add_click({
        (New-Object -Com Shell.Application).Open("https://asu.service-now.com/kb_view_customer.do?sysparm_article=KB0015918")
    })

    # Form visibility changed
    $Global:bulkGroupForm.window.add_IsVisibleChanged({
        if ($Global:bulkGroupForm.window.isVisible -eq $true) {
            $Global:bulkGroupForm.window.topmost = $true
            $Global:bulkGroupForm.window.topmost = $false
            $Global:bulkGroupForm.window.focus()
        }
    })

    $global:bulkGroupForm.window.showdialog()

    if ($global:bulkGroupForm.okButtonClicked -eq $true) {

        $script:ticketNumer = ''
        $script:departmentName = ''
        $script:updateGroup = ''
        $script:supportGroup = ''
        try {
            $script:ticketNumer = $Global:bulkGroupForm.txtTicketNumber.Text
            $script:departmentName = $Global:bulkGroupForm.txtDepartment.Text
            $script:updateGroup = $Global:bulkGroupForm.cmbUpdateGroup.SelectedValue.Content.ToString()
            $script:supportGroup = $Global:bulkGroupForm.cmbSupportGroup.SelectedValue.ToString()
        } catch {
            $script:ticketNumer = ''
            $script:departmentName = ''
        }

        $Global:bulkGroupForm.groupList | ForEach-Object {

            # Test if group already exists
            $groupExists = $false
            $qResult = 0
            Try {
                $qResult = Get-ADGroup -Identity $_.Name -Server $global:ADserver
                if ($qResult -is [Microsoft.ActiveDirectory.Management.ADGroup]) {
                    $groupExists = $true
                }
            } Catch {
                $groupExists = $false
            }

            $createResult = 0
            if (!$groupExists) {
                $createResult = CreateGroup -Name $_.name -Ticket "${script:ticketNumer}" -Department $script:departmentName -UpdateGroup $script:updateGroup -SupportGroup $script:supportGroup
            } else {
                $syncHash.print("[ WRN ] Skipping creation of $($_.Name) because it already exists")
                $createResult = 0
            }


            if ($createResult -gt 0) {
                $createConfirm = PopupBox 'The previous window was cancelled. Would you like to continue processing additional groups?' "Confirm cancel" "yn"
                if ($createConfirm -ne 'yes') {
                    $syncHash.print("[ WRN ] Bulk group creation aborted")
                    exit 99
                }
            }
        }
    }

    $syncHash.print("[ OK! ] Bulk import completed!")
}
