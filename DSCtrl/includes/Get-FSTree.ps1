#REQUIRES -Version 5.0

class FSTreeNode {
    static [Int]$maxDepth = 8
    static [Int]$ageOver = 0
    static [Int]$sizeOver = 0
    [String]$name
    [String]$fullName
    [datetime]$created
    [datetime]$modified
    [datetime]$accessed
    [Boolean]$isContainer
    [Boolean]$isLink

    hidden [Int64]$byteLength
    hidden [System.Collections.ArrayList]$children
    hidden [System.Security.AccessControl.NativeObjectSecurity]$acl
    hidden [Boolean]$accessError = $false
    hidden [int]$depth

    FSTreeNode ([System.IO.FileSystemInfo]$dirInfo) {
        $this.Construct($dirInfo, 0)
    }

    FSTreeNode ([System.IO.FileSystemInfo]$dirInfo, [Int]$instanceDepth) {
        $this.Construct($dirInfo, $instanceDepth)
    }

    # PS 5.0 does not support constructor chaining, so we need to use a helper
    hidden [void]Construct ([System.IO.FileSystemInfo]$dirInfo, [Int]$depth) {
        $this.name = $dirInfo.Name
        $this.fullName = $dirInfo.FullName
        $this.created = $dirInfo.CreationTime
        $this.modified = $dirInfo.LastWriteTime
        $this.accessed = $dirInfo.LastAccessTime

        if ($dirInfo -is [System.IO.DirectoryInfo]) {
            $this.isContainer = $true
        }
        else {
            $this.isContainer = $false
        }

        # Determine if the object is a link
        if ($dirInfo.mode -imatch 'l$') {
            $this.isLink = $true
        }
        else {
            $this.isLink = $false
        }

        if ( ($dirInfo.length) -and ($dirInfo.length -gt 0) ) {
            $this.byteLength = $dirInfo.length
        }
        else {
            $this.byteLength = 0
        }

        # Only pull children if maxDepth has not been reached
        # and this node is a container
        $this.depth = $depth
        if (
            ($this.depth -lt [FSTreeNode]::maxDepth) -and
            ($this.isContainer -eq $true) -and 
            ($this.isLink -eq $false)
        ) {
            $this.accessError = $this.GetChildren()
        }
        else {
            $this.children = $null
        }

        # Pull ACL's for containers so we can display that information later
        # We avoid pulling file ACL's because that could be a lot of system
        # calls, and file-level access control is less common than container level
        if ($this.isContainer -eq $true) {
            try {
                $this.acl = Get-ACL -Path $this.fullName
            }
            catch {
                $this.acl = $null
            }
        }
        else {
            $this.acl = $null
        }
    }

    hidden [Boolean]GetChildren() {
        $success = $false
        $childObjects = $null
        try {
            $childObjects = Get-ChildItem -Path $this.fullName
        }
        catch {
            $childObjects = $null
        }

        if ( ($childObjects) -and ($childObjects.count -gt 0)) {
            $this.children = New-Object System.Collections.ArrayList
            $childObjects | ForEach-Object {
                $newNode = [FSTreeNode]::new($_, $this.depth + 1)
                $this.children.add($newNode)
            }
            $success = $true
        }
        elseif ($childObjects) {

            # We were able to query the resource, but no children were found
            $success = $true
        }
        else {
            $success = $false
        }
        return !$success
    }

    [Int64]GetByteLength() {
        [Int64]$totalLength = 0
        $totalLength += $this.byteLength
        if (
            ($null -ne $this.children) -and 
            ($this.children.count -gt 0) -and
            ($this.accessError -ne $true)
        ) {
            $this.children | Foreach-Object {
                $totalLength += $_.GetByteLength()
            }
        }
        return $totalLength
    }

    [String]ToString() {
        $outString = (' ' * $this.Depth * 2) + $this.name + "`n"
        if (
            ($null -ne $this.children) -and 
            ($this.children.count -gt 0) -and
            ($this.accessError -ne $true)
        ) {
            $outString += ($this.children | ForEach-Object { $_.ToString() })
        }
        return $outString
    }

    [String]ToCSV() {
        return $this.ToCSV(0)
    }

    [String]ToCSV([int]$verbosity) {
        $CSVdata = ''

        $folderString = ''

        # Build a giant list of all nodes in the object
        if ($this.depth -eq 0) {
            $CSVdata = """Path"",""FileName"",""Created"",""Modified"",""Accessed"",""Size"",""Permissions"",""Messages"""
        }

        # Get size in bytes (this is expensive, because recursion)
        #TODO: Be smared about how we calculate this
        $mySize = $this.GetByteLength()
        if ($mySize -ge 1000000) {
            $mySize = [Math]::Round($mySize / 1000000)
            $mySize = "${mySize} MB"
        }
        elseif ($mySize -ge 1000) {
            $mySize = [Math]::Round($mySize / 1000)
            $mySize = "${mySize} KB"
        }
        else {
            $mySize = "${mySize}  B"
        }
        
        # Generate data based on node type
        if ($this.isContainer) {

            # Add permissions information to container header
            if ($null -ne $this.acl) {

                # Build string for errors and warnings
                $errorStr = ''
                $accessStr = ''

                # Count inherited ACL's to determine if inheritance is in effect
                $inheritanceCount = $this.acl.access | Where-Object -Property 'IsInherited' -eq $true | Measure-Object
                if ($inheritanceCount -eq 0) {
                    $errorStr += "Warning: security inheritance for this object is disabled`n"
                }
                
                # Warn the user if the contents of this folder are deeper than max depth
                if ($this.depth -eq [FSTreeNode]::maxDepth) {
                    $errorStr += "Warning: The contents of this container were not scanned, because it is beyond the current maximum depth setting`n"
                }
                try {
                    $accessStr = $this.acl.access | Where-Object -Property 'IsInherited' -eq $false | ForEach-Object {
                        "$($_.IdentityReference) $($_.AccessControlType): $($_.FileSystemRights)`n"
                        if ($($_.IdentityReference) -eq 'CREATOR OWNER') {
                            $errorStr += "Warning: Account CREATOR OWNER has explicit permission on folder.`n"
                        } elseif ($_.IdentityReference.AccountDomainSid) { 
                            $errorStr += "Warning: User account ($($_.IdentityReference.value)) has explicit permission on folder.`n"
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
                                }
                            }
                        }
                    }
                } catch {
                    $errorStr += "Error: Unknown problem account permission.`n"
                }

                if ($accessStr.length -eq 0) {
                    $accessStr = 'Inherited'
                }
                
            } else {
                $errorStr += "ERROR: Could not read security information`n"
                $accessStr = 'Error'
            }

            if ($this.accessError) {
                $errorStr += "Warning: Empty folder`n"
                $folderString += "`n""$($this.fullName)"","""",""$($this.created)"",""$($this.modified)"",""$($this.accessed)"",""$($mySize.toString())"",""$accessStr"",""$errorStr"""
                $accessStr = 'empty'
            }
            elseif ( 
                ($this.children -is [System.Collections.ArrayList]) -and
                ($this.children.count -gt 0)
            ) {
                $childString = ($this.children | Foreach-Object { $_.ToCSV($verbosity) })
                if ($childString -ne '') {
                    $folderString += "`n""$($this.fullName)"","""",""$($this.created)"",""$($this.modified)"",""$($this.accessed)"",""$($mySize.toString())"",""$accessStr"",""$errorStr""$childString"
                    $accessStr = 'Display'
                }
            } else {
                $folderString += "`n""$($this.fullName)"","""",""$($this.created)"",""$($this.modified)"",""$($this.accessed)"",""$($mySize.toString())"",""$accessStr"",""$errorStr"""
            }

            if (
               ($accessStr -ne 'Inherited') -or
               ($this.depth -eq 0)
             ) {
                $CSVdata += $folderString
            }
        }
        else { 

            # If this is not a container, then it is likely a file
            # Only add file information if verbosity is higher than 0
            if (
                ($verbosity -gt 0) -and
                ($this.modified -lt (Get-Date).AddYears(-1 * [FSTreeNode]::ageOver)) -and
                ($this.accessed -lt (Get-Date).AddYears(-1 * [FSTreeNode]::ageOver)) -and
                ((1000000 * [FSTreeNode]::sizeOver) -lt $this.GetByteLength())
            ) {
                $CSVdata += "`n""$($this.fullName)"",""$($this.name)"",""$($this.created)"",""$($this.modified)"",""$($this.accessed)"",""$($mySize.toString())"""
            }
        }

        return $CSVdata
    }
    <#
    .SYNOPSIS
    Generates an HTML report using the lowest verbosity level
    .NOTES
    To specify a higher verbosity level, call the overloaded function with an integer
    #>
    [string]ToHTML() {
        return $this.ToHTML(0)
    }

    <#
    .SYNOPSIS
    Returns a representation of the data in an HTML report
    .DESCRIPTION
    Generates an HTML report based on two possible verbosity settings
    0 = Only display containers and related information
    1 = Display containers and files
    #>
    [String]ToHTML([int]$verbosity) {
        $returnString = ''

        # Add headers if root node
        if ($this.depth -eq 0) {
            $returnString += @"
<!DOCTYPE html>
<html>
<head>
<style>
* {
    margin: 0px;
    padding: 0px;
  }
  
  body {
    margin: 50px;
    font-family: Arial;
  }
    ul {
    margin: 0px 0px 0px 20px;
    list-style: none;
    line-height: 2em;
    font-family: Arial;
  }
  ul li {
    font-size: 16px;
    position: relative;
  }
  ul li:before {
    position: absolute;
    left: -15px;
    top: 0px;
    content: '';
    display: block;
    border-left: 1px solid #ddd;
    height: 1em;
    border-bottom: 1px solid #ddd;
    width: 10px;
  }
  ul li:after {
    position: absolute;
    left: -15px;
    bottom: -7px;
    content: '';
    display: block;
    border-left: 1px solid #ddd;
    height: 100%;
  }
  ul li.root {
    margin: 0px 0px 0px -20px;
  }
  ul li.root:before {
    display: none;
  }
  ul li.root:after {
    display: none;
  }
  ul li:last-child:after {
    display: none;
  }
  div.security {
    background-color: #d6d6c2;
    margin-left: 5px;
    padding: 10px;
  }
  div.warning {
    background-color: #b3ecff;
    margin-left: 5px;
    padding: 10px;
  }
  div.error {
    background-color: #ffcccc;
    margin-left: 5px;
    padding: 10px;
  }
</style>
</head>
<body>
<h1>Filesystem Report ($(Get-Date))</h1>
<p>$($this.fullName)</p>
<br />
<p><b>Search Settings:</b></p>
<p> 
"@

if ($verbosity -gt 0) {
    $returnString += "&nbsp;&nbsp;Folders & Files<br />&nbsp;&nbsp;Files older than $([FSTreeNode]::ageOver) yrs.<br />&nbsp;&nbsp;Files larger than $([FSTreeNode]::sizeOver) MB."
} else {
    $returnString += "&nbsp;&nbsp;Folders only"
}

$returnString += "</p><br /><p><b>Legend:</b><br />&nbsp;&nbsp;<b>c</b> = Created<br />&nbsp;&nbsp;<b>a</b> = Accessed<br />&nbsp;&nbsp;<b>m</b> = Modified<br />&nbsp;&nbsp;<b>sz</b> = Size</p><br /><ul>"

        }

        # Get size in bytes (this is expensive, because recursion)
        #TODO: Be smared about how we calculate this
        $mySize = $this.GetByteLength()
        if ($mySize -ge 1000000) {
            $mySize = [Math]::Round($mySize / 1000000)
            $mySize = "${mySize} MB"
        }
        elseif ($mySize -ge 1000) {
            $mySize = [Math]::Round($mySize / 1000)
            $mySize = "${mySize} KB"
        }
        else {
            $mySize = "${mySize}  B"
        }
        
        # Generate HTML based on node type
        if ($this.isContainer) {

            # Add permissions information to container header
            if ($null -ne $this.acl) {

                # Build string for errors and warnings
                $errorHTML = ''
                $accessHTML = ''

                # Count inherited ACL's to determine if inheritance is in effect
                $inheritanceCount = $this.acl.access | Where-Object -Property 'IsInherited' -eq $true | Measure-Object
                if ($inheritanceCount -eq 0) {
                    $errorHTML += "<div class='warning'><b><u>Warning</u></b>: security inheritance for this object is disabled</div>"
                }

                # Warn the user if the contents of this folder are deeper than max depth
                if ($this.depth -eq [FSTreeNode]::maxDepth) {
                    $errorHTML += "<div class='warning'><b><u>Warning</u></b>: The contents of this container were not scanned, because it is beyond the current maximum depth setting</div>"
                }

                try {
                    $accessHTML = $this.acl.access | Where-Object -Property 'IsInherited' -eq $false | ForEach-Object {
                        "<br />$($_.IdentityReference) $($_.AccessControlType): $($_.FileSystemRights)"
                        if ($($_.IdentityReference) -eq 'CREATOR OWNER') {
                            $errorHTML += "<div class='warning'><b><u>Warning</u></b>: Account CREATOR OWNER has explicit permission on folder.</div>"
                        } elseif ($_.IdentityReference.AccountDomainSid) { 
                            $errorHTML += "<div class='warning'><b><u>Warning</u></b>: User account ($($_.IdentityReference.value)) has explicit permission on folder.</div>"
                        } else {
                            $userAccount = ($_.IdentityReference.value).Split('\')
                            if ( (-NOT $userAccount[1].Contains('.')) -and ($userAccount[1].length -le 8)) {
                                if ($userAccount[0] -eq 'ASURITE') {
                                    $domainServer = "asurite.ad.asu.edu"
                                } else {
                                    $domainServer = "ad.asu.edu"
                                }
                                if ($(get-ADObject -Server $domainServer -Filter "Name -eq '$($userAccount[1])'").ObjectClass -eq 'user') {
                                    $errorHTML += "<div class='warning'><b><u>Warning</u></b>: User account ($($userAccount[1])) has explicit permission on folder.</div>"
                                }
                            }
                        }
                    }
                } catch {
                    $errorHTML += "<div class='Error'><b><u>ERROR</u></b>: Unknown problem account permission.</div>"
                }

                if ($accessHTML.length -eq 0) {
                    $accessHTML = 'Inherited'
                }
                $aclHTML += @"
${errorHTML}
<div class='security'>
Permissions applied to this container:
${accessHTML}
</div>                
"@
            }
            else {
                $aclHTML += "<div class='error'><b><u>ERROR</u></b>: Could not read security information</div>"
                $accessHTML = 'Error'
            }

            if ($this.depth -eq 0) {
                $folderString += '<li class="root">'
                $folderString += "&#128194;<b>$($this.name)</b> <font color='#808080'>[c:$($this.created)] [m:$($this.modified)] [a:$($this.accessed)] [sz:$($mySize.toString())]</font>"
                $folderString += "</li><li>"
            }
            else {
                $folderString += "<li>"
                $folderString += "&#128194;<b>$($this.name)</b> <font color='#808080'>[c:$($this.created)] [m: $($this.modified)] [a:$($this.accessed) [sz:$($mySize.toString())]</font>"
            }

            $folderString += $aclHTML

            if ($this.accessError) {
                $folderString += "<div class='warning'><b><u>Warning</u></b>: Empty folder</div>"
                $accessHTML = 'empty'
            }
            elseif ( 
                ($this.children -is [System.Collections.ArrayList]) -and
                ($this.children.count -gt 0)
            ) {
                $childString = ($this.children | Foreach-Object { $_.ToHTML($verbosity) })
                if ($childString -ne '') {
                    $folderString += "<ul> $childString </ul>"
                    $accessHTML = 'Display'
                }
            } 
            
            if ($this.depth -ne 0) {
                $folderString += "</li>"
            }

            if (
               ($accessHTML -ne 'Inherited') -or
               ($this.depth -eq 0)
             ) {
                $returnString += $folderString
            }
        }
        else { 

            # If this is not a container, then it is likely a file
            # Only add file information if verbosity is higher than 0
            if (
                ($verbosity -gt 0) -and
                ($this.modified -lt (Get-Date).AddYears(-1 * [FSTreeNode]::ageOver)) -and
                ($this.accessed -lt (Get-Date).AddYears(-1 * [FSTreeNode]::ageOver)) -and
                ((1000000 * [FSTreeNode]::sizeOver) -lt $this.GetByteLength())
            ) {
                $returnString += "<li>$($this.name) <font color='#808080'>[c:$($this.created)] [m:$($this.modified)] [a:$($this.accessed)] [sz:${mySize} ]</font></li>"
            }
        }
        # Add footers if root node
        if ($this.depth -eq 0) {
            $returnString += "</ul></body></html>"
        }
        return $returnString
    }
}


function Get-FSTree {
    [CmdletBinding()]
    param (
        [Parameter(Mandatory = $False,
            ValueFromPipeline = $True,
            ValueFromPipelineByPropertyName = $True,
            HelpMessage = 'String to display')]
        [Alias('String')]
        [String]$HTML
    )

    # Form defenition
    $inputXML = @"
<Window
xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
Title="Share Report" Height="600" MinHeight="480" Width="1000" MinWidth="1000">
    <DockPanel Margin="16">
        <StackPanel Orientation="Horizontal" Margin="4" HorizontalAlignment="Right" DockPanel.Dock="Bottom" Height="45">
            <Button Name="btnSaveHTML" Width="100" Margin="8" IsEnabled="False">Save HTML</Button>
            <Button Name="btnSaveCSV" Width="100" Margin="8" IsEnabled="False">Save CSV</Button>
        </StackPanel>
        <StackPanel Orientation="Vertical" Margin="4" HorizontalAlignment="Left" DockPanel.Dock="Top">
            <StackPanel Orientation="Horizontal">
                <Label>Share path: </Label>
                <TextBox Name="txtPath" Width="256" Margin="8,0,8,0"></TextBox>
                <Label>Maximum search depth: </Label>
                <TextBox Name="txtMaxDepth" TextAlignment="Right" Width="64" Margin="8,0,0,0">999</TextBox>
                <Label> (1-999)</Label>
                <Button Name="btnStart" Width="80">Calculate!</Button>
                <ProgressBar Name="progress" IsIndeterminate="False" Margin="15,0,8,0" Width="200" Height="20" />
            </StackPanel>
            <GroupBox Header="Report on">
                <StackPanel Orientation="Horizontal">
                    <RadioButton Name="radFolders" Margin="8" Content="Folders only" HorizontalAlignment="Left" VerticalAlignment="Top" GroupName="verbosity" IsChecked="True"/>
                    <RadioButton Name="radFoldersAndFiles" Margin="8" Content="Folders and files" HorizontalAlignment="Left" VerticalAlignment="Top" GroupName="verbosity"/>
                </StackPanel>
            </GroupBox>
            <GroupBox Header="File settings">
                <StackPanel Orientation="Horizontal">
                    <Label>Not modified and accessed over (yrs.): </Label>
                    <TextBox Name="txtAgeOver" TextAlignment="Right" Width="45" IsEnabled="{Binding ElementName=radFoldersAndFiles, Path=IsChecked}">0</TextBox>
                    <Label>Larger than (MB): </Label>
                    <TextBox Name="txtsizeOver" TextAlignment="Right" Width="45" IsEnabled="{Binding ElementName=radFoldersAndFiles, Path=IsChecked}">0</TextBox>
                </StackPanel>
            </GroupBox>
        </StackPanel>
        <GroupBox Header="Report" VerticalAlignment="Stretch" HorizontalAlignment="Stretch" DockPanel.Dock="Bottom">
            <WebBrowser Name="reportView" HorizontalAlignment="Left" VerticalAlignment="Top" />
        </GroupBox>
    </DockPanel>
</Window>
"@

    # Load XML; prep for deserialization
    $inputXML = $inputXML -replace 'mc:Ignorable="d"', '' -replace "x:N", 'N' -replace '^<Win.*', '<Window'
    [void][System.Reflection.Assembly]::LoadWithPartialName('Microsoft.VisualBasic')
    Add-Type -AssemblyName PresentationCore, PresentationFramework, WindowsBase, system.windows.forms
    [xml]$XAML = $inputXML

    # Deserialize the form into an object
    $form = $null
    try {
        $Form = [Windows.Markup.XamlReader]::Load((new-object System.Xml.XmlNodeReader $xaml))
    }
    catch [System.Management.Automation.MethodInvocationException] {
        Write-Warning "We ran into a problem with the XAML code.  Check the syntax for this control..."
        write-host $error[0].Exception.Message -ForegroundColor Red
        if ($error[0].Exception.Message -like "*button*") {
            write-warning "Ensure your &lt;button in the `$inputXML does NOT have a Click=ButtonClick property.  PS can't handle this`n`n`n`n"
        }
    }
    catch {
        #if it broke some other way :D
        Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."
    }

    # Load form components into a hash table
    # This must be stored in global scope so event calls can reach it
    $Global:shareForm = [hashtable]::Synchronized(@{})
    $Global:shareForm.Window = $Form
    $xaml.SelectNodes("//*[@Name]") | ForEach-Object {
        $Global:shareForm.add($_.Name, $Form.FindName($_.Name))
    }
    $Global:shareForm.add('fsPath', '')
    $Global:shareForm.add('maxDepth', 64)
    $Global:shareForm.add('verbosity', 0)
    $Global:shareForm.add('FileTree', $null)
    $Global:shareForm.add('ageOver', 0)
    $Global:shareForm.add('sizeOver', 0)

    # Create runspace for long-running tasks
    $Runspace = [runspacefactory]::CreateRunspace()
    $Runspace.ApartmentState = "STA"
    $Runspace.ThreadOptions = "ReuseThread"
    $Runspace.Open()
    $Runspace.SessionStateProxy.SetVariable("syncHash", $Global:shareForm)
    $Runspace.SessionStateProxy.SetVariable("includePath", $includePath)

    # Form visibility changed
    $Global:shareForm.window.add_IsVisibleChanged( {
            if ($Global:shareForm.window.isVisible -eq $true) {
                $Global:shareForm.window.topmost = $true
                $Global:shareForm.window.topmost = $false
                $Global:shareForm.window.focus()
            }
        })

    $Global:shareForm.btnStart.Add_Click( {
        
            # Check quality of provided path
            $Global:shareForm.fsPath = $Global:shareForm.txtPath.Text
            if ($Global:shareForm.fsPath.length -lt 2) {
                PopupBox "The provided path does not appear to be valid, check that a path was entered"
                $Global:shareForm.fsPath = $null
            }

            # Check that provided depth is an integer
            $providedDepth = $Global:shareForm.txtmaxDepth.Text
            if ($providedDepth -imatch '^\d{1,3}$') {
                $Global:shareForm.maxDepth = [int]($providedDepth)
            }
            else {
                PopupBox "The provided maximum depth must be between 1 and 999"
            }

            # Check that provided size is an integer
            $providedSize = $Global:shareForm.txtsizeOver.Text
            if ($providedSize -match '^[0-9]{1,7}$') {
                $Global:shareForm.sizeOver = ($providedSize)
            }
            else {
                PopupBox "The provided size must be between 0 and 9999999"
            }

            # Check that provided age is an integer
            $providedAge = $Global:shareForm.txtAgeOver.Text
            if ($providedAge -match '^[0-9]{1,2}$') {
                $Global:shareForm.ageOver = ($providedAge)
            }
            else {
                PopupBox "The provided age must be between 0 and 99"
            }

            # Configure verbosity settings
            if ($Global:shareForm.radFoldersAndFiles.isChecked) {
                $Global:shareForm.verbosity = 1
            }
            else {
                $Global:shareForm.verbosity = 0
            }

            if ( ($Global:shareForm.fsPath) -and ($Global:shareForm.maxDepth) -and ($Global:shareForm.ageOver) -and ($Global:shareForm.sizeOver)) {
                $code = {
                    . "${includePath}\global.ps1"
                    . "${includePath}\Get-FSTree.ps1"
                    Start-Transcript -Path (Join-Path $Global:logdir "shareReport-worker-$((get-date).ToFileTime())") -IncludeInvocationHeader
                    $syncHash.Window.Dispatcher.invoke(
                        [action] {
                            $syncHash.reportView.NavigateToString(@"
<HTML>
<head>
<style>
</style>
</head>
<body>
<h3>Collecting data for the report, please wait...</h3>
<p>As soon as everything is ready, you will be able to save the report.</p>
<p>This could take quite awhile...</p>
</body>
</HTML>
"@)
                            $syncHash.btnSaveHTML.IsEnabled = $False
                            $syncHash.btnSaveCSV.IsEnabled = $False
                            $syncHash.btnStart.IsEnabled = $False
                            $syncHash.progress.IsIndeterminate = $True
                        }
                    )

                
                    $fsRoot = $null
                    try {
                        $fsRoot = Get-Item -Path $syncHash.fsPath
                    }
                    catch {
                        $syncHash.Window.Dispatcher.invoke(
                            [action] {$syncHash.reportView.NavigateToString(@"
                        <HTML><body>
                        <h1>ERROR! - Could not access specified path</h1>
                        <p>The first step in generating your report requires us to access the container at the location you specified. When we tried to do that, we encountered an error. Any specific error information will be displayed below. Check that the path you specified was valid, and try again.</p>
                        <p>$($_.message)</p>
                        <p>$($_.toString())</p>
                        </body></HTML>
"@)}
                        )
                        $syncHash.btnStart.IsEnabled = $True
                        $syncHash.progress.IsIndeterminate = $False
                        return
                    }

                    try {
                        [FSTreeNode]::maxDepth = $syncHash.maxDepth
                        #$syncHash.ageOver = $Global:shareForm.ageOver
                        #[FSTreeNode]::ageOver = $Global:shareForm.ageOver
                        [FSTreeNode]::ageOver = $syncHash.ageOver
                        [FSTreeNode]::sizeOver = $syncHash.sizeOver
                        $syncHash.fileTree = [FSTreeNode]::new($fsRoot)
                    }
                    catch {
                        $syncHash.Window.Dispatcher.invoke(
                            [action] {$syncHash.reportView.NavigateToString(@"
                        <HTML><body>
                        <h1>ERROR! - Something went wrong while collecting data</h1>
                        <p>This reporting process queries the child nodes of the location you specified and their security settings. If something goes wrong, such as network instability, or we hit a resource we don't know how to deal with, things can go sideways quickly. Any information about the specific error will be displayed below. Examining the application logs can also give you an idea of what happened. From the DSCtrl main window, select Help > View Application Logs and look for files starting with 'sharereport-worker'</p>
                        <p>$($_.message)</p>
                        <p>$($_.toString())</p>
                        </body></HTML>
"@)}
                        )
                        $syncHash.btnStart.IsEnabled = $True
                        $syncHash.progress.IsIndeterminate = $False
                        return
                    }

                    $script:reportHTML = $syncHash.fileTree.ToHTML($syncHash.verbosity)
                    $syncHash.Window.Dispatcher.invoke(
                        [action] {
                            $syncHash.reportView.NavigateToString(@"
                        <HTML>
                          <head>
                          </head>
                          <body>
                            <h1>Processed folder structure</h1>
                            <p>Please save a copy of the report by click to view.</p>
                          </body>
                        </HTML>
"@)
                            #$syncHash.reportView.NavigateToString($script:reportHTML)
                            $syncHash.btnSaveHTML.IsEnabled = $True
                            $syncHash.btnSaveCSV.IsEnabled = $True
                            $syncHash.btnStart.IsEnabled = $True
                            $syncHash.progress.IsIndeterminate = $False
                        })
                    Stop-Transcript
                }
                $PSinstance = [powershell]::Create().AddScript($Code)
                $PSinstance.Runspace = $Runspace
                $PSinstance.BeginInvoke()
            }
        })

    $Global:shareForm.btnSaveHTML.Add_Click( {
            $Global:shareForm.btnSaveCSV.IsEnabled = $False
            $Global:shareForm.btnSaveHTML.IsEnabled = $False
            $Global:shareForm.btnSaveCSV.IsEnabled = $False
            $Global:shareForm.btnStart.IsEnabled = $False
            $Global:shareForm.progress.IsIndeterminate = $True

            $saveDialog = New-Object windows.forms.savefiledialog
            $saveDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
            $saveDialog.title = "Save Share Report"
            $saveDialog.filter = "HTML|*.html|All Files|*.*"
            $result = $saveDialog.ShowDialog()
            if ($result -ne "OK") {
                # canceled so stop
                return
            }
        
            $global:shareForm.fileTree.toHTML($global:shareForm.verbosity) | Set-Content -Path $saveDialog.filename
        
            # open html
            Invoke-Item $saveDialog.filename

            $Global:shareForm.btnSaveHTML.IsEnabled = $True
            $Global:shareForm.btnSaveCSV.IsEnabled = $True
            $Global:shareForm.btnStart.IsEnabled = $True
            $Global:shareForm.progress.IsIndeterminate = $False
        })

    $Global:shareForm.btnSaveCSV.Add_Click( {
            $Global:shareForm.btnSaveHTML.IsEnabled = $False
            $Global:shareForm.btnSaveCSV.IsEnabled = $False
            $Global:shareForm.btnStart.IsEnabled = $False
            $Global:shareForm.progress.IsIndeterminate = $True

            $saveDialog = New-Object windows.forms.savefiledialog
            $saveDialog.initialDirectory = [System.IO.Directory]::GetCurrentDirectory()
            $saveDialog.title = "Save Share Report"
            $saveDialog.filter = "CSV|*.csv|All Files|*.*"
            $result = $saveDialog.ShowDialog()
            if ($result -ne "OK") {
                # canceled so stop
                return
            }
        
            $global:shareForm.fileTree.toCSV($global:shareForm.verbosity) | Set-Content -Path $saveDialog.filename
            # open csv
            Invoke-Item $saveDialog.filename

            $Global:shareForm.btnSaveHTML.IsEnabled = $True
            $Global:shareForm.btnSaveCSV.IsEnabled = $True
            $Global:shareForm.btnStart.IsEnabled = $True
            $Global:shareForm.progress.IsIndeterminate = $False
        })

    $global:shareForm.reportView.NavigateToString(@"
<HTML>
  <body>
    <h2>Report generator is ready.</h2>
    <p>To get started, enter the path to a network share, any settings and click the 'Calculate!' button.</p>
    <p>Just keep in mind that this might take several minutes, depending on the maximum depth value, and the number of objects in the share.</p>
  </body>
</HTML>
"@)
    $global:shareForm.Window.showDialog() | Out-Null
}
