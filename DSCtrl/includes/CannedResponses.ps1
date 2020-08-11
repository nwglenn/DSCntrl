[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Email Template" Height="254.167" Width="663">
    <Grid HorizontalAlignment="Left" Height="218" VerticalAlignment="Top" Width="653">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Label Content="To:" HorizontalAlignment="Left" Margin="22,32,0,0" Grid.RowSpan="2" VerticalAlignment="Top" RenderTransformOrigin="1.48,0.654" Height="26" Width="25"/>
        <TextBox Name="toField" HorizontalAlignment="Left" Height="23" Margin="47,35,0,0" Grid.RowSpan="2" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
        <Label Content="From:" HorizontalAlignment="Left" Margin="7,68,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Height="26" Width="40"/>
        <TextBox Name="fromField" HorizontalAlignment="Left" Height="23" Margin="47,71,0,0" Grid.RowSpan="2" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
        <Button Name="CopyBtn" Content="Copy" HorizontalAlignment="Left" Margin="22,144,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
        <Button Name="CloseBtn" Content="Close" HorizontalAlignment="Left" Margin="287,182,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="74" Height="20"/>
        <ComboBox Name="Template" HorizontalAlignment="Left" Margin="47,106,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="120">
            <ComboBoxItem>Closing</ComboBoxItem>
            <ComboBoxItem>Appointment</ComboBoxItem>
            <ComboBoxItem>Testing</ComboBoxItem>
        </ComboBox>
        <Label Content="Type:" HorizontalAlignment="Left" Margin="10,102,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Button Name="PreviewBtn" Content="Preview" HorizontalAlignment="Left" Margin="110,144,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75" Height="20"/>
        <Button Name="SendBtn" Content="Send" HorizontalAlignment="Left" Margin="315,144,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <Label Content="Template" HorizontalAlignment="Left" Margin="66,0,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Label Content="Send Email" HorizontalAlignment="Left" Margin="270,0,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Separator HorizontalAlignment="Left" Height="28" Margin="109,76,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="181" RenderTransformOrigin="0.5,0.5">
            <Separator.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="90"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Separator.RenderTransform>
        </Separator>
        <Label Content="Address:" HorizontalAlignment="Left" Margin="210,32,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="Address" HorizontalAlignment="Left" Height="22" Margin="270,36,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="CC:" HorizontalAlignment="Left" Margin="238,68,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="ccAddress"  HorizontalAlignment="Left" Height="23" Margin="270,71,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="Subject:" HorizontalAlignment="Left" Margin="213,102,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="Subject" HorizontalAlignment="Left" Height="23" Margin="270,105,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Label Content="GAL Search" HorizontalAlignment="Left" Margin="497,0,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Label Content="First Name:" HorizontalAlignment="Left" Margin="434,32,0,0" Grid.RowSpan="2" VerticalAlignment="Top" RenderTransformOrigin="0.29,-0.269"/>
        <TextBox Name="FirstNameGAL" HorizontalAlignment="Left" Height="23" Margin="504,35,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Name="SearchBtn" Content="Search" HorizontalAlignment="Left" Margin="556,182,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <Separator HorizontalAlignment="Left" Height="28" Margin="324,74,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="181" RenderTransformOrigin="0.5,0.5">
            <Separator.RenderTransform>
                <TransformGroup>
                    <ScaleTransform/>
                    <SkewTransform/>
                    <RotateTransform Angle="90"/>
                    <TranslateTransform/>
                </TransformGroup>
            </Separator.RenderTransform>
        </Separator>
        <ListBox Name="GALResults" HorizontalAlignment="Left" Height="100" Margin="434,71,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="209"/>
        <Button Name="SelectBtn" Content="Select" HorizontalAlignment="Left" Margin="452,182,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
    </Grid>
</Window>
'@

#I added this comment.

#Read XAML
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Some possible causes for this problem include: .NET Framework is missing PowerShell must be launched with PowerShell -sta, invalid XAML code was encountered."; exit}

#===========================================================================
# Store Form Objects In PowerShell
#===========================================================================

$xaml.SelectNodes("//*[@Name]") | % {Set-Variable -Name ($_.Name) -Value $Form.FindName($_.Name)}

# function for building the email from the body and signature templates, the SelectedTemplate dropdown must match the template file name (other than the .txt)
# Returns the built message
function Build-EmailMessage($toPerson, $fromPerson) {
    $salutation = "Hi {0}," -f $toPerson
    $body = Get-Content "$settingsPath\$script:selectedTemplate.txt"
    $sigName = $fromPerson
    $signature = Get-Content "$settingsPath\Signature.txt"
    $message = ($salutation, "", ($body -join "`n"), "", "Thanks,", "", $sigName, ($signature -join "`n")) -join "`n"
    return $message
}

# Finding the include path and then back tracking up to the folder above it, then down into the settings folder has worked so far
# POTENTIAL REVISION
if ($includePath -eq $null) {
    $includePath = split-path $SCRIPT:MyInvocation.MyCommand.Path -parent
}
$settingsPath = ("{0}\settings" -f (Split-path ${includePath} -parent))

# Event for when the drop down closes to update what is currently selected
$Template.add_DropDownCLosed({
    $script:selectedTemplate = $Template.text
})

$GALResults.add_SelectionChanged({
    $script:SelectedUser = $GALResults.SelectedItem
})

# When this button is clicked it builds the email from the fields filled in and copies it to the clipboard directly
$CopyBtn.add_click({
    $fromPerson = $fromField.text
    $toPerson = $toField.text
    Set-Clipboard -value (Build-EmailMessage -toPerson $toPerson -fromPerson $fromPerson)
})

# This will open a new window with the email message constructed and allows for editing before re-copy
$PreviewBtn.add_click({
    $fromPerson = $fromField.text
    $toPerson = $toField.text
    $script:message = Build-EmailMessage -toPerson $toPerson -fromPerson $fromPerson

    . "${includePath}\CannedResponsesPreviewWindow.ps1"
})

# This will send the template as an email directly with the information filled in the "Send Email" section as well as the Template information
$SendBtn.add_click({
    $fromPerson = $fromField.text
    $toPerson = $toField.text
    $toAddress = $Address.text
    $cc = $ccAddress.text
    $script:message = Build-EmailMessage -toPerson $toPerson -fromPerson $fromPerson

    Add-Type -assembly "Microsoft.Office.Interop.Outlook"
    $o = New-Object -ComObject Outlook.Application
    $mail = $o.CreateItem(0)
    $mail.subject = $subject.text
    $mail.body = $message
    $mail.to = $toAddress
    $mail.CC = $cc
    $mail.send()
})

$SearchBtn.add_click({
    $GALResults.items.Clear()
    $fName = $FirstNameGAL.text
    $result = Get-ADUser -filter "GivenName -eq '$fName'"  -Properties "ExtensionAttribute8" | select GivenName, Surname, ExtensionAttribute8
    foreach($value in $result) {
        $GALResults.items.add(("{0} {1} | {2}" -f $value.givenname, $value.surname, $value.ExtensionAttribute8))
    }
})

$SelectBtn.add_click({
    $fName = ($script:SelectedUser -split " ")[0]
    $lName = ($script:SelectedUser -split " ")[1]
    $email = ($script:SelectedUser -split " ")[3]
    $toField.text = $fName
    $address.text = $email
})


$CloseBtn.add_click({
    $Form.close()
})

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null

