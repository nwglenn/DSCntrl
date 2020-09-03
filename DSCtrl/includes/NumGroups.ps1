[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = @'
<Window
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:Testing_Forms_Again_For_Real"
        Title="Number of Groups" Height="222.167" Width="235.667">
    <Grid HorizontalAlignment="Left" Height="224" VerticalAlignment="Top" Width="412" Margin="0,0,0,-1">
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
            <ColumnDefinition Width="0*"/>
        </Grid.ColumnDefinitions>
        <Grid.RowDefinitions>
            <RowDefinition Height="0*"/>
            <RowDefinition/>
        </Grid.RowDefinitions>
        <Label Content="ASURITE:" HorizontalAlignment="Left" Margin="22,19,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="ASURITE" HorizontalAlignment="Left" Height="23" Margin="82,23,0,0" Grid.RowSpan="2" TextWrapping="Wrap" VerticalAlignment="Top" Width="120"/>
        <Button Name="GetBtn" Content="Get" HorizontalAlignment="Left" Margin="22,62,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <Label Content="Answer:" HorizontalAlignment="Left" Margin="102,59,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Label Name="AnswerLabel" Content="0" HorizontalAlignment="Left" Margin="158,59,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <Label Content="Group:" HorizontalAlignment="Left" Margin="31,105,0,0" Grid.RowSpan="2" VerticalAlignment="Top"/>
        <TextBox Name="GroupName" HorizontalAlignment="Left" Height="23" Margin="82,108,0,0" Grid.RowSpan="2" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="120"/>
        <Button Name="WhatIfBtn" Content="What-If" HorizontalAlignment="Left" Margin="22,149,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
        <Button Name="CloseBtn" Content="Close" HorizontalAlignment="Left" Margin="121,149,0,0" Grid.RowSpan="2" VerticalAlignment="Top" Width="75"/>
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

# Button for getting the number of groups the user is in
$GetBtn.add_click({
    $username = Get-ADUser $ASURITE.Text

    $token = (Get-ADUser $username -Properties tokengroups).tokengroups

    $NumGroups = $token.count

    $AnswerLabel.Content = $NumGroups
    $Global:SyncHash.print(("{0} is in {1} groups." -f $username.name, $NumGroups), $false)
})

# Button to calculate the number of groups WILL be in if they are added to the group specified. This is to make sure they wouldn't go over the max number of groups prior to adding them
$WhatIfBtn.add_click({
    $Global:SyncHash.print(("Calculating number of groups after adding to {0}, this may take some time..." -f $GroupName.Text), $false)

    $username = Get-ADUser $ASURITE.Text
    $token = (Get-ADUser $username -Properties tokengroups).tokengroups
    $NumGroupsBefore = $token.count

    $script:count = 0

    function GetParentGroups {
        param(
            $groupName
        )
    
        $MemberOf = Get-ADPrincipalGroupMembership $groupName
    
        foreach($group in $MemberOf) {
            $script:count++ 
            GetParentGroups -groupName $group
        }
    }
    
    GetParentGroups -groupName $GroupName.Text
    
    $NumGroupsAfter = $NumGroupsBefore + $script:count

    $AnswerLabel.Content = $NumGroupsAfter
    $Global:SyncHash.print(("{0} will be in {1} groups after being added to {2}." -f $username.name, $NumGroupsAfter, $GroupName.Text), $false)


    ### Calculation by adding and removing (OLD) ###
    # $timeOut = 5
    # $counter = 0



    # Add-ADGroupMember -Identity $GroupName.Text -Members $ASURITE.Text

    # $token = (Get-ADUser $username -Properties tokengroups).tokengroups
    # $NumGroupsAfter = $token.count

    # while(($NumGroupsBefore -eq $NumGroupsAfter) -and ($counter -lt $timeOut)) {
    #     Start-Sleep -seconds 1
    #     $token = (Get-ADUser $username -Properties tokengroups).tokengroups
    #     $NumGroupsAfter = $token.count
    #     $counter = $counter + 1
    # }

    # Remove-ADGroupMember -Identity $GroupName.Text -Members $ASURITE.Text -Confirm:$false

    # if($NumGroupsBefore -eq $NumGroupsAfter) {
    #     $AnswerLabel.Content = "Error"
    #     $Global:SyncHash.print("There was an error counting the number of groups. Please try again.")
    # }

    # else{
    #     $AnswerLabel.Content = $NumGroupsAfter
    #     $Global:SyncHash.print(("{0} will be in {1} groups after being added to {2}." -f $username.name, $NumGroupsAfter, $GroupName.Text), $false)
    # }
})

$CloseBtn.add_click({
    $Form.close()
}),

$Form.Topmost = $true

#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null