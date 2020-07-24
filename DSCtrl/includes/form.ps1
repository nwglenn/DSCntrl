<#
.SYNOPSIS
Main form function.
.DESCRIPTION
Main form function that displays all the actions for the script.
#>

Add-Type -Assembly System.Windows.Forms     ## Load the Windows Forms assembly 

function MainForm {
    <#
    .SYNOPSIS
    Displays the main form for the script
    .DESCRIPTION
    Displays all the action buttons for the script.

    This function has been modified from it's original form located at https://gallery.technet.microsoft.com/GUI-Popup-Custom-Form-with-bf6c4141
    .PARAMETER listing
    List of buttons and their actions
    .PARAMETER title
    The name form
    .EXAMPLE
    MainFrom $buttons 'UTO Deskside Powershell'
    #>
    
    param ($listing, [string]$title)

    ## Create the main form  
    $form = New-Object Windows.Forms.Form 
    $form.FormBorderStyle = "FixedToolWindow" 
    $form.Text = $title 
    $form.StartPosition = "CenterScreen" 
    $form.Width = 740 ; $form.Height = 380  # Make the form wider 
    #Add Buttons- ## Create the button panel to hold the OK and Exit buttons 
    $buttonPanel = New-Object Windows.Forms.Panel  
    $buttonPanel.Size = New-Object Drawing.Size @(400,40) 
    $buttonPanel.Dock = "Bottom"    
    $exitButton = New-Object Windows.Forms.Button  
    $exitButton.Top = $buttonPanel.Height - $exitButton.Height - 10; $exitButton.Left = $buttonPanel.Width - $exitButton.Width - 10 
    $exitButton.Text = "Exit" 
    $exitButton.DialogResult = "Exit" 
    $exitButton.Anchor = "Right" 
    ## Create the OK button, which will anchor to the left of Exit 
    $okButton = New-Object Windows.Forms.Button   
    $okButton.Top = $exitButton.Top ; $okButton.Left = $exitButton.Left - $okButton.Width - 5 
    $okButton.Text = "Ok" 
    $okButton.DialogResult = "Ok" 
    $okButton.Anchor = "Right" 
    ## Add the buttons to the button panel 
    $buttonPanel.Controls.Add($okButton) 
    $buttonPanel.Controls.Add($exitButton) 
    ## Add the button panel to the form 
    $form.Controls.Add($buttonPanel) 
    ## Set Default actions for the buttons 
    $form.AcceptButton = $okButton          # ENTER = Ok 
    $form.CancelButton = $exitButton      # ESCAPE = Exit 
 
    ## Label and TextBox  
    ## Computer/Host Name 
    $lblHost = New-Object System.Windows.Forms.Label   
    $lblHost.Text = "Host Name:"  
    $lblHost.Top = 10 ; $lblHost.Left = 5; $lblHost.Width=150 ;$lblHost.AutoSize = $true 
    $form.Controls.Add($lblHost)    # Add to Form 
    # 
    $txtHost = New-Object Windows.Forms.TextBox  
    $txtHost.TabIndex = 0 # set Tab Order 
    $txtHost.Top = 10; $txtHost.Left = 160; $txtHost.Width = 120;  
    $txtHost.Text = $env:computername   # Use Corrent computer name as default 
    $form.Controls.Add($txtHost)    # Add to Form 
    # Obtain Value with: $txtHost.Text 
 
    ## ListBox - Fill with Data From Azure Location Name 
    $lblLoc = New-Object System.Windows.Forms.Label   
    $lblLoc.Text = "Azure Loation:"; $lblLoc.Top = 50; $lblLoc.Left = 5; $lblLoc.Autosize = $true  
    $form.Controls.Add($lblLoc)  
    Write-Host "Building List of available Locations" (Get-Date) -ForegroundColor Green 
    # Listbox for Location Name 
    $locListBox = New-Object System.Windows.Forms.ListBox  
    $locListBox.Top = 50; $locListBox.Left = 160; $locListBox.Height = 120 
    $locListBox.TabIndex = 1 
    # we need to populate the listbox... Example: $objListBox.Items.Add("Item 1 Test Do NOT USE") 
    # in our case, we will use a call to Azure for our "list" 
    $LocArray = @("a1", "a2", "a3") #Get-AzureLocation # | Format-list SubscriptionName, IsDefault, SubscriptionId 
    $i=0   # Counter 
    foreach ($element in $LocArray) {
        # Loop through Azure list and add to listbox 
        [void] $locListBox.Items.Add($element.name)  # Add element to listbox 
        # Using this looping, You can also do other line item processing if needed :) 
        if ($element.name -eq "East US 2") {
            [void] $locListBox.SetSelected($i,$true) 
        } # Set Default 
        $i ++ 
    } 
         
    $form.Controls.Add($locListBox) #Add listbox to form 
    Write-Host "Finished Getting Locations" (Get-Date) -ForegroundColor Green 
    # Obtain Value with: $locListBox.SelectedItem 
 
    ## CheckBox 
    $chkThis = New-Object Windows.Forms.checkbox 
    $chkThis.Left = 5; $chkThis.Width = 280; $chkThis.Top = 190  
    $chkThis.Text = "PowerShell CheckBox" 
    $chkThis.Checked = $true   # set a default value 
    $chkThis.TabIndex = 2 
    $form.Controls.Add($chkThis) 
    # Obtain Value with: $chkThis.Checked 
 
 
    ## Create the OS Image ComboBox 
    $lblImage = New-Object System.Windows.Forms.Label; $lblImage.Text = "VM Template:"; $lblImage.Top = 230; $lblImage.Left = 5; $lblImage.Autosize = $true  
    $form.Controls.Add($lblImage)    # Add to Form 
    $cbImage1 = New-Object Windows.Forms.ComboBox ; $cbImage1.Top = 230; $cbImage1.Left = 160; $cbImage1.Width = 550 
    Write-Host "Building list of available VM images" (Get-Date) -ForegroundColor Green 
    $ArrayImage = @("t1", "t2", "t3") #Get-AzureVMImage  # Download a list of VM OS Images from the Azure Portal 
    [void] $cbImage1.BeginUpdate() # This tells the control to not update the display while processing (saves time) 
    $i = 0 ; $iSelect = -1 
    foreach ($element in $ArrayImage) {  
        $thisElement = $i.ToString() +"::" + $element.label 
        [void] $cbImage1.Items.Add($thisElement) 
        if ($element.label -eq "Windows Server 2012 R2 Datacenter, April 2015") {
            $cbImage1.Text = $thisElement; $iSelect = $i 
        } # Set Default     $cbImage1.Text = $i.ToString() +"::" +$element.label 
        $i ++ 
    } 
    $cbImage1.SelectedIndex = $iSelect  # Set the default SelectedIndex 
    [void] $cbImage1.EndUpdate()  # update the control with all the data that was added 
    $form.Controls.Add($cbImage1) 
    Write-Host "Finished building list of available VM Images" (Get-Date) -ForegroundColor Green 
    # Obtain Value with: $cbImage1.SelectedItem 
    # Obtain index: $cbImage1.SelectedIndex 
    # Obtain Access to entire SELECTED element from original Array with: $ArrayImage[$cbImage1.SelectedIndex].????? 
 
    ## Finalize Form and Show Dialog 
    Write-Host "Show form" (Get-Date)  
    $form.Add_Shown( { $form.Activate(); $okButton.Focus() } )  #Activate and Set Focus 
    $result = $form.ShowDialog()          ## Show the form, and wait for the response 
 
    # Finished with Dialog Box, Now let's see what the user did... 
    $Result 
    if($result -eq "OK") {
        # Copy variables and use them as you desire... 
        $txtHost.Text 
        $locListBox.SelectedItem 
        $chkThis.Checked 
        $cbImage1.SelectedItem 
        Write-Host "VM Image:" $cbImage1.SelectedItem 
        Write-Host $cbImage1.SelectedIndex 
        $ArrayImage[($cbImage1.SelectedIndex)].ImageName 
        $ArrayImage[($cbImage1.SelectedIndex)].Label 
    } 
    else {
        Write-Host "Exit"
    } 

}

MainForm "hi" "UTO Deskside - DSCTRL"