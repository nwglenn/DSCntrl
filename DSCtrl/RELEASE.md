# DSCtrl Release Notes

## Version 1.3.2 April 21st , 2021

- **UPDATED**: Main Interface
    - The interface has changed from a dropdown menu driven interface to a tabs and button interface.

- **UPDATED**: File Share Report
    - The report generated should correctly display any errors and warnings when viewing the HTML report.

- **REMOVED**: LAPS retrieval module

- **REMOVED**: Establish BOMGAR connection
    - This function was not working as intended removing it from the UI for now

- **REMOVED**: Clean OU Module
    - Removing this until the need arises to bring it back, seems like a module that could do some accidental damage without proper testing

## Version 1.3.1 February 19th, 2021

**NOTE**
    - Version number has been updated to reflect development off of the original DSCtrl program instead of starting over.

- **REMOVED**: Move and Disable Module

- **REMOVED**: Update Dell BIOS Module

- **NEW**: Get Last Logon CSV Module
    - This module was a script Jake Hartman had provided to us to adapt to the DSCtrl UI

- **UPDATED**: Num Groups Module
    - Updated the interface for displaying the resulting number of groups for each button
    - Added the option to add the user to the group selected (with confirmation)

- **NEXT STEPS**
    - Update the Share Reporting module to include the errors correctly (as it has been before)
    - Continue to develop a TreeView UI to easier specify files/folders/OUs/AD Groups, etc instead of typing in a name
    - Add tooltips to the dropdown menu items for more clarity on functionality

## Version: 1.0 Januray 29th, 2021

- **UPDATED**: File Share Report  
    - Added "Include Empty Folders" Checkbox  
        - This will hide any folders that do not have child items in the report.  
    - HTML Report: Changed folders that only had inherited permissions, but had child items to take up less space in the report  
        - Previously, if a folder only had inherited permissions but had children with non-inherited permissions it would show in the report which was a lot of wasted space in the report. Now, they still show but do not show any permissions and only show the folder name.  
    - **NOTE**: Exporting to CSV has not been chaned, and the "Include Empty Folders" checkbox currently does not affect CSV.   
  
- **NEW**: Num Groups  
    - Module to check the number of groups a user is in  
    - Includes a "What if" to check if someone may go over the maximum allowed groups if you were to place them in specified group  
  
- **NEW**: Get TPM  
    - Module to get TPM information from a remote computer (uses the computer listed from the main control panel)  
        - This module requires that PSRemoting be enabled on the remote computer. As of right now, this is not a default setting so this module may not return information on the provided computer.  
  
- **NEW**: Get BitLocker  
    - Module to retrieve a BitLocker passowrd via either the computer name or by the BitLocker ID  
        - Note about searching by BitLocker ID: If the searchbase is set to the default (the entire ASURITE domain) the search will most likely take a minute or two  
  
- **NEW**: Test Remote  
    - Button to test the status of whether or not the listed computer is able to be PSRemoted into.  
  
- **Known Issues**:  
    - Get BitLocker has no input validation, if the user inputs an invalid computer name or BitLocker ID, the script will return nothing instead of informing the user that something is wrong.  
