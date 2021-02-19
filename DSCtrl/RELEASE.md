# DSCtrl Release Notes

## Version: 1.0 Januray 29th, 2020

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
