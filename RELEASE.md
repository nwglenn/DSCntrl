# DSCtrl Release Notes

## 1.3.0  Wednesday, August 22, 2018
- **NEW**: The menus have been re-arranged to make finding specific tools easier
- **NEW**: The Share Report has been significantly improved! Huge shout out to Orion Eckstrom for all of his feedback.
  - **NEW**: UI & performance improvements
  - **NEW**: Improved HTML report layout that is easier to read
  - **NEW**: The report now hides folders that inherit all of their permissions from a parent, reducing clutter in the report
  - **NEW**: Folders and files now display the created, last accessed and modify dates
  - **NEW**: Added ability to only display files and folders that are older than a specified number of years, or are larger than a specific size
  - **NEW**: The report now more clearly shows when a user account has explicit permissions to a folder
  - **FIX**: Script produced an error when the permissions on a folder contained "CREATOR Owner"
  - **FIX**: When a folder did not contain any items, the report would incorrectly show an 'access denied' message. The report now shows; “Warning: Empty folder”
  - **FIX**: CSV export was saving a blank file

## 1.2.8  Friday, July 13, 2018
* NEW: Several changes to the _Computer Clean-Up_ function
  * Corrected typo in comments on disabled systems and formated the date
  * Provided the ability to search and display systems not in ServiceNow
  * Added ability to search either current OU or all children.
  * You can now change the number of months used in the search for stale objects

## 1.2.7  Friday, June 29, 2018
* NEW: Computer Clean-Up now prompts for the number of months to use when searching for old comuter accounts
  * For example, entering 1 will flag any Computer Accounts that have not communicated with AD in more than 1 month. 6 or 12 months are recommended

## 1.2.6  Wednesday, May 16, 2018
* NEW: Computer clean-up now detects machines that were never joined to AD
* NEW: Computer clean-up can now be run on any OU - rather than only staging OUs
* NEW: Computer clean-up tool now also displays LAPS expiration dates
* FIX: When a DL is removed, the list is now updated to reflect the change
* FIX: When removing DL's from a user, there was a case where the group would not actually be removed

## 1.2.5  Wednesday, March 14, 2018
* NEW: Exchange Offboarding form added to Tools menu

## 1.2.4  Wednesday, Feb 28, 2018
* NEW: Redisigned share report tool
* NEW: Migration worksheet now populates Support Group, Update Group and department fields
* NEW: You can now open the Service-Now asset for a specified computer via the Computer menu

## 1.2.3  Wednesday, January 17, 2018
* NEW: The group creation tool has been updated to calculate long resource names with the ~# suffix, where # is the next available number starting with 1
* NEW: When using the bulk import tool, the name of the group being created now appears at the top of each group creation window
* NEW: Completely new update interface which displays full release information in a much more readable format
* NEW: When creating security groups for network shares, Deny-All (DA) groups are enabled by default but can still be disabled if desired
* NEW: When creating share or printer security groups, the 'Authorized aprober' field is now marked as optional
* NEW: A new tools menu has been added which contains a Token Calculator tool
  * This tool estimates the size of a users security token and can be used to identify/troubleshoot login problems related to token size problems
* FIX: When a ticket number was provided on the bulk import form, it would randomly not appear in subsequent group creation windows
* FIX: 1.2.2 introduced a bug in which creating a printer using the bulk import tool would also want to create a parent '.PRT' group
* FIX: Log messages that are generated on successful group creation no longer say 'completed without error' because it was confisong
* FIX: Groups for printers and shares now check for additional characters that are not allowed in security group names and omits them from the generated name

## 1.2.2  Thursday, November 30, 2017
* NEW: Support for bulk renaming has been added to the 'Groups' menu
* FIX: AD queries have been synchronized to use the same AD server between calls, resolving an issue where sequenced calls would fail because of replication delays between servers
* FIX: The bulk import tool will now check for intermediate parent groups and create them if needed

## 1.2.1  Tuesday, November 7, 2017
* NEW: The computer security group report is now much more robust; providing information about nested, software and Windows 10 servicing groups
* NEW: When creating security groups, the form will now remember the last used unit name
* NEW: Security groups are now backed up locally whenever modifications are made within DSCtrl (%localappdata%\ASU\DSCtrl\backups)
* NEW: You can now launch a Bomgar jump session from the Computer menu
* FIX: Corrected a problem that would prevent the bulk group creation form from being displayed properly

## 1.2.0  Tuesday, October 24, 2017
* NEW: You can now automate the creation of security groups by importing a list of groups from a migration worksheet
* FIX: Corrected an issue where share groups could be created with incorrectly formatted type identifiers

## 1.1.4  Tuesday, October 10, 2017
* NEW: When creating a printer security group, the form now displays the driver name, IP address and location information that was queried from the selected printer. This information can also be edited before creating the group
* NEW: The security group creation window now displays the character length of the group name to help avoid names larger than 64 characters
* NEW: It is no longer necessary to have the LAPS PowerShell module installed to obtain administrator passwords
* FIX: When using the security group creation tool to create share groups; the form would fail to provide an accurate name recommendation if the share name contained spaces or several sub-directories
* FIX: When running the computer clean-up script, closing the initial warning message would continue the process. The process now only continues if the 'OK' button is used
* FIX: Error messages generated by the security group tool were appearing underneath the window; hiding them from view

## 1.1.3  Thursday, September 28, 2017
* Fix: When nesting user and computer security groups into their parent groups, this would fail in specific circumstances
* Fix: populating computer groups from a staging area would fail if any warnings were generated on previous steps
* Fix: When launching the group creation screen, the new window would often appear underneath other windows
* Fix: When using the computer cleanup function, queries to Service Now would fail if Internet Explorer was not allowed to render JSON documents. A check has been added to ensure this seting is enabled, and may require one time elevated access to configure this setting.
* New: When creating new security groups, optional fields are now marked with (optional)
* New: When a newer version of the Deskside Control Panel is available, a list of the most recent changes will be displayed
* New: All tasks now log their execution to %LOCALAPPDATA%\ASU\DSCtrl\transcripts\ to aid in troubleshooting and debugging

## 1.1.2  Friday, September 22, 2017
* {e67056edf8} When the computer cleanup process disables 'unused' computer objects, it now populates the _ExtensionAttribute7_ property with the date it was disabled, and the values of any fields used when determining how much time has passed since last contact.
* {fd3443f45a} You can now edit the names of printer groups after selecting a printer
* {fd3443f45a} Fixed an issue where the length of group names was not being calculated correctly, resulting in the app trying to create groups with names larger than 64 characters
* {46fdc5fc70} Fixed spelling errors in printer report
* {05d54983d7} Resolved an issue where the Service-Now login for computer cleanup would not be detected properly under very specific conditions
* {309f2f8403} Small UI tweaks to prevent computer name text box from dissapearing

## 1.1.0
* {UTOITCSSD-67} App now checks for updates at startup and via the 'Check for updates' option within the Help menu. When updates are found, they can be automatically downloaded and installed by clicking 'yes' when prompted.

## 1.0.3
* {c1693d1066} Addressed script not running on machines with restrictive ExecutionPolicy modes.

## 1.0.2
* {UTOITCSSD-108} Remove need for admin elevation
* {245a7c3456} When creating a Computer and User security group at the same time with the option to populate from a staging area enabled, the script would sometimes populate _both_ Computer and User groups with the computers. This has been fixed to only populate the Computer group.
* {b04e701250} Removed requirement for LAPS to be installed before launching DSCtrl. Now, if LAPS is not installed; the option to query a LAPS password will be disabled in the interface.

## 1.0.1
* Initial release to UTO techs involved in first migrations.
