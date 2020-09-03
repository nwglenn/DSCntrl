# ..: rbc.ps1
# ..: Amazing, New and Wonderful!!   Combination of FULT-EMP, update_dl 
# ..: and sec_roles.  Provides an easier way to on-board and exit staff.
# ..: Version 1.0		7/19/2018		Jon Finley	(jfinley1)
# ..: Ver 1.1 fixed an issue where if a DL was created in all lowercase it was skipped for removal. 8/27/2018
# ----------------------------------------------------------------------------------------------------
<#
.SYNOPSIS
.DESCRIPTION
.PARAMETER
.NOTES
.EXAMPLE
rbc add haner UTO   Will Add haner to the default security group allowing access to resources
					and Create the Personal Folder
rbc add haner UTO -rn rolename	Will only add the Customer to the named role.
rbc del jfinley1 B1701 -rn UTO-SupportTeam	Will attempt to remove User ID from the B1701-UTO-SupportTeam role.
rbc del haner UTO	Will attempt to remove the User ID from all departmental DL's and SG's.
#>

param(
	[Parameter(Mandatory=$False,Position=1,HelpMessage="Enter Add, Del, Dept, New, Undo.")]
	[ValidateSet("Add","Del","Dept","New","Undo")]
	[alias("A")]
		[String]$Global:Action,     #Actions=Add, Del, Dept, New, Undo.
	[Parameter(Mandatory=$False,Position=2,HelpMessage="ASURITE ID of User.")]
	[alias("ID")]
		[String]$Global:UserID,     #ASU User ID.
	[Parameter(Mandatory=$False,Position=3,HelpMessage="Department (B0101001, SOLS, etc.")]
	[alias("D")]
		[String]$Global:Dept,       # Department name as set in the CSV file.
	[Parameter(Mandatory=$False,Position=4,HelpMessage="Name of the ROLE to Apply, Remove or Create.")]
	[alias("RN")]
		[String]$Global:RName,      # Role Name
	[Parameter(Mandatory=$False)]
	[alias("R")]
		[Switch]$Global:Role,       # Original switch for using a role
	[Parameter(Mandatory=$False)]
	[alias("NR")]
		[Switch]$Global:NewRole,    # Switch for a New Role creation
	[Parameter(Mandatory=$False)]
	[alias("DR")]
		[Switch]$Global:DeptRoles,  # Switch for collecting all roles for the Department
	[Parameter(Mandatory=$False)]
	[alias("CH")]
		[Switch]$Global:CreateHome, # Switch used to Create the Home folder.
	[Parameter(Mandatory=$False)]
	[alias("ST","SW")]
		[Switch]$Global:Student,    # Switch used to designate a Student Worker account.
	[Parameter(Mandatory=$False)]
	[alias("X","EX")]
		[Switch]$Global:Xternal,    # Switch used to designate RBC was launched externally.
	[Parameter(Mandatory=$False)]
	[alias("H","?")]
		[Switch]$Help               # Help Switch.
)

# testparams  Add asuriteID OU.location -Role Dept RoleName RoleAction -NewRole -OURoles

# -----------------------------------------------------
# -----------  GLOBALS  -------------
$MyCSVfile="\\itfs1.asu.edu\UTO\Transfer\___UTO_TS___\RBC\DeptCSV.CSV"
$MyOUTFile="\\itfs1.asu.edu\UTO\Transfer\___UTO_TS___\RBC\Undo"
$MySRPath="\\itfs1.asu.edu\UTO\Transfer\___UTO_TS___\RBC\Sec-Roles"   # Path to network share for storeage of roles
# The paths above should end with RBC\DeptCSV.CSV, RBC\Undo and RBC\Sec-Roles
# The "pre" path should change to the location that makes sense for your Deskside Team
# -----------------------------------------------------------
$Global:MyServer="asurite.ad.asu.edu"
$MyVer=1.4
# -----------------------------------------------------------

function AddItems($MyRoleLoc){
	$MyDeptRoles=Get-Content $MyRoleLoc | sort
	$MySGs=$MyDeptRoles | where-object{-NOT $_.Contains("CN=DL.")}
	$MySGs | SORT | ForEach-Object{
			write-host "Adding $MyUser to: "$_.split(",")[0].substring(3) -foregroundcolor Green
			if($Student){
				$MyUserID=Get-ADUser -ident $UserID -Server "ad.asu.edu"
				$MyADsg=Get-ADGroup $_ -server "asurite.ad.asu.edu"
			}else{
				$MyUserID=$UserID
				$MyADsg=$_
			}
			try
			{
				Add-ADPrincipalGroupMembership $MyUserID -MemberOf $MyADsg 3> $null -ErrorAction Stop
			}
			catch [System.Exception]
			{switch ($_.FailureCategory){
				MemberAlreadyExistsException	{write-host $UserID "already exists in "$_ -foregroundcolor Yellow}
				}
			}
		}
	$MyDLs=$MyDeptRoles | where-object{$_.Contains("CN=DL.")}
	if(-NOT $MyDLs){write-host "No Departmental DLs found to ADD, exiting.." -foregroundcolor Green;exit}
	ExchLogin
	if($Student){
		$MyUser=$UserID+"@sundevils.asu.edu"
	}else{
		$MyUser=$UserID+"@asurite.asu.edu"
	}
	$MyDLs | foreach{
			write-host "Adding $MyUser to: "$_.split(",")[0].substring(3) -foregroundcolor Green
			try
			{
				Add-DistributionGroupMember -ident $_.split(",")[0].substring(3) -Member $MyUser -Confirm:$false 3> $null -ErrorAction Stop
			}
			catch [System.Exception]
			{switch ($_.FailureCategory){
				MemberAlreadyExistsException	{write-host $UserID "already exists in "$_ -foregroundcolor Yellow}
				OperationRequiresGroupManagerException	{write-host "You are NOT an owner of this DL: "$_ -foregroundcolor Red}
				}
			}

		}
	ExchLogoff
} # --- end of function ---

function AddRole{
# Initial onboarding or job function change
	if ($CreateHome){CreateHome}
	if( -NOT $Global:RName){GetRole}
	write-host "Prog will ADD the $Dept-$RName role to $UserID"
	$MyRole="$MySRPath\$Dept-$RName.txt"
	if (test-path $MyRole){
		write-host "Adding the $Dept-$RName ROLE to "$UserID
	    AddItems($MyRole)
	}else{Write-host "ERROR!  File does not exist or cannot be found!  Check your path."}
} # --- end of function ---

function BackUpRoles{
	write-host "Backing up $UserID's current SG's and DL's to $MyOUTFile\$UserID-undo.txt file"
	$Global:Protected=$null
	$Global:MyUndoFile="$MyOUTFile\$UserID-undo.txt"
	# This function triggers as part of the delete function
	#export out the MemberOF array for possible UNDO or ROLE use (asurite_id-undo.txt)
	GetUserGroups($MyUndoFile)

} # --- end of function ---

function CreateHome{
	write-host "Creating $MyDeptCIFS\$UserID and updating $MyDeptCIFSsg for $UserID"
	write-host "Adding "$UserID" to "$MyDeptCIFSsg " and the "$MyDeptCIFS" share"
	if($MyDeptCIFSsg){
		Add-ADGroupMember -Ident $MyDeptCIFSsg -Member $UserID -Confirm:$false
	}
	#add id to all group DL's (see note on Del below)
    write-host "Creating new Personal folder under: "$MyDeptCIFS" and setting the Modify ACL"
    New-Item "$MyDeptCIFS\$UserID" -type directory | Out-Null
    #set ACLs on personal folder (old way - icacls.exe $folder /grant 'domain\user:(OI)(CI)(M)')
    $newacl=new-object System.Security.AccessControl.FileSystemAccessRule ("ASURITE\$UserID","Modify","ContainerInherit,ObjectInherit","None","Allow")
    $curacl=get-acl $MyDeptCIFS\$UserID
    $curacl.SetAccessRule($newacl)
    $curacl | Set-Acl $MyDeptCIFS\$UserID
} # --- end of function ---

function CreateRole{
	if( -NOT $Global:RName){GetRole}
	write-host "Prog will CREATE the NEW $Dept-$RName role from $UserID"
	$MyRole="$MySRPath\$Dept-$RName.txt"
	#omit the OUusers group in output.
	$Global:Protected="$MyDeptOU.OUusers"
	if (test-path $MyRole){
	   write-host `r`n"The ROLE file already exists!"
	   $MyAns=Read-Host -prompt 'Do you want to over-write the ROLE file? (Y/N)'
	   if ($MyAns -EQ "N"){
	     $RName=Read-Host -prompt 'Please provide the new role name'
		 $MyRole="$MySRPath\$Dept-$RName.txt"
		 }
	}
	GetUserGroups($MyRole)
	write-host `r`n"New ROLE $Dept-$RName created!"`r`n
} # --- end of function ---

function DelAll{
	# This function should trigger at Termination or Transfer.
	write-host "Deleting "$UserID" from "$MyDeptCIFSsg" and the "$MyDeptCIFS" share"
	#move personal share to ..\users\_ARCHIVE folder
	write-host "Moving the "$UserID" personal folder to _ARCHIVE"
	try
	{
		Move-Item "$MyDeptCIFS\$UserID" "$MyDeptCIFS\_ARCHIVE" -ErrorAction Stop 
	}
	catch [System.Exception]
	{switch ($_.FailureCategory){
		PathNotFound	{write-host $UserID "folder was not found at "$_ -foregroundcolor Yellow}
		}
	}	
	BackUpRoles
	#remove account from ALL groups in OU or ALL Groups possible
	write-host "Program will remove all SG's and DL's for $Dept from $UserID"
	DelItems($MyUndoFile)
} # --- end of function ---

function DelItems($MyRoleLoc){
    $MyDeptRoles=Get-Content $MyRoleLoc | SORT
	$MySGs=$MyDeptRoles | where-object{-NOT ($_.toupper()).Contains("CN=DL.")}
	$MySGs | SORT | ForEach-Object{
			write-host "Deleting $UserID from: "$_.split(",")[0].substring(3) -foregroundcolor Yellow
			if($Student){
				$MyUserID=Get-ADUser -ident $UserID -Server "ad.asu.edu"
				try
				{
					$MyADsg=Get-ADGroup $_ -server "asurite.ad.asu.edu" 2>$null
				}
				catch [System.Exception]
				{switch ($_.FailureCategory){
					ObjectNotFoundException {write-host $UserID "is not a member of "$_ -foregroundcolor Yellow}
					}
				}
			}else{
				$MyUserID=$UserID
				$MyADsg=$_
			}
			try
			{
				Remove-ADPrincipalGroupMembership $MyUserID -MemberOf $MyADsg -Confirm:$false 3> $null -ErrorAction Stop
			}
			catch [System.Exception]
			{switch ($_.FailureCategory){
				MemberNotFoundException	{write-host $UserID "is not a member of "$_ -foregroundcolor Yellow}
				}
			}
	}
	$MyDLs=$MyDeptRoles | where-object{($_.toupper()).Contains("CN=DL.")}
	if(-NOT $MyDLs){write-host "No Departmental DLs found to REMOVE, exiting.."}else{
	if(-NOT $Xternal){ExchLogin}
	if($Student){
		$MyUser=$UserID+"@sundevils.asu.edu"
	}else{
		$MyUser=$UserID+"@asurite.asu.edu"
	}
	$MyDLs | foreach{
			write-host "Deleting $MyUser from: "$_.split(",")[0].substring(3) -foregroundcolor Yellow
			try
			{
				Remove-DistributionGroupMember -ident $_.split(",")[0].substring(3) -Member $MyUser -Confirm:$false -ErrorAction Stop
			}
			catch [System.Exception]
			{switch ($_.FailureCategory){
				MemberNotFoundException	{write-host $UserID "is not a member of "$_ -foregroundcolor Yellow}
				OperationRequiresGroupManagerException	{write-host "You are NOT an owner of this DL: "$_ -foregroundcolor Red}
				}
			}
		}
	if(-NOT $Xternal){ExchLogoff}
	}
} # --- end of function ---

function DelRole{
	if( -NOT $Global:RName){GetRole}
	write-host "Prog will DELETE the $Dept-$RName role from $UserID"
	$Global:MyRole="$MySRPath\$Dept-$RName.txt"
	# Job function change
	if (test-path $MyRole){
	     write-host "Removing the $Dept-$RName ROLE from "$UserID
		 DelItems($MyRole)
	}else{Write-host "ERROR!  File does not exist or cannot be found!  Check your path."}
} # --- end of function ---

function ExchLogin {
	$UserCredential = Get-Credential
	$Global:Session = New-PSSession -ConfigurationName Microsoft.Exchange -ConnectionUri https://mail.asu.edu/powershell -Credential $UserCredential -Authentication Basic -AllowRedirection
	Import-PSSession $Session 3> $null
} # --- end of function ---

function ExchLogoff {
	Remove-PSSession $Session
	$Global:Session=$null
} # --- end of function ---

function GetADDept{
	$Global:MyDeptFullOU = (Get-ADOrganizationalUnit -filter 'Name -like $MyDeptOU').DistinguishedName
#return $Global:MyDeptFullOU
} # --- end of function ---

function GetDept{
param(
	[Parameter(Mandatory=$True,HelpMessage="Dept to use")]
	[alias("D")]
	[String]$Global:Dept
	)
  return $Global:Dept = $Global:Dept.toupper()
} # --- end of function ---

function GetDeptDLs{
#	ExchLogin
	write-host "Gathering ALL Departmental DL's..."
	if( $MyDeptDLs.Count -GT 1){
		$MyDL1=$MyDeptDLs[0]+="*"
		$MyDL2=$MyDeptDLs[1]+="*"
		$Global:DeptDLs=(Get-DistributionGroup -filter "Name -like '$MyDL1' -OR Name -like '$MyDL2'").DistinguishedName
	}else{
		$MyDeptDLs+="*"
		$Global:DeptDLs=(Get-DistributionGroup -filter "Name -like '$MyDeptDLs'").DistinguishedName
	}
	write-host "Done gathering Departmental DL's"
#	Remove-PSSession $Session
#return $Global:DeptDLs
} # --- end of function ---

function GetDeptRoles{
	if(-NOT $Dept){GetDept}
	write-host "Prog will: Gather all of the roles available from: $Dept located at: $MyDeptOU `r`nand save it to: $MyDeptFile"
	write-host "Gathering ALL of the Security Groups from the "$MyDept" OU structure."
	write-host "Collection will be saved as: "$MyDeptFile
	$MyDeptOUTFile="$MySRPath\$MyDeptFile"
	GetDeptSGs
	$DeptSGs| Out-File $MyDeptOUTFile
	ExchLogin
	GetDeptDLs
	ExchLogoff
	$DeptDLs | Out-File $MyDeptOUTFile -Append
} # --- end of function ---

function GetDeptSGs{
	GetADDept
	$Global:DeptSGs=(Get-ADObject -Filter 'ObjectClass -like "group"' -Searchbase "$MyDeptFullOU").DistinguishedName
#return $Global:DeptSGs
} # --- end of function ---

function GetRole{
param(
	[Parameter(Mandatory=$True,HelpMessage="Name of the Role to use")]
	[alias("RN")]
	[String]$Global:RName
	)
  return $Global:RName
} # --- end of function ---

function GetServer{
	if($Student){
		$Global:MyServer="ad.asu.edu"
	}else{
		$Global:MyServer="asurite.ad.asu.edu"
	}
} # --- end of function ---

function GetUserGroups($MyRoleLoc){
	if($MyDeptDLs.count -GT 1){
		Get-ADPrincipalGroupMembership -ident $UserID -Server $MyServer | where{($_.DistinguishedName.toupper()).contains("OU=$MyDeptOU") -OR ($_.DistinguishedName.toupper()).contains($MyDeptDLs[0]) -OR ($_.DistinguishedName.toupper()).contains($MyDeptDLs[1]) -AND $_.DistinguishedName -NE $Protected} | select DistinguishedName -expand DistinguishedName | sort | Out-File $MyRoleLoc
	}else{
		Get-ADPrincipalGroupMembership -ident $UserID -Server $MyServer | where{($_.DistinguishedName.toupper()).contains("OU=$MyDeptOU") -OR ($_.DistinguishedName.toupper()).contains($MyDeptDLs) -AND $_.DistinguishedName -NE $Protected} | select DistinguishedName -expand DistinguishedName | sort | Out-File $MyRoleLoc
	}
} # --- end of function ---

function PrintVer{
	$MyProgName = $myInvocation.ScriptName
	write-host "Running Script: $MyProgName `tJon Finley (jfinley1) `tVersion: "$MyVer
}

function ReadCSV{
# the CSV file holds entry and exit information for all supported Departments
	$MyCSV=Get-Content $MyCSVfile
	$MyDeptInfo=$MyCSV | where-object{$_.split(",")[0].Contains(("$Dept").toupper())}
	if(-NOT $MyDeptInfo){
		write-host "TERMINATING Program - Department ID: [ $Dept ] was NOT found in the CSV file." -foregroundcolor Yellow
		exit
	}
	$MyDept=$MyDeptInfo.split(",")
		$Global:MyDeptUnit=$MyDept[0].toupper()
		$Global:MyDeptOU=$MyDept[1]
		$Global:MyDeptFile=($MyDeptOU.split(".") -join "-")+"-OU-List.txt"
		$Global:MyDeptCIFS=$MyDept[2]
		$Global:MyDeptCIFSsg=$MyDept[3]
		$Global:MyDeptsg="T.Dept."+$MyDeptUnit
		$Global:MyDeptDLs=($MyDept[4]).split(";")
	if($MyDeptUnit -NE $Dept){
		write-host "TERMINATING Program - Department ID: [ $Dept ] was NOT found in the CSV file." -foregroundcolor Yellow
		exit
	}
	
} # --- end of function ---

function UndoDel{
	$MyUndo="$MyOUTFile\$UserID-undo.txt"
	write-host "OOPS!  Reading in from $MyUndo file.  Will re-apply rolls now!"
	write-host "Attempting to restore Folder and security groups for: "$UserID
	try
	{
		Move-Item "$MyDeptCIFS\_ARCHIVE\$UserID" "$MyDeptCIFS\$UserID" -ErrorAction Stop
	}
	catch [System.Exception]
	{switch ($_.FailureCategory){
		MemberAlreadyExistsException	{write-host $UserID "already exists in "$_ -foregroundcolor Yellow}
		OperationRequiresGroupManagerException	{write-host "You are NOT an owner of this DL: "$_ -foregroundcolor Red}
		}
	}
	AddItems($MyUndo)
} # --- end of function ---

# ------------------- End of Functions -----------------------

# ----------------------------------------------------------------------------------
# ----------------------  MAIN PROGRAM ---------------------------------------------
# ----------------------------------------------------------------------------------
PrintVer

GetServer
if($Dept){$Dept=$Global:Dept=$Global:Dept.toupper()}else{GetDept}

ReadCSV

switch ($Action){
	Add	{
			if($RName -OR $Role){$Role=$true;AddRole}else{CreateHome}
		}
	Del	{
			if($RName -OR $Role){$Role=$true;DelRole}else{DelAll}
		}
	Dept{
			$Global:DeptRoles=$true
			GetDeptRoles
		}
	New	{
			$Global:NewRole=$true
			if($RName -OR $Role){CreateRole}
		}
	Undo{
			UndoDel
		}
}

# ==================================================================================
# ===============================   End of Program =================================
# ==================================================================================
