<#
Project Title => Active Directory Group Membership Report
Programmer:
     Paul Marcantonio
Date:
     September 12 2018
Version:
     1.8
Objective:
     - Obtain all groups within Active Directory Sort by Name
	 - Obtain all users within Active Directory sort by Name
     - Obtain all members within each Groups (Group and/or Users)
	 - Obtain member attribute (Name, Title)
     - Build Excel file containing (GroupName, MemberType, MemberName, Title)
     - Place a Copy of the excel file in archivelocation on network share
	 - Send Email with excel as attachment
	 - Remove excel files in archive that are older than 7 days from running of this script
Pre-Condition(s):
     - Powershell Execution Policy set to Remove Signed (Get-ExecutionPolicy, Set-ExecutionPolicy RemoteSigned)
     - EPPLUS DLL must be downloaded and ready for loading (https://www.nuget.org/packages/EPPlus/)
     - Import AD module for powershell commandlets to be available https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-pss
     - Folder structure and right to read and write HTML file to
     - The server this script runs on must be on the "allow relay" on any and all load balancers infront of exchange and within exchange
Post-Condition(s):
     - New Excel file created and populated with group and membership infformation
     - Email sent to the desired email list with:
          * Copy of Excel file
Installation:
     - Make sure Powershell version 3 or above is installed on the server (Server Role and Features (Windows Powershell)
     - Make available EPPLUS DLL for loading (https://www.nuget.org/packages/EPPlus/)
Contributing:
     - https://www.nuget.org/packages/EPPlus/
Citations:
     None
Contact:
     Paul Marcantonio
          Email => paulmarcantonio@yahoo.com

#>

<#
.SYNOPSIS => Initalize EPPLUS DLL for use to build excel file
.PARAMETER
     None Needed
.INPUTS
     Path to EPPLUS DLL (I.E "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll")
.OUTPUTS
    None
#>
FUNCTION EXCEL_LOAD_DLL()
{
# Load EPPlus
	$DLLPath = "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll"
	[Reflection.Assembly]::LoadFile($DLLPath) | Out-Null
}

<#
.SYNOPSIS => Return Excel Package Object for use during Excel opperations
.PARAMETER
     Full UNC Path to the None Needed
.INPUTS
     Path to EPPLUS DLL (I.E "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll")
.OUTPUTS
    Excel object of the file path passed
#>
FUNCTION EXCEL_BUILD_FILE_OBJECT ($filePath)
	{
	#CREATE EXCEL OBJECT
		$ExcelPackage = New-Object OfficeOpenXml.ExcelPackage($filePath) 
	RETURN $ExcelPackage
	}

<#
.SYNOPSIS => Adds a worksheet to the excelObject passed 
.PARAMETER
    $excelObject = Excel Object
	$name = Name of the new worksheet 
.INPUTS
     None
.OUTPUTS
    Worksheet within object passed 
#>
FUNCTION EXCEL_ADD_WORKSHEET($excelObject, $name)
{
	$workSheet = $excelObject.Workbook.Worksheets.Add($name)
	$excelObject.Save()
	RETURN $workSheet
}

<#
.SYNOPSIS => Converts data passed and adds data to worksheet within excelObject passed 
.PARAMETER
    $excelObject = Excel Object
	$name = Name of the target worksheet
	$data = data you want to insert worksheet
.INPUTS
     None
.OUTPUTS
    Excel is updated with data and file is saved
#>
FUNCTION EXCEL_ADD_DATA_TO_WORKSHEET ($excelObject, $excelWorkSheet, $data)
{
	#CONVERT DATA TO CSV FORMAT
		$ProcessesString = $data | ConvertTo-Csv -NoTypeInformation | Out-String
	#BUILD HOW THE CSV WILL BE DELIMITED 
		$Format = New-object -TypeName OfficeOpenXml.ExcelTextFormat -Property @{TextQualifier = '"'}
	#LOAD DELIMITED DATA INTO SPREADSHEET 
		$null=$excelWorkSheet.Cells.LoadFromText($ProcessesString,$Format)
	#SAVE TO EXCEL OBJECT
		$excelObject.Save()
}

<#
.SYNOPSIS => Converts data passed and adds data to worksheet within excelObject passed 
.PARAMETER
    $excelObject = Excel Object
	$name = Name of the target worksheet
	$data = data you want to insert worksheet
.INPUTS
     None
.OUTPUTS
    Worksheet is updated with data and file is saved
#>
FUNCTION BUILD_EXCEL_FILE ($data)
{
	#BUILD EXCEL FILE
	# Load EPPlus
		$DLLPath = "E:\MetroITScripts\Bin\DotNet4\EPPlus.dll"
		[Reflection.Assembly]::LoadFile($DLLPath) | Out-Null
	# Clear old File
		#clearCurrentFile $global:ExcelFile
		# Create Excel File
			$ExcelPackage = New-Object OfficeOpenXml.ExcelPackage 
			$Worksheet = $ExcelPackage.Workbook.Worksheets.Add($global:timeStamp)
			$ProcessesString = $data | ConvertTo-Csv -NoTypeInformation | Out-String
			$Format = New-object -TypeName OfficeOpenXml.ExcelTextFormat -Property @{TextQualifier = '"'}
			$null=$Worksheet.Cells.LoadFromText($ProcessesString,$Format)
			$ExcelPackage.SaveAs($global:ExcelFile)
	}

<#
.SYNOPSIS => Set background colors, auto fit, font, and autoFilter in target worksheet  
.PARAMETER
    $excelObject = Excel Object
	$workSheetName = Name of the target worksheet
	$headers = data section that will be formated 
.INPUTS
     specific ranges to target
.OUTPUTS
    Worksheet will be updated with proper formatting along the headers
#>
FUNCTION EXCEL_FORMAT_WORKSHEET_HEADERS($excelObject, $worksheetName, $headers)
{
	#https://www.powershellgallery.com/packages/ImportExcel/5.4.4/Content/Set-Row.ps1
	$headerRange = 'A1:A4'
	$targetWorkSheet = $excelObject.Workbook.Worksheets[$worksheetName]
	#Add Headers to Sheet (Note: $headers array is zero indexed) 
	for ($row = 1; $row -le 1; $row ++)
	{
		for ($column = 1; $column -le $headers.count; $column ++)
		{
			$targetWorkSheet.Cells[$row, $column].Style.Fill.PatternType = 1 # 1 Denotes solid Color 
			$targetWorkSheet.Cells[$row, $column].Style.Fill.BackgroundColor.SetColor([System.Drawing.Color]::LightBlue)
			$targetWorkSheet.Cells[$row, $column].Style.Font.Bold = $true
			$targetWorkSheet.Cells[$row, $column].AutoFitColumns()
		}
	}
	$targetWorkSheet.Cells["A1:D1"].AutoFilter = $true
	$excelObject.Save()
}

<#
.SYNOPSIS => Send email out to receiptents and an excel attachment.  
.PARAMETER
    $attachment = Excel attachments
	$messageBody = Name of the target worksheet
.INPUTS
     $smtpServer = email server that will receive the email
	 $emailFrom = General email address that email will be preceived to come from
	 $emailToSuccess = creates an array of smtp addresses to be sent in the email.
	 $subjectTextSuccess = the text that will be presented on the subject line 
.OUTPUTS
    Worksheet will be updated with proper formatting along the headers
#>
FUNCTION SEND_EMAIL ($attachment, $messageBody)
{
	# --- Set Email Variables ---
	#EMAIL SERVER 
		$smtpServer = "owa.company.org"									#Metro Email server used to process email
		$smtp = New-Object Net.Mail.SmtpClient($smtpServer)								#

	#FROM ADDRESS
		$emailFrom = "Daily_AD_Report_Groups_and_Members_$ENV:COMPUTERNAME@company.org"				#from email address shown in the email
	#TO ADDRESSED
		$emailToSuccess = @("PMarcantonio@company.org") #success email address that will recieve logs
	
	#SUBJECT 
		$subjectTextSuccess = "Daily Extract AD Group and Membership Script completed on server $ENV:COMPUTERNAME at $global:timeStamp"	#subject text for success messages
	#BODY CONTENT 
		$global:bodyTextSuccess += $global:htmlMobileHeader 
	#SEND EMAILS 
		foreach ($rcp in $emailToSuccess)
		{
			Send-MailMessage -from "$emailFrom" -to "$rcp" -subject "$subjectTextSuccess" -body $messageBody -BodyAsHtml -Attachments $attachment -smtpServer "$smtpServer"
		}
}

<#
.SYNOPSIS => Remove excel files that are older than the date passed 
.PARAMETER
    $dir = UNC to the directory that we want to remove files in 
	$termDate = Date witch will be used to compare the lastwritetime of each file
.INPUTS
     None
.OUTPUTS
    None
#>
FUNCTION CLEANUP_ARCHIVES($dir, $termDate)
{
	Get-ChildItem -LiteralPath $dir -File | Where-Object {$_.LastWriteTime -lt $termDate} | Remove-Item -Force
}

<#
.SYNOPSIS => Gather AD Groups and their members, build excel file with group and members, send email to target emails  
.PARAMETER
	None
.INPUTS
    $global:ExcelFileDir = Location of excel files
	$global:ExcelFile = 
.OUTPUTS
    None
#>
FUNCTION MAIN ()
{
	#OBTAIN AD GROUPS
		$groups = Get-ADGroup -Filter * -properties * | sort-object name
	#OBTAIN AD USERS
		$users = Get-ADUser -Filter * -properties * | sort-object name
	#TIME STAMP FOR FILE 
		$timeStamp = Get-Date -uformat "%Y-%m-%dT%H-%M-%S"
	#INSTANTIATE EXCEL FILE
		$global:ExcelFileDir = "\\SERVER\Web Pages\AD Reports\GroupMembership"
		$global:ExcelFile = "$global:ExcelFileDir\ADGroups_and_Membership$timeStamp.xlsx"
			if(Test-Path $global:ExcelFile)
			{
				Remove-Item -LiteralPath $global:ExcelFile -Force
			}
	#TIME STAMP FOR EMAIL 
		$global:timeStamp = Get-Date -UFormat "%Y-%m-%d at %H-%M-%S"
	#INSTANTIATE ARRAY TO HOLD GROUP AND MEMBERSHIP RESULTS 
		$resultsArray = @()
	#LOOP GROUPS AND OBTAIN MEMBERSHIP AND POPULATE RESULTS ARRAY
		FOREACH ($group in $groups)
		{
			#LOOP MEMBERS IN GROUP MEMBERSHIP
			FOREACH ($member in $group.members)
			{
				#IF THE GROUP MEMBER IS A GROUP (SUBGROUP) THEN CREATE GROUP OBJECT
					#IF (($groups.DistinguishedName).contains($member))
					IF ($groups.DistinguishedName -contains $member) 
					{
						#$memberGroup = (Get-ADGroup -Filter{DistinguishedName -eq $member}).Name
						#OBTAIN GROUP NAME FROM PASSED DISTINGUISHED NAME
							$memberGroup = $groups | Where-Object {$_.DistinguishedName -eq $member} | Select-Object Name
						#BUILD SUBGROUP OBJECT 
							$subGroup = [PSCustomObject]@{
								GroupName = $group.Name
								MemberType = "Group"
								MemberName = $memberGroup
								Title = "N/A"}
						#ADD TO ARRAY
							$resultsArray += $subGroup
					}
				#IF THE GROUP MEMBER IS A USER, THEN CREATE USER OBJECT
					ELSEIF (($users.DistinguishedName).contains($member))
					{
						#OBTAIN USER OBJECT FROM USERS ARRAY	
							$memberUser = $users | where-object {$_.DistinguishedName -eq $member} 
						#GATHER ATTRIBUTES	
							$name = $memberUser.Name
							$title = $memberUser.Title
							if ($title -eq $null)
							{
								$title = "(--No Title Listed in AD--)"
							}
						#BUILD USER OBJECT
							$userInGroupObject = [PSCustomObject]@{
								GroupName = $group.Name
								MemberType = "User"
								MemberName = $name
								Title = $title}
						#ADD TO RESULTS ARRAY
							$resultsArray += $userInGroupObject
					}
			}
		}
	#CONSTRUCT EXCEL FILE WITH DETAILS
		#LOAD EPPLUS DLL
			EXCEL_LOAD_DLL
		#EXCEL HEADER DETAILS
			$header = @("GroupName","MemberType","MemberName","Title")
		#BUILD EXCEL OBJECT
			$excelFileObject = EXCEL_BUILD_FILE_OBJECT $global:ExcelFile
		#BUILD EXCEL WORKSHEET WITH TIME STAMP	
			$excelWorkSheet = EXCEL_ADD_WORKSHEET $excelFileObject $global:timeStamp
		#ADD DATA TO EXCEL WORKSHEET
			EXCEL_ADD_DATA_TO_WORKSHEET $excelFileObject $excelWorkSheet $resultsArray
		#FORMAT WORKSHEET
			EXCEL_FORMAT_WORKSHEET_HEADERS $excelFileObject $global:timeStamp $header
	#SEND EMAIL 
		#CONSTRUCT MESSAGE
			$message = '<h3>Metro Active Directory Group Report</h3>'
			$message += '<h3>This report shows groups and its members with job titles.</h3>'
			$message += '<h3>Please find attached excel spreadsheet of it findings.</h3>'
		#SEND EMAIL WITH EXCEL FILE AND MESSAGE
			SEND_EMAIL $global:ExcelFile $message
	#CLEANUP ARCHIVES
		$terminationDate = (Get-Date).AddDays(-7)
		CLEANUP_ARCHIVES $global:ExcelFileDir $terminationDate
}

