# AD_Group_Membership
Project Title => Active Directory Group Membership Report
<br /><strong><u>Programmer</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;Paul Marcantonio
<strong><u>Date</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;September 12 2018
<strong><u>Version</u>:</strong>
     <br/>&nbsp;&nbsp;&nbsp;&nbsp;1.8
<strong><u>Objective</u>:</strong>
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Obtain all groups within Active Directory Sort by Name
	  <br/>&nbsp;&nbsp;&nbsp;&nbsp; Obtain all users within Active Directory sort by Name
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Obtain all members within each Groups (Group and/or Users)
	  <br/>&nbsp;&nbsp;&nbsp;&nbsp; Obtain member attribute (Name, Title)
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Build Excel file containing (GroupName, MemberType, MemberName, Title)
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Place a Copy of the excel file in archivelocation on network share
	  <br/>&nbsp;&nbsp;&nbsp;&nbsp; Send Email with excel as attachment
	  <br/>&nbsp;&nbsp;&nbsp;&nbsp; Remove excel files in archive that are older than 7 days from running of this script
<strong><u>Pre-Condition(s)</u>:</strong>
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Powershell Execution Policy set to Remove Signed (Get-ExecutionPolicy, Set-ExecutionPolicy RemoteSigned)
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; EPPLUS DLL must be downloaded and ready for loading (https://www.nuget.org/packages/EPPlus/)
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Import AD module for powershell commandlets to be available https://docs.microsoft.com/en-us/powershell/module/addsadministration/?view=win10-pss
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Folder structure and right to read and write HTML file to
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; The server this script runs on must be on the "allow relay" on any and all load balancers infront of exchange and within exchange
<strong><u>Post-Condition(s)</u>:</strong>
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; New Excel file created and populated with group and membership infformation
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Email sent to the desired email list with:
          * Copy of Excel file
<strong><u>Installation</u>:</strong>
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Make sure Powershell version 3 or above is installed on the server (Server Role and Features (Windows Powershell)
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; Make available EPPLUS DLL for loading (https://www.nuget.org/packages/EPPlus/)
<strong><u>Contributing</u>:</strong>
      <br/>&nbsp;&nbsp;&nbsp;&nbsp; https://www.nuget.org/packages/EPPlus/
<strong><u>Citations</u>:</strong>
     None
<strong><u>Contact</u>:</strong>
     Paul Marcantonio
          Email => paulmarcantonio@yahoo.com

		
