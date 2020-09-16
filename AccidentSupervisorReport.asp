
<%
Server.ScriptTimeout=10
On Error Resume Next
Response.CacheControl = "no-cache"
Response.AddHeader "Pragma", "no-cache"
Response.Expires = -1
Response.Buffer = true
headerText = "FLS CPH Customer Report"
%>
<!DOCTYPE html>
<html>
<head>
<meta http-equiv="X-UA-Compatible" content="IE=9; IE=8; IE=7; IE=11" > 
<!--#include file ="../username.asp"-->
<!--#include file ="includes/DataStore.asp"-->
<link rel="stylesheet" type="text/css" href="http://fieldreports.fai.fujitsu.com/tables.css">
<title>Accident Superviosr Report</title>
<%
'=====================================================================================================================================================
AccessLevel = 1 'Anyone can submit an accident report

if SSMAccess > 0 then AccessLevel = 2 end if 'SSMs can see accident reports under them and can approve/revert back accident reports. 
if RDAccess > 0 then AccessLevel = 3 end if 'RDs can see accident reports under them. 
if InStr("LMATHIASON;SUNKERRO",UCASE(TRIM(User_Name))) > 0 then AccessLevel = 4 end if 'FSAdmins can see all accident reports and can provide final approval. Len Mathiason
if InStr("MALAPPAB;MARKSKEV;JHELTON",UCASE(TRIM(User_Name))) > 0 then AccessLevel = 5 end if 'Ops can see everything and edit/approve everything 

Select Case UCASE(TRIM(User_Name))
Case "DLONG"
	AccessLevel = 2
End Select
'=====================================================================================================================================================

'Turn off emails here for testing. 1 equals on, 0 equals off
emailsOn = 1
testing = 1 '1 equals testing on, 0 equals testing off


Set objRSCheckAccess=Server.CreateObject ("ADODB.Recordset")
objRSCheckAccess.Open "  SELECT TOP 1 * from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Network_Username] = '"& User_Name &"' ", ConnectSQL, 1, 3
Oracle_Resource_Name = UCASE(TRIM(objRSCheckAccess("Oracle_Resource_Name")))
objRSCheckAccess.Close
set objRSCheckAccess = Nothing

FSAdmin = "Len Mathiason"

ViewingRecord = 0
AddingRecord = 0
recordID = 0
MainBoxTitle = "Add New Record"
SubmitButtonText = "Submit Record"
Supervisor_Name = ""
InstructionsVisibility = "style='display: none;'"
'MainContentVisibility = "style='display: none;'"
MainContentVisibility = ""

Acciddent_Report_ID = 0
if request.querystring("Acciddent_Report_ID") <> "" AND request.querystring("Acciddent_Report_ID") <> 0 then 
	Set objRSReqFields =Server.CreateObject ("ADODB.Recordset")
	objRSReqFields.Open "select * from [10.159.215.4].[PSP_McDonald].[dbo].[Accident_Supervisor_Report] Where [Acciddent_Report_ID] = '"& INT(request.querystring("Acciddent_Report_ID")) &"'", ConnectSQL, 1, 3
	'if Supervisor report exists, redirecting to exisiting Supervisor Record
	if NOT objRSReqFields.BOF and NOT objRSReqFields.EOF then
		recordID = objRSReqFields("ID")
		objRSReqFields.Close
		set objRSReqFields = Nothing
		response.Redirect ("http://fieldreports.fai.fujitsu.com/opsdev/AccidentSupervisorReport.asp?ID=" & recordID )
	end if 
	objRSReqFields.Close
	set objRSReqFields = Nothing
end if

'Getting accident report id 
if request.querystring("ID") <> "" AND request.querystring("ID") <> 0 then 'Supervisor Report Exists
	Set objRSReqFields =Server.CreateObject ("ADODB.Recordset")
	objRSReqFields.Open "select * from [10.159.215.4].[PSP_McDonald].[dbo].[Accident_Supervisor_Report] Where ID = '"& INT(request.querystring("ID")) &"'", ConnectSQL, 1, 3
	Acciddent_Report_ID = objRSReqFields("Acciddent_Report_ID")
	Supervisor_Name = objRSReqFields("Supervisor_Name")
	objRSReqFields.Close
	set objRSReqFields = Nothing
	ViewingRecord = 1
	MainContentVisibility = ""
	SubmitButtonText = "Submit Updates"
	recordID = INT(request.querystring("ID"))
	MainBoxTitle = "Supervisor's and Vehicle Accident Investigation - Viewing Record ID #" & recordID

	Set objRSRecord=Server.CreateObject ("ADODB.Recordset")
	objRSRecord.Open "select * from [10.159.215.4].[PSP_McDonald].[dbo].[Accident_Supervisor_Report] where ID = '"& recordID &"' ", ConnectSQL

else 'Supervisor Report Does Not Exists
	Acciddent_Report_ID = request.querystring("Acciddent_Report_ID")
	'Getting details from Accident Report 
	Set objRSReqFields =Server.CreateObject ("ADODB.Recordset")
	objRSReqFields.Open "select * from [10.159.215.4].[PSP_McDonald].[dbo].[AccidentTracking] Where ID = '"& INT(Acciddent_Report_ID) &"'", ConnectSQL, 1, 3 ' 'Getting details from Accident Report 
	Accident_Report_Created_By = objRSReqFields("CREATED_BY")
	Driver_Name = objRSReqFields("DRIVER_FUJ_NAME")
	Accident_Date = objRSReqFields("DATE")
	Accident_Time = objRSReqFields("TIME")
	Location = objRSReqFields("LOCATION")
	Location = Location & " | " & objRSReqFields("ACCIDENT_CITY")
	Location = Location & " | " & objRSReqFields("ACCIDENT_STATE")
	objRSReqFields.Close
	set objRSReqFields = Nothing
	Set objSupervisor=Server.CreateObject ("ADODB.Recordset")
	'objSupervisor.Open "select isnull(Supervisor,' ') as Supervisor from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Network_Username] = '"& Accident_Report_Created_By &"' ", ConnectSQL, 1, 3
	objSupervisor.Open "select Oracle_Resource_Name, first_name + ' ' + last_name as Full_Name,Resource_Type from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] a left join (select top 1 Supervisor from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Network_Username] = '"& Accident_Report_Created_By &"' )sub on a.last_name = sub.supervisor where last_name = sub.supervisor", ConnectSQL, 1, 3
	Supervisor_Name = UCASE(TRIM(objSupervisor("Oracle_Resource_Name")))
	Supervisor_Title = objSupervisor("Resource_Type")
	objSupervisor.Close
	Set objSupervisor = Nothing
end if 



' restrict report access. Currently no restrictions applied 
if Supervisor_Name = "" then Supervisor_Name = " " end if
if AccessLevel < 1 then ' and UCASE(TRIM(Supervisor)) <> UCASE(TRIM(LastName))
	response.write "<br><br><br>You do not have the necessary permissions to view this supervisor report. Please contact <a href='mailto:"& SupportEmail &"?Subject=Accident Report Access Request ["& User_Name &"] ["& AccessLevel &"] ["& request.querystring("ID") &"] ' target='_top'>"& SupportName &"</a> for help."
	response.write "<br><br><a href = 'http://fieldreports.fai.fujitsu.com/opsdev/AccidentReporting.asp'>Click here</a> if you choose to be redirected to Accident Reporting home page."    
response.end
end if

%>
<link rel="stylesheet" href="jquery/jquery-ui.css">
<script src="jquery/jquery.js"></script>
<script src="jquery/jquery-ui.js"></script>
<script src="OFL_sorttable.js"></script>


<script type="text/javascript">
    function toggleDiv(divId) {
        $("#" + divId).toggle();
    }

    function validateForm() {

        var x = document.forms["mainForm"]["Supervisor_Name"].value;
        if (x == null || x == " ") {
            alert("The Supervisor's Name field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Driver_Name"].value;
        if (x == null || x == "") {
            alert("The Driver Name field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Accident_Date"].value;
        if (x == null || x == "") {
            alert("The Accident Date field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Accident_Time"].value;
        if (x == null || x == "") {
            alert("The Accident Time field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Location"].value;
        if (x == null || x == "") {
            alert("The Location field must be filled out");
            return false;
        }

    	var x = document.forms["mainForm"]["Cause"].value;
       if (x == null || x == "") {
            alert("The Cause field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Pre_Accident"].value;
        if (x == null || x == "") {
            alert("The Pre-Accident field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Correction"].value;
        if (x == null || x == "") {
            alert("The Correction field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Supervisor_Review"].value;
        if (x == null || x == "") {
            alert("Supervisor's Vehicle Review field must be filled out");
            return false;
        }	

        if (confirm("Please confirm to submit")) { }
        else {return false;}

/*
        var x = document.forms["mainForm"]["Date"].value;
        if (x == null || x == "") {
            alert("The Date field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Assigned_Office"].value;
        if (x == null || x == "") {
            alert("The Assigned Office field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Driver_Department"].value;
        if (x == null || x == "") {
            alert("The Driver Department field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Driver_Training_Type"].value;
        if (x == null || x == "") {
            alert("The Driver Training Type field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Driver_Training_Date"].value;
        if (x == null || x == "") {
            alert("The Driver Training Date field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Driver_Employment_Duration"].value;
        if (x == null || x == "") {
            alert("The Driver Employment Duration field must be filled out");
            return false;
        }

        var x = document.forms["mainForm"]["Describe_Vehicle_Accident"].value;
        if (x == null || x == "") {
            alert("Describe Vehicle Accident field must be filled out");
            return false;
        }
*/
 
    }
</script>


<style type="text/css">
  table {
    table-layout:fixed;
}
.btn {
  background: #3498db;
  background-image: -webkit-linear-gradient(top, #3498db, #2980b9);
  background-image: -moz-linear-gradient(top, #3498db, #2980b9);
  background-image: -ms-linear-gradient(top, #3498db, #2980b9);
  background-image: -o-linear-gradient(top, #3498db, #2980b9);
  background-image: linear-gradient(to bottom, #3498db, #2980b9);
  font-family: Arial;
  color: #ffffff;
  font-size: 20px;
  padding: 10px 20px 10px 20px;
  text-decoration: none;
}

.btn:hover {
  background: #3cb0fd;
  background-image: -webkit-linear-gradient(top, #3cb0fd, #3498db);
  background-image: -moz-linear-gradient(top, #3cb0fd, #3498db);
  background-image: -ms-linear-gradient(top, #3cb0fd, #3498db);
  background-image: -o-linear-gradient(top, #3cb0fd, #3498db);
  background-image: linear-gradient(to bottom, #3cb0fd, #3498db);
  text-decoration: none;
}
</style>
</head>

<body>
<%headerText = "Supervisor's Accident Report&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br>"& WelcomeName &" Access Level "& AccessLevel%>
<!--#include file ="../header.asp"-->
<!--Load the AJAX API-->
<!--#include file ="../Google_Charts_Loader.asp"-->

<%'Set the visibilty for the add record box




strSQL = "SELECT * FROM [10.159.215.4].[PSP_McDonald].[dbo].[Accident_Supervisor_Report] "
Select Case AccessLevel
Case 1 'Driver Access
strSQL = strSQL  & "Where 1=1 AND Driver_Name = '" & UCASE(TRIM(Oracle_Resource_Name)) &"'"
Case 2 'SSM Access Level
strSQL = strSQL  & "Where 1=1 AND Supervisor_Name = '" & UCASE(TRIM(Oracle_Resource_Name)) &"'"
Case 3 'RD Access Level
'strSQL = strSQL  & "Where 1=1 AND CREATED_BY IN (Select Network_Username From [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Region LIKE '%" & UCASE(TRIM(LastName)) &"%')"
strSQL = strSQL  & "Where 1=1 AND Supervisor_Name IN (Select Oracle_Resource_Name From [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Region LIKE '%" & UCASE(TRIM(Oracle_Resource_Name)) &"%')"
Case 4 'FSAdmin Access Level
strSQL = strSQL  & "Where 1=1"
Case 5 'Ops Access Level
strSQL = strSQL  & "Where 1=1"
Case Else
response.write "<br>You do not have the necessary permissions to view this report. Please contact <a href='mailto:"& SupportEmail &"?Subject=Accident Request Access ["& User_Name &"] ["& AccessLevel &"]' target='_top'>"& SupportName &"</a> for help."
response.end
End Select
strSQL = strSQL & " Order By ID DESC"
%>

<%


'Common form fields used by add and update
Sub commonFields()
	objRSUpdate("Supervisor_Name") = request.form("Supervisor_Name")
	objRSUpdate("Assigned_Office") = request.form("Assigned_Office")
	objRSUpdate("Driver_Name") = request.form("Driver_Name")
	objRSUpdate("Driver_Department") = request.form("Driver_Department")
	objRSUpdate("Driver_Training_Type") = request.form("Driver_Training_Type")
	objRSUpdate("Driver_Training_Date") = request.form("Driver_Training_Date")
	objRSUpdate("Driver_Employment_Duration") = request.form("Driver_Employment_Duration")
	objRSUpdate("Accident_Date") = request.form("Accident_Date")
	objRSUpdate("Accident_Time") = request.form("Accident_Time")
	objRSUpdate("Location") = request.form("Location")
	objRSUpdate("Describe_Vehicle_Accident") = request.form("Describe_Vehicle_Accident")
	objRSUpdate("Cause") = request.form("Cause")
	objRSUpdate("Pre_Accident") = request.form("Pre_Accident")
	objRSUpdate("Correction") = request.form("Correction")
	objRSUpdate("Seat_Belt") = request.form("Seat_Belt")
	objRSUpdate("Using_Cell_Phone") = request.form("Using_Cell_Phone")
	objRSUpdate("Hands_Free_Device") = request.form("Hands_Free_Device")
	objRSUpdate("Accident_Report_ID") = request.form("Accident_Report_ID")
	objRSUpdate("Supervisor_Review") = request.form("Supervisor_Review")
	objRSUpdate("Review_Date") = request.form("Review_Date")
	objRSUpdate("Supervisor_Title") = request.form("Supervisor_Title")
	objRSUpdate("Additional_Information_Notes") = request.form("Additional_Information_Notes")
	'objRSUpdate("Review_Committee_Decision") = request.form("Review_Committee_Decision")
	'objRSUpdate("Future_Actions") = request.form("Future_Actions")
	'objRSUpdate("Reviewer1_Name") = request.form("Reviewer1_Name")
	'objRSUpdate("Reviewer2_Name") = request.form("Reviewer2_Name")
	'objRSUpdate("Reviewer3_Name") = request.form("Reviewer3_Name")
	'objRSUpdate("Reviewer1_Title") = request.form("Reviewer1_Title")
	'objRSUpdate("Reviewer2_Title") = request.form("Reviewer2_Title")
	'objRSUpdate("Reviewer3_Title") = request.form("Reviewer3_Title")
	'objRSUpdate("Driver_Notified_In_Writing") = request.form("Driver_Notified_In_Writing")
	'objRSUpdate("Driver_Record_File_Noted") = request.form("Driver_Record_File_Noted")
	'objRSUpdate("Course_Assigned") = request.form("Course_Assigned")

End Sub


if request.form("process") = 1 AND request.form("ID") = 0 then 'adding a new record
	if request.form("Return") = "" then
	ReturnValue = "none"
	else
	ReturnValue = request.form("Return")
	end if
	
	Set objRSUpdate=Server.CreateObject ("ADODB.Recordset")
	objRSUpdate.Open "select * from [PSP_McDonald].[dbo].[Accident_Supervisor_Report]", ConnectSQL_Direct, 1, 3
	objRSUpdate.AddNew
	'To call subroutine
	'commonFields()
	objRSUpdate("Accident_Report_ID") = request.form("Accident_Report_ID")
	objRSUpdate("Supervisor_Name") = request.form("Supervisor_Name")
	objRSUpdate("Assigned_Office") = request.form("Assigned_Office")
	objRSUpdate("Driver_Name") = request.form("Driver_Name")
	objRSUpdate("Driver_Department") = request.form("Driver_Department")
	objRSUpdate("Driver_Training_Type") = request.form("Driver_Training_Type")
	objRSUpdate("Driver_Training_Date") = request.form("Driver_Training_Date")
	objRSUpdate("Driver_Employment_Duration") = request.form("Driver_Employment_Duration")
	objRSUpdate("Accident_Date") = request.form("Accident_Date")
	objRSUpdate("Accident_Time") = request.form("Accident_Time")
	objRSUpdate("Location") = request.form("Location")
	objRSUpdate("Describe_Vehicle_Accident") = request.form("Describe_Vehicle_Accident")
	objRSUpdate("Cause") = request.form("Cause")
	objRSUpdate("Pre_Accident") = request.form("Pre_Accident")
	objRSUpdate("Correction") = request.form("Correction")
	objRSUpdate("Seat_Belt") = request.form("Seat_Belt")
	objRSUpdate("Using_Cell_Phone") = request.form("Using_Cell_Phone")
	objRSUpdate("Hands_Free_Device") = request.form("Hands_Free_Device")
	objRSUpdate("Additional_Information_Notes") = request.form("Additional_Information_Notes")
	objRSUpdate("Supervisor_Review") = request.form("Supervisor_Review")
	objRSUpdate("Review_Date") = request.form("Review_Date")
	objRSUpdate("Supervisor_Title") = request.form("Supervisor_Title")
	objRSUpdate("Updated_By") = Oracle_Resource_Name
	objRSUpdate("Updated_DateTime") = Now()


	
	
	objRSUpdate("Supervisor_Report_Status") = "F.S Team Approval Pending"
	objRSUpdate("Comments_FSAdmin") = ""		
	objRSUpdate("Created_By") = USER_NAME
	objRSUpdate("Acciddent_Report_ID") = request.form("Acciddent_Report_ID")
	objRSUpdate("Created_DateTime") = Now()
	objRSUpdate.Update
	objRSUpdate.Close
	Set objRSUpdate = Nothing

	Set objRSMax=Server.CreateObject ("ADODB.Recordset")
	objRSMax.Open "select MAX(ID) AS ID from [PSP_McDonald].[dbo].[Accident_Supervisor_Report]", ConnectSQL_Direct, 1, 3
	if NOT objRSMax.BOF and NOT objRSMax.EOF then
		'Notify supervisor report creation
		msg = ""
		cc = Email_Address
		recipient = "ronnie.sunker@fujitsu.com;len.mathiason@fujitsu.com;"
		msg = "A new supervisor report ID# <a href='http://fieldreports.fai.fujitsu.com/opsdev/AccidentSupervisorReport.asp?ID="& objRSMax("ID") &"'>"& objRSMax("ID") &"</a> has been created by "& Oracle_Resource_Name & " on "& now() &".<br><br>Please click the link above and review this supervisor report."
		'remove after testing
		msg = msg & "<br><br>Testing Phase notes <br> Send to " & recipient & ". Copy (creator) - " & cc
		subject = "New Supervisor Report #" & objRSMax("ID")			
		sendmail msg,recipient,subject,cc

		response.redirect "AccidentSupervisorReport.asp?p=1&ID=" & objRSMax("ID")
	else
		response.redirect "AccidentSupervisorReport.asp?p=1"
	end if	

end if

if request.form("process") = 1 AND request.form("ID") <> 0 then 'updating an existing record
	if request.form("Return") = "" then
	ReturnValue = "none"
	else
	ReturnValue = request.form("Return")
	end if

	Set objRSUpdate=Server.CreateObject ("ADODB.Recordset")
	objRSUpdate.Open "select * from [PSP_McDonald].[dbo].[Accident_Supervisor_Report] Where ID = '"& INT(request.form("ID")) &"'", ConnectSQL_Direct, 1, 3
	if NOT objRSUpdate.BOF and NOT objRSUpdate.EOF then
	'To call subroutine
	'commonFields()
	objRSUpdate("Accident_Report_ID") = request.form("Accident_Report_ID")
	objRSUpdate("Supervisor_Name") = request.form("Supervisor_Name")
	objRSUpdate("Assigned_Office") = request.form("Assigned_Office")
	objRSUpdate("Driver_Name") = request.form("Driver_Name")
	objRSUpdate("Driver_Department") = request.form("Driver_Department")
	objRSUpdate("Driver_Training_Type") = request.form("Driver_Training_Type")
	objRSUpdate("Driver_Training_Date") = request.form("Driver_Training_Date")
	objRSUpdate("Driver_Employment_Duration") = request.form("Driver_Employment_Duration")
	objRSUpdate("Accident_Date") = request.form("Accident_Date")
	objRSUpdate("Accident_Time") = request.form("Accident_Time")
	objRSUpdate("Location") = request.form("Location")
	objRSUpdate("Describe_Vehicle_Accident") = request.form("Describe_Vehicle_Accident")
	objRSUpdate("Cause") = request.form("Cause")
	objRSUpdate("Pre_Accident") = request.form("Pre_Accident")
	objRSUpdate("Correction") = request.form("Correction")
	objRSUpdate("Seat_Belt") = request.form("Seat_Belt")
	objRSUpdate("Using_Cell_Phone") = request.form("Using_Cell_Phone")
	objRSUpdate("Hands_Free_Device") = request.form("Hands_Free_Device")
	objRSUpdate("Additional_Information_Notes") = request.form("Additional_Information_Notes")
	objRSUpdate("Supervisor_Review") = request.form("Supervisor_Review")
	objRSUpdate("Review_Date") = request.form("Review_Date")
	objRSUpdate("Supervisor_Title") = request.form("Supervisor_Title")
	objRSUpdate("Updated_By") = Oracle_Resource_Name



	if UCASE(TRIM(request.form("Supervisor_Name"))) = UCASE(TRIM(Oracle_Resource_Name)) then
		objRSUpdate("Supervisor_Report_Status") = "F.S Team Approval Pending"
	else
		objRSUpdate("Supervisor_Report_Status") = request.form("Supervisor_Report_Status")
	end if	
	objRSUpdate("Comments_FSAdmin") = request.form("Comments_FSAdmin")	
	objRSUpdate("Review_Committee_Decision") = request.form("Review_Committee_Decision")
	objRSUpdate("Future_Actions") = request.form("Future_Actions")
	objRSUpdate("Reviewer1_Name") = request.form("Reviewer1_Name")
	objRSUpdate("Reviewer2_Name") = request.form("Reviewer2_Name")
	objRSUpdate("Reviewer3_Name") = request.form("Reviewer3_Name")
	objRSUpdate("Reviewer1_Title") = request.form("Reviewer1_Title")
	objRSUpdate("Reviewer2_Title") = request.form("Reviewer2_Title")
	objRSUpdate("Reviewer3_Title") = request.form("Reviewer3_Title")
	objRSUpdate("Driver_Notified_In_Writing") = request.form("Driver_Notified_In_Writing")
	objRSUpdate("Driver_Record_File_Noted") = request.form("Driver_Record_File_Noted")
	objRSUpdate("Course_Assigned1") = request.form("Course_Assigned1")
	objRSUpdate("Course_Assigned2") = request.form("Course_Assigned2")
	objRSUpdate("Course_Assigned3") = request.form("Course_Assigned3")



	objRSUpdate("Updated_DateTime") = Now()
	objRSUpdate.Update
	end if
	objRSUpdate.Close
	Set objRSUpdate = Nothing

	'Sending email
	msg = ""
	cc = ""
	msg = "Accident supervisor report ID# <a href='http://fieldreports.fai.fujitsu.com/opsdev/AccidentSupervisorReport.asp?ID="& INT(request.form("ID")) &"'>"& INT(request.form("ID")) &"</a> has been updated by "& Oracle_Resource_Name & " on "& now() &". "
	msg = msg & "<br><br>Updated Status is - " & request.form("Supervisor_Report_Status")
	recipient = "ronnie.sunker@fujitsu.com;len.mathiason@fujitsu.com;"
	Set objSupervisor=Server.CreateObject ("ADODB.Recordset")
	objSupervisor.Open "select top 1 email_address from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Oracle_Resource_Name] = '"& request.form("Supervisor_Name") & "'", ConnectSQL, 1, 3
	Supervisor_Email = objSupervisor("email_address")
	objSupervisor.Close
	Set objSupervisor = Nothing		
	recipient = recipient & Supervisor_Email
	if request.form("Supervisor_Report_Status") = "Supervisor Report Complete" then 'notify the driver when the supervisor report is complete
		Set objDriver=Server.CreateObject ("ADODB.Recordset")
		objDriver.Open "select top 1 email_address from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Oracle_Resource_Name] = '"& request.form("Driver_Name") & "'", ConnectSQL, 1, 3
		Driver_Email = objDriver("email_address")
		objDriver.Close
		Set objDriver = Nothing	
		recipient = recipient & ";" & Driver_Email
	end if
	subject = "Accident Supervisor Report Updated"

	'remove after testing
	msg = msg & "<br><br>Testing Phase notes <br> Send to " & recipient
	sendmail msg,recipient,subject,cc

	response.redirect "AccidentSupervisorReport.asp?p=2&ID=" & request.form("ID") 
end if


'=========================================================================================================================================
'Send email function
Function sendmail(msg,recipient,subject,cc)
	if emailsOn = 1 then
		Dim myMail
		Set myMail=CreateObject("CDO.Message")
		myMail.From = "FNA_field_reports@fujitsu.com"
		
		'Recipient email address
		myMail.to = recipient
		myMail.cc = cc
		myMail.Subject = subject
		myMail.HTMLBody = msg

		'delete after testing	
		myMail.to = "ronnie.sunker@fujitsu.com;len.mathiason@fujitsu.com;"	
		'myMail.to = "abhiraj.malappa@fujitsu.com"
		myMail.cc = ""


		Dim fso
		Set fso = CreateObject("Scripting.FileSystemObject")
		Set cdoConfig = CreateObject("CDO.Configuration")  
		With cdoConfig.Fields 
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserver") = "intr.fnanic.fujitsu.com"
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
			.Item("http://schemas.microsoft.com/cdo/configuration/sendusing") = "2"
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpconnectiontimeout") = 60
			.Item("http://schemas.microsoft.com/cdo/configuration/smtpauthenticate") = "1" 
			.Update
		End With
		myMail.Configuration = cdoConfig
		myMail.Send
		Set myMail = Nothing
		Set fso=Nothing
	end if
End Function
'=========================================================================================================================================
%>

<%
if request.querystring("p") = 1 then
response.write("<center><font color='green'><b>Record Successfully Created</b></font></center><br>")
end if
if request.querystring("p") = 2 then
response.write("<br><center><font color='green'><b>Record Successfully Updated</b></font></center><br>")
end if
%>


<table style="width: 95%" align="center" cellpadding="0" cellspacing="0">	
<tr>
<td>

	<br>
	<div class="headerboxes" id="addContent" <%=MainContentVisibility%>>
	<link href="images/tabcss.css" rel="stylesheet" type="text/css">
	<div class="shadetabs">
	<ul>
	<li> <a STYLE="text-decoration:none" href="AccidentReporting.asp?ID=<%=Acciddent_Report_ID%>">Accident Report</a></li>
	<li class="selected"><a STYLE="text-decoration:none" href="AccidentSupervisorReport.asp?ID=<%=recordID%>">Supervisor Report </a></li>
	</ul>
	</div>

	<table style="width: 100%;border-collapse: collapse;border-style: solid;border: 2px solid #666666;" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td>
	<div id="info"  align="left"><br /><br />
	<table style="width: 97.5%;border-collapse: collapse;border-style: solid;border: 2px solid #666666;" align="center" cellpadding="0" cellspacing="0">
	<tr>
	<td>
		<table style="width: 97.5%;border-collapse: collapse;border-style: solid;border: 0px solid #666666;" align="center" cellpadding="0" cellspacing="0">
		<tr>
		<td><br>
			<img src="images/info.png" border="0"> THIS FORM IS NOT FOR REPORTING OF CLAIMS, BUT IS FOR DETERMINING VEHICLE ACCIDENT CAUSES SO THAT THEY CAN BE ELIMINATED.
			<br />&nbsp; &nbsp; &nbsp; You may use this FORM IN conjunction WITH THE "Fleet Safety Program Appendix 1 Accident Report Form" provided to the driver.	
			<br /><br />
		</td>	
		</tr>
		</table>
	</td>
	</tr>
	</table>
	</div>
	</td>
	</tr>

	<tr>
	<td>
	<table style="width: 97.5%" align="center">
	
	<form name="mainForm" method="post" action="AccidentSupervisorReport.asp" onsubmit="return validateForm()">
	<tr><td><br />
	<b>Supervisor Name:  
	<input name="Supervisor_Name" type="text" id="Supervisor_Name" class="inputField" size="25" <%if recordID <> 0 then%>value="<%=objRSRecord("Supervisor_Name")%>"<%else%>value="<%=Supervisor_Name%>"<%end if%> >
	 &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
	Assigned Office:
	<input name="Assigned_Office" type="text" id="Assigned_Office" class="inputField" size="25" <%if recordID <> 0 then%>value="<%=objRSRecord("Assigned_Office")%>"<%end if%>  >
	</b>
	<br /><br />DateTime of Report: <b><%if recordID <> 0 then response.write(objRSRecord("Created_DateTime")) else response.write(Now()) end if %></b>&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
	Report Status:<b><%if recordID <> 0 then response.write(objRSRecord("Supervisor_Report_Status")) else response.write("Yet to be submitted") end if %></b>
	</td></tr>

	<tr>
	<td>
	<br/>

		<div class="datagrid">
		<table style="width: 100%">
			<thead>
			<tr>
			<th colspan="7" class="alignLeft">Supervisor's Accident Report</th>
			<th colspan="1" class="alignRight" style='border:none;'><img src="images/asterisk.gif" border="0" title="Required">Required Fields</th>

			</tr>
			</thead>

			<tr><td colspan="8" align="Left"><h3> Vehicle Accident Investigation </h3></td></tr>

			<tr>
			<td colspan="1" rowspan="2" class="alignRight"><b>Who</b></td>
			<td colspan="3" class="alignLeft">Driver Information:</td>
			<td colspan="4" class="alignLeft">Driver Training Received (Type & Date):</td>
			</tr>
			<tr>
			<td colspan="1" class="alignLeft"><img src="images/asterisk.gif" border="0" title="Required">
			Name: 
			<input name="Driver_Name" type="text" id="Driver_Name" class="inputField" style="width: 100%" <%if recordID <> 0 then%>value="<%=objRSRecord("Driver_Name")%>"<%else%>value="<%=Driver_Name%>"<%end if%> >
			</td>
			<td colspan="1" class="alignLeft">
			Dept:
			<input name="Driver_Department" type="text" id="Driver_Department" class="inputField" style="width: 100%"<%if recordID <> 0 then%>value="<%=objRSRecord("Driver_Department")%>"<%end if%>  >
			</td>
			<td colspan="1" class="alignLeft">
			How Long Employed:
			<input name="Driver_Employment_Duration" type="text" id="Driver_Employment_Duration" class="inputField" style="width: 100%"<%if recordID <> 0 then%>value="<%=objRSRecord("Driver_Employment_Duration")%>"<%end if%>  >
			</td>
			<td colspan="2" class="alignLeft">
			Type:<br />
			<input name="Driver_Training_Type" type="text" id="Driver_Training_Type" class="inputField" style="width: 100%"<%if recordID <> 0 then%>value="<%=objRSRecord("Driver_Training_Type")%>"<%end if%>  >
			</td>
			<td colspan="2" class="alignLeft">
			Date:<br />
			<input name="Driver_Training_Date" type="text" id="Driver_Training_Date" class="inputField" style="width: 100%"<%if recordID <> 0 then%>value="<%=objRSRecord("Driver_Training_Date")%>"<%end if%>  >
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><b>When and Where</b></td>
			<td colspan="1" class="alignLeft"><img src="images/asterisk.gif" border="0" title="Required">
			Date: 
			<input name="Accident_Date" type="text" id="Accident_Date" class="inputField" style="width: 100%" <%if recordID <> 0 then%>value="<%=objRSRecord("Accident_Date")%>"<%else%>value="<%=Accident_Date%>"<%end if%> >
			</td>
			<td colspan="1" class="alignLeft"><img src="images/asterisk.gif" border="0" title="Required">
			Time: 
			<input name="Accident_Time" type="text" id="Accident_Time" class="inputField" style="width: 100%" <%if recordID <> 0 then%>value="<%=objRSRecord("Accident_Time")%>"<%else%>value="<%=Accident_Time%>"<%end if%> >
			</td>
			<td colspan="5" class="alignLeft"><img src="images/asterisk.gif" border="0" title="Required">
			Location: <br />
			<input name="Location" type="text" id="Location" class="inputField" style="width: 100%" <%if recordID <> 0 then%>value="<%=objRSRecord("Location")%>"<%else%>value="<%=Location%>"<%end if%> >
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><b>How</b></td>
			<td colspan="7" class="alignLeft"> 
			Describe the Vehicle Accident:<br />
			<textarea name="Describe_Vehicle_Accident" ID="Describe_Vehicle_Accident" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Describe_Vehicle_Accident")%><%end if%></textarea> <br />
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><img src="images/asterisk.gif" border="0" title="Required"><b>Cause</b></td>
			<td colspan="7" class="alignLeft"> 
			What do you believe caused the accident? Describe (1) unsafe driving of others (2) unsafe road conditions (3) faulty vehicle equipment (4) other:<br />
			<textarea name="Cause" ID="Cause" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Cause")%><%end if%></textarea> <br />
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><img src="images/asterisk.gif" border="0" title="Required"><b>Pre-Accident</b></td>
			<td colspan="7" class="alignLeft"> 
			Describe all activities, events, and surroundings for the 10 minutes prior to the accident:<br />
			<textarea name="Pre_Accident" ID="Pre_Accident" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Pre_Accident")%><%end if%></textarea> <br />
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><img src="images/asterisk.gif" border="0" title="Required"><b>Correction</b></td>
			<td colspan="7" class="alignLeft"> 
			What could the driver have reasonably done to prevent this accident? (Consider all aspects of Defensive Driving. Did you, the driver, make no errors; allow for weather and traffic road conditions; allow for enough time?):<br />
			<textarea name="Correction" ID="Correction" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Correction")%><%end if%></textarea> <br />
			</td>
			</tr>

			<tr>
			<td colspan="1" rowspan="3" class="alignRight"><b>Additional Information</b></td>
			<td colspan="7" class="alignLeft">Were the driver and passengers wearing seat belts? &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="Seat_Belt" type="checkbox" id="Seat_Belt" class="inputField" value="Yes" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Seat_Belt") = "Yes" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Yes
            <input name="Seat_Belt" type="checkbox" id="Seat_BeltNo" class="inputField" value="No" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Seat_Belt") = "No" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >No
			</td>
			</tr>

			<tr>
			<td colspan="7" class="alignLeft">Was the driver using a cell phone at time of accident?&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
            <input name="Using_Cell_Phone" type="checkbox" id="Using_Cell_Phone" class="inputField" value="Yes" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Using_Cell_Phone") = "Yes" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Yes
            <input name="Using_Cell_Phone" type="checkbox" id="Using_Cell_PhoneNo" class="inputField" value="No" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Using_Cell_Phone") = "No" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >No
			<br />
			-If yes to above, was a Hands Free Device being used?&nbsp;&nbsp;
            <input name="Hands_Free_Device" type="checkbox" id="Hands_Free_Device" class="inputField" value="Yes" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Hands_Free_Device") = "Yes" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Yes
            <input name="Hands_Free_Device" type="checkbox" id="Hands_Free_DeviceNo" class="inputField" value="No" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Hands_Free_Device") = "No" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >No
			</td>
			</tr>

			<tr>
			<td colspan="7" class="alignLeft">
			What else could be done to prevent similar accidents in the future? (Consider changing routes, allow more time, etc.):<br />
			<textarea name="Additional_Information_Notes" ID="Additional_Information_Notes" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Additional_Information_Notes")%><%end if%></textarea> <br />
			</td>
			</tr>

			<tr><td colspan="8" align="Left"><h3> Supervisor's Vehicle Accident Review</h3></td></tr>

			<tr>
			<td colspan="1"  class="alignRight"><img src="images/asterisk.gif" border="0" title="Required">
			<b>Supervisor's Comments</b></td>
			<td class="alignLeft" colspan="7">
			As the driver's supervisor I have reviewd this accident with the driver involved and have the following comments:	
			<textarea name="Supervisor_Review" ID="Supervisor_Review" rows="3" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Supervisor_Review")%><%end if%></textarea> <br />
			DateTime:
			<input name="Review_Date" type="text" id="Review_Date" class="inputField" size="25"<%if recordID <> 0 then%> value="<%=objRSRecord("Review_Date")%>"<%else%>value="<% response.write(Now())%>"<% end if%>  >
			 &nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp&nbsp
			Title:
			<input name="Supervisor_Title" type="text" id="Supervisor_Title" class="inputField" size="25"<%if recordID <> 0 then%>value="<%=objRSRecord("Supervisor_Title")%>"<%else%>value="<%=Supervisor_Title%>"<% end if%>  >
			</td>
			</tr>


<%if recordID <>0 then 'REVIEW COMMITTEE DECISION
	if objRSRecord("Supervisor_Report_Status") = "F.S Team Approved. Review Committee Decision Pending" or objRSRecord("Supervisor_Report_Status") = "Supervisor Report Complete" then %>
			<tr><td colspan="8" align="left"><h3>Review Committee Decision</h3></td></tr>
			<tr>
			<td colspan="1"  class="alignRight"><b>Decision</b></td>
			<td colspan ="7" class="alignLeft">
			The committee has reviewed this accident in accordance with our vehicle Accident Control Program and has found that it should be judged:<br />
            <input name="Review_Committee_Decision" type="checkbox" id="Review_Committee_Decision" class="inputField" value="Preventable" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Review_Committee_Decision") = "Preventable" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Preventable
            <input name="Review_Committee_Decision" type="checkbox" id="Review_Committee_Decision" class="inputField" value="Non-Preventable" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Review_Committee_Decision") = "Non-Preventable" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Non-Preventable
			<br />
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><b>Comments</b></td>
			<td class="alignLeft" colspan ="7">
			Consideration of the facts indicates the following action should be taken to prevent such accidents in the future:<br />
			<textarea name="Future_Actions" ID="Future_Actions" rows="4" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Future_Actions")%><%end if%></textarea>
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><b>Reviewed By</b></td>
			<td class="alignLeft" colspan ="2">
			<table>
			<tr style="border: none;"><td style="border: none;">
			<select style="width:100%;" name="Reviewer1_Name" id="Reviewer1_Name" class="inputField">
			<option value=" " <%if objRSRecord("Reviewer1_Name") = " " then%>selected<%end if%>></option>
			<%
			Set objRSReviewer = Server.CreateObject ("ADODB.Recordset")
			objRSReviewer.open "select case when resource_type = 'ssm' then (First_Name + ' ' + Last_Name + ' | Regional Field Manager') when resource_type = 'rd' then (First_Name + ' ' + Last_Name + ' | Regional Director') when oracle_resource_name = 'mathiason, len' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Manager') when oracle_resource_name = 'sunker, ronnie' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Coordinator') end as Name from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where ((resource_type = 'ssm' or resource_type = 'rd') or oracle_resource_name = 'mathiason, len' or oracle_resource_name = 'sunker, ronnie')and status = 'active' order by resource_type,name", ConnectSQL
			While NOT objRSReviewer.BOF and NOT objRSReviewer.EOF
			%>
			<option value="<%=objRSReviewer("Name")%>" <%if objRSRecord("Reviewer1_Name") = objRSReviewer("Name") then%>selected <% end if%>><%=objRSReviewer("Name")%></option>
			<%
			objRSReviewer.MoveNext
			Wend
			objRSReviewer.Close
			Set objRSReviewer = Nothing
			%>
			</select>
			</td></tr>
			<tr style="border: none;"><td style="border: none;">
			<select style="width:100%;" name="Reviewer2_Name" id="Reviewer2_Name" class="inputField">
			<option value=" " <%if objRSRecord("Reviewer2_Name") = " " then%>selected<%end if%>></option>
			<%
			Set objRSReviewer = Server.CreateObject ("ADODB.Recordset")
			objRSReviewer.open "select case when resource_type = 'ssm' then (First_Name + ' ' + Last_Name + ' | Regional Field Manager') when resource_type = 'rd' then (First_Name + ' ' + Last_Name + ' | Regional Director') when oracle_resource_name = 'mathiason, len' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Manager') when oracle_resource_name = 'sunker, ronnie' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Coordinator') end as Name from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where ((resource_type = 'ssm' or resource_type = 'rd') or oracle_resource_name = 'mathiason, len' or oracle_resource_name = 'sunker, ronnie')and status = 'active' order by resource_type,name", ConnectSQL
			While NOT objRSReviewer.BOF and NOT objRSReviewer.EOF
			%>
			<option value="<%=objRSReviewer("Name")%>" <%if objRSRecord("Reviewer2_Name") = objRSReviewer("Name") then%>selected <% end if%>><%=objRSReviewer("Name")%></option>
			<%
			objRSReviewer.MoveNext
			Wend
			objRSReviewer.Close
			Set objRSReviewer = Nothing
			%>
			</select>
			</td></tr>
			<tr style="border: none;" ><td style="border: none;">

			<select style="width:100%;" name="Reviewer3_Name" id="Reviewer3_Name" class="inputField">
			<option value=" " <%if objRSRecord("Reviewer3_Name") = " " then%>selected<%end if%>></option>
			<%
			Set objRSReviewer = Server.CreateObject ("ADODB.Recordset")
			objRSReviewer.open "select case when resource_type = 'ssm' then (First_Name + ' ' + Last_Name + ' | Regional Field Manager') when resource_type = 'rd' then (First_Name + ' ' + Last_Name + ' | Regional Director') when oracle_resource_name = 'mathiason, len' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Manager') when oracle_resource_name = 'sunker, ronnie' then (First_Name + ' ' + Last_Name + ' | Fleet Safety Coordinator') end as Name from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where ((resource_type = 'ssm' or resource_type = 'rd') or oracle_resource_name = 'mathiason, len' or oracle_resource_name = 'sunker, ronnie')and status = 'active' order by resource_type,name", ConnectSQL
			While NOT objRSReviewer.BOF and NOT objRSReviewer.EOF
			%>
			<option value="<%=objRSReviewer("Name")%>" <%if objRSRecord("Reviewer3_Name") = objRSReviewer("Name") then%>selected <% end if%>><%=objRSReviewer("Name")%></option>
			<%
			objRSReviewer.MoveNext
			Wend
			objRSReviewer.Close
			Set objRSReviewer = Nothing
			%>
			</select>
			</td></tr>
			</table>
			</td>
			<td class="alignLeft" colspan ="5">
			Date:<br><input name="Committee_Review_Date" type="text" id="Committee_Review_Date" class="inputField" size="29"<%if recordID <> 0 then%>value="<%=objRSRecord("Committee_Review_Date")%>"<%end if%>  >
			</td>
			</tr>

			<tr>
			<td colspan="1"  class="alignRight"><b>Course Assigned</b></td>
			<td class="alignLeft" colspan ="7">
			<table>
			<tr style="border: none;"><td style="border: none;">	
			<select name="Course_Assigned1" id="Course_Assigned1" class="inputField">
			<option value=" " <%if objRSRecord("Course_Assigned1") = " " then%>selected<%end if%>></option>
			<option value="DRV-2.2   Distracted Driving Prevention" <%if objRSRecord("Course_Assigned1") = "DRV-2.2   Distracted Driving Prevention" then%>selected <% end if%>>DRV-2.2   Distracted Driving Prevention</option>
			<option value="DRV-6.2   Hazardous Driving Conditions" <%if objRSRecord("Course_Assigned1") = "DRV-6.2   Hazardous Driving Conditions" then%>selected<%end if%>>DRV-6.2   Hazardous Driving Conditions</option>
			<option value="SNP-60.2  Hazardous Driving Conditions: Driving in Severe Weather" <%if objRSRecord("Course_Assigned1") = "SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather" then%>selected<%end if%>>SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather</option>
			<option value="DRV-4.2   Hazards of Speeding" <%if objRSRecord("Course_Assigned1") = "DRV-4.2   Hazards of Speeding" then%>selected<%end if%>>DRV-4.2   Hazards of Speeding</option>
			<option value="DRV-1.2   Driver Safety" <%if objRSRecord("Course_Assigned1") = "DRV-1.2   Driver Safety" then%>selected<%end if%>>DRV-1.2   Driver Safety</option>
			<option value="SNP-50.2 Driver Safety: Safe and Defensive Driving" <%if objRSRecord("Course_Assigned1") = "SNP-50.2 Driver Safety: Safe and Defensive Driving" then%>selected<%end if%>>SNP-50.2 Driver Safety: Safe and Defensive Driving</option>
			</select>
			</td></tr>
			<tr style="border: none;"><td style="border: none;">
			<select name="Course_Assigned2" id="Course_Assigned2" class="inputField">
			<option value=" " <%if objRSRecord("Course_Assigned2") = " " then%>selected<%end if%>></option>
			<option value="DRV-2.2   Distracted Driving Prevention" <%if objRSRecord("Course_Assigned2") = "DRV-2.2   Distracted Driving Prevention" then%>selected<%end if%>>DRV-2.2   Distracted Driving Prevention</option>
			<option value="DRV-6.2   Hazardous Driving Conditions" <%if objRSRecord("Course_Assigned2") = "DRV-6.2   Hazardous Driving Conditions" then%>selected<%end if%>>DRV-6.2   Hazardous Driving Conditions</option>
			<option value="SNP-60.2  Hazardous Driving Conditions: Driving in Severe Weather" <%if objRSRecord("Course_Assigned2") = "SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather" then%>selected<%end if%>>SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather</option>
			<option value="DRV-4.2   Hazards of Speeding" <%if objRSRecord("Course_Assigned2") = "DRV-4.2   Hazards of Speeding" then%>selected<%end if%>>DRV-4.2   Hazards of Speeding</option>
			<option value="DRV-1.2   Driver Safety" <%if objRSRecord("Course_Assigned2") = "DRV-1.2   Driver Safety" then%>selected<%end if%>>DRV-1.2   Driver Safety</option>
			<option value="SNP-50.2 Driver Safety: Safe and Defensive Driving" <%if objRSRecord("Course_Assigned2") = "SNP-50.2 Driver Safety: Safe and Defensive Driving" then%>selected<%end if%>>SNP-50.2 Driver Safety: Safe and Defensive Driving</option>
			</select>
			</td></tr>
			<tr style="border: none;"><td style="border: none;">
			<select name="Course_Assigned3" id="Course_Assigned3" class="inputField">
			<option value=" " <%if objRSRecord("Course_Assigned3") = " " then%>selected<%end if%>></option>
			<option value="DRV-2.2   Distracted Driving Prevention" <%if objRSRecord("Course_Assigned3") = "DRV-2.2   Distracted Driving Prevention" then%>selected<%end if%>>DRV-2.2   Distracted Driving Prevention</option>
			<option value="DRV-6.2   Hazardous Driving Conditions" <%if objRSRecord("Course_Assigned3") = "DRV-6.2   Hazardous Driving Conditions" then%>selected<%end if%>>DRV-6.2   Hazardous Driving Conditions</option>
			<option value="SNP-60.2  Hazardous Driving Conditions: Driving in Severe Weather" <%if objRSRecord("Course_Assigned3") = "SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather" then%>selected<%end if%>>SNP-60.2 Hazardous Driving Conditions: Driving in Severe Weather</option>
			<option value="DRV-4.2   Hazards of Speeding" <%if objRSRecord("Course_Assigned3") = "DRV-4.2   Hazards of Speeding" then%>selected<%end if%>>DRV-4.2   Hazards of Speeding</option>
			<option value="DRV-1.2   Driver Safety" <%if objRSRecord("Course_Assigned3") = "DRV-1.2   Driver Safety" then%>selected<%end if%>>DRV-1.2   Driver Safety</option>
			<option value="SNP-50.2 Driver Safety: Safe and Defensive Driving" <%if objRSRecord("Course_Assigned3") = "SNP-50.2 Driver Safety: Safe and Defensive Driving" then%>selected<%end if%>>SNP-50.2 Driver Safety: Safe and Defensive Driving</option>
			</select>
			</td></tr>
			</table>
			</td>
			</tr>
			<tr>
			<td colspan="1"  class="alignRight"><b> </b></td>
			<td class="alignLeft" colspan ="7">			
            <input name="Driver_Notified_In_Writing" type="checkbox" id="Driver_Notified_In_Writing" class="inputField" value="Yes" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Driver_Notified_In_Writing") = "Yes" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Driver Notified in Writing &nbsp&nbsp&nbsp&nbsp
            <input name="Driver_Record_File_Noted" type="checkbox" id="Driver_Record_File_Noted" class="inputField" value="Yes" 
                <%
                if  recordID <> 0 then
                    if objRSRecord("Driver_Record_File_Noted") = "Yes" then 
                %> checked 
                <%  end if
                end if 
                %> 
            >Driver Record File Noted 
			</td>
			</tr>


	<%end if
end if %>
		</table>
		</div>
</td>
</tr>


<%if recordID <>0 then  %>
<tr>
<td>
<br />
	<table style="width: 100%" align="center" cellpadding="2" cellspacing="" class="tableborders">
	<tr>
	<td class="alignCenter">

	<div class="datagrid" >
		<table style="cellpadding="2" cellspacing="0" class="tableborders">
			<thead>
			<tr>
			<th class="alignLeft" colspan ="8">Supervisor Report Status (FS Admin access only)  
			</th>
			</tr>
			</thead>
			
			<tr>
			<td class="alignRight" colspan ="1"><b>Status</b></td>
			<td colspan ="7" align="left">
			<%
			'FSAdmin, OPSdev will have ability to update the status
			if AccessLevel > 3 then %> 
				<select name="Supervisor_Report_Status" id="Supervisor_Report_Status" class="inputField">
				<option value="Open for Editing" <%if objRSRecord("Supervisor_Report_Status") = "Open for Editing" then%>selected<%end if%>>Open for Editing</option>
				<option value="F.S Team Approval Pending" <%if objRSRecord("Supervisor_Report_Status") = "F.S Team Approval Pending" then%>selected<%end if%>>F.S Team Approval Pending</option>
				<option value="F.S Team Approved. Review Committee Decision Pending" <%if objRSRecord("Supervisor_Report_Status") = "F.S Team Approved. Review Committee Decision Pending" then%>selected<%end if%>>F.S Team Approved. Review Committee Decision Pending</option>
				<option value="Supervisor Report Complete" <%if objRSRecord("Supervisor_Report_Status") = "Supervisor Report Complete" then%>selected<%end if%>>Supervisor Report Complete</option>
				</select>
			<%else
				response.write(objRSRecord("Supervisor_Report_Status"))
				response.write("<input name='Supervisor_Report_Status' type='hidden' id='Supervisor_Report_Status' value='" & objRSRecord("Supervisor_Report_Status") & "'>")
			end if%>
			</td>
			</tr>

			<tr>
			<td class="alignRight" colspan ="1"><b>Comments (FSAdmin):</b></td>

			<td colspan ="7" align="left"> 			
			<%if AccessLevel > 3  then%>
				<textarea name="Comments_FSAdmin" ID="Comments_FSAdmin" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("Comments_FSAdmin")%><%end if%></textarea>
			<% else
				if objRSRecord("Comments_FSAdmin") = "" then 
					response.write("FSAdmin has not entered comments yet")
				else 
					response.write(objRSRecord("Comments_FSAdmin"))
				end if
				response.write("<input name='Comments_FSAdmin' type='hidden' id='Comments_FSAdmin' value='" & objRSRecord("Comments_FSAdmin") & "'>")
			end if%>
			</td>
			</tr>

			<tr>
			<td class="alignRight" colspan ="1"><b>Last Updated:</b></td>
			<td colspan ="7" align="left">
			Updated by: <%=objRSRecord("Updated_By") %>&nbsp&nbsp&nbsp&nbsp
			Updated DateTime: <%=objRSRecord("Updated_DateTime") %>
			</td>
			</tr>

		</table>
	</div>
	<br />
	</td>
	</tr>
	</table>
</td>
</tr>
<%end if %>

<tr>
<td>
<center>
<%if recordID = 0 then 'supervisor/FSAdmin/OPSdev can submit a new record %> 
	<input type="submit" class="btn" value="<%=SubmitButtonText %>"/>
<%else 'if viewing an existing record
	select case objRSRecord("Supervisor_Report_Status")
	case "Open for Editing" 'if report in open status, supervisor/FSAdmin/OPSdev can edit%>
		<%if UCASE(TRIM(Supervisor_Name)) = UCASE(TRIM(Oracle_Resource_Name)) or AccessLevel > 3 then %>
			<input type="submit" class="btn" value="<%=SubmitButtonText %>"/>
		<%end if%>
	<%case "Supervisor Report Complete"'if report is in complete status, only OPSdev can edit%>
		<%if AccessLevel = 5 then %>
			<input type="submit" class="btn" value="<%=SubmitButtonText %>"/>
		<%end if%>
	<%case else%>
		<%if AccessLevel > 3 then %>
			<input type="submit" class="btn" value="<%=SubmitButtonText %>"/>
		<%end if
	end select
end if
%>
<br /><br />
</center>
</td>
</tr>
<input type="hidden" name="process" value="1">
<input type="hidden" name="ID" value="<%=recordID%>"> 
<input type="hidden" name="Acciddent_Report_ID" value="<%=Acciddent_Report_ID%>"> 
</form>

</table>
</td>
</tr>
</table>

	</div>



	<br>
	<%
	Set objRS=Server.CreateObject("ADODB.Recordset")
	objRS.Open strSQL, ConnectSQL
	%>
	<table style="width: 100%" align="center">
	<tr>
	<td>	
		<div class="datagrid" >
		<table style="width: 100%" class="sortable">

			<thead><tr>
				<th class="alignCenter">ID</th>	
				<th class="alignCenter">Supervisor Report Submission Date</th>				
				<th class="alignCenter">Acciddent Report ID</th>
				<th class="alignCenter">Driver Name</th>
				<th class="alignCenter">Supervisor Name</th>
				<th class="alignCenter">Review Committee Decision</th>
			</tr></thead>
			<%
			if objRS.BOF and objRS.EOF then
			response.write("<tr><td>No Records Found</td></tr>")
			end if
			
			'Loop through the recordset
			While (NOT objRS.BOF) AND (NOT objRS.EOF)
			RecordCounter = RecordCounter + 1
			trcolor = ""
			if RecordCounter Mod 2 = 1 then
			trcolor = "class='alt'"
			end if
			%>
			<tr <%=trcolor%>>
				<td class="alignCenter"><a href="AccidentSupervisorReport.asp?ID=<%=objRS("ID")%>"><%=objRS("ID")%></a> </td>
				<td class="alignCenter"><%=objRS("Created_DateTime")%></td>
				<td class="alignCenter"><a href="AccidentReporting.asp?ID=<%=objRS("Acciddent_Report_ID")%>"><%=objRS("Acciddent_Report_ID")%></a> </td>
				<td class="alignCenter"><%=objRS("Driver_Name")%></td>
				<td class="alignCenter"><%=objRS("Supervisor_Name")%></td>
				<td class="alignCenter"><%=objRS("Review_Committee_Decision")%></td>						
			</tr>
			<%
			objRS.MoveNext
			Wend
			objRS.Close
			Set objRS = Nothing
			%>
		</table>
		</div>				
	
<br>	
<br>For support with this page please contact <a href='mailto:<%=SupportEmail%>?Subject=Install Audit' target='_top'><%=SupportName%></a>.<br><br>

</td>
</tr>
</table>
	
</body>

</html>