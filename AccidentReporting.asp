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
<title>Accident Reporting</title>
<%
'=====================================================================================================================================================
AccessLevel = 1 'Anyone can submit an accident report

if SSMAccess > 0 then AccessLevel = 2 end if 'SSMs can see accident reports under them and can approve/revert back accedent reports. 
if RDAccess > 0 then AccessLevel = 3 end if 'RDs can see accident reports under them. 
if InStr("LMATHIASON",UCASE(TRIM(User_Name))) > 0 then AccessLevel = 4 end if 'FSAdmins can see all accident reports and can provide final approval. Len Mathiason
if InStr("MALAPPAB;MARKSKEV;JHELTON",UCASE(TRIM(User_Name))) > 0 then AccessLevel = 5 end if 'Ops can see everything and edit/approve everything 

Select Case UCASE(TRIM(User_Name))
Case "DLONG"
	AccessLevel = 2
End Select
'=====================================================================================================================================================

'check if the user has permission to see an exisiting record 
if request.querystring("ID") <> "" AND request.querystring("ID") <> 0 then 
	Set objRSCheckAccess=Server.CreateObject ("ADODB.Recordset")
	objRSCheckAccess.Open "select * from [PSP_McDonald].[dbo].[AccidentTracking] Where ID = '"& INT(request.querystring("ID")) &"'", ConnectSQL_Direct, 1, 3
	Report_Created_By = objRSCheckAccess("CREATED_BY")
	objRSCheckAccess.Close
	set objRSCheckAccess = Nothing
	Set objSupervisor=Server.CreateObject ("ADODB.Recordset")
	objSupervisor.Open "select isnull(Supervisor,' ') as Supervisor from [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] where [Network_Username] = '"& Report_Created_By &"' ", ConnectOpsDev
	Supervisor = objSupervisor("Supervisor")
	objSupervisor.Close
	Set objSupervisor = Nothing
	if Supervisor = "" then Supervisor = " " end if
	if UCASE(TRIM(User_Name)) <> UCASE(TRIM(Report_Created_By)) and UCASE(TRIM(Supervisor)) <> UCASE(TRIM(LastName)) and AccessLevel < 4 then 
		response.write "<br><br><br>You do not have the necessary permissions to view this accident record. Please contact <a href='mailto:"& SupportEmail &"?Subject=Accident Report Access Request ["& User_Name &"] ["& AccessLevel &"] ["& request.querystring("ID") &"] ' target='_top'>"& SupportName &"</a> for help."
		response.write "<br><br><a href = 'http://fieldreports.fai.fujitsu.com/opsdev/AccidentReporting.asp'>Click here</a> if you choose to be redirected to Accident Reporting home page."    
	response.end
	end if
end if

'Turn off emails here for testing. 1 equals on, 0 equals off
emailsOn = 0
testing = 1 '1 equals testing on, 0 equals testing off

FSAdmin = "Len Mathiason "

%>
<link rel="stylesheet" href="jquery/jquery-ui.css">
<script src="jquery/jquery.js"></script>
<script src="jquery/jquery-ui.js"></script>
<script src="OFL_sorttable.js"></script>


<script type="text/javascript">
function toggleDiv(divId) {
   $("#"+divId).toggle();
}

function validateForm() {

	//Validate the Time field
    var x = document.forms["mainForm"]["TIME"].value;
    if (x==null || x=="") {
        alert("The TIME field must be filled out");
        return false;
		}
		
	//Validate the DATE field
    var x = document.forms["mainForm"]["DATE"].value;
    if (x==null || x=="") {
        alert("The DATE field must be filled out");
        return false;
		}
		
	//Validate the Location field
	var x = document.forms["mainForm"]["LOCATION"].value;
    if (x==null || x=="") {
        alert("The LOCATION field must be filled out");
        return false;
		}	

	//Validate the Accident City field
	var x = document.forms["mainForm"]["ACCIDENT_CITY"].value;
    if (x==null || x=="") {
        alert("The Accident City field must be filled out");
        return false;
		}	

	//Validate the Accident State field
	var x = document.forms["mainForm"]["ACCIDENT_STATE"].value;
    if (x==null || x=="") {
        alert("The Accident State field must be filled out");
        return false;
		}	
		
	//Validate the PAVEMENT field		
    var x = document.forms["mainForm"]["PAVEMENT"].value;
    if (x==null || x=="") {
        alert("The PAVEMENT field must be filled out");
        return false;
		}
		
	//Validate the TRAFFIC_CONTROL field		
	var x = document.forms["mainForm"]["TRAFFIC_CONTROL"].value;
    if (x==null || x=="") {
        alert("The TRAFFIC CONTROL field must be filled out");
        return false;
		}
				
	//Validate the WEATHER field		
	var x = document.forms["mainForm"]["WEATHER"].value;
    if (x==null || x=="") {
        alert("The WEATHER field must be filled out");
        return false;
		}

	//Validate the ACCIDENT_DESC field		
	var x = document.forms["mainForm"]["ACCIDENT_DESC"].value;
    if (x==null || x=="") {
        alert("The ACCIDENT DESCRIPTION field must be filled out");
        return false;
		}

	//Validate the ESTIMATE field		
	var x = document.forms["mainForm"]["ESTIMATE"].value;
    if (x==null || x=="") {
        alert("The ESTIMATE field must be filled out");
        return false;
		}
	//Validate the POLICE_REPORT field		
	var x = document.forms["mainForm"]["POLICE_REPORT"].value;
    if (x==null || x=="") {
        alert("The POLICE REPORT field must be filled out");
        return false;
		}
		
	//Validate the POLICE_DEPT field		
	var x = document.forms["mainForm"]["POLICE_DEPT"].value;
    if (x==null || x=="") {
        alert("The POLICE DEPT field must be filled out");
        return false;
		}
		
	//Validate the POLICE_PHONE field		
	var x = document.forms["mainForm"]["POLICE_PHONE"].value;
    if (x==null || x=="") {
        alert("The POLICE PHONE# field must be filled out");
        return false;
		}

	//Validate the CITATION_ISSUED field		
	var x = document.forms["mainForm"]["CITATION_ISSUED"].value;
    if (x==null || x=="") {
        alert("The CITATION field must be filled out");
        return false;
		}

	//Validate the DRIVER_FUJ_NAME field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_NAME"].value;
    if (x==null || x=="") {
        alert("The Fuijitsu Driver field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_ADDRESS field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_ADDRESS"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Address field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_CITY field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_CITY"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver City field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_ST field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_ST"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver State field must be filled out");
        return false;
		}

	//Validate the DRIVER_FUJ_ZIP field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_ZIP"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Zip Code field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_DOB field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_DOB"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Date of Birth field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_LICENSE field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_LICENSE"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver License Number field must be filled out");
        return false;
		}		

		//Validate the DRIVER_FUJ_EMAIL field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_EMAIL"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver E-Mail field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_TELE_DAY field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_TELE_DAY"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Daytime Phone# field must be filled out");
        return false;
		}		

	//Validate the DRIVER_FUJ_TELE_NIGHT field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_TELE_NIGHT"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Nighttime Phone# field must be filled out");
        return false;
		}
		
	//Validate the DRIVER_FUJ_INJURED field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_INJURED"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Driver Injured field must be filled out");
        return false;
		}

	//Validate the VEHICLE_FUJ_YEAR field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_YEAR"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle Year field must be filled out");
        return false;
		}

	//Validate the VEHICLE_FUJ_MAKE field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_MAKE"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle Make field must be filled out");
        return false;
		}

	//Validate the VEHICLE_FUJ_MODEL field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_MODEL"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle Model field must be filled out");
        return false;
		}

	//Validate the VEHICLE_FUJ_VIN field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_VIN"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle VIN field must be filled out");
        return false;
		}

	//Validate the VEHICLE_FUJ_PLATE field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_PLATE"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle Plate Number field must be filled out");
        return false;
		}		

	//Validate the VEHICLE_FUJ_DAMAGE field		
	var x = document.forms["mainForm"]["VEHICLE_FUJ_DAMAGE"].value;
    if (x==null || x=="") {
        alert("The Fujitsu Vehicle Damage field must be filled out");
        return false;
		}

	//Validate the DRIVER_FUJ_CLAIM field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_CLAIM"].value;
    if (x==null || x=="") {
        alert("The Insurance Claim Number field must be filled out");
        return false;
		}
		
		//Validate the DRIVER_FUJ_REPORT field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_REPORT"].value;
    if (x==null || x=="") {
        alert("The Insurance Report Number field must be filled out");
        return false;
		}

	//Validate the DRIVER_FUJ_OFFICE field		
	var x = document.forms["mainForm"]["DRIVER_FUJ_OFFICE"].value;
    if (x==null || x=="") {
        alert("The Office Handling field must be filled out");
        return false;
    }
	<% if AccessLevel = 1 then %>
    if (confirm("Please confirm to submit. You will not be able to make any updates on submit. ")) { }
    else { return false; }	
	<% else %>
    if (confirm("Please confirm to submit")) { }
    else { return false; }	
	<%end if%>
	
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
<%headerText = "Accident Tracking&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <br>"& WelcomeName &" Access Level "& AccessLevel%>
<!--#include file ="../header.asp"-->
<!--Load the AJAX API-->
<!--#include file ="../Google_Charts_Loader.asp"-->

<%'Set the visibilty for the add record box

ViewingRecord = 0
AddingRecord = 0
recordID = 0
MainBoxTitle = "Add New Record"
SubmitButtonText = "Add New Record"
InstructionsVisibility = "style='display: none;'"
MainContentVisibility = "style='display: none;'"

strSQL = "SELECT * FROM [PSP_McDonald].[dbo].[AccidentTracking] "

Select Case AccessLevel
Case 1 'FE Access Level
strSQL = strSQL  & "Where 1=1 AND CREATED_BY = '" & UCASE(TRIM(User_Name)) &"'"
Case 2 'SSM Access Level
'strSQL = strSQL  & "Where 1=1 AND CREATED_BY IN (Select Network_Username From [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Supervisor = '" & UCASE(TRIM(LastName)) &"' OR Supervisor = '" & UCASE(TRIM(Supervisor)) &"')"
strSQL = strSQL  & "Where 1=1 AND CREATED_BY IN (Select Network_Username From [10.159.165.179].[ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Supervisor = '" & UCASE(TRIM(LastName)) &"' OR Supervisor = '" & UCASE(TRIM(Supervisor)) &"')"
Case 3 'RD Access Level
'strSQL = strSQL  & "Where 1=1 AND CREATED_BY IN (Select Network_Username From [ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Region LIKE '%" & UCASE(TRIM(LastName)) &"%')"
strSQL = strSQL  & "Where 1=1 AND CREATED_BY IN (Select Network_Username From [10.159.165.179].[ORACLE_ADHOC].[dbo].[FIELD_RESOURCES] Where Region LIKE '%" & UCASE(TRIM(LastName)) &"%')"
Case 4 'FSAdmin Access Level
strSQL = strSQL  & "Where 1=1"
Case 5 'Ops Access Level
strSQL = strSQL  & "Where 1=1"
Case Else
response.write "<br>You do not have the necessary permissions to view this report. Please contact <a href='mailto:"& SupportEmail &"?Subject=Accident Request Access ["& User_Name &"] ["& AccessLevel &"]' target='_top'>"& SupportName &"</a> for help."
response.end
End Select

strSQL = strSQL & " Order By ID DESC"

Set objRS=Server.CreateObject("ADODB.Recordset")
objRS.Open strSQL, ConnectSQL_Direct
%>

<%
if request.querystring("ID") <> "" AND request.querystring("ID") <> 0 then
	ViewingRecord = 1
	MainContentVisibility = ""
	SubmitButtonText = "Update Record"
	
	recordID = INT(request.querystring("ID"))
	Set objRSRecord=Server.CreateObject ("ADODB.Recordset")
	objRSRecord.Open "select * from [PSP_McDonald].[dbo].[AccidentTracking] where ID = '"& recordID &"' ", ConnectSQL_Direct
	if objRSRecord.BOF and objRSRecord.EOF then
	recordID = 0
	else
	MainBoxTitle = "Accident Details - Viewing Record ID #" & objRSRecord("ID")
	end if
end if

'Common form fields used by add and update
Sub commonFields()

	if request.form("TIME") = "" then
		response.write("The TIME field was left blank.")
		response.end
	end if
	if request.form("LOCATION") = "" then
		response.write("The LOCATION field was left blank.")
		response.end
	end if
	if request.form("ACCIDENT_CITY") = "" then
		response.write("The ACCIDENT_CITY field was left blank.")
		response.end
	end if
	if request.form("ACCIDENT_STATE") = "" then
		response.write("The ACCIDENT_STATE field was left blank.")
		response.end
	end if	
	if request.form("DATE") = "" then
		response.write("The DATE field was left blank.")
		response.end
	end if
	if request.form("PAVEMENT") = "" then
		response.write("The PAVEMENT field was left blank.")
		response.end
	end if
	if request.form("WEATHER") = "" then
		response.write("The WEATHER field was left blank.")
		response.end
	end if
	if request.form("TRAFFIC_CONTROL") = "" then
		response.write("The TRAFFIC CONTROL field was left blank.")
		response.end
	end if
	if request.form("ACCIDENT_DESC") = "" then
		response.write("The ACCIDENT DESCRIPTION field was left blank.")
		response.end
	end if
	if request.form("ESTIMATE") = "" then
		response.write("The ESTIMATE field was left blank.")
		response.end
	end if
	if request.form("POLICE_REPORT") = "" then
		response.write("The POLICE REPORT field was left blank.")
		response.end
	end if
	if request.form("POLICE_DEPT") = "" then
		response.write("The POLICE DEPT field was left blank.")
		response.end
	end if
	if request.form("POLICE_PHONE") = "" then
		response.write("The POLICE PHONE# field was left blank.")
		response.end
	end if
	if request.form("CITATION_ISSUED") = "" then
		response.write("The CITATION ISSUED field was left blank.")
		response.end
	end if	
	if request.form("DRIVER_FUJ_NAME") = "" then
		response.write("The Fuijitsu Driver field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_ADDRESS") = "" then
		response.write("The Fuijitsu Driver Address field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_CITY") = "" then
		response.write("The Fuijitsu Driver City field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_ST") = "" then
		response.write("The Fuijitsu Driver State field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_ZIP") = "" then
		response.write("The Fuijitsu Driver Zip Code field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_TELE_DAY") = "" then
		response.write("The Fuijitsu Driver Daytime Phone# field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_TELE_NIGHT") = "" then
		response.write("The Fuijitsu Driver Nighttime Phone# field was left blank.")
		response.end
	end if	
	if request.form("DRIVER_FUJ_EMAIL") = "" then
		response.write("The Fuijitsu Driver E-Mail field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_LICENSE") = "" then
		response.write("The Fuijitsu Driver'd License# field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_DOB") = "" then
		response.write("The Fuijitsu Driver Date of Birth field was left blank.")
		response.end
	end if	
	if request.form("DRIVER_FUJ_INJURED") = "" then
		response.write("The Fuijitsu Driver Injured field was left blank.")
		response.end
	end if	
	if request.form("VEHICLE_FUJ_YEAR") = "" then
		response.write("The Fujitsu Vehicle Year field was left blank.")
		response.end
	end if	
	if request.form("VEHICLE_FUJ_MAKE") = "" then
		response.write("The Fujitsu Vehicle Make field was left blank.")
		response.end
	end if
	if request.form("VEHICLE_FUJ_MODEL") = "" then
		response.write("The Fujitsu Vehicle Model field was left blank.")
		response.end
	end if
	if request.form("VEHICLE_FUJ_VIN") = "" then
		response.write("The Fujitsu Vehicle VIN field was left blank.")
		response.end
	end if
	if request.form("VEHICLE_FUJ_PLATE") = "" then
		response.write("The Fujitsu Vehicle Plate Number field was left blank.")
		response.end
	end if
	if request.form("VEHICLE_FUJ_DAMAGE") = "" then
		response.write("The Fujitsu Vehicle Damage field was left blank.")
		response.end
	end if	
	if request.form("DRIVER_FUJ_REPORT") = "" then
		response.write("The Insurance Report Number field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_CLAIM") = "" then
		response.write("The Insurance Claim Number field was left blank.")
		response.end
	end if
	if request.form("DRIVER_FUJ_OFFICE") = "" then
		response.write("The Insurance Office field was left blank.")
		response.end
	end if
	
	objRSUpdate("DATE") = request.form("DATE")
	objRSUpdate("TIME") = request.form("TIME")
	objRSUpdate("LOCATION") = request.form("LOCATION")
	objRSUpdate("ACCIDENT_CITY") = request.form("ACCIDENT_CITY")
	objRSUpdate("ACCIDENT_STATE") = request.form("ACCIDENT_STATE")
	objRSUpdate("PAVEMENT") = request.form("PAVEMENT")
	objRSUpdate("WEATHER") = request.form("WEATHER")
	objRSUpdate("TRAFFIC_CONTROL") = request.form("TRAFFIC_CONTROL")
	objRSUpdate("ACCIDENT_DESC") = request.form("ACCIDENT_DESC")
	objRSUpdate("VEHICLE_FUJ_DAMAGE") = request.form("VEHICLE_FUJ_DAMAGE")
	objRSUpdate("ESTIMATE") = request.form("ESTIMATE")
	objRSUpdate("POLICE_REPORT") = request.form("POLICE_REPORT")
	objRSUpdate("POLICE_DEPT") = request.form("POLICE_DEPT")
	objRSUpdate("POLICE_PHONE") = request.form("POLICE_PHONE")
	objRSUpdate("CITATION_ISSUED") = request.form("CITATION_ISSUED")
	objRSUpdate("DRIVER_FUJ_NAME") = request.form("DRIVER_FUJ_NAME")
	objRSUpdate("DRIVER_FUJ_ADDRESS") = request.form("DRIVER_FUJ_ADDRESS")
	objRSUpdate("DRIVER_FUJ_CITY") = request.form("DRIVER_FUJ_CITY")
	objRSUpdate("DRIVER_FUJ_ST") = request.form("DRIVER_FUJ_ST")
	objRSUpdate("DRIVER_FUJ_ZIP") = request.form("DRIVER_FUJ_ZIP")
	objRSUpdate("DRIVER_FUJ_TELE_DAY") = request.form("DRIVER_FUJ_TELE_DAY")
	objRSUpdate("DRIVER_FUJ_TELE_NIGHT") = request.form("DRIVER_FUJ_TELE_NIGHT")
	objRSUpdate("DRIVER_FUJ_EMAIL") = request.form("DRIVER_FUJ_EMAIL")
	objRSUpdate("DRIVER_FUJ_LICENSE") = request.form("DRIVER_FUJ_LICENSE")
	objRSUpdate("DRIVER_FUJ_DOB") = request.form("DRIVER_FUJ_DOB")
	objRSUpdate("VEHICLE_FUJ_YEAR") = request.form("VEHICLE_FUJ_YEAR")
	objRSUpdate("VEHICLE_FUJ_MAKE") = request.form("VEHICLE_FUJ_MAKE")
	objRSUpdate("VEHICLE_FUJ_MODEL") = request.form("VEHICLE_FUJ_MODEL")
	objRSUpdate("VEHICLE_FUJ_VIN") = request.form("VEHICLE_FUJ_VIN")
	objRSUpdate("VEHICLE_FUJ_PLATE") = request.form("VEHICLE_FUJ_PLATE")
	objRSUpdate("DRIVER_FUJ_REPORT") = request.form("DRIVER_FUJ_REPORT")
	objRSUpdate("DRIVER_FUJ_CLAIM") = request.form("DRIVER_FUJ_CLAIM")
	objRSUpdate("DRIVER_FUJ_OFFICE") = request.form("DRIVER_FUJ_OFFICE")
	objRSUpdate("DRIVER_FUJ_INJURED") = request.form("DRIVER_FUJ_INJURED")
	objRSUpdate("DRIVER_OTHER_NAME") = request.form("DRIVER_OTHER_NAME")
	objRSUpdate("DRIVER_OTHER_ADDRESS") = request.form("DRIVER_OTHER_ADDRESS")
	objRSUpdate("DRIVER_OTHER_CITY") = request.form("DRIVER_OTHER_CITY")
	objRSUpdate("DRIVER_OTHER_ST") = request.form("DRIVER_OTHER_ST")
	objRSUpdate("DRIVER_OTHER_PHONE") = request.form("DRIVER_OTHER_PHONE")
	objRSUpdate("DRIVER_OTHER_LICENSE") = request.form("DRIVER_OTHER_LICENSE")
	objRSUpdate("VEHICLE_OTHER_OWNER_NAME") = request.form("VEHICLE_OTHER_OWNER_NAME")
	objRSUpdate("VEHICLE_OTHER_OWNER_ADDRESS") = request.form("VEHICLE_OTHER_OWNER_ADDRESS")
	objRSUpdate("VEHICLE_OTHER_OWNER_CITY") = request.form("VEHICLE_OTHER_OWNER_CITY")
	objRSUpdate("VEHICLE_OTHER_OWNER_ST") = request.form("VEHICLE_OTHER_OWNER_ST")
	objRSUpdate("VEHICLE_OTHER_OWNER_PHONE") = request.form("VEHICLE_OTHER_OWNER_PHONE")
	objRSUpdate("VEHICLE_OTHER_PLATE") = request.form("VEHICLE_OTHER_PLATE")
	objRSUpdate("VEHICLE_OTHER_YEAR") = request.form("VEHICLE_OTHER_YEAR")
	objRSUpdate("VEHICLE_OTHER_MAKE") = request.form("VEHICLE_OTHER_MAKE")
	objRSUpdate("VEHICLE_OTHER_MODEL") = request.form("VEHICLE_OTHER_MODEL")
	objRSUpdate("VEHICLE_OTHER_INSURANCE") = request.form("VEHICLE_OTHER_INSURANCE")
	objRSUpdate("VEHICLE_OTHER_INS_POLICY") = request.form("VEHICLE_OTHER_INS_POLICY")
	objRSUpdate("VEHICLE_OTHER_INS_PHONE") = request.form("VEHICLE_OTHER_INS_PHONE")
	objRSUpdate("VEHICLE_OTHER_DAMAGE") = request.form("VEHICLE_OTHER_DAMAGE")
	objRSUpdate("PROPERTY_DAMAGE") = request.form("PROPERTY_DAMAGE")
	objRSUpdate("PROPERTY_OWNER") = request.form("PROPERTY_OWNER")
	objRSUpdate("PROPERTY_ADDRESS") = request.form("PROPERTY_ADDRESS")
	objRSUpdate("PROPERTY_CITY") = request.form("PROPERTY_CITY")
	objRSUpdate("PROPERTY_ST") = request.form("PROPERTY_ST")
	objRSUpdate("PROPERTY_PHONE") = request.form("PROPERTY_PHONE")
	objRSUpdate("WITNESS1_NAME") = request.form("WITNESS1_NAME")
	objRSUpdate("WITNESS1_ADDRESS") = request.form("WITNESS1_ADDRESS")
	objRSUpdate("WITNESS1_CITY") = request.form("WITNESS1_CITY")
	objRSUpdate("WITNESS1_ST") = request.form("WITNESS1_ST")
	objRSUpdate("WITNESS1_TYPE") = request.form("WITNESS1_TYPE")
	objRSUpdate("WITNESS1_LOCATION") = request.form("WITNESS1_LOCATION")
	objRSUpdate("WITNESS1_INJURED") = request.form("WITNESS1_INJURED")
	objRSUpdate("WITNESS2_NAME") = request.form("WITNESS2_NAME")
	objRSUpdate("WITNESS2_ADDRESS") = request.form("WITNESS2_ADDRESS")
	objRSUpdate("WITNESS2_CITY") = request.form("WITNESS2_CITY")
	objRSUpdate("WITNESS2_ST") = request.form("WITNESS2_ST")
	objRSUpdate("WITNESS2_TYPE") = request.form("WITNESS2_TYPE")
	objRSUpdate("WITNESS2_LOCATION") = request.form("WITNESS2_LOCATION")
	objRSUpdate("WITNESS2_INJURED") = request.form("WITNESS2_INJURED")
	objRSUpdate("WITNESS3_NAME") = request.form("WITNESS3_NAME")
	objRSUpdate("WITNESS3_ADDRESS") = request.form("WITNESS3_ADDRESS")
	objRSUpdate("WITNESS3_CITY") = request.form("WITNESS3_CITY")
	objRSUpdate("WITNESS3_ST") = request.form("WITNESS3_ST")
	objRSUpdate("WITNESS3_TYPE") = request.form("WITNESS3_TYPE")
	objRSUpdate("WITNESS3_LOCATION") = request.form("WITNESS3_LOCATION")
	objRSUpdate("WITNESS3_INJURED") = request.form("WITNESS3_INJURED")
	objRSUpdate("Accident_report_Status") = request.form("Accident_report_Status")
	if request.form("Accident_report_Status") = "Supervisor Review Complete. F.S. Admin Review Pending" then
		objRSUpdate("SUPERVISOR_REPORT") = "Open"
	end if
	objRSUpdate("COMMENTS_Supervisor") = request.form("COMMENTS_Supervisor")
	objRSUpdate("COMMENTS_FSAdmin") = request.form("COMMENTS_FSAdmin")

End Sub


if request.form("process") = 1 AND request.form("ID") = 0 then 'adding a new record
	if request.form("Return") = "" then
	ReturnValue = "none"
	else
	ReturnValue = request.form("Return")
	end if
	
	Set objRSUpdate=Server.CreateObject ("ADODB.Recordset")
	objRSUpdate.Open "select * from [PSP_McDonald].[dbo].[AccidentTracking]", ConnectSQL_Direct, 1, 3
	objRSUpdate.AddNew
	'To call subroutine
	commonFields()
	objRSUpdate("CREATED_BY") = USER_NAME
	objRSUpdate("CREATED_BY_RESOURCE_ID") = Oracle_Resource_ID
	objRSUpdate("CREATED_DATE") = Now()

	objRSUpdate.Update
	objRSUpdate.Close
	Set objRSUpdate = Nothing
	
	Set objRSUpdate=Server.CreateObject ("ADODB.Recordset")
	objRSUpdate.Open "select MAX(ID) AS ID from [PSP_McDonald].[dbo].[AccidentTracking]", ConnectSQL_Direct, 1, 3
	if NOT objRSUpdate.BOF and NOT objRSUpdate.EOF then
		'Notify the user of the new accident report
		ToRcpt = ToRcpt & Email_Address &"; "
		
		if AccessLevel < 3 then 'Only generate manager emails when an FE or SSM submits the accident report
			'Notify the Manager of the new accident report
			Set objRSEmail=Server.CreateObject ("ADODB.Recordset")
			objRSEmail.Open "SELECT Email_Address,Supervisor FROM FIELD_RESOURCES Where Last_Name = '"& Supervisor &"' AND Resource_Type IN('SSM','FS','RD') AND Status IN('ACTIVE','MLOA')", ConnectSQL, 1, 3
			if NOT objRSEmail.BOF and NOT objRSEmail.EOF then
				RDName = objRSEmail("Supervisor") 'this is the RD's name
				ToRcpt = ToRcpt & objRSEmail("Email_Address") &"; "
			end if
			objRSEmail.Close
			Set objRSEmail = Nothing

			'Notify the RD of the new accident report
			if RDName <> Supervisor then 'Don't send the email if the RD was emailed above
				Set objRSEmail=Server.CreateObject ("ADODB.Recordset")
				objRSEmail.Open "SELECT Email_Address FROM FIELD_RESOURCES Where Last_Name = '"& RDName &"' AND Resource_Type IN('RD') AND Status IN('ACTIVE','MLOA')", ConnectSQL, 1, 3
				if NOT objRSEmail.BOF and NOT objRSEmail.EOF then
					ToRcpt = ToRcpt & objRSEmail("Email_Address") &"; "
				end if
				objRSEmail.Close
				Set objRSEmail = Nothing
			end if
		end if

		if ToRcpt <> "" then 'There are people to be notified
			msg = "A new accident report ID# <a href='http://fieldreports.fai.fujitsu.com/opsdev/AccidentReporting.asp?ID="& objRSUpdate("ID") &"'>"& objRSUpdate("ID") &"</a> has been created by "& User_Name &" on "& now() &".<br><br><b>Please click the link above and review this accident report.</b>"
			subject = "New Accident Report #" & objRSUpdate("ID")
			recipient = ToRcpt
			
			sendmail msg,recipient,subject
		else
			msg = "A new accident report ID# <a href='http://fieldreports.fai.fujitsu.com/opsdev/AccidentReporting.asp?ID="& objRSUpdate("ID") &"'>"& objRSUpdate("ID") &"</a> has been created by "& User_Name &" on "& now() &".<br><br><b>Please click the link above and review this accident report.</b><br><br>No other notifications were sent out for this new accident report because the user's manager could not be found."
			subject = "New Accident Report #" & objRSUpdate("ID") &" (No Notifications)"
			recipient = "CIBUSA@fujitsu.com"
		
			sendmail msg,recipient,subject
		end if
	
		response.redirect "AccidentReporting.asp?p=1&ID=" & objRSUpdate("ID") 
	else
		response.redirect "AccidentReporting.asp?p=1"
	end if
end if

if request.form("process") = 1 AND request.form("ID") <> 0 then 'updating an existing record
	if request.form("Return") = "" then
	ReturnValue = "none"
	else
	ReturnValue = request.form("Return")
	end if

	Set objRSUpdate=Server.CreateObject ("ADODB.Recordset")
	objRSUpdate.Open "select * from [PSP_McDonald].[dbo].[AccidentTracking] Where ID = '"& INT(request.form("ID")) &"'", ConnectSQL_Direct, 1, 3
	if NOT objRSUpdate.BOF and NOT objRSUpdate.EOF then
	'To call subroutine
	commonFields()
	objRSUpdate("UPDATED_BY") = USER_NAME
	objRSUpdate("UPDATED_DATE") = Now()
	if request.form("Accident_report_Status") = "Open For Editing" then
		objRSUpdate("Accident_report_Status") = "Supervisor Review Pending"
	end if
	objRSUpdate.Update
	end if
	objRSUpdate.Close
	Set objRSUpdate = Nothing
		
	response.redirect "AccidentReporting.asp?p=2&ID=" & request.form("ID")
end if

'=========================================================================================================================================
'Send email function
Function sendmail(msg,recipient,subject)
	if emailsOn = 1 then
		Dim myMail
		Set myMail=CreateObject("CDO.Message")
		myMail.From = "FNA_field_reports@fujitsu.com"
		
		'Recipient email address
		myMail.to = recipient
		
		'Add cc recipients from the distribution list
		cc = ""
		Set objRSEmail=Server.CreateObject ("ADODB.Recordset")
		objRSEmail.Open "SELECT [EMAIL_ADDRESS] FROM [AUDIT].[dbo].[DISTRIBUTION] where [ITEM_NAME] = 'Accident_Tracking' AND Active = 'Y'", ConnectSQL
		if NOT objRSEmail.BOF and NOT objRSEmail.EOF then
			cc = cc & objRSEmail("Email_Address")
		end if
		objRSEmail.Close
		Set objRSEmail = Nothing
		if cc <> "" then
		myMail.cc = cc
		end if
		
		'Add CIBUSA as BCC
		if recipient <> "CIBUSA@fujitsu.com" then
			myMail.bcc = "CIBUSA@fujitsu.com"
		end if

		myMail.Subject = subject
		myMail.HTMLBody = msg
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
response.write("<br><center><font color='green'><b>Task Successfully Updated</b></font></center><br>")
end if
%>

<script type="text/javascript">
function toggleDiv(divId) {
   $("#"+divId).slideToggle();
   
}

function showonlyone(thechosenone) {
     $('.headerboxes').each(function(index) {
          if ($(this).attr("id") == thechosenone) {
               $(this).slideToggle();
          }
          else {
               $(this).hide(0);
          }
     });
}

$(function() {
$('#WEATHER').change( function() {
          var value = $(this).val();
          if (!value || value == '') {
             var other = prompt( "Please indicate other value:" );
             if (!other) return false;
             $(this).append('<option value="'
                               + other
                               + '" selected="selected">'
                               + other
                               + '</option>');
          }
      });
	  
$('#TRAFFIC_CONTROL').change( function() {
          var value = $(this).val();
          if (!value || value == '') {
             var other = prompt( "Please indicate other value:" );
             if (!other) return false;
             $(this).append('<option value="'
                               + other
                               + '" selected="selected">'
                               + other
                               + '</option>');
          }
      });	  
	  
$('#PAVEMENT').change( function() {
          var value = $(this).val();
          if (!value || value == '') {
             var other = prompt( "Please indicate other value:" );
             if (!other) return false;
             $(this).append('<option value="'
                               + other
                               + '" selected="selected">'
                               + other
                               + '</option>');
          }
      });	  
	  
});
</script>


<table style="width: 95%" align="center" cellpadding="0" cellspacing="0">	
<tr>
<td>

	<table cellpadding="0" cellspacing="0" style="width: 50%" align="center">
		<tr>
			<td class="alignCenter">
			<input type="button" class="btn" value="Accident Report Instructions" onclick="javascript:showonlyone('instructionsContent');" />
			</td>	
		</tr>
	</table>
	<br>
	
	
	
	<div class="headerboxes" id="instructionsContent" <%=InstructionsVisibility%>>
		<div class="datagrid">
			<table style="width: 100%">
			<thead><tr>
			<th class="alignLeft">Instructions</th>
			</tr></thead>
			<tr>
			<td class="alignCenter">
				<table style="width: 100%">
				<tr>
				<td width="50%" valign="top">
				<h2>Fujitsu America Inc.<br>
				Insured Vehicles</h2><br>
				<b><i>Policy Number:<br>
				All States - #CA 640290401</i></b>
				<br>
				<br>
				<br>
				<b>Retain original copy and photos for your records.</b><p>
				<b>Once you submit your claim, information cannot be changed. Contact your supervisor if there is an issue with any data you submit. You may be asked to resubmit the form with all of the correct information filled out.</b><p>
				
				</td>
				<td width="50%" valign="top">
				<b>IF YOU HAVE AN ACCIDENT</b><br>
				<ul>
				<li>Remain Calm</li>
				<li>Do not argue or admit liability.</li>
				<li>Gather the facts outlined on the Accident Report Form</li>
				<li>Take photos of both vehicles and the scene of the accident.</li>
				<li>Fill out the online Accident Report within 24 hours and attach* a copy of the Accident Report Form and photos.</li>
				<li>Call Tokyo Marine Reporting at (800) 948-6546 to report the loss.</li>
				<li>Call Mike Jeziorski, Fleet Administrator  (813) 505-1634</li>
				<li>Report all accidents to the local police department, even if they were not at the scene</li>
				<li>Secure a local police report number.</li>
				<li>Do not give written or recorded statements to anyone other than a Police Officer.</li>
				</ul>
				<b>* Files may be attached to the Accident Report once the initial report has been saved</b>
				</td>
				</tr>
				<tr>
				<td colspan="3" align="Center">
				<a href="attachments/Accident_Report_Form.zip">Accident Report Form (MS Word/PDF Version)</a>
				</td>
				</tr>
				</table>
			<br>
			<input type="button" class="btn" value="Proceed to Create Accident Report" onclick="javascript:showonlyone('addContent');" />
			<br>
			<br>
			</td>
			</tr>
			</table>
		</div>	
	</div>

	<div class="headerboxes" id="addContent" <%=MainContentVisibility%>>
	<link href="images/tabcss.css" rel="stylesheet" type="text/css">
	<div class="shadetabs">
	<ul>
	<li class="selected"> <a STYLE="text-decoration:none" href="AccidentReporting.asp?ID=<%=recordID%>">Accident Report</a></li>
	<%if recordID <> 0 and objRSRecord("SUPERVISOR_REPORT") = "Open" and AccessLevel > 1 then %>
	<li><a STYLE="text-decoration:none" href="AccidentSupervisorReport.asp?Acciddent_Report_ID=<%=recordID%>">Supervisor Report </a></li>
	<% end if %>
	</ul>
	</div>
	<table style="width: 100%;border-collapse: collapse;border-style: solid;border: 2px solid #666666;" align="center" cellpadding="0" cellspacing="0">


	<tr>
	<td>
	<table style="width: 97.5%" align="center">
	<tr>
	<td>
	<br/>

		<form name="mainForm" method="post" action="AccidentReporting.asp" onsubmit="return validateForm()">
		<div class="datagrid">
		<table style="width: 100%">
			<thead><tr>
			<%if recordID = 0 then%>
			<th colspan="3" class="alignLeft"><%=MainBoxTitle%></th>
			<%else%>
			<th colspan="3" class="alignLeft">
				<table>
				<th class="alignLeft"><%=MainBoxTitle%></th>
				<th class="alignRight">
				<%
				Set objAttachment=Server.CreateObject ("ADODB.Recordset")
				objAttachment.Open "select * from [AccidentTrackingAttachments] where AccidentID = '"& recordID &"' ", ConnectOpsDev
				if NOT objAttachment.BOF and NOT objAttachment.EOF then
				response.write("<a href='AccidentTracking_attachments.asp?ID="& recordID &"'><img src='images/attachment_on_bright.png' border='0' title='Click to view/add attachments'></a>")
				else
				response.write("<a href='AccidentTracking_attachments.asp?ID="& recordID &"'><img src='images/attachment.png' border='0' title='Click to add attachments'></a>")
				end if
				objAttachment.Close
				Set objAttachment = Nothing
				%>
			    <a href="AccidentTracking_attachments.asp?ID=<%=recordID%>">Attachments</a></th>
				</table>
			</th>
			<%end if%>
			</tr></thead>

			<tr><td colspan="3" align="Center"><img src="images/asterisk.gif" border="0" title="Required">Required Fields</td></tr>

			<tr>
				<td width="33%" valign="top">
				
				<table border="0" width="100%">
				<tr>
				<td colspan="4" align="Center"><h3>The Accident</h3></td>
				</tr>
				<tr>
				<td><img src="images/asterisk.gif" border="0" title="Required">Date:</td>
				<td>
				<script>
				$(function() {
					$( "#datepicker1" ).datepicker({
					changeMonth: true,
					changeYear: true,
					dateFormat: "m/d/yy"
					});
				});
				</script>
		
				<%
				if recordID <> 0 then
				theDate = objRSRecord("DATE")
				else
				theDate = date()
				end if
				%>
				
				<input name="DATE" type="text" style="width: 100%" id="datepicker1" value="<%=theDate%>" class="inputField">
				</td>
				<td><img src="images/asterisk.gif" border="0" title="Required">Time:</td>
				<td><input type="text" name="TIME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("TIME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Location:</td>
				<td colspan="3">
				<textarea rows="2" style="width:100%" name="LOCATION" ID="LOCATION" class="inputField"><%if recordID <> 0 then%><%=objRSRecord("LOCATION")%><%end if%></textarea>
				</td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">City:</td>
				<td colspan="1"><input type="text" name="ACCIDENT_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("ACCIDENT_CITY")%>"<%end if%>></td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">State:</td>
				<td colspan="1"><input type="text" name="ACCIDENT_STATE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("ACCIDENT_STATE")%>"<%end if%>></td>
				</tr>				
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Pavement Condition:</td>
				<td colspan="1"><select name="PAVEMENT" id="PAVEMENT" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("PAVEMENT") = "" then%>selected<%end if%>></option>
				<%end if%>
				
				<%
				inList = 0
				if recordID <> 0 then
					If objRSRecord("PAVEMENT") = "Dry" OR objRSRecord("PAVEMENT") = "Wet" OR objRSRecord("PAVEMENT") = "Ice/Snow" OR objRSRecord("PAVEMENT") = "Dirt/Sand" then inlist = 1 end if 
				end if
				%>				
				<option value="Dry"<%if recordID <> 0 then%><%if objRSRecord("PAVEMENT") = "Dry" then%>selected<%end if%><%end if%>>Dry</option>
				<option value="Wet"<%if recordID <> 0 then%><%if objRSRecord("PAVEMENT") = "Wet" then%>selected<%end if%><%end if%>>Wet</option>
				<option value="Ice/Snow"<%if recordID <> 0 then%><%if objRSRecord("PAVEMENT") = "Ice/Snow" then%>selected<%end if%><%end if%>>Ice/Snow</option>
				<option value="Dirt/Sand"<%if recordID <> 0 then%><%if objRSRecord("PAVEMENT") = "Dirt/Sand" then%>selected<%end if%><%end if%>>Dirt/Sand</option>	
				<%if recordID <> 0 AND inList = 0 then%>
				<option value="<%=objRSRecord("PAVEMENT")%>" selected><%=objRSRecord("PAVEMENT")%></option>
				<%end if%>
				<option value="">Specify Other</option>					
				</select>
				</td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Traffic Control:</td>
				<td colspan="1">
				<select name="TRAFFIC_CONTROL" id="TRAFFIC_CONTROL" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("TRAFFIC_CONTROL") = "" then%>selected<%end if%>></option>
				<%end if%>
				<%
				inList = 0
				if recordID <> 0 then
					If objRSRecord("TRAFFIC_CONTROL") = "Barriers" OR objRSRecord("TRAFFIC_CONTROL") = "Cones" OR objRSRecord("TRAFFIC_CONTROL") = "Human" OR objRSRecord("TRAFFIC_CONTROL") = "Lights" OR objRSRecord("TRAFFIC_CONTROL") = "Signs" OR objRSRecord("TRAFFIC_CONTROL") = "None" then inlist = 1 end if 
				end if
				%>
				
				<option value="Barriers"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "Barriers" then%>selected<%end if%><%end if%>>Barriers</option>
				<option value="Cones"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "Cones" then%>selected<%end if%><%end if%>>Cones</option>
				<option value="Human"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "Human" then%>selected<%end if%><%end if%>>Human</option>				
				<option value="Lights"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "Lights" then%>selected<%end if%><%end if%>>Lights</option>
				<option value="Signs"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "Signs" then%>selected<%end if%><%end if%>>Signs</option>
				<option value="None"<%if recordID <> 0 then%><%if objRSRecord("TRAFFIC_CONTROL") = "None" then%>selected<%end if%><%end if%>>None</option>
				<%if recordID <> 0 AND inList = 0 then%>
				<option value="<%=objRSRecord("TRAFFIC_CONTROL")%>" selected><%=objRSRecord("TRAFFIC_CONTROL")%></option>
				<%end if%>
				<option value="">Specify Other</option>				
				</select>
				</td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Weather:</td>
				<td colspan="3">
				<select name="WEATHER" id="WEATHER" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WEATHER") = "" then%>selected<%end if%>></option>
				<%end if%>
				
				<%
				inList = 0
				if recordID <> 0 then
					If objRSRecord("WEATHER") = "Clear" OR objRSRecord("WEATHER") = "Dry" OR objRSRecord("WEATHER") = "Hail" OR objRSRecord("WEATHER") = "Ice" OR objRSRecord("WEATHER") = "Rain" OR objRSRecord("WEATHER") = "Snow" OR objRSRecord("WEATHER") = "Wind" OR objRSRecord("WEATHER") = "Sleet" then inlist = 1 end if 
				end if
				%>
				<option value="Clear"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Clear" then%>selected<%end if%><%end if%>>Clear</option>
				<option value="Dry"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Dry" then%>selected<%end if%><%end if%>>Dry</option>
				<option value="Hail"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Hail" then%>selected<%end if%><%end if%>>Hail</option>
				<option value="Ice"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Ice" then%>selected<%end if%><%end if%>>Ice</option>				
				<option value="Rain"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Rain" then%>selected<%end if%><%end if%>>Rain</option>
				<option value="Sleet"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Sleet" then%>selected<%end if%><%end if%>>Sleet</option>
				<option value="Snow"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Snow" then%>selected<%end if%><%end if%>>Snow</option>
				<option value="Wind"<%if recordID <> 0 then%><%if objRSRecord("WEATHER") = "Wind" then%>selected<%end if%><%end if%>>Wind</option>
				<%if recordID <> 0 AND inList = 0 then%>
				<option value="<%=objRSRecord("WEATHER")%>" selected><%=objRSRecord("WEATHER")%></option>
				<%end if%>
				<option value="">Specify Other</option>
				</select>
				</td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Description of Accident:</td>
				<td colspan="3"><textarea rows="4" style="width:100%" name="ACCIDENT_DESC" ID="ACCIDENT_DESC" class="inputField"><%if recordID <> 0 then%><%=objRSRecord("ACCIDENT_DESC")%><%end if%></textarea></td>
				</tr>

				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Estimate:</td>
				<td colspan="3">
				<select name="ESTIMATE" id="ESTIMATE" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("ESTIMATE") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("ESTIMATE") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("ESTIMATE") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Police Report #:</td>
				<td colspan="3"><input type="text" name="POLICE_REPORT" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("POLICE_REPORT")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Police Dept:</td>
				<td colspan="3"><input type="text" name="POLICE_DEPT" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("POLICE_DEPT")%>"<%end if%>>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Police Phone #:</td>
				<td colspan="3"><input type="text" name="POLICE_PHONE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("POLICE_PHONE")%>"<%end if%>>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Citation Issued:</td> 
				<td colspan="3">
				<select name="CITATION_ISSUED" id="CITATION_ISSUED" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("CITATION_ISSUED") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("CITATION_ISSUED") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("CITATION_ISSUED") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>
				</tr>
				<tr>
				<td colspan="4" align="center">
				<h3>Your Information</h3>
				</td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Name:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Address:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">City:</td>
				<td colspan="1"><input type="text" name="DRIVER_FUJ_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_CITY")%>"<%end if%>></td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">State:</td>
				<td colspan="1"><input type="text" name="DRIVER_FUJ_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_ST")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Zip Code:</td>
				<td colspan="1"><input type="text" name="DRIVER_FUJ_ZIP" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_ZIP")%>"<%end if%>></td>
				<td colspan="2"></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Date of Birth:</td>
				<td colspan="3">
				<script>
				$(function() {
					$( "#datepicker2" ).datepicker({
					changeMonth: true,
					changeYear: true,
					yearRange: "-100:+0",
					dateFormat: "m/d/yy"
					});
				});
				</script>
		
				<%
				if recordID <> 0 then
				theDate = objRSRecord("DRIVER_FUJ_DOB")
				else
				theDate = date()
				end if
				%>
				<input name="DRIVER_FUJ_DOB" type="text" style="width: 100%" id="datepicker2" value="<%=theDate%>" class="inputField"></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Driver's License #:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_LICENSE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_LICENSE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">E-Mail:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_EMAIL" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_EMAIL")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="4" align="center"><b>Phone</b></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Day</td>
				<td colspan="1"><input type="text" name="DRIVER_FUJ_TELE_DAY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_TELE_DAY")%>"<%end if%>></td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Evening</td>
				<td colspan="1"><input type="text" name="DRIVER_FUJ_TELE_NIGHT" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_TELE_NIGHT")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Injured? </td>
				<td colspan="3">
				<select name="DRIVER_FUJ_INJURED" id="DRIVER_FUJ_INJURED" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("DRIVER_FUJ_INJURED") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("DRIVER_FUJ_INJURED") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("DRIVER_FUJ_INJURED") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>
				</tr>				
				</table>
				</td>

				<td width="33%" valign="top">
				<table border="0" width="100%">
				<tr>
				<td colspan="4" align="Center"><h3>Fujitsu Vehicle Information</h3></td>
				</tr>

				<tr>				
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Year:</td>
				<td colspan="1"><input type="text" name="VEHICLE_FUJ_YEAR" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_FUJ_YEAR")%>"<%end if%>></td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Make:</td>
				<td colspan="1"><input type="text" name="VEHICLE_FUJ_MAKE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_FUJ_MAKE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Model:</td>
				<td colspan="1"><input type="text" name="VEHICLE_FUJ_MODEL" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_FUJ_MODEL")%>"<%end if%>></td>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">VIN:</td>
				<td colspan="1"><input type="text" name="VEHICLE_FUJ_VIN" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_FUJ_VIN")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">License Plate #:</td>
				<td colspan="3"><input type="text" name="VEHICLE_FUJ_PLATE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_FUJ_PLATE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Damage to your Vehicle:</td>
				<td colspan="3"><textarea rows="4" style="width:100%" name="VEHICLE_FUJ_DAMAGE" ID="VEHICLE_FUJ_DAMAGE" class="inputField"><%if recordID <> 0 then%><%=objRSRecord("VEHICLE_FUJ_DAMAGE")%><%end if%></textarea></td>
				</tr>
				<tr>
				<td colspan="4" align="Center"><h3>Other Driver Information</h3></td>
				</tr>
				<tr>
				<td colspan="1">Name:</td>
				<td colspan="3"><input type="text" name="DRIVER_OTHER_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="DRIVER_OTHER_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="DRIVER_OTHER_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="DRIVER_OTHER_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_ST")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Phone:</td>
				<td colspan="3"><input type="text" name="DRIVER_OTHER_PHONE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_PHONE")%>"<%end if%>></td>				
				</tr>
				<tr>
				<td colspan="1">Driver's License #:</td>
				<td colspan="3"><input type="text" name="DRIVER_OTHER_LICENSE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_OTHER_LICENSE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="4" align="Center"><h3>Other Vehicle Owner Information</h3></td>
				</tr>
				<tr>
				<td colspan="1">Name:</td>
				<td colspan="3"><input type="text" name="VEHICLE_OTHER_OWNER_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_OWNER_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="VEHICLE_OTHER_OWNER_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_OWNER_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_OWNER_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_OWNER_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_OWNER_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_OWNER_ST")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Phone:</td>
				<td colspan="3"><input type="text" name="VEHICLE_OTHER_OWNER_PHONE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_OWNER_PHONE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">License Plate #:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_PLATE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_PLATE")%>"<%end if%>></td>
				<td colspan="1">Year:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_YEAR" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_YEAR")%>"<%end if%>></td>				
				</tr>
				<tr>				
				<td colspan="1">Make:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_MAKE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_MAKE")%>"<%end if%>></td>
				<td colspan="1">Model:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_MODEL" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_MODEL")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="2">Insurance Company:</td>
				<td colspan="2"><input type="text" name="VEHICLE_OTHER_INSURANCE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_INSURANCE")%>"<%end if%>></td>
				</tr>				
				<tr>				
				<td colspan="1">Policy #:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_INS_POLICY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_INS_POLICY")%>"<%end if%>></td>
				<td colspan="1">Phone #:</td>
				<td colspan="1"><input type="text" name="VEHICLE_OTHER_INS_PHONE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("VEHICLE_OTHER_INS_PHONE")%>"<%end if%>></td>
				</tr>				
				<tr>
				<td colspan="1">Damage to other Vehicle:</td>
				<td colspan="3"><textarea rows="2" style="width:100%" name="VEHICLE_OTHER_DAMAGE" ID="VEHICLE_OTHER_DAMAGE" class="inputField"><%if recordID <> 0 then%><%=objRSRecord("VEHICLE_OTHER_DAMAGE")%><%end if%></textarea></td>
				</tr>
				</table>
				</td>

					
				<td width="33%" valign="top">
				<table border="0" width="100%">

				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Claim #:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_CLAIM" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_CLAIM")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Report #:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_REPORT" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_REPORT")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1"><img src="images/asterisk.gif" border="0" title="Required">Office Handling:</td>
				<td colspan="3"><input type="text" name="DRIVER_FUJ_OFFICE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("DRIVER_FUJ_OFFICE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="4" align="Center"><h3>Other Damage</h3></td>
				</tr>
				<tr>
				<td colspan="1">Damage to Property (Non-Vehicle):</td>
				<td colspan="3"><textarea rows="2" style="width:100%" name="PROPERTY_DAMAGE" ID="PROPERTY_DAMAGE" class="inputField"><%if recordID <> 0 then%><%=objRSRecord("PROPERTY_DAMAGE")%><%end if%></textarea></td>
				</tr>	
				<tr>
				<td colspan="1">Property Owner:</td>
				<td colspan="3"><input type="text" name="PROPERTY_OWNER" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("PROPERTY_OWNER")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="PROPERTY_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("PROPERTY_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="PROPERTY_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("PROPERTY_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="PROPERTY_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("PROPERTY_ST")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Phone:</td>
				<td colspan="3"><input type="text" name="PROPERTY_PHONE" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("PROPERTY_PHONE")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="4" align="Center"><h3>Passengers & Witnesses</h3></td>
				</tr>
				<tr>
				<td colspan="1">Name:</td>
				<td colspan="3"><input type="text" name="WITNESS1_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS1_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="WITNESS1_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS1_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="WITNESS1_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS1_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="WITNESS1_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS1_ST")%>"<%end if%>></td>
				</tr>				
				<tr>
				<td colspan="1">Type:</td> 
				<td colspan="1">
				<select name="WITNESS1_TYPE" id="WITNESS1_TYPE" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS1_TYPE") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Passenger"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_TYPE") = "Passenger" then%>selected<%end if%><%end if%>>Passenger</option>
				<option value="Witness"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_TYPE") = "Witness" then%>selected<%end if%><%end if%>>Witness</option>
				</select>
				</td>
				<td colspan="1">Location:</td> 
				<td colspan="1">
				<select name="WITNESS1_LOCATION" id="WITNESS1_LOCATION" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS1_LOCATION") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Your Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_LOCATION") = "Your Vehicle" then%>selected<%end if%><%end if%>>Your Vehicle</option>
				<option value="Other Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_LOCATION") = "Other Vehicle" then%>selected<%end if%><%end if%>>Other Vehicle</option>
				</select>
				</td>
				</tr>				
				<tr>
				<td colspan="1">Injured? </td> 
				<td colspan="1">
				<select name="WITNESS1_INJURED" id="WITNESS1_INJURED" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS1_INJURED") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_INJURED") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("WITNESS1_INJURED") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>				
				</tr>	
				<tr>
				<td colspan="1">Name:</td>
				<td colspan="3"><input type="text" name="WITNESS2_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS2_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="WITNESS2_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS2_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="WITNESS2_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS2_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="WITNESS2_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS2_ST")%>"<%end if%>></td>
				</tr>				
				<tr>
				<td colspan="1">Type:</td> 
				<td colspan="1">
				<select name="WITNESS2_TYPE" id="WITNESS2_TYPE" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS2_TYPE") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Passenger"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_TYPE") = "Passenger" then%>selected<%end if%><%end if%>>Passenger</option>
				<option value="Witness"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_TYPE") = "Witness" then%>selected<%end if%><%end if%>>Witness</option>
				</select>
				</td>
				<td colspan="1">Location:</td> 
				<td colspan="1">
				<select name="WITNESS2_LOCATION" id="WITNESS2_LOCATION" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS2_LOCATION") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Your Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_LOCATION") = "Your Vehicle" then%>selected<%end if%><%end if%>>Your Vehicle</option>
				<option value="Other Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_LOCATION") = "Other Vehicle" then%>selected<%end if%><%end if%>>Other Vehicle</option>
				</select>
				</td>
				</tr>				
				<tr>
				<td colspan="1">Injured? </td> 
				<td colspan="1">
				<select name="WITNESS2_INJURED" id="WITNESS2_INJURED" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS2_INJURED") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_INJURED") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("WITNESS2_INJURED") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>				
				</tr>	
				<tr>
				<td colspan="1">Name:</td>
				<td colspan="3"><input type="text" name="WITNESS3_NAME" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS3_NAME")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">Address:</td>
				<td colspan="3"><input type="text" name="WITNESS3_ADDRESS" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS3_ADDRESS")%>"<%end if%>></td>
				</tr>
				<tr>
				<td colspan="1">City:</td>
				<td colspan="1"><input type="text" name="WITNESS3_CITY" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS3_CITY")%>"<%end if%>></td>
				<td colspan="1">State:</td>
				<td colspan="1"><input type="text" name="WITNESS3_ST" style="width: 100%" class="inputField" <%if recordID <> 0 then%>value="<%=objRSRecord("WITNESS3_ST")%>"<%end if%>></td>
				</tr>				
				<tr>
				<td colspan="1">Type:</td> 
				<td colspan="1">
				<select name="WITNESS3_TYPE" id="WITNESS3_TYPE" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS3_TYPE") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Passenger"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_TYPE") = "Passenger" then%>selected<%end if%><%end if%>>Passenger</option>
				<option value="Witness"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_TYPE") = "Witness" then%>selected<%end if%><%end if%>>Witness</option>
				</select>
				</td>
				<td colspan="1">Location:</td> 
				<td colspan="1">
				<select name="WITNESS3_LOCATION" id="WITNESS3_LOCATION" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS3_LOCATION") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Your Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_LOCATION") = "Your Vehicle" then%>selected<%end if%><%end if%>>Your Vehicle</option>
				<option value="Other Vehicle"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_LOCATION") = "Other Vehicle" then%>selected<%end if%><%end if%>>Other Vehicle</option>
				</select>
				</td>
				</tr>				
				<tr>
				<td colspan="1">Injured? </td> 
				<td colspan="1">
				<select name="WITNESS3_INJURED" id="WITNESS3_INJURED" style="width: 100%" class="inputField">
				<%if recordID = 0 then%>
				<option value="" selected></option>
				<%end if%>
				<%if recordID <> 0 then%>
				<option value=""<%if objRSRecord("WITNESS3_INJURED") = "" then%>selected<%end if%>></option>
				<%end if%>
				<option value="Yes"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_INJURED") = "Yes" then%>selected<%end if%><%end if%>>Yes</option>
				<option value="No"<%if recordID <> 0 then%><%if objRSRecord("WITNESS3_INJURED") = "No" then%>selected<%end if%><%end if%>>No</option>
				</select>
				</td>				
				</tr>					
				</table>
				</td>
				</tr>
		</table>
		<br>
		<center>
		<%
		if objRSRecord("Accident_Report_Status") <> "Approval Process Complete"  then 
			if recordID <> 0 then 'They are looking at an existing record
				if  AccessLevel >= 5 or ( objRSRecord("Accident_Report_Status") = "Open for Editing" and objRSRecord("CREATED_BY") = UCASE(TRIM(User_Name)) ) then
				%>
				<input type="submit" class="btn" value="Submit Updates"/>
				<%end if%>
			<font style="font-size: 12px;"><a href="AccidentReporting.asp">Clear Form</a></font>
			<%
			else 'Anyone can submit a new accident report
			%>
			<input type="hidden" name="Accident_report_Status" value="Supervisor Review Pending">
			<input type="hidden" name="COMMENTS_Supervisor" value="">
			<input type="hidden" name="COMMENTS_FSAdmin" value="">

			<input type="submit" class="btn" value="Submit For Review"/>
			<%end if%>
			<p><b>Remember to attach any files related to your accident after you have saved your report.</b>
		<%end if%>
		</center>
		</div>

</td>
</tr>


	<%if recordID <> 0 then  %>
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
			<th class="alignLeft" colspan="13">Accident Report Review:  
			</th>
			</tr>
			</thead>
			
			<tr>
			<td class="alignRight" colspan="1"><b>Status</b></td>
			<td colspan="5" align="left"><br />
			<%
			'Supervisor, FSAdmin, OPSdev will have ability to update the status
			if AccessLevel >3 or UCASE(TRIM(Supervisor)) = UCASE(TRIM(LastName)) then %> 
				<select name="Accident_Report_Status" id="Accident_Report_Status" class="inputField">
				<option value="Open for Editing" <%if objRSRecord("Accident_Report_Status") = "Open for Editing" then%>selected<%end if%>>Open for Editing. Driver can update.</option>
				<option value="Supervisor Review Pending" <%if objRSRecord("Accident_Report_Status") = "Supervisor Review Pending" then%>selected<%end if%>>Supervisor Review Pending</option>
				<option value="Supervisor Review Complete. F.S. Admin Review Pending" <%if objRSRecord("Accident_Report_Status") = "Supervisor Review Complete. F.S. Admin Review Pending" then%>selected<%end if%>>Supervisor Review Complete. F.S. Admin Review Pending</option>
					<option value="Approval Process Complete" <%if objRSRecord("Accident_Report_Status") = "Approval Process Complete" then%>selected<%end if%>>Approval Process Complete</option>
			<%else
				response.write(objRSRecord("Accident_Report_Status"))
				response.write("<input name='Accident_Report_Status' type='hidden' id='Accident_Report_Status' value='" & objRSRecord("Accident_Report_Status") & "'>")
			end if%>
			<br /><br />
			</td>

			<td colspan="7" align="left">
			<ul>
			<li>Accident Report is reviewed and approved in two stages.</li>
			<li>FE fills in the accident form and will submit for review. The supervisor/SSM is notified of the new submission. (Status: SSM Review Pending) </li>
			<li>Supervisor/SSM reviews the form and approves it. Accident report moves to the next stage of approval. FSAmin is notified to review the report. (Status: SSM Review Complete. F.S.Admin Review Pending)</li>
			<li>With FSAdmin's approval, the accident reporting is complete. (Status: Approval process complete).</li>
			<li>At either approval stage (supervisor/FSAdmin), the reviewer will have the ability to send back the report for changes. The report will be open for FE to edit and make updates. (Status: Open for editing)</li>
			</ul>
			</td>
			</tr>

			<tr>
			<td class="alignRight" colspan="1"><b>Comments</b></td>
			<td colspan="5" align="left"> 
			<b>Supervisor: <%=Supervisor%></b><br>
			Comments:
			<%if AccessLevel = 5 or UCASE(TRIM(Supervisor)) = UCASE(TRIM(LastName)) then %>
				<textarea name="COMMENTS_Supervisor" ID="COMMENTS_Supervisor" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("COMMENTS_Supervisor")%><%end if%></textarea>
			<% else
				if objRSRecord("COMMENTS_Supervisor") = "" then 
					response.write("Supervisor has not entered comments yet")
				else 
					response.write(objRSRecord("COMMENTS_Supervisor"))
				end if
				response.write("<input name='COMMENTS_Supervisor' type='hidden' id='COMMENTS_Supervisor' value='" & objRSRecord("COMMENTS_Supervisor") & "'>")
			end if%>
			</td>

			<td colspan="7" align="left"> 
			<b>FSAdmin: <%=FSAdmin%></b><br>
			Comments:
			<%if AccessLevel > 3  then%>
				<textarea name="COMMENTS_FSAdmin" ID="COMMENTS_FSAdmin" rows="2" style="width: 100%" class="inputField" ><%if recordID <> 0 then%><%=objRSRecord("COMMENTS_FSAdmin")%><%end if%></textarea>
			<% else
				if objRSRecord("COMMENTS_FSAdmin") = "" then 
					response.write("FSAdmin has not entered comments yet")
				else 
					response.write(objRSRecord("COMMENTS_FSAdmin"))
				end if
				response.write("<input name='COMMENTS_FSAdmin' type='hidden' id='COMMENTS_FSAdmin' value='" & objRSRecord("COMMENTS_FSAdmin") & "'>")
			end if%>
			</td>
			</tr>
			
			<%if AccessLevel = 5 or ( (AccessLevel = 4 or UCASE(TRIM(Supervisor)) = UCASE(TRIM(LastName)) ) and objRSRecord("Accident_Report_Status") <> "Approval Process Complete") then %>
			<tr>
			<td colspan="1"> </td>
			<td colspan="12">
				<%if UCASE(TRIM(Supervisor)) = UCASE(TRIM(LastName)) then %>
					<% if objRSRecord("Accident_Report_Status") <> "Supervisor Review Complete. F.S. Admin Review Pending" and objRSRecord("Accident_Report_Status") <> "Approval Process Complete" then %>
						<input type="submit" value="Submit" >
					<% end if %>
				<%else %>
					<input type="submit" value="Submit" >
				<%end if %>
			</td>
			</tr>
			<% end if %>
		</table>
	</div>
	<br />
	</td>
	</tr>
	</table>

	</td>
	</tr>
	<%end if %>
	<input type="hidden" name="process" value="1">
	<input type="hidden" name="ID" value="<%=recordID%>">
	</form>

</table>
</td>
</tr>
</table>

	</div>



	<br>
	
	<table style="width: 100%" align="center">
	<tr>
	<td>
	
	<%
	'if OPSAccess > 0 OR RDAccess > 0 OR SSMAccess > 0 then
	%>

		<div class="datagrid" >
		<table style="width: 100%" class="sortable">

			<thead><tr>
				<th class="alignCenter">ID</th>	
				<th class="alignCenter">Accident Date</th>
				<th class="alignCenter">Reported Submission Date</th>				
				<th class="alignCenter">Driver</th>
				<th class="alignCenter">Claim #</th>
				<th class="alignCenter">Attachments</th>				
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
				<td class="alignCenter"><a href="AccidentReporting.asp?ID=<%=objRS("ID")%>"><%=objRS("ID")%></a> </td>
				<td class="alignCenter"><%=objRS("DATE")%></td>
				<td class="alignCenter"><%=objRS("CREATED_DATE")%> PST</td>
				<td class="alignCenter"><%=objRS("DRIVER_FUJ_NAME")%></td>
				<td class="alignCenter"><%=objRS("DRIVER_FUJ_CLAIM")%></td>		
				<td class="alignCenter">
				<%
				Set objAttachment=Server.CreateObject ("ADODB.Recordset")
				objAttachment.Open "select * from [AccidentTrackingAttachments] where AccidentID = '"& objRS("ID") &"' ", ConnectOpsDev
				if NOT objAttachment.BOF and NOT objAttachment.EOF then
				response.write("<a href='AccidentTracking_attachments.asp?ID="& objRS("ID") &"'><img src='images/attachment_on.png' border='0' title='Click to view/add attachments'></a>")
				else
				response.write("<a href='AccidentTracking_attachments.asp?ID="& objRS("ID") &"'><img src='images/attachment.png' border='0' title='Click to add attachments'></a>")
				end if
				objAttachment.Close
				Set objAttachment = Nothing
				%>
				</td>					
			</tr>
			<%
			objRS.MoveNext
			Wend
			objRS.Close
			Set objRS = Nothing
			%>
		</table>
		</div>
		
	<%
	'end if
	%>
	
	
	
	<%
	if AccessLevel >= 3 then
	
	if request("MetricView") = "" then
	MetricView = "Accidents"
	else
	MetricView = request("MetricView")
	end if
	%>

	<br> <br>
	<img src="images/toggle.png"> <a href="javascript:toggleDiv('metricsDiv');" id="metrics">Toggle Metrics</a>
	<br>
	<div id="metricsDiv" <%if request("MetricView") = "" then%>style="display: none;"<%end if%>>
	<br><br>
<div class="datatable" >
<table style="width: 100%" align="center">
<thead><tr>
	<th align="center">
	<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
	<script type="text/javascript">
	google.charts.load('current', {packages: ['corechart', 'bar']});
	google.charts.setOnLoadCallback(drawBasic);

	function drawBasic() {

      var data = google.visualization.arrayToDataTable([
		['<%=chartXName%>', '<%=chartYName%>'],
	  <%
		tempStr = ""
		sqlQuery = "Select TOP 3 Count([ID]) AS Totals, [Weather] FROM [OpsDev].[dbo].[AccidentTracking] Where DATEDIFF(month,[DATE], GETDATE()) <= 12 Group by [Weather] Order by [Totals] DESC"

		Set objRSBar = Server.CreateObject("ADODB.Recordset")
		objRSBar.open sqlQuery, ConnectOpsDev

		While NOT objRSBar.BOF and NOT objRSBar.EOF
		tempStr = tempStr & "['" & objRSBar("WEATHER") & "', " & objRSBar("TOTALS") & "],"

		objRSBar.MoveNext
		Wend

		response.write(tempStr)
		
		objRSBar.close
		Set objRSBar = Nothing
	  %>
      ]);

      var options = {
        title: 'Accidents by Weather Type',
        chartArea: {width: '50%'},
        hAxis: {
          title: 'Accidents',
		  minValue: 0
        },
		legend: {position: 'none'},
        vAxis: {
          title: 'Weather'
        }
      };

      var chart = new google.visualization.BarChart(document.getElementById('chart_div'));

      chart.draw(data, options);
    }
	
	</script>
	<div id="chart_div"></div>
	</th>
	<th align="center">
	<script type="text/javascript" src="https://www.gstatic.com/charts/loader.js"></script>
	<script type="text/javascript">
	google.charts.load('current', {packages: ['corechart', 'bar']});
	google.charts.setOnLoadCallback(drawBasic);

	function drawBasic() {

      var data = google.visualization.arrayToDataTable([
		['<%=chartXName%>', '<%=chartYName%>'],
	  <%
		tempStr = ""
		sqlQuery = " Select Count([ID]) AS Totals, [Driver_Resource_Type] FROM [OpsDev].[dbo].[AccidentTracking] Where DATEDIFF(month,[DATE], GETDATE()) <= 12 Group by [Driver_Resource_Type] Order by [Totals] DESC"

		Set objRSBar = Server.CreateObject("ADODB.Recordset")
		objRSBar.open sqlQuery, ConnectOpsDev

		While NOT objRSBar.BOF and NOT objRSBar.EOF
		tempStr = tempStr & "['" & objRSBar("Driver_Resource_Type") & "', " & objRSBar("TOTALS") & "],"

		objRSBar.MoveNext
		Wend

		response.write(tempStr)
		
		objRSBar.close
		Set objRSBar = Nothing
	  %>
      ]);

      var options = {
        title: 'Accidents by Technician Type',
        chartArea: {width: '50%'},
        hAxis: {
          title: 'Accidents',
		  minValue: 0
        },
		legend: {position: 'none'},
        vAxis: {
          title: 'Type'
        }
      };

      var chart = new google.visualization.BarChart(document.getElementById('chart_div1'));

      chart.draw(data, options);
    }
	
	</script>
	<div id="chart_div1"></div>	
	</th>
	</tr>
</table>
</div>	
	
	
		<br>
		<br>
		View Metrics By <select name="Action" onchange="window.open(this.options[this.selectedIndex].value,'_top')" class="inputField">
			<option value="AccidentTracking.asp?MetricView=SR_Count#metrics" <%if MetricView = "SR_Count" then%>selected<%end if%>>SR Count</option>
			<option value="AccidentTracking.asp?MetricView=Travel_Miles#metrics" <%if MetricView = "Travel_Miles" then%>selected<%end if%>>Travel Miles</option>
			<option value="AccidentTracking.asp?MetricView=Travel_Miles_avg#metrics" <%if MetricView = "Travel_Miles_avg" then%>selected<%end if%>>Travel Miles Avg</option>
			<option value="AccidentTracking.asp?MetricView=Accidents#metrics" <%if MetricView = "Accidents" then%>selected<%end if%>>Accident Counts</option>
			</select>
		<br>
		<div class="datagrid" >
		<table style="width: 100%" class="sortable">
			<thead><tr>
				<th class="alignLeft">Manager</th>
				<th class="alignLeft">Area</th>
				<th class="alignLeft">Region</th>
				<%
				'counter = 0
				Set objMetrics=Server.CreateObject ("ADODB.Recordset")
				objMetrics.Open "select Distinct Date From [OpsDev].[dbo].[AccidentTracking_Metrics] Order By Date ", ConnectOpsDev
				if NOT objMetrics.BOF and NOT objMetrics.EOF then
				FirstMonth = DateDiff("m",objMetrics("Date"),date())
				end if
				
				While Not objMetrics.BOF and NOT objMetrics.EOF
				'counter = counter + 1
				%>
				<th class="alignCenter"><%=Left(MonthName(MONTH(objMetrics("Date"))),3)%>&nbsp;<%=YEAR(objMetrics("Date"))%></th>	
				<%
				if NOT objMetrics.BOF and NOT objMetrics.EOF then
				LastMonth = DateDiff("m",objMetrics("Date"),date())
				end if
				
				objMetrics.MoveNext
				Wend

				objMetrics.Close
				Set objMetrics = Nothing
				%>
				<th class="alignCenter">Total</th>
			</tr></thead>
			<%
			Set objMetrics=Server.CreateObject ("ADODB.Recordset")
			objMetrics.Open "select Distinct SortOrder, Manager, Area, Region, Country From [OpsDev].[dbo].[AccidentTracking_Metrics] Order By SortOrder,Region,Manager", ConnectOpsDev
			While Not objMetrics.BOF and NOT objMetrics.EOF
				RecordCounter = RecordCounter + 1
				trcolor = ""
				if RecordCounter Mod 2 = 1 then
				trcolor = "class='alt'"
				end if
				
				if SortOrder = 1 AND SortOrder <> objMetrics("SortOrder") then
				response.write("<tfoot>")
				end if
				
				if objMetrics("SortOrder") = 2 then
					trcolor="style='background-color: #b1b1b1'"
				end if
				
				if objMetrics("SortOrder") = 3 then
					trcolor="style='background-color: #9c9898'"
				end if
				%>
				<tr <%=trcolor%>>
				<td class="alignLeft"><%=objMetrics("Manager")%></td>
				<td class="alignLeft"><%=objMetrics("Area")%></td>
				<td class="alignLeft"><%=objMetrics("Region")%></td>
					<%
					total = 0
					ycounter = FirstMonth
					While ycounter > LastMonth-1
					
					sql = "select ["& MetricView &"] From [OpsDev].[dbo].[AccidentTracking_Metrics] Where Manager = '"& objMetrics("Manager") &"' AND Date =  DATEADD(MONTH, DATEDIFF(MONTH, 0, GETDATE())-"& ycounter &", 0) "
					Set objValue=Server.CreateObject ("ADODB.Recordset")
					objValue.Open sql, ConnectOpsDev
					total = total + objValue(MetricView)
					%>
					<td class="alignCenter">
					<%
					'response.write(sql)
					response.write(FormatNumber(objValue(MetricView),0))
					%>
					</td>
					<%
					objValue.Close
					Set objValue = Nothing
					
					ycounter = ycounter - 1
					Wend
					
					if MetricView = "Travel_Miles_avg" then
					total = total / 12
					end if
					%>
				<td class="alignCenter" style='background-color: #9c9898 !important'><%=FormatNumber(total,0)%></td>	
				</tr>
				<%
				SortOrder = objMetrics("SortOrder")
			objMetrics.MoveNext
			Wend
			objMetrics.Close
			Set objMetrics = Nothing
			%>
			
			</tfoot>
		</table>
		</div>
		
		<%
		response.write("<br><i>Metrics are compiled monthly for the previous month during the monthly metrics reports run.</i>")
		%>
		
		
	</div>	
	<%end if%>
	
	
<br>	
<br>For support with this page please contact <a href='mailto:<%=SupportEmail%>?Subject=Install Audit' target='_top'><%=SupportName%></a>.<br><br>

</td>
</tr>
</table>
	
</body>

</html>