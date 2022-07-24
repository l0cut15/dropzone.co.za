<html>

<head>
<meta http-equiv="author" content="Brett Samuel khaoz@khaoz.co.za">
<meta http-equiv="content" content="Skydiving in South Africa">
<meta http-equiv="keywords" content="Witbank, Skydive, Skydiving , Club, tandem , AFF, cypress, South Africa, static line, RSL, freefall, dope rope, paracute, Skymaster, Tempo, Dropzone, ZA">
<meta http-equiv="pragma" content="nocache">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
<title>Display Event / Add event</title>
<style TYPE="text/css">
<!--
H1 {font-size:26pt;font-family:  Arial, sans-serif}
H2 {font-size:20pt;font-family:  Arial, sans-serif}
H3 {font-size:16pt;font-family:  Arial, sans-serif}
H4 {font-size:12pt;font-family:  Arial, sans-serif}
H5 {font-size:10pt;font-family:  Arial, sans-serif}
H6 {font-size:8pt;font-family: Arial, sans-serif}
P {font-size:10pt;font-family:  Verdana,Arial, sans-seriff}
TD {font-size:10pt;font-family:  Verdana,Arial, sans-seriff}
A.menu {text-decoration: none}
A:hover {color:yellow}
-->
</style>
</head>

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<% 
 ' =======================================================================================================================
'	S E C T I O N  1    ( D I S P L A Y   T H E   E V E N T )
 ' =======================================================================================================================


 ' What type of request is this 
If Request.Form("RequestType") = "show" then 
%>
<div><div align="center"><center>

<table border="1" cellspacing="0" cellpadding="2" bordercolor="#0080FF">
  <tr>
    <td nowrap colspan="4" bgcolor="#400080"><h4 align="center">Events for <%Response.Write FormatDateTime(Request.Form("EventDate"), 1) %></h4>
    </td>
  </tr>
<%
fp_sQry = "SELECT * FROM Events WHERE (((EventDate)=#" & Request.Form("EventDate") & "#));"
fp_sDefault = ""
fp_sNoRecords = ""
fp_iMaxRecords = 0
fp_iTimeout = 0
fp_iCurrent = 1
fp_fError = False
fp_bBlankField = False
If fp_iTimeout <> 0 Then Server.ScriptTimeout = fp_iTimeout
Do While (Not fp_fError) And (InStr(fp_iCurrent, fp_sQry, "%%") <> 0)
	' found a opening quote, find the close quote
	fp_iStart = InStr(fp_iCurrent, fp_sQry, "%%")
	fp_iEnd = InStr(fp_iStart + 2, fp_sQry, "%%")
	If fp_iEnd = 0 Then
		fp_fError = True
		Response.Write "<B>Database Region Error: mismatched parameter delimiters</B>"
	Else
		fp_sField = Mid(fp_sQry, fp_iStart + 2, fp_iEnd - fp_iStart - 2)
		If Mid(fp_sField,1,1) = "%" Then
			fp_sWildcard = "%"
			fp_sField = Mid(fp_sField, 2)
		Else
			fp_sWildCard = ""
		End If
		fp_sValue = Request.Form(fp_sField)

		' if the named form field doesn't exist, make a note of it
		If (len(fp_sValue) = 0) Then
			fp_iCurrentField = 1
			fp_bFoundField = False
			Do While (InStr(fp_iCurrentField, fp_sDefault, fp_sField) <> 0) _
				And Not fp_bFoundField
				fp_iCurrentField = InStr(fp_iCurrentField, fp_sDefault, fp_sField)
				fp_iStartField = InStr(fp_iCurrentField, fp_sDefault, "=")
				If fp_iStartField = fp_iCurrentField + len(fp_sField) Then
					fp_iEndField = InStr(fp_iCurrentField, fp_sDefault, "&")
					If (fp_iEndField = 0) Then fp_iEndField = len(fp_sDefault) + 1
					fp_sValue = Mid(fp_sDefault, fp_iStartField+1, fp_iEndField-1)
					fp_bFoundField = True
				Else
					fp_iCurrentField = fp_iCurrentField + len(fp_sField) - 1
				End If
			Loop
		End If

		' this next finds the named form field value, and substitutes in
		' doubled single-quotes for all single quotes in the literal value
		' so that SQL doesn't get confused by seeing unpaired single-quotes
		If (Mid(fp_sQry, fp_iStart - 1, 1) = """") Then
			fp_sValue = Replace(fp_sValue, """", """""")
		ElseIf (Mid(fp_sQry, fp_iStart - 1, 1) = "'") Then
			fp_sValue = Replace(fp_sValue, "'", "''")
		ElseIf Not IsNumeric(fp_sValue) Then
			fp_sValue = ""
		End If

		If (len(fp_sValue) = 0) Then fp_bBlankField = True

		fp_sQry = Left(fp_sQry, fp_iStart - 1) + fp_sWildCard + fp_sValue + _
			Right(fp_sQry, Len(fp_sQry) - fp_iEnd - 1)
		
		' Fixup the new current position to be after the substituted value
		fp_iCurrent = fp_iStart + Len(fp_sValue) + Len(fp_sWildCard)
	End If
Loop

If Not fp_fError Then
	' Use the connection string directly as entered from the wizard
	On Error Resume Next
	set fp_rs = CreateObject("ADODB.Recordset")
	If fp_iMaxRecords <> 0 Then fp_rs.MaxRecords = fp_iMaxRecords
	fp_rs.Open fp_sQry, "DSN=wsc"
	If Err.Description <> "" Then
		Response.Write "<B>Database Error: " + Err.Description + "</B>"
		if fp_bBlankField Then
			Response.Write "  One or more form fields were empty."
		End If
	Else
		' Check for the no-record case
		If fp_rs.EOF And fp_rs.BOF Then
			Response.Write fp_sNoRecords
		Else
			' Start a while loop to fetch each record in the result set
			Do Until fp_rs.EOF %>
<%' If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventWho")) %>
  <tr>
    <td nowrap bordercolor="#0080FF"><p align="left"><strong>Type</strong></td>
    <td nowrap bordercolor="#0080FF"><p align="left"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventType")) %></td>
    <td nowrap bordercolor="#0080FF"><strong>Submitted by</strong></td>
    <td nowrap bordercolor="#0080FF"><p align="left"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("SubmittedBy")) %></td>
  </tr>
  <tr>
    <td nowrap bordercolor="#0080FF"><p align="left"><strong>Time</strong></td>
    <td nowrap bordercolor="#0080FF"><p align="left"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventTime")) %></td>
    <td nowrap bordercolor="#0080FF"><strong>Submitted date</strong></td>
    <td nowrap bordercolor="#0080FF"><p align="left"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("SubmittedDate")) %></td>
  </tr>
  <tr>
    <td nowrap bordercolor="#0080FF"><strong>Location</strong></td>
    <td nowrap bordercolor="#0080FF"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventWhere")) %>
</td>
    <td nowrap bordercolor="#0080FF"><strong>Who Should Attend</strong></td>
    <td nowrap bordercolor="#0080FF"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventWho")) %>
</td>
  </tr>
  <tr>
    <td nowrap bordercolor="#0080FF"><strong>Details</strong></td>
    <td colspan="3" bordercolor="#0080FF"><%If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventDetails")) %>
</td>
  </tr>
  <tr>
    <td nowrap bgcolor="#400080" colspan="4">&nbsp;</td>
  </tr>
<%
				' Close the loop iterating records
				fp_rs.MoveNext
			Loop
		End If
		fp_rs.Close
	' Close the If condition checking for a connection error
	End If
' Close the If condition checking for a parse error when replacing form field params
End If
set fp_rs = Nothing
%>
  <tr>
    <td nowrap colspan="4"><form method="POST" action="eventdb.asp">
      <input type="hidden" name="EventDate" value="<%Response.Write Request.Form("EventDate")%>"><input type="hidden" name="RequestType" value="add"><div align="center"><center><p><input type="submit" value="Add event for this date" name="B1"><br>
      or<br>
      <a href="displaypage.asp">Return to Events Home</a><br>
      </p>
      </center></div>
    </form>
    </td>
  </tr>
</table>
</center></div></div><% 

 ' What type of request is this 
elseif Request.Form("RequestType") = "add" then 



 ' =======================================================================================================================
'   S E C T I O N  2    ( A D D  A N  E V E N T )
 ' =======================================================================================================================



%>


<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form2_Validator(theForm)
{

  if (theForm.SubmittedBy.value == "")
  {
    alert("Please enter a value for the \"Submitted By\" field.");
    theForm.SubmittedBy.focus();
    return (false);
  }

  if (theForm.SubmittedBy.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Submitted By\" field.");
    theForm.SubmittedBy.focus();
    return (false);
  }

  if (theForm.EventWhere.value == "")
  {
    alert("Please enter a value for the \"Location\" field.");
    theForm.EventWhere.focus();
    return (false);
  }

  if (theForm.EventWhere.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Location\" field.");
    theForm.EventWhere.focus();
    return (false);
  }

  if (theForm.EventType.value == "")
  {
    alert("Please enter a value for the \"Type Of Event\" field.");
    theForm.EventType.focus();
    return (false);
  }

  if (theForm.EventType.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Type Of Event\" field.");
    theForm.EventType.focus();
    return (false);
  }

  if (theForm.EventWho.value == "")
  {
    alert("Please enter a value for the \"Who Should Attend\" field.");
    theForm.EventWho.focus();
    return (false);
  }

  if (theForm.EventWho.value.length < 3)
  {
    alert("Please enter at least 3 characters in the \"Who Should Attend\" field.");
    theForm.EventWho.focus();
    return (false);
  }

  if (theForm.EventDetails.value == "")
  {
    alert("Please enter a value for the \"Details\" field.");
    theForm.EventDetails.focus();
    return (false);
  }

  if (theForm.EventDetails.value.length < 5)
  {
    alert("Please enter at least 5 characters in the \"Details\" field.");
    theForm.EventDetails.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="POST" action="eventdb.asp" onsubmit="return FrontPage_Form2_Validator(this)" name="FrontPage_Form2">
  <input type="hidden" name="EventDate" value="<%Response.Write Request.Form("EventDate")%>"><input type="hidden" name="RequestType" value="add_db"><input type="hidden" name="SubmittedDate" value="<%Response.Write Date%>"><div align="center"><center><table border="1" cellpadding="2" bordercolor="#0080FF" cellspacing="0" height="258">
    <tr>
      <td colspan="4" bgcolor="#400080" align="left" height="13"><div align="center"><center><h4>Please
      enter the details of the Event to add</h4>
      </center></div></td>
    </tr>
    <tr>
      <td align="left" height="23"><div align="left"><p><strong>Event submitted by</strong></td>
      <td align="left" height="23"><!--webbot bot="Validation" S-Display-Name="Submitted By" B-Value-Required="TRUE" I-Minimum-Length="3" --><input type="text" name="SubmittedBy" size="20" tabindex="1"></td>
      <td align="left" height="23"><strong>Location </strong></td>
      <td align="left" height="23"><!--webbot bot="Validation" S-Display-Name="Location" B-Value-Required="TRUE" I-Minimum-Length="3" --><input type="text" name="EventWhere" size="20" tabindex="2"></td>
    </tr>
    <tr>
      <td align="left" height="24"><div align="left"><p><strong>Date of event </strong></td>
      <td align="left" height="24"><div align="left"><p><%Response.Write FormatDateTime(Request.Form("EventDate"), 1) 
%></td>
      <td align="left" height="24"><div align="left"><p><strong>Starting time</strong></td>
      <td align="left" height="24"><select name="Hours" size="1">
        <option value="1">1</option>
        <option value="2">2</option>
        <option value="3">3</option>
        <option value="4">4</option>
        <option value="5">5</option>
        <option value="6">6</option>
        <option value="7">7</option>
        <option value="8">8</option>
        <option value="9">9</option>
        <option selected value="10">10</option>
        <option value="11">11</option>
        <option value="12">12</option>
      </select><select name="Minutes" size="1">
        <option value="00">00</option>
        <option value="15">15</option>
        <option value="30">30</option>
        <option value="45">45</option>
      </select><select name="AMPM" size="1">
        <option value="AM">AM</option>
        <option value="PM">PM</option>
      </select></td>
    </tr>
    <tr>
      <td align="left" height="23"><div align="left"><p><strong>Type of event</strong></td>
      <td align="left" height="23"><!--webbot bot="Validation" S-Display-Name="Type Of Event" B-Value-Required="TRUE" I-Minimum-Length="3" --><input type="text" name="EventType" size="20" tabindex="3"></td>
      <td align="left" height="23"><div align="left"><p><strong>Who should attend</strong></td>
      <td align="left" height="23"><!--webbot bot="Validation" S-Display-Name="Who Should Attend" B-Value-Required="TRUE" I-Minimum-Length="3" --><input type="text" name="EventWho" size="20" tabindex="4"></td>
    </tr>
    <tr>
      <td align="left" colspan="4" height="10"></td>
    </tr>
    <tr>
      <td valign="top" align="left" height="98"><div align="left"><p><strong>Details</strong></td>
      <td colspan="3" align="left" height="98"><div align="left"><p><!--webbot bot="Validation" S-Display-Name="Details" B-Value-Required="TRUE" I-Minimum-Length="5" --><textarea rows="4" name="EventDetails" cols="40" tabindex="5"></textarea></td>
    </tr>
    <tr>
      <td colspan="4" align="left" height="25"><div align="center"><center><p><input type="submit" value="Add this event" name="B1"></td>
    </tr>
  </table>
  </center></div>
</form>
<% 
 ' What type of request is this 
elseif Request.Form("RequestType") = "add_db" then 

 ' =======================================================================================================================
 '  S E C T I O N  3    A D D   E V E N T   T O   D A T A B A S E
 ' =======================================================================================================================


%>
<div align="center">

<p align="center">Event added </p>
<%
'Compile Event Time into Useable Format in variable EventTime
EventTime = Request.Form("Hours") & ":" & Request.Form("Minutes") & ":00 "& Request.Form("AMPM") 

 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================

fp_sQry = "SELECT EventDetails FROM Events WHERE (((EventDetails)='%%EventDetails%%') AND ((EventDate)=#" & Request.Form("EventDate") & "#));"

fp_sDefault = ""
fp_sNoRecords = ""
fp_iMaxRecords = 0
fp_iTimeout = 0
fp_iCurrent = 1
fp_fError = False
fp_bBlankField = False
If fp_iTimeout <> 0 Then Server.ScriptTimeout = fp_iTimeout
Do While (Not fp_fError) And (InStr(fp_iCurrent, fp_sQry, "%%") <> 0)
	' found a opening quote, find the close quote
	fp_iStart = InStr(fp_iCurrent, fp_sQry, "%%")
	fp_iEnd = InStr(fp_iStart + 2, fp_sQry, "%%")
	If fp_iEnd = 0 Then
		fp_fError = True
		Response.Write "<B>Database Region Error: mismatched parameter delimiters</B>"
	Else
		fp_sField = Mid(fp_sQry, fp_iStart + 2, fp_iEnd - fp_iStart - 2)
		If Mid(fp_sField,1,1) = "%" Then
			fp_sWildcard = "%"
			fp_sField = Mid(fp_sField, 2)
		Else
			fp_sWildCard = ""
		End If
		fp_sValue = Request.Form(fp_sField)

		' if the named form field doesn't exist, make a note of it
		If (len(fp_sValue) = 0) Then
			fp_iCurrentField = 1
			fp_bFoundField = False
			Do While (InStr(fp_iCurrentField, fp_sDefault, fp_sField) <> 0) _
				And Not fp_bFoundField
				fp_iCurrentField = InStr(fp_iCurrentField, fp_sDefault, fp_sField)
				fp_iStartField = InStr(fp_iCurrentField, fp_sDefault, "=")
				If fp_iStartField = fp_iCurrentField + len(fp_sField) Then
					fp_iEndField = InStr(fp_iCurrentField, fp_sDefault, "&")
					If (fp_iEndField = 0) Then fp_iEndField = len(fp_sDefault) + 1
					fp_sValue = Mid(fp_sDefault, fp_iStartField+1, fp_iEndField-1)
					fp_bFoundField = True
				Else
					fp_iCurrentField = fp_iCurrentField + len(fp_sField) - 1
				End If
			Loop
		End If

		' this next finds the named form field value, and substitutes in
		' doubled single-quotes for all single quotes in the literal value
		' so that SQL doesn't get confused by seeing unpaired single-quotes
		If (Mid(fp_sQry, fp_iStart - 1, 1) = """") Then
			fp_sValue = Replace(fp_sValue, """", """""")
		ElseIf (Mid(fp_sQry, fp_iStart - 1, 1) = "'") Then
			fp_sValue = Replace(fp_sValue, "'", "''")
		ElseIf Not IsNumeric(fp_sValue) Then
			fp_sValue = ""
		End If

		If (len(fp_sValue) = 0) Then fp_bBlankField = True

		fp_sQry = Left(fp_sQry, fp_iStart - 1) + fp_sWildCard + fp_sValue + _
			Right(fp_sQry, Len(fp_sQry) - fp_iEnd - 1)
		
		' Fixup the new current position to be after the substituted value
		fp_iCurrent = fp_iStart + Len(fp_sValue) + Len(fp_sWildCard)
	End If
Loop

If Not fp_fError Then
	' Use the connection string directly as entered from the wizard
	On Error Resume Next
	set fp_rs = CreateObject("ADODB.Recordset")
	If fp_iMaxRecords <> 0 Then fp_rs.MaxRecords = fp_iMaxRecords
	fp_rs.Open fp_sQry, "DSN=wsc"
	If Err.Description <> "" Then
		Response.Write "<B>Database Error: " + Err.Description + "</B>"
		if fp_bBlankField Then
			Response.Write "  One or more form fields were empty."
		End If
	Else
		' Check for the no-record case
		If fp_rs.EOF And fp_rs.BOF Then
			Response.Write fp_sNoRecords
		Else
			RecordExists=True

		End If
		fp_rs.Close
	' Close the If condition checking for a connection error
	End If
' Close the If condition checking for a parse error when replacing form field params
End If
set fp_rs = Nothing

 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================
 ' ======================================================================================================

' Check if data Already Exists 

If RecordExists=True then 

Response.Write "Record Already Exists"

else

' Insert Entry Into Database
' Query Consists of for fields and EventTime Variable
fp_sQry = "INSERT INTO Events ( EventDate, SubmittedDate, SubmittedBy, EventTime, EventType, EventWho, EventWhere, EventDetails ) VALUES ('%%EventDate%%', '%%SubmittedDate%%', '%%SubmittedBy%%', '" & EventTime & "', '%%EventType%%', '%%EventWho%%', '%%EventWhere%%', '%%EventDetails%%');"
fp_sDefault = ""
fp_sNoRecords = ""
fp_iMaxRecords = 0
fp_iTimeout = 0
fp_iCurrent = 1
fp_fError = False
fp_bBlankField = False
If fp_iTimeout <> 0 Then Server.ScriptTimeout = fp_iTimeout
Do While (Not fp_fError) And (InStr(fp_iCurrent, fp_sQry, "%%") <> 0)
	' found a opening quote, find the close quote
	fp_iStart = InStr(fp_iCurrent, fp_sQry, "%%")
	fp_iEnd = InStr(fp_iStart + 2, fp_sQry, "%%")
	If fp_iEnd = 0 Then
		fp_fError = True
		Response.Write "<B>Database Region Error: mismatched parameter delimiters</B>"
	Else
		fp_sField = Mid(fp_sQry, fp_iStart + 2, fp_iEnd - fp_iStart - 2)
		If Mid(fp_sField,1,1) = "%" Then
			fp_sWildcard = "%"
			fp_sField = Mid(fp_sField, 2)
		Else
			fp_sWildCard = ""
		End If
		fp_sValue = Request.Form(fp_sField)

		' if the named form field doesn't exist, make a note of it
		If (len(fp_sValue) = 0) Then
			fp_iCurrentField = 1
			fp_bFoundField = False
			Do While (InStr(fp_iCurrentField, fp_sDefault, fp_sField) <> 0) _
				And Not fp_bFoundField
				fp_iCurrentField = InStr(fp_iCurrentField, fp_sDefault, fp_sField)
				fp_iStartField = InStr(fp_iCurrentField, fp_sDefault, "=")
				If fp_iStartField = fp_iCurrentField + len(fp_sField) Then
					fp_iEndField = InStr(fp_iCurrentField, fp_sDefault, "&")
					If (fp_iEndField = 0) Then fp_iEndField = len(fp_sDefault) + 1
					fp_sValue = Mid(fp_sDefault, fp_iStartField+1, fp_iEndField-1)
					fp_bFoundField = True
				Else
					fp_iCurrentField = fp_iCurrentField + len(fp_sField) - 1
				End If
			Loop
		End If

		' this next finds the named form field value, and substitutes in
		' doubled single-quotes for all single quotes in the literal value
		' so that SQL doesn't get confused by seeing unpaired single-quotes
		If (Mid(fp_sQry, fp_iStart - 1, 1) = """") Then
			fp_sValue = Replace(fp_sValue, """", """""")
		ElseIf (Mid(fp_sQry, fp_iStart - 1, 1) = "'") Then
			fp_sValue = Replace(fp_sValue, "'", "''")
		ElseIf Not IsNumeric(fp_sValue) Then
			fp_sValue = ""
		End If

		If (len(fp_sValue) = 0) Then fp_bBlankField = True

		fp_sQry = Left(fp_sQry, fp_iStart - 1) + fp_sWildCard + fp_sValue + _
			Right(fp_sQry, Len(fp_sQry) - fp_iEnd - 1)
		
		' Fixup the new current position to be after the substituted value
		fp_iCurrent = fp_iStart + Len(fp_sValue) + Len(fp_sWildCard)
	End If
Loop

If Not fp_fError Then
	' Use the connection string directly as entered from the wizard
	On Error Resume Next
	set fp_rs = CreateObject("ADODB.Recordset")
	If fp_iMaxRecords <> 0 Then fp_rs.MaxRecords = fp_iMaxRecords
	fp_rs.Open fp_sQry, "DSN=wsc"
	If Err.Description <> "" Then
		Response.Write "<B>Database Error: " + Err.Description + "</B>"
		if fp_bBlankField Then
			Response.Write "  One or more form fields were empty."
		End If
	Else
		' Check for the no-record case
		If fp_rs.EOF And fp_rs.BOF Then
			Response.Write fp_sNoRecords
		Else
			' Start a while loop to fetch each record in the result set
			Do Until fp_rs.EOF
 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then Response.Write CStr(fp_rs("EventWho")) 

				' Close the loop iterating records
				fp_rs.MoveNext
			Loop
		End If
		fp_rs.Close
	' Close the If condition checking for a connection error
	End If
' Close the If condition checking for a parse error when replacing form field params
End If
set fp_rs = Nothing

' check if record exists and add is nessisary
end if
%>

<form method="POST" action="eventdb.asp">
  <input type="hidden" name="EventDate" value="<%Response.Write Request.Form("EventDate")%>"><input type="hidden" name="RequestType" value="show"><p><input type="submit" value="Show Events" name="B3"></p>
</form>
</div><% 
 ' What type of request is this 
else

 ' =======================================================================================================================
 '  S E C T I O N  4    I F   D A T A   N O T   G I V E N 
 ' =======================================================================================================================

%>


<h3 align="center">Error , </h3>

<h4 align="center">Required data not passed</h4>

<p align="center">Perhaps you have followed a link of bookmark directly to this page, if
so please return to the events page and try again.</p>
<div align="center"><center>

<table border="0">
  <tr>
    <td width="50%" colspan="2"><h4>Options :</h4>
    </td>
  </tr>
  <tr>
    <td width="50%"><a href="displaypage.asp">Return to events</a></td>
    <td width="50%">Email: the <a href="mailto:blueskies@khaoz.co.za">webmaster</a></td>
  </tr>
</table>
</center></div><%
 ' What type of request is this 
End if

%>

</body>
</html>
