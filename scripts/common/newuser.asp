<!-- #include file ="../../lib/include.inc" -->

<%
If Request.QueryString("QueryType") = "CheckContact" then

'If Request.QueryString("SubmitYes") <> "" then		' Do not show display any headers if DB page only
'CheckContact
 else
 
 
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="nocache">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../lib/styles/dz.css" rel="stylesheet" type="text/css">
<title>About you</title>
</head>
<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

<%end if

Select Case Request.QueryString("QueryType")

' -----------------------------------------N E X T  -  C A S E-----------------------------------------

' No input  / Present Form
Case ""
%>


<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.FirstName.value == "")
  {
    alert("Please enter a value for the \"FirstName\" field.");
    theForm.FirstName.focus();
    return (false);
  }

  if (theForm.FirstName.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"FirstName\" field.");
    theForm.FirstName.focus();
    return (false);
  }

  if (theForm.LastName.value == "")
  {
    alert("Please enter a value for the \"LastName\" field.");
    theForm.LastName.focus();
    return (false);
  }

  if (theForm.LastName.value.length < 1)
  {
    alert("Please enter at least 1 characters in the \"LastName\" field.");
    theForm.LastName.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form method="GET" action="newuser.asp" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
  <table border="0">
    <tr>
      <td colspan="5">
        <h4>Step 1, About yourself:</h4>
      </td>
    </tr>
    <tr>
      <td height="15" colspan="5"></td>
    </tr>
    <tr>
      <td><b>FirstName</b></td>
      <td><!--webbot bot="Validation" S-Display-Name="FirstName"
        B-Value-Required="TRUE" I-Minimum-Length="1" --><input type="text" name="FirstName" size="20" tabindex="1"><i>Required</i></td>
      <td width="20"></td>
      <td><b>e-mail</b></td>
      <td><input type="text" name="email" size="26" tabindex="4"></td>
    </tr>
    <tr>
      <td><b>LastName</b></td>
      <td><!--webbot bot="Validation" S-Display-Name="LastName"
        B-Value-Required="TRUE" I-Minimum-Length="1" --><input type="text" name="LastName" size="20" tabindex="2"><i>Required</i></td>
      <td width="20"></td>
      <td><b>Web Site</b></td>
      <td>http://<input type="text" name="URL" size="20" tabindex="5"></td>
    </tr>
    <tr>
      <td><b>NickName</b></td>
      <td><input type="text" name="NickName" size="20" tabindex="3"></td>
      <td width="20"></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td height="10"></td>
      <td height="10"></td>
      <td height="10" width="20"></td>
      <td height="10"></td>
      <td height="10"></td>
    </tr>
    <tr>
      <td><b>Dropzone</b></td>
      <td><select size="1" name="Dropzone" tabindex="6">
          <option selected value="">--- Choose one ---</option>
          <%
           DropZones = CreateRS(RecordSet,"Select * From Dropzones Order by Dropzone;",DSN)	' Display options for all dropzones
           Call RSDisplay(RecordSet,"DZOption","","","")					' in DZ's Table
           Call CloseRS(RecordSet)
          %>
          </select></td>
      <td width="20"></td>
      <td colspan="2">Include me in the Skydiver Database</td>
    </tr>
    <tr>
      <td>
        <p align="right"><b>other</b></p>
      </td>
      <td><input type="text" name="DZOther" size="20"></td>
      <td width="20"></td>
      <td><b>Yes</b></td>
      <td><input type="radio" value="yes" checked name="Display" tabindex="7"></td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td width="20"></td>
      <td><b>No</b></td>
      <td><input type="radio" name="Display" value="no"></td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td width="20"></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td colspan="5">
        <p align="center"><input type="submit" value="Proceed to Step 2  > >" name="B1"></td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
      <td></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <input type="hidden" name="QueryType" value="CheckContact">
  <%Call GenTranIDInput("Contacts",DSN)%>
  <input type="hidden" name="exitpage" value="<%= Request.QueryString("exitpage")%>">
</form>

<%

' -----------------------------------------N E X T  -  C A S E-----------------------------------------


Case "CheckContact"

Dim Records,Record,LastRecord,RSRow,InsertAction

Records = CreateRS(RecordSet,"Select * From Contacts where ((FirstName like '%%%FirstName%%%') and (LastName like '%%%LastName%%%'));",DSN) 
LastRecord = Int(Request.QueryString("LastRecord"))' Define Lastrecord String of none available

If LastRecord = "0" then Record = 1 'If restload assign Record Number

If Request.QueryString("SubmitYes") <> "" then	' Determine output of last form
 InsertAction = "Found"
 Record = LastRecord  ' Set workning record
elseif Request.QueryString("SubmitNo") <> "" then
 Record = LastRecord + 1 ' Set workning record
end if


If Records <> "Nothing" and Record <= Records then	' Record Possibly Exists
RSRow = 1

	Do Until RSRow = Record ' Move RecordSet To Current Record
	 RecordSet.MoveNext
	 RSRow = RSRow + 1
	Loop
		
	If InsertAction = "Found" then
	 ContactID = RSColData("ContactID",RecordSet)
	else

%>

<!-- #include file ="../../lib/headers.inc" -->

<form method="GET" action="newuser.asp">
<input type="hidden" name="QueryType" value="CheckContact">
<input type="hidden" name="FirstName" value="<%Response.Write Request.QueryString("FirstName")%>">
<input type="hidden" name="LastName" value="<%Response.Write Request.QueryString("LastName")%>">
<input type="hidden" name="NickName" value="<%Response.Write Request.QueryString("NickName")%>">
<input type="hidden" name="Email" value="<%Response.Write Request.QueryString("Email")%>">
<input type="hidden" name="URL" value="<%Response.Write Request.QueryString("URL")%>">
<input type="hidden" name="Dropzone" value="<%Response.Write Request.QueryString("Dropzone")%>">
<input type="hidden" name="DZOther" value="<%Response.Write Request.QueryString("DZOther")%>">
<input type="hidden" name="Display" value="<%Response.Write Request.QueryString("Display")%>">
<input type="hidden" name="TranID" value="<%Response.Write Request.QueryString("TranID")%>">
<input type="hidden" name="LastRecord" value="<%Response.Write Record%>">
<input type="hidden" name="exitpage" value="<%= Request.QueryString("exitpage")%>">

<table border="0">
  <tr>
    <td colspan="3">
      <h4>Possible match found, Is this you:?</h4>
    </td>
  </tr>
  <tr>
    <td colspan="3" height="10"></td>
  </tr>
  <tr>
    <td colspan="3">
      <p><%Response.Write "<B>" & "Name: </B>" & RSColData("FirstName",RecordSet)& " " & RSColData("LastName",RecordSet)
      If RSColData("NickName",RecordSet) <> "" then
      Response.Write " (" & RSColData("NickName",RecordSet) & ")"
      end IF
      Response.Write "<BR>" 
      If RSColData("Email",RecordSet) <> "" then
      Response.Write "<B>e-mail address: </B>" & RSColData("Email",RecordSet) & "<BR>" 
      end IF
      If RSColData("Dropzone",RecordSet) <> "" then
      Response.Write " <B>Skydives at: </B>" & RSColData("Dropzone",RecordSet)
      end IF
            
      %> </p>
    </td>
  </tr>
  <tr>
    <td colspan="3" height="10">
    </td>
  </tr>
  <tr>
    <td>
      <p align="right"><input type="submit" value="Yes" name="SubmitYes"></p>
    </td>
    <td>
      <p align="center"><b>&lt; &lt; Choose One &gt; &gt;</b></p>
    </td>
    <td>
      <input type="submit" value="No" name="SubmitNo">
    
    </td>
  </tr>
</table>
</form>
<p>


<%

	end if ' End InsertAction IF
End If

Call CloseRS(RecordSet)


If  Records = "Nothing" or Record > Records then	' Definately not in DB

' Output Form for record
' Add Record To Database and retrieve contact ID

	If Request.QueryString("DZOther") = "" then ' Check to see if other DZ was stipulated
		ContactID = InsertFormData("ContactID","'FirstName','LastName','NickName','Email','URL','Dropzone','Display'","Contacts",DSN)
	Else
		DZ = InsertFormData("Dropzone","'DZOther--Dropzone'","Dropzones",DSN)
		ContactID = InsertFormData("ContactID","'FirstName','LastName','NickName','Email','URL','DZOther--Dropzone','Display'","Contacts",DSN)
	end If
end If


If ContactID <> "" then		' Redirect							
	Response.Redirect(Request.QueryString("exitpage")) & "?ContactID=" & ContactID
end if

' -----------------------------------------N E X T  -  C A S E-----------------------------------------



end Select%>







</body>

</html>
