<html>

<head>
<title>Are You A Skydiving Freak</title>
<meta http-equiv="author" content="Brett Samuel blueskies@dropzone.co.za">
<meta http-equiv="content" content="Skydiving in South Africa">
<meta http-equiv="keywords"
content="Witbank, Skydive, Skydiving Club, training, tandem , AFF, cypress, South Africa, static line, RSL, freefall, dope rope, paracute, Skymaster, Tempo, Dropzone, ZA, Dytter, Cypress, Time Out, Reserve, PISA, PASA">
<meta name="GENERATOR" content="Microsoft FrontPage 3.0">
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
<%

'
' Obtain Question Number

If Request.Form("QuestionNumber") = "" then
QuestionNumber = 1
else
QuestionNumber = Request.Form("QuestionNumber") + 1
end if


' Substitute in form parameters into the query string
fp_sQry = "SELECT* FROM Freak_Questions WHERE (((Freak_Questions.QuestionNumber)=" & QuestionNumber & "));"
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
			TestStatus = "Completed"

		Else
			TestStatus = "Active"
			' Start a while loop to fetch each record in the result set
			Do Until fp_rs.EOF
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionString = CStr(fp_rs("QuestionString")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionA = CStr(fp_rs("A")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionB = CStr(fp_rs("B")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionC = CStr(fp_rs("C")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionD = CStr(fp_rs("D")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionE = CStr(fp_rs("E")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then A_A = CStr(fp_rs("A_A")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then A_B = CStr(fp_rs("A_B")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then A_C = CStr(fp_rs("A_C")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then A_D = CStr(fp_rs("A_D")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then A_E = CStr(fp_rs("A_E"))
 If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionType = CStr(fp_rs("QuestionType")) 
If Not IsEmpty(fp_rs) And Not (fp_rs Is Nothing) Then QuestionCount = CStr(fp_rs("QuestionCount")) 
				
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

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399"
text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%Dim Score 
If Request.Form("LastQuestionType") = "Radio" then
LastScore = Request.Form("Question")
else

LastScore = Request.Form("A_A") + 1 +Request.Form("A_B") + 1 + Request.Form("A_C") + 1 + Request.Form("A_D") + 1 + Request.Form("A_E") + 1
LastScore = LastScore - 5

end if


'calculate latest score
Score = Request.Form ("Score")
If Score = "" then Score = 0
Score = Score + 1
Score = LastScore + Score -1

%>

<h3 align="center">Are You A Skydiving Freak? </h3>

<p>&nbsp;</p>
<% 
' Check to see if questions have ended
if TestStatus = "Active" then

If QuestionType = "Radio" then %>

<form method="POST" action="sky_freak.asp">
  <input type="hidden" name="LastQuestionType" value="<%Response.Write QuestionType%>"><input
  type="hidden" name="QuestionNumber" value="<%Response.Write QuestionNumber%>"><input
  type="hidden" name="Score" value="<%Response.Write Score%>"><strong><p><%Response.Write QuestionNumber%>) <%Response.Write QuestionString%> </strong></p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="radio" name="Question" value="<%Response.Write A_A%>"> <%Response.Write QuestionA%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="radio" name="Question" value="<%Response.Write A_B%>"> <%Response.Write QuestionB%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="radio" name="Question" value="<%Response.Write A_C%>"> <%Response.Write QuestionC%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="radio" name="Question" value="<%Response.Write A_D%>"> <%Response.Write QuestionD%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="radio" name="Question" value="<%Response.Write A_E%>"> <%Response.Write QuestionE%><br>
<%end if%>  </p>
  <p><input type="submit" value="Submit" name="B1"> </p>
</form>
<%else%>

<form method="POST" action="sky_freak.asp">
  <input type="hidden" name="QuestionNumber" value="<%Response.Write QuestionNumber%>"><input
  type="hidden" name="QuestionType" value="Checkbox"><input type="hidden" name="Score"
  value="<%Response.Write Score%>"><strong><p><%Response.Write QuestionNumber%>) <%Response.Write QuestionString%> </strong></p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="checkbox" value="<%Response.Write A_A%>" name="A_A"> <%Response.Write QuestionA%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="checkbox" value="<%Response.Write A_B%>" name="A_B"> <%Response.Write QuestionB%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input t type="checkbox" value="<%Response.Write A_C%>" name="A_C"> <%Response.Write QuestionC%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input t type="checkbox" value="<%Response.Write A_D%>" name="A_D"> <%Response.Write QuestionD%><br>
<%end if%>  </p>
<%If QuestionCount > 0 then
QuestionCount = QuestionCount - 1 %>
  <p><input type="checkbox" value="<%Response.Write A_E%>" name="A_E"> <%Response.Write QuestionE%><br>
<%end if%>  </p>
  <p><input type="submit" value="Submit" name="B1"> </p>
</form>
<%end if%>
<%else
	
   	If Score <= 4 then
            ResultsTitle = "Wuffo !!!"
			  ResultsDesciption = "Very disappointing! Do you jump more than twice a year? Skydive yer Freak ! "
		elseif score <=11  then
           ResultsTitle = "Thinking About It ?"
			 ResultsDesciption = "Come on, you can do better than that. Financial difficulties? No excuse! Get your ass in the sky!"
		elseif score <= 20 then
          ResultsTitle= "Getting There !"
			 ResultsDesciption = "Potential is there, you just need more commitment. Keep it up!"
		elseif score <= 30 then
			ResultsTitle= "Definite Potential !"
			 ResultsDesciption  ="Nearly there, a good healthy balance of 90% skydiving 10% life."

		elseif score <= 40 then
			ResultsTitle= "Seriously Disturbed !!"
			 ResultsDesciption  ="There really is more to life than skydiving (not much more though). Seek medical attention before it’s too late."
         else  
			ResultsTitle = "Get Commited !!"
			ResultsDesciption  ="Truly a twisted individual, its probably far too late to seek medical help. Don’t panic! There are worse ways to go! "

      End IF
%>

<table border="0" width="400">
  <tr>
    <td colspan="2" nowrap><h3 align="center"><font color="#FFFF80"><%Response.Write ResultsTitle%></font></h3>
    </td>
    <td nowrap valign="top"><p align="right"><font color="#FFFF80">Your Score: <%response.Write Score%></font></td>
  </tr>
  <tr>
    <td width="10"></td>
    <td width="100%" colspan="2"><%Response.Write ResultsDesciption %>
</td>
  </tr>
</table>

<p>&nbsp;<%end if%> </p>
</body>
</html>
