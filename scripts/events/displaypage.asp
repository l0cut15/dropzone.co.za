<html>

<head>
<title>Events</title>
<meta http-equiv="author" content="Brett Samuel khaoz@khaoz.co.za">
<meta http-equiv="content" content="Skydiving in South Africa">
<meta http-equiv="keywords"
content="Witbank, Skydive, Skydiving , Club, tandem , AFF, cypress, South Africa, static line, RSL, freefall, dope rope, paracute, Skymaster, Tempo, Dropzone, ZA">
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
P.small_text_p {font-size:8pt;font-family:  Verdana,Arial, sans-seriff}
td {font-size:10pt;font-family:  Verdana,Arial, sans-seriff}
td.heading {font-size:12pt;font-family:  Verdana,Arial, sans-seriff; font-weight: bold}
A.menu {text-decoration: none}
A:hover {color:yellow}
INPUT.Events {font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold; float: right;background-color: rgb(0,51,153); color: rgb(255,255,255);border: 10px rgb(0,51,153)}
INPUT.Days {font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold; float: right;background-color: rgb(0,51,153); color: rgb(255,255,0);border: 10px rgb(0,51,153)}

-->
</style>
<base target="main">
</head>

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399"
text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<%

' ---------------------------------------------
' Dimension any variables here 

dim EventTrue(31)


' ---------------------------------------------
'Set date of form to be dispalyed

If Request.Form("ReqDate") = "" then 	' First time for is loaded
ReqDate = Date
else

' Adjust Date According to form input

Select case Request.Form("MoveType")
					Case "1" ReqDate = DateAdd("yyyy", Request.Form("Inc")*-1, Request.Form("ReqDate"))  ' sub 1 year
					Case "2" ReqDate = DateAdd("m", Request.Form("Inc")*-1, Request.Form("ReqDate"))	' sub 1 month
					Case "3" ReqDate = DateAdd("m", Request.Form("Inc"), Request.Form("ReqDate"))	' add 1 month	
					Case "4" ReqDate = DateAdd("yyyy", Request.Form("Inc"), Request.Form("ReqDate"))	' add 1 year
end select

	
end if 
' ---------------------------------------------

' Determine Day, Month, Year 
ReqMonth = Month(ReqDate)
ReqDay = Day(ReqDate) 
ReqYear = Year(ReqDate)
ReqWeekDay = weekday(ReqDate)
' ---------------------------------------------

' Determine Starting Weekday Of Month
FirstDayOfMonth = ReqMonth & " 01 " & ReqYear
FirstWeekDayOfMonth = weekday(FirstDayOfMonth)
' ---------------------------------------------

' Set Inital Values for Drawing Table

DaysInMonth= 1
CellDay = 2 - FirstWeekDayOfMonth
DateOffset = FirstWeekDayOfMonth ' Number of blank spaces 
' ---------------------------------------------
'   Populate Event By Name Array
' ---------------------------------------------


' Query Database to create list of events
sQry = Events_month_qry(ReqDate)

On Error Resume Next
	set rs = CreateObject("ADODB.Recordset") ' Open Database
	rs.Open sQry, "DSN=wsc"

	
			Do Until rs.EOF

If Not IsEmpty(rs) And Not (rs Is Nothing) Then EventTrue(Day(CStr(rs("EventDate")))) = True

				rs.MoveNext
			Loop
rs.Close
set rs = Nothing

' ---------------------------------------------






%>

<h3 align="center">Events</h3>
<div align="center"><center>

<table border="1" cellpadding="2" bordercolor="#0080FF" cellspacing="0" height="184">
  <tr>
    <td rowspan="3"><table border="1" cellpadding="2" bordercolor="#0080FF" cellspacing="0">
      <tr>
        <td class="heading" colspan="5" height="1"><%Response.Write MonthName(ReqDate)%>
</td>
        <td class="heading" colspan="2" height="1"><%Response.Write Year(ReqDate)%>
</td>
      </tr>
      <tr>
        <td align="center"><p align="center"><strong>Sun</strong></td>
        <td align="center"><p align="center"><strong>Mon</strong></td>
        <td align="center"><p align="center"><strong>Tue</strong></td>
        <td align="center"><p align="center"><strong>Wed</strong></td>
        <td align="center"><p align="center"><strong>Thur</strong></td>
        <td align="center"><p align="center"><strong>Fri</strong></td>
        <td align="center"><p align="center"><strong>Sat</strong></td>
      </tr>
<%
Do While  DaysInMonth <= MonthDays(ReqDate)

%>
      <tr align="center">
<%
DayColum = 1
Do while DayColum <=7
%>
        <td><%
If CellDay < 1 then 
Response.Write ""
DayInc = 0
elseif CellDay > MonthDays(ReqDate) then 
Response.Write ""
else%>
<form method="POST" action="eventdb.asp">
          <input type="hidden" name="EventDate" value="<%Response.Write CellDate(CellDay)%>"><input
          type="hidden" name="RequestType" value="show"><p><input type="submit"
          value="<%Response.Write CellDay%>" name="B1"
          class="<% If EventTrue(CellDay) = True then 
																			Response.Write "days"
																			else
																			Response.Write "events"
																			End IF
%>"></p>
        </form>
<%

DayInc = 1
end if
%>
        </td>
<%
DayColum = DayColum  + 1
DaysInMonth = DaysInMonth + DayInc
CellDay = CellDay + 1
Loop
%>
<%
Loop
%>
      </tr>
      <tr>
        <td colspan="7" height="87"><form method="POST" action="displaypage.asp"
        name="FrontPage_Form2">
          <input type="hidden" name="ReqDate" value="<%Response.Write ReqDate%>"><table border="0"
          width="100%" cellspacing="3" cellpadding="2">
            <tr>
              <td align="center" valign="top" nowrap><input type="radio" value="1" name="MoveType"><br>
              <font size="1">-Years</font></td>
              <td align="center" valign="top" nowrap><input type="radio" name="MoveType" value="2"><br>
              <font size="1">-Months</font></td>
              <td align="center" valign="top" nowrap><select name="Inc" size="1">
                <option selected value="1">1</option>
                <option value="2">2</option>
                <option value="3">3</option>
                <option value="4">4</option>
                <option value="5">5</option>
                <option value="6">6</option>
                <option value="7">7</option>
                <option value="8">8</option>
                <option value="9">9</option>
                <option value="10">10</option>
                <option value="11">11</option>
                <option value="12">12</option>
              </select></td>
              <td align="center" valign="top" nowrap><input type="radio" name="MoveType" value="3"
              checked><br>
              <font size="1">+Months</font></td>
              <td align="center" valign="top" nowrap><input type="radio" name="MoveType" value="4"><br>
              <font size="1">+Years</font></td>
            </tr>
            <tr>
              <td colspan="5" align="center" valign="top" nowrap><div align="center"><center><p><input
              type="submit" value="Go there !" name="B1"></td>
            </tr>
          </table>
        </form>
        </td>
      </tr>
    </table>
    </td>
    <td align="center" valign="top"><table border="1" width="100%" bordercolor="#0080FF"
    cellspacing="0">
      <tr>
        <td nowrap bgcolor="#800080"><p align="center"><strong>Month's Highlights<br>
        </strong><small><small>Click date for details</small></small></td>
      </tr>
<%' Query Database to create list of events
sQry = Events_month_qry(ReqDate)

On Error Resume Next
	set rs = CreateObject("ADODB.Recordset") ' Open Database
	rs.Open sQry, "DSN=wsc"

	
			Do Until rs.EOF
%>
      <tr>
        <td><% 
If Not IsEmpty(rs) And Not (rs Is Nothing) Then Response.Write Day(CStr(rs("EventDate")))
%>
<%Response.Write ", "%>
<% 
If Not IsEmpty(rs) And Not (rs Is Nothing) Then Response.Write CStr(rs("EventType")) 
%>
</td>
      </tr>
<%
				' Close the loop iterating records
				rs.MoveNext
			Loop
rs.Close
set rs = Nothing
%>
    </table>
    </td>
  </tr>
</table>
</center></div>

<h6 align="center">Click the date to add an event.</h6>
</body>
<%
' -------------------------------------------------------------------
'   F U N C T I O N S 
' -------------------------------------------------------------------
Function Leapyear(ReqDate)

' Name - Leapyear Function by Brett Samuel
' Last Modified - 03-04-1999
' Check Leap Year Finction
' Returns True or False values for Leapyear
' 
' Inputs - ReqYear 
'

CheckReqYear= Year(ReqDate) / 400					' Divide Year by 400 to check for millenium leap
CheckFixReqYear = fix(CheckReqYear)

If CheckReqYear = CheckFixReqYear then
Leapyear = True
else

	CheckReqYear= Year(ReqDate) / 100					' Divide Year by 100 to check for cent leap
	CheckFixReqYear = fix(CheckReqYear)
	If CheckReqYear = CheckFixReqYear then
	Leapyear = False

	else
			CheckReqYear= Year(ReqDate) / 4				' Divide Year by 4 to check for leap
			CheckFixReqYear = fix(CheckReqYear)
			
			If CheckReqYear = CheckFixReqYear then
			Leapyear = True
			else
			Leapyear = False	
			end if  	
	end if 
End If

End Function

' -------------------------------------------------------------------
'  N E X T   F U N C T I O N . . . . .  
' -------------------------------------------------------------------

Function Monthdays(ReqDate)

' Name - Monthdays Finction by Brett Samuel
' Last Modified - 02-08-1999
' Determine Number of days in a month
' Returns number of days (Integer)
'
' Uses leapyear(function)
' 
' Inputs - ReqDate (mm/dd/yy)
'

Select case Month(ReqDate)
					Case "1" Monthdays = 31
					Case "2" 
								If leapyear(ReqDate) = True then 
								Monthdays = 29 
								else 
								Monthdays = 28
					         end if				
					Case "3" Monthdays = 31
					Case "4" Monthdays = 30
					Case "5" Monthdays = 31
					Case "6" Monthdays = 30
					Case "7" Monthdays = 31
					Case "8" Monthdays = 31
					Case "9" Monthdays = 30
					Case "10" Monthdays = 31
					Case "11" Monthdays = 30
					Case "12" Monthdays = 31
end select

End Function

' -------------------------------------------------------------------
'  N E X T   F U N C T I O N . . . . .  
' -------------------------------------------------------------------

Function MonthName(ReqDate)

' Name - MonthName Finction by Brett Samuel
' Last Modified - 03-03-1999
' Determine Number of days in a month
' Returns name of month based on the date (string)
' 
' Inputs - Reqdate (mm/dd/yy)
'

Select case Month(ReqDate)
					Case "1" MonthName = "January"
					Case "2" MonthName = "February"
					Case "3" MonthName = "March"
					Case "4" MonthName = "April"
					Case "5" MonthName = "May"
					Case "6" MonthName = "June"
					Case "7" MonthName = "July"
					Case "8" MonthName = "August"
					Case "9" MonthName = "September"
					Case "10" MonthName = "October"
					Case "11" MonthName = "November"
					Case "12" MonthName = "December"
end select

End Function

' -------------------------------------------------------------------
'  N E X T   F U N C T I O N . . . . .  
' -------------------------------------------------------------------

Function CellDate(CellDay)

' Name - CellDate Function by Brett Samuel
' Last Modified - 04-03-1999
' Determine Number of days in a month
' Return Datestring for Calender Cell(string)
' 
' Outpub - CellDate (mm/dd/yy)
' Inputs - CellDay (dd)


CellDate = Month(ReqDate)& "/" & CellDay & "/" & Year(ReqDate)


End Function

' -------------------------------------------------------------------
'  N E X T   F U N C T I O N . . . . .  
' -------------------------------------------------------------------




' -------------------------------------------------------------------
'  N E X T   F U N C T I O N . . . . .  
' -------------------------------------------------------------------

Function Events_month_qry(ReqDate)

' Name - Events_month_qry Function by Brett Samuel
' Last Modified - 04-03-1999
' Generate Query For Events this Montg
' Return String 
' 
' Inputs - ReqDate
Events_month_qry = "SELECT EventDate, EventType FROM Events WHERE (((Events.EventDate)>=#" & Month(ReqDate) & "/"& "1" & "/" & Year(ReqDate) &"# And (Events.EventDate)<#" & Month(ReqDate)& "/" & MonthDays(ReqDate) &"/"&Year(ReqDate)&"#));"

End Function

' -------------------------------------------------------------------
'  E N D  F U N C T I O N S
' -------------------------------------------------------------------
%>
</html>
