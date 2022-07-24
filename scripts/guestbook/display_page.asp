
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
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
A.NavBar {text-decoration: none}
P.NavBar {text-decoration: none}
P.NavStatTop {font-size:11pt;font-family:  Arial, sans-serif; font-weight: bold} 
A:hover {color:yellow}
INPUT.Events {font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold; float: right;background-color: rgb(0,51,153); color: rgb(255,255,255);border: 10px rgb(0,51,153)}
INPUT.Days {font-family: Arial, sans-serif; font-size: 8pt; font-weight: bold; float: right;background-color: rgb(0,51,153); color: rgb(255,255,0);border: 10px rgb(0,51,153)}

-->
</style>

<!-- #include file ="../../lib/include.inc" -->

<base target="main">

</head>


<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">


<%

DSN="DSN=WSC"


Call CreateRS(RecordSet,"SELECT * FROM Contacts INNER JOIN GuestBook ON Contacts.ContactID = GuestBook.ContactID ORDER BY GuestBook.MessageID DESC;",DSN) 


%>
<h6 align="center"><img src="../../images/titles/guestbook.gif" alt="Guestbook"><br>
<br>
</h6>
<h4 align="center">There are currently <% Response.Write RSDisplay(RecordSet,"","","","0")%> entries.</h4>
<h4 align="center">So whats stopping you , <a href="../common/newuser.asp?exitpage=../guestbook/addmessage.asp">ADD</a>
yours now.</h4>

<%

call RSDisplay(RecordSet,"GuestBook","NavGuestHead","NavGuestTail","10") ' Display Contents of RS in HTML

%>

<%Call CloseRS(RecordSet)%>





</body>

</html>
