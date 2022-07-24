<!-- #include file ="../../lib/include.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta http-equiv="Pragma" content="nocache">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../lib/styles/dz.css" rel="stylesheet" type="text/css">



<title>Your Message</title>

</head>


<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">


<%

Select Case Request.QueryString("QueryType")

Case "" %>


<form method="GET" action="addmessage.asp">
  <table border="0">
    <tr>
      <td colspan="3">
        <h4>Step 2, Enter Message:</h4>
      </td>
    </tr>
    <tr>
      <td colspan="3" height="20"></td>
    </tr>
    <tr>
      <td valign="top"><b> Message:</b></td>
      <td><textarea rows="4" name="Message" cols="40"></textarea></td>
      <td><i>Required</i></td>
    </tr>
    <tr>
      <td colspan="3">
        <p align="center"><input type="submit" value="Add to guestbook  > >" name="Add"></td>
    </tr>
  </table>
  <p>&nbsp;</p>
  <input type="hidden" name="ContactID" value="<%Response.Write Request.QueryString("ContactID")%>"><input type="hidden" name="TranDate" value="<%Response.Write Date%>"><input type="hidden" name="QueryType" value="AddMessage"><input type="hidden" name="TranTime" value="<%Response.Write Time%>">
<%Call GenTranIDInput("GuestBook",DSN)%>
</form>


<%Case "AddMessage"
Call InsertFormData("MessageID","'Message','TranDate','TranTime',ContactID","GuestBook",DSN)

%>

<table border="0" width="100%">
  <tr>
    <td width="100%">
      <h3 align="center">Message successfully added ...</h3>
      <h4 align="center">Thank you for your submission, </h4>
      <h4 align="center">To view the guestbook click <a href="display_page.asp">here</a>.</h4>
      <h4 align="center">If your entry doesn't show up immediately, try pressing
      the refresh / reload button in your browser.</h4>
      <h4 align="center">or</h4>
      <h4 align="center">e-mail the webmaster for assistance</h4>
      <p align="center"><a href="mailto:freakfaller@dropzone.co.za">freakfaller@dropzone.co.za</a></p>
      <p>&nbsp;</td>
  </tr>
</table>
<p>

<%

' -----------------------------------------N E X T  -  C A S E-----------------------------------------


end Select%>



</p>



</body>

</html>
