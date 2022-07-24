<html>

<head>
<title>Witbank Skydiving Club</title>
<meta http-equiv="Author" content="Brett Samuel">
<meta http-equiv="content" content="Skydiving and Paracuting in South Africa">
<meta http-equiv="Description " content="Skydiving and Paracuting in South Africa">
<meta http-equiv="keywords" content=" ZA, skydiving in South Africa, skydiving, south africa">
<meta http-equiv="keywords"
content="rave, extacy, nude skydivers, free beer, Extreme Sports, Skysurfing, Free Fall">
<meta http-equiv="keywords"
content="Witbank, Skydive, Skydiving, Club, tandem, training, AFF, cypress, South Africa, static line, RSL, freefall, dope rope, paracute, Skymaster, Tempo, reserve, PISA, Dropzone">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<base target="main">
<style TYPE="text/css">
<!--
H1 {font-size:26pt;font-family:  Arial, sans-serif}
H2 {font-size:20pt;font-family:  Arial, sans-serif}
H3 {font-size:16pt;font-family:  Arial, sans-serif}
H4 {font-size:12pt;font-family:  Arial, sans-serif}
H5 {font-size:10pt;font-family:  Arial, sans-serif}
H6 {font-size:8pt;font-family: Arial, sans-serif}
P {font-size:10pt;font-family:  Verdana,Arial, sans-seriff}
A.menu {text-decoration: none}
A:hover {color:yellow}
-->
</style>
<!-- #include file ="../lib/include.inc" -->
<%

REMOTE_ADDR =  Request.ServerVariables ("REMOTE_ADDR")
REMOTE_HOST =  Request.ServerVariables ("ALL_RAW")
HTTP_USER_AGENT = Request.ServerVariables ("HTTP_USER_AGENT")
HitDate = Date
HitTime = Time

QS = "Insert Into Hits (IP,HOST,HitDate,HitTime,AGENT) values ('"&REMOTE_ADDR&"','"&REMOTE_HOST&"','"&HitDate&"','"&HitTime&"','"&HTTP_USER_AGENT&"');"

Call CreateRS(RecordSet,QS,DSN)
Call CloseRS(RecordSet)
%>

</head>

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" leftmargin="3" topmargin="0">

<table border="0" width="141">
  <tr>
    <td width="100%"><img src="../images/icons/wsclogo.gif" alt="Home" border="0" WIDTH="123" HEIGHT="105"></td>
  </tr>
  <tr>
    <td width="100%"><a href="../theclub.htm" target="_top"><img src="../images/titles/theclub.gif" alt="The Club" border="0" WIDTH="113" HEIGHT="27"></a> </td>
  </tr>
  <tr>
    <td width="100%"><a href="skydiving.htm" target="_self"><img src="../images/titles/skydiving.gif" alt="Skydiving" border="0" WIDTH="127" HEIGHT="30"></a> </td>
  </tr>
  <tr>
    <td width="100%"><a href="../gallery.htm" target="_top"><img src="../images/titles/gallery.gif" alt="Gallery" border="0" WIDTH="94" HEIGHT="30"></a> </td>
  </tr>
  <tr>
    <td width="100%"><a href="funstuff.htm" target="_self"><img src="../images/titles/funstuff.gif" alt="Fun Stuff" border="0" WIDTH="113" HEIGHT="27"></a> </td>
  </tr>
  <tr>
    <td width="100%"><a href="../events.htm" target="_top"><img src="../images/titles/events.gif" alt="Events" border="0" WIDTH="83" HEIGHT="26"></a></td>
  </tr>
  <tr>
    <td width="100%"><a href="../guestbook.htm" target="_top"><img src="../images/titles/guestbook.gif" alt="Guestbook" border="0" WIDTH="141" HEIGHT="27"></a></td>
  </tr>
  <tr>
    <td width="100%" height="30"></td>
  </tr>
  <tr>
    <td width="100%"><h6 align="center"><!--webbot bot="HitCounter" i-image="0" i-digits="5" PREVIEW="&lt;strong&gt;[Hit Counter]&lt;/strong&gt;" u-custom i-resetvalue="4877" startspan --><img src="../_vti_bin/fpcount.exe/?Page=lists/default.asp|Image=0|Digits=5" alt="Hit Counter"><!--webbot bot="HitCounter" endspan i-checksum="13815" --><br>
    Skydivers and Counting</h6>
    </td>
  </tr>
  <tr>
    <td width="100%"><p align="center"><small><a href="../content/about_this_site.htm">About
    this site</a></small></td>
  </tr>
</table>
</body>
</html>
