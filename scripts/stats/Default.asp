
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<!-- #include file ="../../lib/include.inc" -->
<link href="/lib/styles/dz.css" rel="stylesheet" type="text/css">
<title></title>
</head>
<body link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">
<% ReqString = "Select * From Hits;" %>
<h4 align="center">Number of Hits : <%Response.Write CreateRS(RecordSet,ReqString,DSN)%></h4>


<%Call RSDisplay(RecordSet,"Stats","NavHead","NavTail","30")
CloseRS(RecordSet)

%>
</body>

</html>
