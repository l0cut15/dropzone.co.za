
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

<script language="JavaScript">

                <!-- Activate Stealth Mode

       browserName = navigator.appName;
       browserVer = parseInt(navigator.appVersion);
             if (browserVer >= 3) version = "n3";
               else version = "n2";
               

      
    function img_act(imageurl) {
               if (version == "n3") {
               window.open(imageurl,'gallery',config='height=460,width=620,toolbar=0,menubar=0,scrollbars=0,resizable=0,location=0,directories=0,status=0')
               }
       }
       
// -->

</script>

</head>
<!-- #include file ="../../lib/include.inc" -->

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">



<h3 align="center">



<%

Dim ReturnedRows
GalQuery = "SELECT Contacts.*, Galleries.GalleryName, Contacts.*, Images.* FROM (Images INNER JOIN Galleries ON Images.GalleryID = Galleries.GalleryID) INNER JOIN Contacts ON Images.ContactID = Contacts.ContactID WHERE (((Galleries.GalleryID)=%%gallery%%) AND ((Images.Uploaded)='YES')) ORDER BY Images.ImageID DESC;"

ReturnedImages = CreateRS(RecordSet,GalQuery,DSN)

%>

 There are <%= ReturnedImages%> images the  <%Response.Write RSColData("GalleryName",RecordSet)%> gallery</h3>



<p align="center">



<a href="../common/newuser.asp?exitpage=../upload/npost.asp"><b>Upload an Image</b></a></p>
<p align="center">
<%


Call RSDisplay(RecordSet,"Gallery","NavPixHead","NavPixTail","6") ' Display Contents of RS in HTML
Call CloseRS(RecordSet)

%>




</p>


</body>

</html>
