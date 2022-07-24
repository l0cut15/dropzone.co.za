<html>

<head>
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
<title>Skydive Witbank Gallery</title>
<script LANGUAGE="JavaScript">
	<!-- hide from old browsers
	

		self.focus()

	 function closePopUp()
	{
		self.close()
	}
	// hide  -->
	</script>
</head>

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399"
text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF" topmargin="1" leftmargin="1">
<%If Request.QueryString("imagename")= "" then 'bad query string %>

<h3 align="center">Oops,</h3>

<h5 align="center">We are unable to display the page that you requested.<br>
<br>
If you believe that this is a problem with our site then pleas e-mail the <a
href="mailto:freakfaller@dropzone.co.za">webmaster</a> so that this problem can be
resolved.<br>
</h5>
<%else 'bad query string %>
<div align="center"><center>

<table border="0">
  <tr>
    <td width="100%"><h5 align="center"><%Response.Write Request.QueryString("imagetitle")%></h5>
    </td>
  </tr>
  <tr>
    <td width="100%"><img src="<%Response.Write Request.QueryString("imagename")%>"
    onClick="closePopUp()" alt="Images Loading ..."></td>
  </tr>
  <tr>
    <td width="100%"><p align="right"><a href="javascript:self.close()"><small
    onClick="closePopUp()">Click image to close window.</small></a></td>
  </tr>
</table>
</center></div><%end if%>

</body>
</html>
