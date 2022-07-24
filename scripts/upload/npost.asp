<% Response.Buffer = TRUE %>
<!-- #include file ="../../lib/include.inc" -->
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../lib/styles/dz.css" rel="stylesheet" type="text/css">
</head>



<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

<!--webbot BOT="GeneratedScript" PREVIEW=" " startspan --><script Language="JavaScript"><!--
function FrontPage_Form1_Validator(theForm)
{

  if (theForm.GalleryID.selectedIndex < 0)
  {
    alert("Please select one of the \"Gallery\" options.");
    theForm.GalleryID.focus();
    return (false);
  }

  if (theForm.GalleryID.selectedIndex == 0)
  {
    alert("The first \"Gallery\" option is not a valid selection.  Please choose one of the other options.");
    theForm.GalleryID.focus();
    return (false);
  }

  if (theForm.ImageTitle.value == "")
  {
    alert("Please enter a value for the \"Image Title\" field.");
    theForm.ImageTitle.focus();
    return (false);
  }

  if (theForm.ImageTitle.value.length < 2)
  {
    alert("Please enter at least 2 characters in the \"Image Title\" field.");
    theForm.ImageTitle.focus();
    return (false);
  }

  if (theForm.CameraPerson.value == "")
  {
    alert("Please enter a value for the \"Camera Person\" field.");
    theForm.CameraPerson.focus();
    return (false);
  }
  return (true);
}
//--></script><!--webbot BOT="GeneratedScript" endspan --><form enctype="multipart/form-data"  action="http://<%Response.Write Request.ServerVariables("SERVER_NAME")%>/wpa/cpshost.dll?PUBLISH?http://<%Response.Write Request.ServerVariables("SERVER_NAME")%>/scripts/upload/repost.asp" method="post" onsubmit="return FrontPage_Form1_Validator(this)" name="FrontPage_Form1">
<% TranID = GenTranIDInput("Images",DSN) ' Generated Transaction ID in HTML
 

InsertQuery = "insert into Images (TranID,ContactID,UploadDate) values (" & TranID & "," & Request.QueryString("ContactID") & ",'"& Date &"');"
Call CreateRS(RecordSet,InsertQuery,DSN) ' Add Entry to DB
Call CloseRS(RecordSet)

Call CreateRS(RecordSet,"Select ImageID from Images WHERE ((TranID)=" & TranID & ");",DSN) ' Add Entry to DB
ImageID = RSColData("ImageID",RecordSet)
Call CloseRS(RecordSet)
%>
  <table border="0">
    <tr>
      <td colspan="3">
        <h3 align="center">Upload Image To Gallery</h3>
      </td>
    </tr>
    <tr>
      <td>Gallery</td>
      <td>
          <!--webbot bot="Validation" S-Display-Name="Gallery"
          B-Value-Required="TRUE" B-Disallow-First-Item="TRUE" -->
          <select size="1" name="GalleryID" tabindex="6">
          <option selected value="">--- Choose one ---</option>
          <%
           DropZones = CreateRS(RecordSet,"Select * From Galleries Order by GalleryName;",DSN)	' Display options for all dropzones
           Call RSDisplay(RecordSet,"GalleryOption","","","")					' in DZ's Table
           Call CloseRS(RecordSet)
          %>
          </select>
</td>
      <td>
          <i>*Required</i>
</td>
    </tr>
    <tr>
      <td>Image Title</td>
      <td><!--webbot bot="Validation" S-Display-Name="Image Title"
        B-Value-Required="TRUE" I-Minimum-Length="2" --><input type="text" name="ImageTitle" size="40"></td>
      <td><i>*Required</i></td>
    </tr>
    <tr>
      <td>Camera Person</td>
      <td><!--webbot bot="Validation" S-Display-Name="Camera Person"
        B-Value-Required="TRUE" --><input type="text" name="CameraPerson" size="20"></td>
      <td><i>*Required</i></td>
    </tr>
    <tr>
      <td></td>
      <td></td>
      <td></td>
    </tr>
    <tr>
      <td>Select An Image</td>
      <td><input type="file" name="My_File"></td>
      <td><i>*Required</i></td>
    </tr>
    <tr>
      <td colspan="3">
        <p align="center"><input type="submit" value="Start Uploading Image  *" name="B1"></td>
    </tr>
  </table>
  <input name = "TargetURL" type="hidden" value="http://<%Response.Write Request.ServerVariables("SERVER_NAME")%>/uploads/images/<%Response.Write ImageID%>" size = "40">
    <input name = "ContactID" type="hidden" value="<%Response.Write Request.QueryString("ContactID")%>" size = "40">
 
 

</form>

<p align="left">* This page requires Internet Explorer 4.01 and above. </p>
<p align="left">The next page may take some time to load depending on the size
of your image.</p>

<p align="left">All images upload to this site are considered public domain, if
your image have been uploaded without your permission please contact the <a href="mailto:webmaster@dropzone.co.za">webmaster</a>
and it will be immediately removed.</p>

</body>

</html>
