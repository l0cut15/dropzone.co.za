
<%Response.Buffer = TRUE%>
<!-- #include file ="../../lib/include.inc" -->

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link href="../../lib/styles/dz.css" rel="stylesheet" type="text/css">
</head>

<body style="font-family: Verdana, Arial, sans-serif; font-size: 10pt" bgcolor="#003399" text="#FFFFFF" link="#FFFFFF" vlink="#FFFFFF" alink="#FFFFFF">

<%




'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'|
'|    A V A I L A B L E   S U B S   A N D  F U N C T I O N S
'|
'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'| TableArrayContents(QueryResults) 'Draw table from array
'| Displays(DisplayName,RecordSet) ' Display RecordSet Data in Displa type set by displayname
'| DisplayNavBar(PageNumber,MaxPages) ' Display NavBar (HTML and code) -- Customizable
'|
'| DeformatIntoArray(FFArray,Delimiter,FFArrayDestRow) ' Remore Delimiters From Row 0
'| ProcessFormFields(FormFields,FFArray) Create Fields Array
'| Delimit(FFString,Delimiter) ' Delimit Strings
'|
'| 
'|
'| CreateRS(RecordSet,ReqString,DSN) 'Create recordset from query string in Null then = "Nothing"
'| CloseRS(RecordSet) ' Delete RS object
'| CursorLocaton(RecordSet,IndexName) ' Find Current location of RAS cursor
'| SetCursor(RecordSet,RecordPos) ' Move Cursor to Positon
'|
'|
'| QueryString(ReqString) (Function) ' Insert form values into query string
'| Function RmvCharStr(Str,Chars) ' Remove selected Chars from string
'| GenSqlInsert(TableName,FormFields) ' Generate insert String
'| GenSqlSelect(TableName,FormFields,TranID) ' Generated select String
'| GenTranIDInput(TableName,DSN) ' Generated Transaction ID in HTML
'| InsertFormData(IDName,FormFields,TableName,DSN)
'| 
'| PageRanges(RecordSet,ByVal DisplayLines) ' Returns String containg renges required to construct pages
'|
'| RSDisplay(RecordSet,ContentName,HeaderName,TailName,DisplayLines) ' Display Contents of RS in HTML
'| DisplayNavBar(PageNumber,MaxPages,DisplayName)
'| NavDisplays(DisplayName,DisplaySeg,PageNumber,MaxPages)
'|
'| RSColData(ColName,RecordSet) 'Return colum data from recordset
'|
'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\



TranID = Request.QueryString("TranID")
GalleryID = Request.QueryString("GalleryID")
ImageTitle = Request.QueryString("ImageTitle")
CameraPerson = Request.QueryString("CameraPerson")
Uploaded = "YES"
FileName = Request.QueryString("FileName") & Request.QueryString("FileExtention")

If Request.QueryString("FileSize") > 0 then

UpdateP = "Update Images SET "						' ommonly used String
UpdateS =  " WHERE (((TranID)=" & TranID & "));"


ReqString = UpdateP & "GalleryID=" & GalleryID & UpdateS
Call CreateRS(RecordSet,ReqString,DSN)
Call CloseRS(RecordSet)


ReqString = UpdateP & "CameraPerson='" & CameraPerson & "'" & UpdateS
Call CreateRS(RecordSet,ReqString,DSN)
Call CloseRS(RecordSet)


ReqString = UpdateP & "Uploaded='" & Uploaded & "'" & UpdateS
Call CreateRS(RecordSet,ReqString,DSN)
Call CloseRS(RecordSet)


ReqString = UpdateP & "FileName='" & FileName & "'" & UpdateS
Call CreateRS(RecordSet,ReqString,DSN)
Call CloseRS(RecordSet)

ReqString = UpdateP & "ImageTitle='" & ImageTitle & "'" & UpdateS
Call CreateRS(RecordSet,ReqString,DSN)
Call CloseRS(RecordSet)

End If

%>

<table border="0">
  <tr>
    <td width="100%" colspan="2" height="21">
      <h4 align="center">File Uploaded: </h4>
    </td>
  </tr>
  <tr>
    <td width="50%" height="18"></td>
    <td width="50%" height="18"></td>
  </tr>
  <tr>
    <td width="50%" height="18">File Name Uploaded:</td>
    <td width="50%" height="18"><%Response.Write Request.QueryString("FileName")
    Response.Write Request.QueryString("FileExtention")
    %></td>
  </tr>
  <tr>
    <td width="50%" height="18">Uploaded Size:</td>
    <td width="50%" height="18"><%Response.Write Request.QueryString("FileSize")%></td>
  </tr>
  <tr>
    <td height="18"></td>
    <td width="50%" height="18"></td>
  </tr>
  <tr>
    <td width="100%" height="18" colspan="2">
      <p align="center"><a href="/scripts/gallery/ShowGallery.asp?gallery=<%Response.Write Request.QueryString("GalleryID")%>">Go
      to Gallery</a></td>
  </tr>
</table>

</body>

</html>
