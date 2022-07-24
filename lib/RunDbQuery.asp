<%@ Language=VBScript%>

<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=windows-1252">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
</head>
<!-- #include file ="Include.inc" -->
<body>


<%
'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'|
'|    A V A I L A B L E   S U B S   A N D  F U N C T I O N S
'|
'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\
'| 
'|
'|
'| TableArrayContents(QueryResults) 'Draw table from array
'| 
'| RSColData(ColName,RecordSet) 'Return colum data from recordset
'| CloseRS(RecordSet) ' Delete RS object
'| Displays(DisplayName,RecordSet) ' Display RecordSet Data in Displa type set by displayname
'| DisplayNavBar(PageNumber,MaxPages) ' Display NavBar (HTML and code) -- Customizable
'|
'| DeformatIntoArray(FFArray,Delimiter,FFArrayDestRow) ' Remore Delimiters From Row 0
'| ProcessFormFields(FormFields,FFArray) Create Fields Array
'| Delimit(FFString,Delimiter) ' Delimit Strings
'|
'| CreateRS(RecordSet,ReqString,DSN) 'Create recordset from query string in Null then = "Nothing"
'| QueryString(ReqString) (Function) ' Insert form values into query string
'| Function RmvCharStr(Str,Chars) ' Remove selected Chars from string
'| GenSqlInsert(TableName,FormFields) ' Generate insert String
'| GenSqlSelect(TableName,FormFields,TranID) ' Generated select String
'| GenTranIDInput(TableName,DSN) ' Generated Transaction ID in HTML
'| InsertFormData(IDName,FormFields,TableName,DSN)
'| RSDisplay(RecordSet,ContentName,HeaderName,TailName,DisplayLines) ' Display Contents of RS in HTML
'|
'/////////////////////////////\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\\

DSN = "DSN=Access"
FormFields = "!'FirstName'!,<!'LastName'!>"
TableName = "Contacts"


StringVal =  Request.QueryString("Message")





' Call CreateRS(RecordSet,"Insert into Contacts (Message) values ('%%Message%%');",DSN)
'Call  RSDisplay(RecordSet,"FSL1","StatusBar","NavBar","7")
' Call CloseRS(RecordSet)




'
Call CreateRS(RecordSet,"Select * from Contacts",DSN)
Call  RSDisplay(RecordSet,"FSL1","StatusBar","NavBar","7")
Call CloseRS(RecordSet)




'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------
%>









</body>
</html>


