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



Private Sub TableArrayContents(ByVal QueryResults)

'-----------------------------------------------------------------
' Name : TableArrayContents(QueryResults) , sub
'
' Usage : Create a HTML of Values in a BiDimensional Array
' 
' Returns : Nothing
' Requires : QueryResults(*.*) Array
'
'
' Version 1.0
' Author : Brett Samuel
' Last Revision Date : 14-04-2000
'
'---------------------------------------------------------------

Dim IndexX
Dim IndexY

Response.Write "<Table>"	' Start Table

For IndexX = 1 to Ubound(QueryResults,1)

	Response.Write "<TR>"	 ' Start Row

		For IndexY = 1 to Ubound(QueryResults,2)
		Response.Write "<TD>"	 ' Start Col
		    
			If  QueryResults(IndexX-1,IndexY-1) = "" then
				Response.Write "Null"
			else
				Response.Write QueryResults(IndexX-1,IndexY-1)
			end if
		
		Response.Write "</TD>"	 ' end Col
		Next

	Response.Write "</TR>"	'End Row

Next

Response.Write "</Table>"	' End Table

end Sub


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Function RSColData(ByVal ColName,ByVal RecordSet)


'-----------------------------------------------------------------
' Name : RSColData(ColName,RecordSet) , function
'
' Usage : Returns colum data from record set at current row
' 
' Returns : Columset Data for current row
' Requires : Colname (Str) Containing name of required colum.
'			- RecordSet (RS) Recordset to be interograted
'
'
' Version 1.0
' Author : Brett Samuel
' Last Revision Date : 16-04-2000
'
'---------------------------------------------------------------
On error resume next

If Not IsEmpty(RecordSet) And Not (RecordSet Is Nothing) Then RSColData = CStr(RecordSet(ColName))

End Function

 

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Sub CloseRS(RecordSet)

'-----------------------------------------------------------------
' Name : CloseRS(RecordSet)
'
' Usage : Delete RecordSet
' 
' Input : RecordSet
' 
'
' Version 1.0
' Author : Brett Samuel
' Last Revision Date : 16-04-2000
'
'---------------------------------------------------------------


Set RecordSet = Nothing	'Destroy Created Objects
Set Connection = Nothing 

end Sub


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------


Private Function RmvCharStr(ByVal Str,Chars)
'
' RmvCharStr(Str,Chars)
'
'
' Remove Selected Charactors From String
' Usefull for removing Special Formatting
'
' Output : New String without selectedchars
' Inputs : 
'
'
' Version 1.0
' Author : Brett Samuel
' Last Revised : 16-04-2000
'


StrPosition = 0						' Set Starting Position In String
StrTmp = ""

Do until StrPosition = Len(Str)
	StrPosition = StrPosition + 1 'Inc

    CharsPositon = 0					' Loop through possible chars				'
	SearchChar = ""
	
	CharFound = "FALSE"
  Do until CharsPositon = Len(Chars)

	
   	 CharsPositon = CharsPositon + 1
	 SearchChar = Mid(Chars,CharsPositon,1)
	 

	  If Mid(Str,StrPosition,1) = SearchChar then	' If No match found add character to string
		
	   CharFound = "TRUE"
	  end If
Loop
	
   If CharFound = "FALSE" then  
    StrTmp=StrTmp&Mid(Str,StrPosition,1)
   End If
  
Loop

RmvCharStr = StrTmp ' Replace Origional String with New one

End Function



'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------



Private Sub DeformatIntoArray(byref FFArray,byVal Delimiter,ByVal FFArrayDestRow)

' Remove Delimiters from Fields in Row and Set the destination Fow to True if Found
'
'
' Inputs:	FFAray - Array with Formatted Fields in Row 0
'			Delimiter - Surround Delimiter to Be removed
'			FFArrayDestRow - Destination Row For True / False Marker
'
' Outputs:	FFAray with outer delimiters removed
'			True / False in destination Row of Array
'
'
'
' Version 1.0
' Author Brett Samuel 
' Last Modified 18-04-2000
'



Do While FFArrayCol < Ubound(FFArray,2)

If FFArray(FFArrayDestRow,FFArrayCol) <> True then ' test if field has been already checked

	If Mid(FFArray(0,FFArrayCol),1,1) = Delimiter then	'If marker found
		FFArray(FFArrayDestRow,FFArrayCol) = True				'Note Field Type 
	
		  	'Remove Marker
			FFArray(0,FFArrayCol) = Mid(FFArray(0,FFArrayCol),2,(Len(FFArray(0,FFArrayCol))-2))
	
			else
		FFArray(FFArrayDestRow,FFArrayCol) = False
	End If

end If

FFArrayCol = FFArrayCol + 1

Loop

end Sub


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Sub ProcessFormFields(ByVal FormFields,ByRef FFArray)


' Create FormFieldsArray containg form fields passed and information about them
' in rows as documented below:
'
'  Dependencies:	 DeformatIntoArray()
'
'
'
' Inputs : FormFields (Str) String containg fields to be processed, 
'							Delimiters <> then ! = unique, then ' = FormField 
'							ie : "'firstane',<!'lastname'!>,id"
'
'			FFArray - Array must already be dimetioned
'
'
'
' 
' Outputs:	FFArray: (x,y)
'						x
'					  012345678
'					 0vut<>
'				y	 1vut<>
'					 2vut<>
'					 3vut<>
'					 4vut<>				
'					 5vut<>
'					
'
'					 v: Form Filelds Broken Into Elements
'					 u: v = Unique
'					 t: v = text 
'					 <: v = Sort Ascending
'					 >: v = Sort Descending
'

'
' Version : 1.0
' Author : Brett Samuel
' Last Modified : 18-04-2000
'

Dim FFLen,FFPos,MarkerPos,FFStrLen, FFArrayCol, Rows

FFPos=1
FFArrayCol = 0
Attribs = 5 ' Number of Attributes ( x axiz)

ReDim FFArray(Attribs,0) 'Dimension FFAray


FFLen = Len(Formfields) ' Determine Len Of FF


' -= Popluate Row 0

Do until MarkerPos = FFLen

MarkerPos = Instr(FFPos,Formfields,",")

FFStrLen = MarkerPos - FFPos

	If MarkerPos = 0 then ' No Marker Found
	 MarkerPos = FFLen
	 FFStrLen = (MarkerPos - FFPos) + 1
	end if
	
FFArrayCol =  FFArrayCol + 1

Redim Preserve FFArray(Attribs,FFArrayCol) ' Add Element to Array
FFArray(0,FFArrayCol-1) = Mid(Formfields,FFPos,FFStrLen) 'Write String To Array

FFPos = MarkerPos + 1 ' Inc 




loop

' Remove Text / unique Delimiters from Array

Call DeformatIntoArray(FFArray,">","4") ' Outers First
Call DeformatIntoArray(FFArray,"<","3") ' Then Inners
Call DeformatIntoArray(FFArray,"!","1")
Call DeformatIntoArray(FFArray,"'","2")


end sub

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------
Private Function Delimit(ByVal FFString,ByVal Delimiter)

' Add Delimiters to beginning and end of string
'
' Inputs -	FFstring ' String to be Processed
'			Delimiter ' String to add
'
' outputs-	Delimited String
'
'
' Version : 1.0
' Author : Brett Samuel
' Last Modified : 18-04-2000
'

Delimit =  Delimiter&FFString 'Add Delimiter to beginning
Delimit = Delimit& Delimiter ' Add Delimiter to end

end Function


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------


Private Function GenTranIDInput(TableName,DSN)

' Generate a random number for use as a transaction ID and check that the number
' is unique within the table that is being queried.
'
' Format the Resulting data as HTML for use as a hidden input in a form.
'
'
' Dependencies : CreateRS()
'
' Inputs:	TableName (Str) ' Name of table to be queried
'			DSN (Str)	' Connection String
'
' Output	HTML > <input type='hidden' name='TranID' value='xxxxxx'>"
'			GenTranIDInput = TranID
' 
'
' Author : Brett Samuel
' Last Modified : 19-04-2000
' Version : 1.0
'

Dim TranID
Dim Unique
Dim ReqString

Unique = False

Randomize			' Initialize random-number generator.


Do Until Unique = True

TranID = Int((200000 * Rnd) + 1)	' Generate random value between 1 and 200000

ReqString = "Select * from Contacts Where (((TranID)=" & TranID & "));"


	If CreateRS(RecordSet,ReqString,DSN) = "Nothing" then
	 Unique = True
	end if

loop

Response.Write "<input type='hidden' name='TranID' value='" & TranID & "'>"
GenTranIDInput = TranID
end Function


 
'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Function GenSqlSelect(TableName,FormFields,TranID)
'
' Generated SlectString For basic searches. 
'
'
' Dependencies :Delimit()
'				ProcessFormFields()
'
' inputs:	TableName ( Str) name of table to be searched                 
'			FormFields (Str) String containg fields to be processed, 
'							Delimiters <> then ! = unique, then ' = FormField 
'							ie : "'firstane',<!'lastname'!>,id"
'
'			TranID (Str) Name of form field containig TrainID.
'
' Outputs:	SqlSelectString and GenSqlSelect
'
'
'
' Version 1.03
' Author : Brett Samuel
' Last Modified : 26-04-2000
'
Dim FFArray() ' Dimension array as Required by ProcessFormFields()
Call ProcessFormFields(FormFields,FFArray) ' Unwrap Colums String into FormColumsArray
Dim Sep
Dim FFcol
Dim WhereSeg , Where , WhereSegVal
Dim SortField , SortType , SortString
Dim WhereStrStart, WhereStrEnd

TranID = Request.QueryString(TranID) ' Check for Transaction ID

FirstValue = True '
FFcol = 0

If TranID <> "" then
FirstValue = False	' Set First falose to that Seperator will be generated
Where = "((TranID)=%%TranID%%)"
end if

Do Until FFrow = (Ubound(FFArray,2)) ' Startting at colum zero until last colum loop

	If FirstValue = True then 'If not first Value , as Sep
	 Sep = ""
	 FirstValue = False
	else 
	 Sep = " AND "
	end If

IF 	FFArray(1,FFrow) = True then ' Check if Field must be unique if not .. skip
	
WhereSegVal = FFArray(0,FFrow) ' Get Value From DB Row 0
WhereSeg = Sep & "((" & WhereSegVal & ")=" ' Create First Part of String
WhereSegVal = Delimit(WhereSegVal,"%%") ' Add Form Field Delimiters

If FFArray(2,FFrow) = True then ' Check if Vaule was text
WhereSegVal = Delimit(WhereSegVal,"'")	' add delimters if true
end If


WhereSeg = WhereSeg & WhereSegVal & ")"' Combine and Add Bracket
Where = Where + WhereSeg ' Add this Segement to where String

end if ' end of required check

If FFArray(3,FFrow) = True then ' Check for Sort Field
SortField = FFArray(0,FFrow)
SortType = ""
end If

If FFArray(4,FFrow) = True then ' Check for Sort Field
SortField = FFArray(0,FFrow)
SortType = "DESC"
end If

FFrow = FFrow + 1 ' Inc Counter

loop

If SortField <> "" then
SortString = "ORDER BY " & SortField & " " & SortType

end if

If Where = "" then			'Create String Containign the WHERE keyowrk and surrounign brackets 
WhereStrStart = " "			'only if where data available
WhereStrEnd = " "
else 
WhereStrStart = " WHERE ("
WhereStrEnd = ") "
end if


GenSqlSelect = "SELECT * FROM " & TableName &  WhereStrStart & Where & WhereStrEnd & SortString & ";"


end Function




'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Function CreateRS(RecordSet,ByVal ReqString,ByVal ConnString)

' CreateRS(RecordSet,ReqString,ConnString)
'
' Create Record Set Based On Query String. Query string can contain form fields if 
' delimited by %%Field%%. Query string should also include the following substitutions:
' " = '
' * = %

' Requires : QueryString()
'
'
'
'
' Inputs	:
'			- ReqString (Str) 
'			- ConnString (Str) ' Db Connection string ie: (DSN=MyDB)
'
' Outputs	:
'			- RecordSet (RS Object)
'			- CreateRS = Number of lines in RS or Nothing If DB Empty
'
' Version 1.01
' Author : Brett Samuel
' Last modified : 29-04-2000

On error resume next	' Allow program to continue in error state



'Set Inital Values

set RecordSet = server.createobject("adodb.recordset") ' Set Objects
set Connection = server.createobject("adodb.connection")

Connection.open ConnString
RecordSet.open QueryString(ReqString),Connection


If Err.Description <> "" Then					' Error Reporting
		Response.Write "<B>Database Error: " + Err.Description + "</B>"
End If ' Close err <> = "" loop

If RecordSet.EOF And RecordSet.BOF Then 
   CreateRS =  "Nothing"
else

CreateRS = 0

	Do until RecordSet.EOF	' Walk RecordSet to obtain number of records	
							' **** THERE HAS TO BE A BETTER WAY OF DOING THIS ****
	RecordSet.MoveNext
	CreateRS = CreateRS + 1
	Loop

RecordSet.MoveFirst ' Reset RecordSet


end If

End Function




'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Function QueryString(ByVal ReqString)

'
'  Strip Away All %% from String and SubStitute Valued from Form Fields
'  Note: Form Must Use 'Get' method
' 
' QueryString(ReqString) (Function)
'
'  Reqires : ReqString - Formatted SQL statement using %% for fiels delimeters 
'			- % can be used as wildcard
'			- ' must be used in preference to " for text fields
' 
' Returns : Formatted SQl Statement with ASP GET form files substituted for %% delited Fileds
'
' Version : 1.02
' Author : Brett Samuel ( Adapted from FP98 Built in Script)
' Last Modified : 03-05-2000
'
' BUGS: Removes all ' chars from request.
' Solution ... massive workaround .... or dont send that charactor.
'


Dim iCurrent ' Dimension Indexes
Dim iStart
Dim iEnd
Dim iError
Dim sField 'Searched Field
Dim sWildCard
Dim sValue 'Value Returned From Form
Dim bFoundField


sDefault = ""
iMaxRecords = 0
iTimeout = 0
iCurrent = 1
fError = False
bBlankField = False

Do While (InStr(iCurrent, ReqString, "%%") <> 0)
	
	iStart = InStr(iCurrent, ReqString, "%%")	' Something Found
	iEnd = InStr(iStart + 2, ReqString, "%%")	' Set Start And End Locations
	
	
	If iEnd = 0 Then	' Something is Wring With Delimeters
		 iError = True
		Response.Write "<B>Delimeters in querystring are Bad</B>"
	Else
		sField = Mid(ReqString, iStart + 2, iEnd - iStart - 2) ' Set Sfield to Searched Field
		If Mid(sField,1,1) = "%" Then ' If Wild Card Exists Set sWildcard to %
			sWildcard = "%"
			sField = Mid(sField, 2)
		Else
			sWildCard = ""
		End If
		
		sValue = RmvCharStr(Request.QueryString(sField),"'") ' Get Value from form
															 ' But remove all "'"'s
		
		

		' if the named form field doesn't exist, make a note of it
		If (len(sValue) = 0) Then
			iCurrentField = 1
			bFoundField = False
			
			Do While (InStr(iCurrentField, sDefault, sField) <> 0) And Not bFoundField
				iCurrentField = InStr(iCurrentField, sDefault, sField)
				iStartField = InStr(iCurrentField, sDefault, "=")
				If iStartField = iCurrentField + len(sField) Then
					iEndField = InStr(iCurrentField, sDefault, "&")
					If (iEndField = 0) Then iEndField = len(sDefault) + 1
					sValue = Mid(sDefault, iStartField+1, iEndField-1)
					bFoundField = True
				Else
					iCurrentField = iCurrentField + len(sField) - 1
				End If
			Loop
			
		End If

		If (len(sValue) = 0) Then bBlankField = True

		ReqString = Left(ReqString, iStart - 1) + sWildCard + sValue + Right(ReqString, Len(ReqString) - iEnd - 1)
		
		' Fixup the new current position to be after the substituted value
		iCurrent = iStart + Len(sValue) + Len(sWildCard)
	End If
Loop


QueryString = ReqString	' Write Value out to function

end Function


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------


Private Function InsertFormData(ByVal IDName,ByVal FormFields,ByVal TableName,ByVal DSN)



' Function To Insert Data from Form into DB and return ID
'
' Dependencies:	DeformatIntoArray()
'				Delimit()
'				GelSqlSelect()
'				GenSqlInsert()
'				ProcessFormFields()
'
'
' Inputs : FormFields (Str) String containg fields to be processed, 
'							Delimiters <> then ! = unique, then ' = FormField 
'							ie : "'firstane',<!'lastname'!>,id"
'
'
' Output : ID of inserted Line 
'
' Format Insert String
' Insert Line into Database
' Search for Transaction ID in order to return ID

'
' Version : 1.01
' Author : Brett Samuel
' Last Modified : 30-04-2000



' Check if Transaction ID already Exist
If CreateRS(RecordSet,GenSqlSelect(TableName,"","TranID"),DSN) = "Nothing" Then
	Call CloseRS(RecordSet) ' Delete RS object

		' Transaction has not been added, checked for uniqueness and then add
		' Do this bysearching for curret field based on unique fields,
		' If no unique fields then proceed to insert into DB
		If CreateRS(RecordSet,GenSqlSelect(TableName,FormFields,""),DSN) = "Nothing" or Instr(1,FormFields,"!") = 0 Then
		Call CloseRS(RecordSet) ' Delete RS object
		
		Call CreateRS(RecordSet,GenSqlInsert(TableName,FormFields),DSN) ' Insert Line Into DB
		Call CloseRS(RecordSet)
		
		Call CreateRS(RecordSet,GenSqlSelect(TableName,"","TranID"),DSN) ' Get ID of Inserted Line
		InsertFormData = RSColData(IDName,RecordSet)
		Call CloseRS(RecordSet)
	
	
		 
		else
		 ' record Already Exists, return ID
		 InsertFormData = RSColData(IDName,RecordSet)
		 Call CloseRS(RecordSet) ' Delete RS object	
		end if
		
else
	
	' Transaction has already been added to Db , obtaib ID
	InsertFormData = RSColData(IDName,RecordSet)
	Call CloseRS(RecordSet) ' Delete RS 
	
end if


end Function

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------
Function GenSqlInsert(TableName,FormFields)
'
' Dependencies :Delimit()
'				ProcessFormFields()
'				CallNavBar()
'

' Function To Generate SQL insert String for form fields.
'
' Inputs : FormFields (Str) String containg fields to be processed, 
'							Delimiters <> then ! = unique, then ' = FormField 
'							ie : "'firstane',<!'lastname'!>,id"
'
'							Hyphen (-) in fieldname means Redirestion 
'							Format: FORMFIELD--DBFIELD
'							Ie: <!'formlastname-dblastame'!>
'
'							WARNING : Redirected fields ore not check for NON-Blank 
'
'
'			TableName ( Str) name of Table
'			
'			TranID (Form Field) Unique Transaction ID to be added to Tran ID Col
'					This field should be imbedded into a from 
'
'
'
'
'
'
' Version 1.10
' Author : Brett Samuel
' Last Modified : 01-05-2000
'
Dim FFArray() ' Dimension array as Required by ProcessFormFields()
Call ProcessFormFields(FormFields,FFArray) ' Unwrap Colums String into FormColumsArray
TranID = Request.QueryString("TranID")

Dim Fields, Values, FirstValue, FFrow, CurField, CurValue, Sep

FirstValue = True '
FFcol = 0

If TranID <> "" then ' If transaction value sent throught then add transaction filld to db
FirstValue = False ' Transaction ID will be first Field
Fields = "TranID"
Values = TranID
end if


Do Until FFrow = (Ubound(FFArray,2)) ' Startting at colum zero until last colum loop


If Request.QueryString(FFArray(0,FFrow)) <> "" or Instr(1,FFArray(0,FFrow),"--") > 0 then 'Only populate array if vaule in form or FF contains -- 


	If FirstValue = True then 'If not first Value , as Sep
	 Sep = ""
	 FirstValue = False
	else 
	 Sep = ","
	end If

'Format Colum Name
CurField = FFArray(0,FFrow) 'Get Value From Array
	
	If Instr(1,CurField,"--") > 0 then 'CHeck for Redirection Marker
	CurField = Mid(CurField,(Instr(1,CurField,"--")+2),Len(CurField))
	end If

'Format Colum Value
CurVal = FFArray(0,FFrow)'Get Value From Array

	If Instr(1,CurVal,"--") > 0 then 'CHeck for Redirection Marker
	CurVal = Mid(CurVal,1,(Instr(1,CurVal,"--")-1))
	end If
	
CurVal = Delimit(CurVal,"%%")	' add form field delimiters


If FFArray(2,FFrow) = True then ' Check if Vaule was text
CurVal = Delimit(CurVal,"'")	' add delimters if true
end If

Fields = Fields & Sep & CurField ' Compile into string
Values = Values & Sep & CurVal


End if ' End form value check

FFrow = FFrow + 1 ' Inc Counter
loop


 GenSqlInsert = "Insert into " & TableName & "(" & Fields & ") values (" & Values & ");"


end Function


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------


Private Function RSDisplay(RecordSet,ByVal Template,ByVal HeaderName,ByVal TailName,ByVal DisplayLines)


' Display The Contents of a record set in HTML
'
'  RSDisplay(RecordSet,ContentName,Header,TailName)
'
'
' Requires : Displays() Sub
'			DisplayNavBar()
'
' Inputs	:
'			- RecordSet (RS Object) ' RS to be queried
'			- HeaderName (Str) ' Name of header in Displays()
'			- Template (Str) ' Name of Content Template in Displays()
'			- TailName (Str) ' Name of Tail  in Displays() use NavBar as tail for Scrolling pages
'			- DisplayLines (Str) ' Number of lines to display / page
'
'			If Header / Tail Start with "Nav" then then DisplayNavBar will be called
'
' Outputs	: To Display 
'
'			- ==> Header
'			- ==> Template
'			- ==> Template ...
'			- ==> Footer
'
'	RSDISPLAY = Nothing If nothing Returned or Number of lines in recordset
'
' Version 1.2
' Author Brett Samuel skydive@dropzone.co.za
' Last UpDate 11-05-2000



'----------------------------------------------------------------
' S T A R T   D E T E R M I N E   P A G E   R A N G E S
'----------------------------------------------------------------

' Values to output
Dim PageNumber
Dim LastPageNumber
Dim NumberOfRecords
Dim FirstRowOnCurrentPage
Dim LastRowOnCurrentPage

' Working Variables

Dim PrevPageNumber

IF DisplayLines = "" then DisplayLines = 0	' Avoid null value generating error

'----------------------------------------------------------------

' Determine number of lines in DB

NumberOfRecords = 0 'Set to Zero

Do While Not RecordSet.EOF 
 NumberOfRecords = NumberOfRecords + 1 ' Count number of lines in recordset
 RecordSet.MoveNext
Loop

RecordSet.MoveFirst ' Reset RecordSet


'----------------------------------------------------------------

' Determine that Last page number

If DisplayLines = 0 then ' DisplayLInes = 0 then all lines must me included on one page

LastPageNumber = 1

else

	LastPageNumber = Int(NumberOfRecords / DisplayLines)  'Determine Number of full pages

	If (NumberOfRecords / DisplayLines) > LastPageNumber then ' Add an extra page if not completely full
	 LastPageNumber = LastPageNumber + 1
	end If	

End If

'----------------------------------------------------------------

' Find Current Page

'First Find Last Page 
If Request.QueryString("PageNumber") = "" then	' If no page number in querystring then assume first
 PrevPageNumber = 0
else
 PrevPageNumber = Request.QueryString("PageNumber") ' Else add one to this request
end if

' Find Direction of Movement adjust Value Accordingly
Select Case Request.QueryString("Move")
	Case "|<"
	 PageNumber = 1
	Case "<<"
	 PageNumber = PrevPageNumber -1
	Case ">>"
	 PageNumber = PrevPageNumber + 1
	Case ">|"
	 PageNumber =  LastPageNumber
	case Else
	 PageNumber = 1
end Select

'----------------------------------------------------------------

' Determine The First Row on Current Page 

If DisplayLines = 0 then ' DisplayLInes = 0 then all lines must me included on one page

	FirstRowOnCurrentPage = 1
else
	FirstRowOnCurrentPage = ((PageNumber - 1) * DisplayLines) + 1
End If	
	
'----------------------------------------------------------------

' Determine The Last Row on Current Page 

If DisplayLines = 0 then ' DisplayLInes = 0 then all lines must me included on one page

	LastRowOnCurrentPage = NumberOfRecords
else
	LastRowOnCurrentPage = (FirstRowOnCurrentPage + DisplayLines) - 1

	If LastRowOnCurrentPage >  NumberOfRecords then	' Last Page in DB may not be full
	 LastRowOnCurrentPage = NumberOfRecords
	End If
	
End If


'----------------------------------------------------------------
'  E N D    D E T E R M I N E   P A G E   R A N G E S			|
'																|
'    S T A R T   D I S P L A Y   D A T A						|
'----------------------------------------------------------------
Dim LineNumber

If RecordSet.EOF And RecordSet.BOF Then ' If nothing returned dont query RS
   RSDisplay =  "Nothing"
  else
	'---------------H E A D   S T A R T -------------------------------
	If HeaderName <> "" then 
		If Mid(HeaderName,1,3) = "Nav" then 'Check if Headername starts with NAV
		 Call DisplayNavBar(PageNumber,LastPageNumber,HeaderName)
		else
		 Call Displays(HeaderName,RecordSet,DisplayLines) ' Display Header
		End If
	End IF
	'--------------H E A D   E N D -------------------------------------

	'--------------C O N T E N T   S T A R T -----------------------------
			Do While Not RecordSet.EOF 
		 LineNumber = LineNumber + 1 
		 
		 If LineNumber >= FirstRowOnCurrentPage and LineNumber <= LastRowOnCurrentPage then
			Call Displays(Template,RecordSet) 'Display Tempate / Loop
		 end If

		 RecordSet.MoveNext
		Loop
		RecordSet.MoveFirst ' Reset RecordSet	
	'------------C O N T E N T   E N D ----------------------------------
	
	'----------------T A I L   S T A R T---------------------------------
		If TailName <> "" then 
			If Mid(TailName,1,3) = "Nav" then 'Check if Headername starts with NAV
			 Call DisplayNavBar(PageNumber,LastPageNumber,TailName)
			else
			 Call Displays(TailName,RecordSet,DisplayLines) ' Display Tail
			End If
		End IF
	'----------------T A I L  E N D----------------------------------------
		
  End if	' Close No Data Loop	
  
 RSDisplay = NumberOfRecords ' Return Number Of Records For Backwards Compatibility
 

'----------------------------------------------------------------
'   E N D   D I S P L A Y   D A T A
'----------------------------------------------------------------
 
end Function

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------


Private Sub DisplayNavBar(ByVal PageNumber,ByVal MaxPages,ByVal DisplayName)

' Display NavBar '
'
' Requires : NavDisplays() 
'			
'
' Inputs	:
'			- PageNumber ( Current Page Number)
'			- MaxPages ( Last Page)
'			- DisplayName (Str) ' Name of header in Displays()
'
' Outputs	: To Display 
'
' Version 1.1
' Author Brett Samuel skydive@dropzone.co.za
' Last UpDate 11-05-2000



QS = Request.QueryString ' Create QS as copy of origional query String

If InStr(1,QS,"Move") = 1 then ' Remove Last Move Position from QueryString
QS = Mid(QS,1,InStr(1,QS,"Move") - 1)
elseIf InStr(1,QS,"&Move") > 0 then
QS = Mid(QS,1,InStr(1,QS,"&Move") - 1)
end if

if QS = "" then	' If no data in remainign query string DO NOT prefix with & symbol
Delimiter = ""
else
Delimiter = "&"
end if

%>
  
<Table width="100%">
<TR>
<TD colspan="4">
<P align="center" class="NavStatTop">
<%Call NavDisplays(DisplayName,"STATUSBARTOP",PageNumber,MaxPages)%>
</TD>
</TR>
<TR>
<TD align="center" width="25%">
<%If PageNumber > 1 then%>
<a href="<%Response.Write Request.ServerVariables("URL")&"?" & QS &  Delimiter & "Move=|<" & "&PageNumber=" & PageNumber%>">
<%Call NavDisplays(DisplayName,"ON_FIRST",PageNumber,MaxPages)%></a>
<%else%>
<%Call NavDisplays(DisplayName,"OFF_FIRST",PageNumber,MaxPages)%>
<%End If%>
</TD>
<TD align="center" width="25%">
<%If PageNumber > 1 then%>
<a href="<%Response.Write Request.ServerVariables("URL")&"?" & QS & Delimiter & "Move=<<" & "&PageNumber=" & PageNumber%>">
<%Call NavDisplays(DisplayName,"ON_PREV",PageNumber,MaxPages)%></a>
<%else%>
<%Call NavDisplays(DisplayName,"OFF_PREV",PageNumber,MaxPages)%>
<%End If%>
</TD>
<TD align="center" width="25%">
<%If PageNumber < MaxPages then%>
<a href="<%Response.Write Request.ServerVariables("URL")&"?" & QS & Delimiter & "Move=>>" & "&PageNumber=" & PageNumber%>">
<%Call NavDisplays(DisplayName,"ON_NEXT",PageNumber,MaxPages)%></a>
<%else%>
<%Call NavDisplays(DisplayName,"OFF_NEXT",PageNumber,MaxPages)%>
<%End If%>
</TD>
<TD align="center" width="25%">
<%If PageNumber < MaxPages then%>
<a href="<%Response.Write Request.ServerVariables("URL")&"?"& QS &  Delimiter & "Move=>|" & "&PageNumber=" & PageNumber%>">
<%Call NavDisplays(DisplayName,"ON_LAST",PageNumber,MaxPages)%></a>
<%else%>
<%Call NavDisplays(DisplayName,"OFF_LAST",PageNumber,MaxPages)%>
<%End If%>
</TD>
</TR>
<TR>
<TD colspan="4">
<P align="center" class="NavStatBottom">
<%Call NavDisplays(DisplayName,"STATUSBARBOTTOM",PageNumber,MaxPages)%>
</TD>
</TR>
</Table>
    
<%

End Sub

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Function CursorLocaton(RecordSet,IndexName)


' Find the Cursor Location of a RecordSet
'
' Inputs :	RecordSet - The Recordset in question
'			IndexName - Name if Index Colum in Array
'
' Outputs:	Current Row Number of RS
'
'
'
' Version: 1.0
' By: Brett Samuel
' Date: 13-05-2000


Dim PositionDelta
Dim IDMarker


IDMarker = RSColData(IndexName,RecordSet) ' Find Marker
RecordSet.MoveFirst ' Move To Beginning of RS
PositionDelta = 1 'Set Starting Position

Do Until RSColData(IndexName,RecordSet) = IDMarker  ' Walk through RS
 PositionDelta = PositionDelta + 1
 RecordSet.MoveNext
Loop

CursorLocaton = PositionDelta

End Function

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Sub SetCursor(RecordSet,ByVal RecordPos)


' Move RecordSet to Specific Locatin
'
' Inputs :	RecordPos - Location to move recordSet
'			RecordSet - RS to move
'
' By : Brett Samuel
' Ver: 1.0
' Date 13-05-2000


Dim PositionDelta
PositionDelta = 1

RecordSet.MoveFirst ' Move To Beginning of RS



Do Until PositionDelta >= Int(RecordPos) or RecordSet.EOF
PositionDelta = PositionDelta + 1
 RecordSet.MoveNext
Loop

End Sub


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

%>
