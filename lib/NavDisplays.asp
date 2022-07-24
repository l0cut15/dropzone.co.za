<%

'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------

Private Sub NavDisplays(ByVal DisplayName,ByVal DisplaySeg,PageNumber,MaxPages)

' Editable Subs containing components to make up navbars
'
' Inputs :	DisplayName - Group of componets selected
'			DisplaySeg - Subcomponents consisting of
'						- ON_FIRST
'						- ON_PREV
'						- ON_NEXT
'						- ON_LAST
'						- OFF_FIRST
'						- OFF_PREV
'						- OFF_NEXT
'						- OFF_LAST
'						- STATUSBARTOP
'						- STATUSBARBOTTOM
'
'						- PageNumber
'						- Maxpages
'
' Version : 1.0
' Author : Brett Samuel
' Date : 11-05-2000
'
'
'-----------------------------------------------------------------------------
Select Case DisplayName
'-----------------------------------------------------------------------------


'-----------------------------------------------------------------------------
Case "NavTail"
'-----------------------------------------------------------------------------


Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

Response.Write "|<"

	Case "OFF_FIRST"

Response.Write "|<"

	'        --------------          -----------
	Case "ON_PREV"


Response.Write "<<"

	Case "OFF_PREV"


Response.Write "<<"

	'        --------------          -----------
	Case "ON_NEXT"

Response.Write ">>"
	'        --------------          -----------
	Case "OFF_NEXT"

Response.Write ">>"

	'        --------------          -----------
	Case "ON_LAST"

Response.Write ">|"

	'        --------------          -----------
	Case "OFF_LAST"


Response.Write ">|"

	'        --------------          -----------
	Case "STATUSBARTOP"



	'        --------------          -----------

	Case "STATUSBARBOTTOM"
Response.Write "Page " & PageNumber & " of " & MaxPages

	'        --------------          -----------

	End Select


'-----------------------------------------------------------------------------
Case "NavHead"
'-----------------------------------------------------------------------------




Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

Response.Write "|<"

	Case "OFF_FIRST"

Response.Write "|<"

	'        --------------          -----------
	Case "ON_PREV"


Response.Write "<<"

	Case "OFF_PREV"


Response.Write "<<"

	'        --------------          -----------
	Case "ON_NEXT"

Response.Write ">>"
	'        --------------          -----------
	Case "OFF_NEXT"

Response.Write ">>"

	'        --------------          -----------
	Case "ON_LAST"

Response.Write ">|"

	'        --------------          -----------
	Case "OFF_LAST"


Response.Write ">|"

	'        --------------          -----------
	Case "STATUSBARTOP"


Response.Write "Page " & PageNumber & " of " & MaxPages
	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select


'-----------------------------------------------------------------------------
Case "NavGuestHead"

'-----------------------------------------------------------------------------

Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

%><img border="0" src="../../images/NavBar/first.gif"><%


	Case "OFF_FIRST"


	'        --------------          -----------
	Case "ON_PREV"

%><img border="0" src="../../images/NavBar/prev.gif"><%


	Case "OFF_PREV"




	'        --------------          -----------
	Case "ON_NEXT"

%><img border="0" src="../../images/NavBar/next.gif"><%
	'        --------------          -----------
	Case "OFF_NEXT"



	'        --------------          -----------
	Case "ON_LAST"

%><img border="0" src="../../images/NavBar/last.gif"><%


	'        --------------          -----------
	Case "OFF_LAST"


	'        --------------          -----------
	Case "STATUSBARTOP"

Response.Write "Page " & PageNumber & " of " & MaxPages


	'        --------------          -----------

	Case "STATUSBARBOTTOM"


Response.Write "<HR>"
	'        --------------          -----------

	End Select

'-----------------------------------------------------------------------------
Case "NavGuestTail"

'-----------------------------------------------------------------------------

Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

%><img border="0" src="../../images/NavBar/first.gif"><%


	Case "OFF_FIRST"


	'        --------------          -----------
	Case "ON_PREV"

%><img border="0" src="../../images/NavBar/prev.gif"><%


	Case "OFF_PREV"




	'        --------------          -----------
	Case "ON_NEXT"

%><img border="0" src="../../images/NavBar/next.gif"><%
	'        --------------          -----------
	Case "OFF_NEXT"



	'        --------------          -----------
	Case "ON_LAST"

%><img border="0" src="../../images/NavBar/last.gif"><%


	'        --------------          -----------
	Case "OFF_LAST"




	'        --------------          -----------
	Case "STATUSBARTOP"



	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select


'-----------------------------------------------------------------------------
Case "NavHead"
'-----------------------------------------------------------------------------




Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

Response.Write "|<"

	Case "OFF_FIRST"

Response.Write "|<"

	'        --------------          -----------
	Case "ON_PREV"


Response.Write "<<"

	Case "OFF_PREV"


Response.Write "<<"

	'        --------------          -----------
	Case "ON_NEXT"

Response.Write ">>"
	'        --------------          -----------
	Case "OFF_NEXT"

Response.Write ">>"

	'        --------------          -----------
	Case "ON_LAST"

Response.Write ">|"

	'        --------------          -----------
	Case "OFF_LAST"


Response.Write ">|"

	'        --------------          -----------
	Case "STATUSBARTOP"


Response.Write "Page " & PageNumber & " of " & MaxPages
	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select


'-----------------------------------------------------------------------------
Case "NavPixHead"

'-----------------------------------------------------------------------------

Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

%><img border="0" src="../../images/NavBar/first.gif"><%


	Case "OFF_FIRST"


	'        --------------          -----------
	Case "ON_PREV"

%><img border="0" src="../../images/NavBar/prev.gif"><%


	Case "OFF_PREV"




	'        --------------          -----------
	Case "ON_NEXT"

%><img border="0" src="../../images/NavBar/next.gif"><%
	'        --------------          -----------
	Case "OFF_NEXT"



	'        --------------          -----------
	Case "ON_LAST"

%><img border="0" src="../../images/NavBar/last.gif"><%


	'        --------------          -----------
	Case "OFF_LAST"


	'        --------------          -----------
	Case "STATUSBARTOP"

Response.Write "Page " & PageNumber & " of " & MaxPages


	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select

'-----------------------------------------------------------------------------
Case "NavPixTail"

'-----------------------------------------------------------------------------

Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"

%><img border="0" src="../../images/NavBar/first.gif"><%


	Case "OFF_FIRST"


	'        --------------          -----------
	Case "ON_PREV"

%><img border="0" src="../../images/NavBar/prev.gif"><%


	Case "OFF_PREV"




	'        --------------          -----------
	Case "ON_NEXT"

%><img border="0" src="../../images/NavBar/next.gif"><%
	'        --------------          -----------
	Case "OFF_NEXT"



	'        --------------          -----------
	Case "ON_LAST"

%><img border="0" src="../../images/NavBar/last.gif"><%


	'        --------------          -----------
	Case "OFF_LAST"




	'        --------------          -----------
	Case "STATUSBARTOP"



	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select
'-----------------------------------------------------------------------------
Case Else

'-----------------------------------------------------------------------------

Select Case DisplaySeg

	'        --------------          -----------
	Case "ON_FIRST"



	Case "OFF_FIRST"


	'        --------------          -----------
	Case "ON_PREV"




	Case "OFF_PREV"




	'        --------------          -----------
	Case "ON_NEXT"


	'        --------------          -----------
	Case "OFF_NEXT"



	'        --------------          -----------
	Case "ON_LAST"



	'        --------------          -----------
	Case "OFF_LAST"




	'        --------------          -----------
	Case "STATUSBARTOP"



	'        --------------          -----------

	Case "STATUSBARBOTTOM"



	'        --------------          -----------

	End Select




'-----------------------------------------------------------------------------
End Select
'-----------------------------------------------------------------------------



end Sub


'-------------------------------------------------------------------------------------
'***************************************************************************************
'-------------------------------------------------------------------------------------
%>