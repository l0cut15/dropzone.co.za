
<%

Cursor = CursorLocaton(RecordSet,"ImageID")
CursorDiv = Cursor / 2

If CursorDiv > Int(CursorDiv) then


%>

<div align="center">

<center>

<table border="0" width="100%">
  <tr>
    <td>
      <div align="left">
      <table border="0">
        <tr>
          <td width="100%" colspan="2" class="headercell" bgcolor="#800080">
            <p align="center" style="margin-left: 0"><b><%Response.Write RSColData("ImageTitle",RecordSet)%></b></p>
          </td>
        </tr>
        <tr>
          <td width="50%" rowspan="3">
           <a href="./enlarge.asp?filename=/uploads/Images/<%Response.Write RSColData("ImageID",RecordSet)%>/<%Response.Write RSColData("FileName",RecordSet)%>" alt="<%Response.Write RSColData("ImageTitle",RecordSet)%>" target="_blank"><img border="0" src="/uploads/Thumbs/<%
           If RSColData("Thumbnail",RecordSet) = "YES" then
            Response.Write RSColData("ImageID",RecordSet)%>/<%Response.Write RSColData("FileName",RecordSet)%>" alt="<%Response.Write RSColData("ImageTitle",RecordSet)
           else
            Response.Write "placeholder.gif"
           end if 
           %>"></a></td>
          <td width="50%" align="left"><b>Cameraperson</b><br>
            <%Response.Write RSColData("Cameraperson",RecordSet)%></td>
        </tr>
        <tr>
          <td width="50%" align="left"><b>Uploaded By</b><br>
            &nbsp; <%
       
    
    If RSColData("NickName",RecordSet) = "" then	' Display Nick if available
    Response.Write RSColData("FirstName",RecordSet)& " " & RSColData("LastName",RecordSet)
    else
    Response.Write RSColData("NickName",RecordSet)
    end if  
%></td>
        </tr>
        <tr>
          <td width="50%" align="left"><b>Date of Upload</b><br>
            &nbsp; <%Response.Write RSColData("UploadDate",RecordSet)%></td>
        </tr>
      </table>
      </div>
    </td>
<%
else
%>
    <td>
      <div align="left">
      <table border="0">
          <tr>
          <td width="100%" colspan="2" bgcolor="#800080">
            <p align="center" style="margin-left: 0"><b><%Response.Write RSColData("ImageTitle",RecordSet)%></b></p>
          </td>
        </tr>
        <tr>
                 <td width="50%" rowspan="3">
           <a href="./enlarge.asp?filename=/uploads/Images/<%Response.Write RSColData("ImageID",RecordSet)%>/<%Response.Write RSColData("FileName",RecordSet)%>" alt="<%Response.Write RSColData("ImageTitle",RecordSet)%>" target="_blank"><img border="0" src="/uploads/Thumbs/<%
           If RSColData("Thumbnail",RecordSet) = "YES" then
            Response.Write RSColData("ImageID",RecordSet)%>/<%Response.Write RSColData("FileName",RecordSet)%>" alt="<%Response.Write RSColData("ImageTitle",RecordSet)
           else
            Response.Write "placeholder.gif"
           end if 
           %>"></a></td>
          <td width="50%" align="left"><b>Cameraperson</b><br>
 <%Response.Write RSColData("Cameraperson",RecordSet)%></td>
        </tr>
        <tr>
          <td width="50%" align="left"><b>Uploaded By</b><br>
 <%
 
 
       
    
    If RSColData("NickName",RecordSet) = "" then	' Display Nick if available
    Response.Write RSColData("FirstName",RecordSet)& " " & RSColData("LastName",RecordSet)
    else
    Response.Write RSColData("NickName",RecordSet)
    end if
 
 
 %></td>
        </tr>
        <tr>
          <td width="50%" align="left"><b>Date of Upload</b><br>
            &nbsp; <%Response.Write RSColData("UploadDate",RecordSet)%></td>
        </tr>
      </table>
  
      </div>
  
    </td>
  </tr>
</table>

</center>

</div>

<%
End If

%>