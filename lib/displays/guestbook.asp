 <table border="0" width="100%">
  <tr>
    <td colspan="4">
      <p align="right"><b>Date: </b><%Response.Write RSColData("TranDate",RecordSet) & " " & RSColData("TranTime",RecordSet)%>
    </td>
  </tr>
  <tr>
    <td width="10%" nowrap><b>Name:&nbsp;</b></td>
    <td><%
    
    If RSColData("NickName",RecordSet) = "" then	' Display Nick if available
    Response.Write RSColData("FirstName",RecordSet)& " " & RSColData("LastName",RecordSet)
    else
    Response.Write RSColData("NickName",RecordSet)
    end if  
    %></td>
    <td colspan="2"></td>
  </tr>
  <tr>
      <td width="10%" nowrap><b>e-mail:&nbsp;</b></td>
      <%If RSColData("Email",RecordSet) <> "" then%>
    <td><a href="mailto:<%Response.Write RSColData("Email",RecordSet)%>"><%Response.Write RSColData("Email",RecordSet)%></a></td>
     
      <%else%>
        
     <td>None</td>
     <%End IF   %>
     <td></td>
  </tr>
  <tr>
    <td width="10%" nowrap><b>Website:&nbsp;</b></td>
    <%If RSColData("URL",RecordSet) <> "" then%>
    <td><a href="http://<%Response.Write RSColData("URL",RecordSet)%>">http://<%Response.Write RSColData("URL",RecordSet)%></a></td>
          
      <%else%>
     <td>None</td>
     <%End IF   %>
      <td></td>
  </tr>
  <tr>
    <td width="10%" nowrap><b>Dropzone:&nbsp;</b></td>
    <%If RSColData("Dropzone",RecordSet) <> "" then%>
    <td><p><%Response.Write RSColData("Dropzone",RecordSet)%></p></td>
          
      <%else%>
     <td>None</td>
     <%End IF   %>
      <td></td>
  </tr>

  <tr>
    <td width="10%" nowrap></td>
    <td></td>
    <td colspan="2"></td>
  </tr>
  <tr>
    <td colspan="4">
      <h4>Message:</h4>
    </td>
  </tr>
  <tr>
    <td colspan="4"><%Response.Write RSColData("Message",RecordSet)%></td>
  </tr>
</table>
 <HR>