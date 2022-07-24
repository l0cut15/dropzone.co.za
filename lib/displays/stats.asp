
<table border="0" width="100%">
  <tr>
    <td valign="middle"><%Response.Write RSColData("HitDate",RecordSet)%>-<%Response.Write RSColData("HitTime",RecordSet)%></td>
    <td><p><%Response.Write RSColData("IP",RecordSet)%></p></td>
    <td valign="middle">
        
      <p align="right">
        
      <textarea rows="2" name="S1" cols="60"><%Response.Write RSColData("AGENT",RecordSet)%>
<%Response.Write RSColData("HOST",RecordSet)%></textarea>
   
   
      </p>
   
   
      </td>
       
  </tr>
</table>
 <HR>