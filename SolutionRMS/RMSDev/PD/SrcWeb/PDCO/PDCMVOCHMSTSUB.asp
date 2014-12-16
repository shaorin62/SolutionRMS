<% Option Explicit%>

<%
Response.ContentType ="application/x-msexcel"
Dim Sql,rs,temp_filename,temp_fileno

'조건날리기 - By. 해줘버려

temp_filename = request("temp_filename")
temp_fileno = temp_filename
temp_filename = temp_filename & ".xls"

Response.AddHeader "Content-Disposition","attachment; filename=" & temp_filename 
Const conn ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev; user id=advsa; password = advsa1234"
%>

<%


Sql = "select postingdate,customercode,summ,ba,costcenter,sumamt,vat,semu,bp,demandday,taxyearmon,taxno,gbn, attr01, attr02 from PD_VOCH_MST where RMSNO= '" & temp_fileno & "'"
Set rs=Createobject("adodb.recordset")
rs.Open Sql, Conn,1
Dim Code,Code_name

Dim fso,act,objStream,download
Set fso = Server.Createobject("Scripting.FileSystemObject")
Set act = fso.CreateTextFile(Server.MapPath("\Excel") & "\" & temp_filename,true)

act.WriteLine "<html xmlns:x=""urn:schemas-microsoft-com:office:excel"">"
act.WriteLine "<Head>"
act.WriteLine "<!--<xml>"
act.WriteLine "<x:ExcelWorkbook>"
act.WriteLine "<x:ExcelWorksheets>"
act.WriteLine "<x:ExcelWorksheet>"
act.WriteLine "<x:Name>Members</x:Name>"
act.WriteLine "<x:worksheetOptions>"
act.WriteLine "<x:print>"
act.WriteLine "<x:validPrinterInfo/>"
act.WriteLine "</x:Print>"
act.WriteLine "</x:worksheetOption>"
act.WriteLine "</x:ExcelWorksheet>"
act.WriteLine "</x:ExcelWorksheets>"
act.WriteLine "</x:ExcelWorkbook>"
act.WriteLine "</xml>"
act.WriteLine "<-->"
act.WriteLine "</head>"
act.WriteLine "<body>"
act.WriteLine "<table>"
act.WriteLine "<tr>"
act.WriteLine "<td>postingdate</td>"
act.WriteLine "<td>customercode</td>"
act.WriteLine "<td>summ</td>"
act.WriteLine "<td>ba</td>"
act.WriteLine "<td>costcenter</td>"
act.WriteLine "<td>sumamt</td>"
act.WriteLine "<td>vat</td>"
act.WriteLine "<td>semu</td>"
act.WriteLine "<td>bp</td>"
act.WriteLine "<td>demandday</td>"
act.WriteLine "<td>taxyearmon</td>"
act.WriteLine "<td>taxno</td>"
act.WriteLine "<td>gbn</td>"
act.WriteLine "<td>account</td>"
act.WriteLine "<td>documentdate</td>"
act.WriteLine "</tr>"
Do until rs.Eof
act.WriteLine "<tr>"
act.WriteLine "<td>"
act.WriteLine " "& rs("postingdate")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("customercode")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("summ")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("ba")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("costcenter")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("sumamt")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("vat")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("semu")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("bp")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("demandday")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("taxyearmon")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("taxno")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("gbn")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("attr01")
act.WriteLine "</td>"
act.WriteLine "<td>"
act.WriteLine " "& rs("attr02")
act.WriteLine "</td>"
act.WriteLine "</tr>"
rs.movenext
Loop
act.WriteLine "</table>"
act.WriteLine "</body>"
act.WriteLine "</html>"
act.close
Set objStream = Server.CreateObject("ADODB.Stream")
objStream.Open
objStream.Type=1
objStream.LoadFromFile Server.MapPath("\Excel") & "\" & temp_filename

download = objStream.Read
Response.BinaryWrite download 
Set objStream = nothing
%>
<script>
this.close();
</script>