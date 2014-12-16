<% Option Explicit%>

<%

'Response.AddHeader = "content-Disposition"
Dim Sql,rs,temp_filename,temp_fileno,mstrGUBUN

'조건날리기 - By. 해줘버려
mstrGUBUN = request("mstrGUBUN")
temp_filename = request("temp_filename")
temp_fileno = temp_filename
temp_filename = temp_filename 
'Response.ContentType = "application/unknown"

'Response.AddHeader "Content-Disposition","attachment; filename=" & temp_filename 




Dim fso,act

Const conn ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"
''strCon  ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev; user id=advsa; password = advsa1234"
Sql = " select  "
Sql = Sql & "convert(char(8),ISNULL(postingdate,'')) postingdate, "
Sql = Sql & "convert(char(10),replace(ISNULL(customercode,''),'-','')) customercode, "
Sql = Sql & "convert(char(25),ISNULL(summ,'')) summ, "
Sql = Sql & "convert(char(4),ISNULL(ba,'')) ba, "
Sql = Sql & "'53105     ' costcenter, "

'차감전표일때 abs함수 절대값으로 +금액으로 가져온다. 뒤에서 gbn 에 -로 넘겨주면됨
Sql = Sql & "case when isnull(amt,0) < 0 then abs(isnull(amt,0)) else  dbo.lpad(amt,13,'0') end  amt, "
Sql = Sql & "case when isnull(vat,0) < 0 then  abs(isnull(vat,0)) else dbo.lpad(vat,13,'0') end  vat, "

Sql = Sql & "convert(char(2),ISNULL(semu,'')) semu, "
Sql = Sql & "convert(char(4),ISNULL(bp,'')) bp, "
Sql = Sql & "convert(char(8),ISNULL(demandday,'')) demandday, "
If mstrGUBUN = "_P" Then
Sql = Sql & "convert(char(10),replace(ISNULL(vendor,''),'-','')) vendor, "
ElseIf mstrGUBUN = "_B" Then
Sql = Sql & "convert(char(10),replace(ISNULL(customercode,''),'-','')) vendor, "
End If
Sql = Sql & "convert(char(6),ISNULL(taxyearmon,'')) taxyearmon, "
Sql = Sql & "dbo.lpad(taxno,4,'0') taxno, "
Sql = Sql & "'' GFLAG, "
If mstrGUBUN = "_P" Then
Sql = Sql & "convert(char(1),ISNULL(GBN,'P')) GBN, "
ElseIf mstrGUBUN = "_B" Then
Sql = Sql & "convert(char(1),ISNULL(GBN,'B')) GBN, "
End If
If mstrGUBUN = "_P" Then
Sql = Sql & "convert(char(10),ISNULL(DEBTOR,''))  DEBTOR, "
Sql = Sql & "convert(char(10),ISNULL(ACCOUNT,'')) ACCOUNT, " '바뀌여야됨
ElseIf mstrGUBUN = "_B" Then
Sql = Sql & "convert(char(10),ISNULL(DEBTOR,''))  ACCOUNT, "
Sql = Sql & "convert(char(10),ISNULL(ACCOUNT,'')) DEBTOR, " '바뀌여야됨
End If

Sql = Sql & "convert(char(8),ISNULL(DOCUMENTDATE,'')) DOCUMENTDATE, "
Sql = Sql & "convert(char(1),ISNULL(PREPAYMENT,'')) PREPAYMENT, "
Sql = Sql & "convert(char(8),ISNULL(FROMDATE,'')) FROMDATE, "
Sql = Sql & "convert(char(8),ISNULL(TODATE,'')) TODATE, "
Sql = Sql & "convert(char(50),ISNULL(SUMMTEXT,'')) SUMMTEXT, "

'차감전표일시 - 
Sql = Sql & " case when isnull(amt,0) < 0 then '-' else '+' end AMTGBN, "

Sql = Sql & "convert(char(1),ISNULL(PAYCODE,'')) PAYCODE, "
Sql = Sql & "convert(char(8),ISNULL(DUEDATE,'')) DUEDATE "

Sql = Sql & " from PD_VOCH_MST where RMSNO= '" & temp_fileno & "'"
Set rs=Createobject("adodb.recordset")
rs.Open Sql, Conn,1


Set fso = Server.CreateObject("Scripting.FileSystemObject")

'response.Write Server.MapPath("\Excel") & "\" & temp_filename
response.Write Server.MapPath("\Excel") & "\" & temp_filename
Set act = fso.CreateTextFile(Server.MapPath("\Excel") & "\" & temp_filename,true)
%>
<%
Do until rs.Eof
	act.WriteLine ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|" & rs("PAYCODE") & "|" & rs("DUEDATE")

	rs.movenext
Loop

act.close
rs.close
'response.Write act
Response.Redirect "http://10.110.10.89/fileftp.asp?temp_filename=" &  temp_filename
%>
