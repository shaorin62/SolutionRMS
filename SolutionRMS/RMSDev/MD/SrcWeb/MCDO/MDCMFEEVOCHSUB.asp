<% Option Explicit%>

<%

'Response.AddHeader = "content-Disposition"
Dim Sql,rs,temp_filename,temp_fileno,mstrGUBUN

'���ǳ����� - By. �������
mstrGUBUN = request("mstrGUBUN")
'A: ����Ź ��޾� , S:������ , G: �Ϲ�
temp_filename = request("temp_filename")
temp_fileno = temp_filename
temp_filename = temp_filename 
'Response.ContentType = "application/unknown"

'Response.AddHeader "Content-Disposition","attachment; filename=" & temp_filename 

Dim fso,act
'	taxyearmon	taxno		gbn	account		documentdate	prepayment	fromdate	todate	summtext
'Const conn ="provider=sqloledb; data source=10.110.10.86; initial catalog=MCDEV_NEW; user id=devadmin; password = password"
Const conn ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"

Sql = " select  "
Sql = Sql & "convert(char(8),ISNULL(postingdate,'')) postingdate, "
Sql = Sql & "convert(char(10),replace(ISNULL(customercode,''),'-','')) customercode, "
Sql = Sql & "convert(char(25),ISNULL(summ,'')) summ, "
Sql = Sql & "convert(char(4),ISNULL(ba,'')) ba, "
Sql = Sql & "'53105     ' costcenter, "
Sql = Sql & "dbo.lpad(sumamt,13,'0') amt, "
Sql = Sql & "null vat, "

If mstrGUBUN = "A" Then
Sql = Sql & "'  ' semu, "
Else
Sql = Sql & "'' semu, "
End If

Sql = Sql & "convert(char(4),ISNULL(bp,'')) bp, "
Sql = Sql & "convert(char(8),ISNULL(demandday,'')) demandday, "
If mstrGUBUN = "A" Then
Sql = Sql & "convert(char(10),replace(ISNULL(vendor,''),'-','')) vendor, "
Else
Sql = Sql & "convert(char(10),replace(ISNULL(customercode,''),'-','')) vendor, "
End If
Sql = Sql & "convert(char(6),ISNULL(taxyearmon,'')) taxyearmon, "
Sql = Sql & "dbo.lpad(taxno,4,'0') taxno, "

if mstrGUBUN = "A" or mstrGUBUN = "S" then
Sql = Sql & "convert(char(1),ISNULL(GFLAG,'')) GFLAG, "
else 
Sql = Sql & "'' GFLAG, "
end if
If mstrGUBUN = "A" Then
Sql = Sql & "convert(char(1),ISNULL(GBN,'T')) GBN, "
ElseIf mstrGUBUN = "S" Then
Sql = Sql & "convert(char(1),ISNULL(GBN,'S')) GBN, "
Else
Sql = Sql & "convert(char(1),ISNULL(GBN,'F')) GBN, "
End If

'�ٲ�����
If mstrGUBUN = "A" or mstrGUBUN = "S" Then
Sql = Sql & "convert(char(10),ISNULL(ACCOUNT,'')) ACCOUNT, " 
Sql = Sql & "convert(char(10),ISNULL(DEBTOR,''))  DEBTOR, "
else
Sql = Sql & "convert(char(10),ISNULL(ACCOUNT,'')) DEBTOR , " 
Sql = Sql & "convert(char(10),ISNULL(DEBTOR,''))  ACCOUNT, "
end if 

Sql = Sql & "convert(char(8),ISNULL(DOCUMENTDATE,'')) DOCUMENTDATE, "
Sql = Sql & "'' PAYCODE, " 
Sql = Sql & "convert(char(1),ISNULL(PREPAYMENT,'')) PREPAYMENT, "
Sql = Sql & "convert(char(8),ISNULL(FROMDATE,'')) FROMDATE, "
Sql = Sql & "convert(char(8),ISNULL(TODATE,'')) TODATE, "
Sql = Sql & "convert(char(50),ISNULL(SUMMTEXT,'')) SUMMTEXT, "
Sql = Sql & "'+' AMTGBN "
Sql = Sql & " from MD_TRUVOCH_MST where RMSNO= '"&temp_fileno&"'"
Set rs=Createobject("adodb.recordset")
rs.Open Sql, Conn,1

'response.Write "" & rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|"

Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set act = fso.CreateTextFile(Server.MapPath("\Excel") & "\" & temp_filename,true)
%>
<%
Do until rs.Eof
	act.WriteLine ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|"

	rs.movenext
Loop
set rs = nothing
act.close

'response.Write act
Response.Redirect "http://10.110.10.89/fileftp.asp?temp_filename=" &  temp_filename
%>
