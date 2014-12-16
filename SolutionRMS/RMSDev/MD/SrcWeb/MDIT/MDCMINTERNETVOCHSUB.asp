<% Option Explicit%>

<%

'Response.AddHeader = "content-Disposition"
Dim Sql,rs,temp_filename,temp_fileno,mstrGUBUN


'A: 위수탁 취급액 , S:수수료 , C: 일반
mstrGUBUN = request("mstrGUBUN")
temp_filename = request("temp_filename")
temp_fileno = temp_filename
temp_filename = temp_filename 

Dim fso,act
'Const conn ="provider=sqloledb; data source=10.110.10.86; initial catalog=MCDEV_NEW; user id=devadmin; password = password"
Const conn ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"

Sql = " select  "
Sql = Sql & "convert(char(8),ISNULL(postingdate,'')) postingdate, "
Sql = Sql & "convert(char(10),replace(ISNULL(customercode,''),'-','')) customercode, "
Sql = Sql & "convert(char(25),ISNULL(summ,'')) summ, "
Sql = Sql & "convert(char(4),ISNULL(ba,'')) ba, "
Sql = Sql & "'53105     ' costcenter, "


If mstrGUBUN = "A" Then 
	Sql = Sql & " dbo.lpad(sumamt,13,'0') amt, "
elseif mstrGUBUN = "S"  then
	Sql = Sql & " dbo.lpad(sumamt,13,'0') amt, "
elseif mstrGUBUN = "C" then
	Sql = Sql & " case medflag when 'OA' then dbo.lpad((isnull(sumamt,0)+isnull(vat,0)),13,'0')  else dbo.lpad(sumamt,13,'0') end as amt, "
end if 


Sql = Sql & "dbo.lpad(vat,13,'0') vat, "

If mstrGUBUN = "A" Then 
	Sql = Sql & "'  ' semu, " '매출은 세무코드 없음
Elseif mstrGUBUN = "S" or mstrGUBUN = "C"  then
	Sql = Sql & "'B5' semu, " '수수료 , 일반 둘다 해당
End If

Sql = Sql & "convert(char(4),ISNULL(bp,'')) bp, "
Sql = Sql & "convert(char(8),ISNULL(demandday,'')) demandday, "

If mstrGUBUN = "A" Then
	Sql = Sql & "convert(char(10),replace(ISNULL(vendor,''),'-','')) vendor, "
ElseIF mstrGUBUN = "S" Then
	Sql = Sql & "'          ' vendor, "
ElseIF mstrGUBUN = "C" Then
	Sql = Sql & "convert(char(10),replace(ISNULL(vendor,''),'-','')) vendor, "
End If

Sql = Sql & "convert(char(6),ISNULL(taxyearmon,'')) taxyearmon, "
Sql = Sql & "dbo.lpad(taxno,4,'0') taxno, "
Sql = Sql & "convert(char(1),ISNULL(GFLAG,'')) GFLAG, "

If mstrGUBUN = "A" Then
	Sql = Sql & "convert(char(1),ISNULL(GBN,'T')) GBN, "
ElseIf mstrGUBUN = "S" Then
	Sql = Sql & "convert(char(1),ISNULL(GBN,'S')) GBN, "
ElseIF mstrGUBUN = "C" Then
	Sql = Sql & "convert(char(1),ISNULL(GBN,'C')) GBN, "
End If

Sql = Sql & "convert(char(10),ISNULL(ACCOUNT,'')) ACCOUNT, "
Sql = Sql & "convert(char(10),ISNULL(DEBTOR,''))  DEBTOR, "
Sql = Sql & "convert(char(8),ISNULL(DOCUMENTDATE,'')) DOCUMENTDATE, "
Sql = Sql & "convert(char(1),ISNULL(PREPAYMENT,'')) PREPAYMENT, "
Sql = Sql & "convert(char(8),ISNULL(FROMDATE,'')) FROMDATE, "
Sql = Sql & "convert(char(8),ISNULL(TODATE,'')) TODATE, "
Sql = Sql & "convert(char(50),ISNULL(SUMMTEXT,'')) SUMMTEXT, "
Sql = Sql & "'+' AMTGBN, "
Sql = Sql & "'' PAYCODE, "
Sql = Sql & "convert(char(8),ISNULL(DUEDATE,'')) DUEDATE "

Sql = Sql & " from MD_TRUVOCH_MST where RMSNO= '"&temp_fileno&"'"


response.Write sql
Set rs=Createobject("adodb.recordset")
rs.Open Sql, Conn,1

'response.Write ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|" & rs("PAYCODE") & "|" & rs("DUEDATE")

Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set act = fso.CreateTextFile(Server.MapPath("\Excel") & "\" & temp_filename,true)
%>
<%
Do until rs.Eof
	if mstrGUBUN = "A" then
		act.WriteLine ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|" & rs("PAYCODE") & "|" & rs("DUEDATE")
	elseif mstrGUBUN = "S" then
		act.WriteLine ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|" & rs("PAYCODE") & "|" & rs("DUEDATE")
	elseif mstrGUBUN = "C" then
		act.WriteLine ""& rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|"
	end if 
	
	rs.movenext
Loop
set rs = nothing
act.close

'response.Write act
Response.Redirect "http://10.110.10.89/fileftp.asp?temp_filename=" &  temp_filename
%>
