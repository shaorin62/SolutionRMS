<% Option Explicit%>

<%

'Response.AddHeader = "content-Disposition"
Dim Sql,rs,temp_filename,temp_fileno,mstrGUBUN

temp_filename = request("temp_filename")
temp_fileno = temp_filename
temp_filename = temp_filename 
'Response.ContentType = "application/unknown"

'Response.AddHeader "Content-Disposition","attachment; filename=" & temp_filename 

Dim fso,act
'Const conn ="provider=sqloledb; data source=10.110.10.86; initial catalog=MCDEV_NEW; user id=devadmin; password = password"
Const conn ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"



Sql = "  SELECT "
Sql = Sql & "  CASE ISNULL(RTRIM(COMPANYCD),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(COMPANYCD),'') END AS COMPANYCD, "
Sql = Sql & "  CASE ISNULL(RTRIM(BILLNO),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BILLNO),'') END AS BILLNO, "
Sql = Sql & "  CASE ISNULL(RTRIM(FISCALLYY),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(FISCALLYY),'') END AS FISCALLYY, "
Sql = Sql & "  CASE ISNULL(RTRIM(BILLFLAG),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BILLFLAG),'') END AS BILLFLAG, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPBSN),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPBSN),'') END AS SUPPBSN, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPLDSCR),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPLDSCR),'') END AS SUPPLDSCR, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPCEO),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPCEO),'') END AS SUPPCEO, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPADDR),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPADDR),'') END AS SUPPADDR, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPBUSICOND),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPBUSICOND),'') END AS SUPPBUSICOND, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPBUSIITEM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPBUSIITEM),'') END AS SUPPBUSIITEM, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYBSN),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYBSN),'') END AS BUYBSN, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYLDSCR),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYLDSCR),'') END AS BUYLDSCR, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYCEO),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYCEO),'') END AS BUYCEO, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYADDR),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYADDR),'') END AS BUYADDR, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYBUSICOND),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYBUSICOND),'') END AS BUYBUSICOND, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYBUSIITEM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYBUSIITEM),'') END AS BUYBUSIITEM, "
Sql = Sql & "  CASE ISNULL(RTRIM(REGDATE),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(REGDATE),'') END AS REGDATE, "
Sql = Sql & "  CASE ISNULL(RTRIM(TOTAMT),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(TOTAMT),'') END AS TOTAMT, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPAMT),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPAMT),'') END AS SUPPAMT, "
Sql = Sql & "  CASE ISNULL(RTRIM(VATAMT),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(VATAMT),'') END AS VATAMT, "
Sql = Sql & "  CASE ISNULL(RTRIM(BILLRMRK),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BILLRMRK),'') END AS BILLRMRK, "
Sql = Sql & "  CASE ISNULL(RTRIM(TITLE),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(TITLE),'') END AS TITLE, "
Sql = Sql & "  CASE ISNULL(RTRIM(REQFLAG),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(REQFLAG),'') END AS REQFLAG, "
Sql = Sql & "  CASE ISNULL(RTRIM(NORMFLAG),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(NORMFLAG),'') END AS NORMFLAG, "
Sql = Sql & "  CASE ISNULL(RTRIM(RECEIPTID),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(RECEIPTID),'') END AS RECEIPTID, "
Sql = Sql & "  CASE ISNULL(RTRIM(RECEIPTNM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(RECEIPTNM),'') END AS RECEIPTNM, "
Sql = Sql & "  CASE ISNULL(RTRIM(PURTEAMCD),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(PURTEAMCD),'') END AS PURTEAMCD, "
Sql = Sql & "  CASE ISNULL(RTRIM(INSDATE),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(INSDATE),'') END AS INSDATE, "
Sql = Sql & "  CASE ISNULL(RTRIM(BILLSEQ),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BILLSEQ),'') END AS BILLSEQ, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPDATE),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPDATE),'') END AS SUPPDATE, "
Sql = Sql & "  CASE ISNULL(RTRIM(ITEMNM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(ITEMNM),'') END AS ITEMNM, "
Sql = Sql & "  CASE ISNULL(RTRIM(SIZE),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SIZE),'') END AS SIZE, "
Sql = Sql & "  CASE ISNULL(RTRIM(QTY),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(QTY),'') END AS QTY, "
Sql = Sql & "  CASE ISNULL(RTRIM(UNITPRC),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(UNITPRC),'') END AS UNITPRC, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPAMT),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPAMT),'') END AS SUPPAMT, "
Sql = Sql & "  CASE ISNULL(RTRIM(VATAMT),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(VATAMT),'') END AS VATAMT, "
Sql = Sql & "  CASE ISNULL(RTRIM(ITEMRMRK),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(ITEMRMRK),'') END AS ITEMRMRK, "
Sql = Sql & "  CASE ISNULL(RTRIM(EDITTYPECD),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(EDITTYPECD),'') END AS EDITTYPECD, "
Sql = Sql & "  CASE ISNULL(RTRIM(FIRSTBILLNO),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(FIRSTBILLNO),'') END AS FIRSTBILLNO, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYEMAIL),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYEMAIL),'') END AS BUYEMAIL, "
Sql = Sql & "  CASE ISNULL(RTRIM(BUYNM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(BUYNM),'') END AS BUYNM, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPEMAIL),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPEMAIL),'') END AS SUPPEMAIL, "
Sql = Sql & "  CASE ISNULL(RTRIM(SUPPNM),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SUPPNM),'') END AS SUPPNM, "
Sql = Sql & "  CASE ISNULL(RTRIM(CANCEL_YN),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(CANCEL_YN),'') END AS CANCEL_YN, "
Sql = Sql & "  CASE ISNULL(RTRIM(SENDNTS_YN),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(SENDNTS_YN),'') END AS SENDNTS_YN, "
Sql = Sql & "  CASE ISNULL(RTRIM(ISTRUST_YN),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(ISTRUST_YN),'') END AS ISTRUST_YN, "
Sql = Sql & "  CASE ISNULL(RTRIM(TRUST_CUSCD),'') WHEN '' THEN ' ' ELSE ISNULL(RTRIM(TRUST_CUSCD),'') END AS TRUST_CUSCD "
Sql = Sql & "  FROM MD_SENDTAX_MST "
Sql = Sql & "  where RMSNO= '"&temp_fileno&"'"
Set rs=Createobject("adodb.recordset")
rs.Open Sql, Conn,1

'response.Write "" & rs("postingdate") &"|" & rs("customercode") &"| "& rs("summ") & "|" & rs("ba") & "|" & rs("costcenter") & "|" & rs("amt") &  "|" & rs("vat") & "|" & rs("semu") & "|" & rs("bp") & "|" & rs("demandday") & "|" & rs("vendor") & "|" & rs("taxyearmon") & "|" & rs("taxno") & "|" & rs("GFLAG") & "|" & rs("GBN") & "|" & rs("ACCOUNT") & "|" & rs("DEBTOR") & "|" & rs("DOCUMENTDATE") & "|" & rs("PREPAYMENT") & "|" & rs("FROMDATE") & "|" & rs("TODATE") & "|" & rs("SUMMTEXT") & "|" & rs("AMTGBN") & "|"

Set fso = Server.CreateObject("Scripting.FileSystemObject")

Set act = fso.CreateTextFile(Server.MapPath("\SENDTAX") & "\" & temp_filename,true)
%>
<%
Do until rs.Eof
	act.WriteLine "CR//"& rs("COMPANYCD") &"//" & rs("BILLNO") &"//" & rs("FISCALLYY") &"//" & rs("BILLFLAG") &"//" & rs("SUPPBSN") &"//" & rs("SUPPLDSCR") &"//" & rs("SUPPCEO") &"//" & rs("SUPPADDR") &"//" & rs("SUPPBUSICOND") &"//" & rs("SUPPBUSIITEM") &"//" & rs("BUYBSN") &"//" & rs("BUYLDSCR") &"//" & rs("BUYCEO") &"//" & rs("BUYADDR") &"//" & rs("BUYBUSICOND") &"//" & rs("BUYBUSIITEM") &"//" & rs("REGDATE") &"//" & rs("TOTAMT") &"//" & rs("SUPPAMT") &"//" & rs("VATAMT") &"//" & rs("BILLRMRK") &"//" & rs("TITLE") &"//" & rs("REQFLAG") &"//" & rs("NORMFLAG") &"//" & rs("RECEIPTID") &"//" & rs("RECEIPTNM") &"//" & rs("PURTEAMCD") &"//" & rs("INSDATE") &"//" & rs("BILLSEQ") &"//" & rs("SUPPDATE") &"//" & rs("ITEMNM") &"//" & rs("SIZE") &"//" & rs("QTY") &"//" & rs("UNITPRC") &"//" & rs("SUPPAMT") &"//" & rs("VATAMT") &"//" & rs("ITEMRMRK") &"//" & rs("EDITTYPECD") &"//" & rs("FIRSTBILLNO") &"//" & rs("BUYEMAIL") &"//" & rs("BUYNM") &"//" & rs("SUPPEMAIL") &"//" & rs("SUPPNM") &"//" & rs("CANCEL_YN") &"//" & rs("SENDNTS_YN") &"//" & rs("ISTRUST_YN") &"//" & rs("TRUST_CUSCD")

	rs.movenext
Loop
set rs = nothing
act.close

'response.Write act
Response.Redirect "http://10.110.10.89/fileftpsendtax.asp?temp_filename=" &  temp_filename

%>
