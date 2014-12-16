<!--METADATA TYPE= "typelib"  NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<%

Dim MSTMSG, FromUserName , FromUserPhone, ToUserPhone,SQL,strSql,strTxtSql,USENO,JOBNO,MANAGER,CONFIRMFLAG,SENDLOG
Dim Msg,Flag

Const conn ="provider=sqloledb; data source=10.110.10.55,1433; initial catalog=SMS; user id=rmsuser; password = rms12#$"
'Const Conn2 = "provider=sqloledb; data source=10.110.10.86;Initial catalog=mcdev_new; user id=devadmin; password = password"
Const conn2 ="provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"

set objConn = server.createobject("ADODB.Connection")




objConn.open Conn2    

SQL = " SELECT USENO,JOBNO,MANAGER,CONFIRMFLAG, "
SQL = SQL & "dbo.SC_EMPNAME_FUN(manager) fromusername,  "
SQL = SQL & "dbo.SC_SMS_FUN(manager) fromuserphone, "
SQL = SQL & "dbo.SC_SMS_FUN(useno) touserphone, "
SQL = SQL & "dbo.SC_EMPNAME_FUN(manager) +'님 께서 '+CAST(DBO.PD_JOBNAME_FUN(JOBNO) AS NVARCHAR(10))+CASE COUNT(USENO) WHEN 1 THEN '' ELSE ' 외'+CAST( COUNT(USENO)-1 AS VARCHAR(20))+' 건' END+ "
SQL = SQL & "CASE CONFIRMFLAG  "
SQL = SQL & "WHEN '3' THEN '승인'  "
SQL = SQL & "WHEN '2' THEN '승인취소'  "
SQL = SQL & "WHEN '0' THEN '반려' END+' 하셨습니다' AS MSTMSG, "
SQL = SQL & "GETDATE() SENDLOG "
SQL = SQL & "FROM PD_SMS_TEMP "
SQL = SQL & "GROUP BY USENO,JOBNO,MANAGER,CONFIRMFLAG "

Set rs=Createobject("adodb.recordset")
rs.Open SQL, Conn2,1

Do until rs.Eof

Set adocmd = Server.CreateObject("ADODB.Command")
MSTMSG = rs("MSTMSG")
FromUserName = rs("fromusername")
FromUserPhone = rs("fromuserphone")
ToUserPhone = rs("touserphone")
USENO = rs("USENO")
JOBNO = rs("JOBNO")
MANAGER = rs("MANAGER")
CONFIRMFLAG = rs("CONFIRMFLAG")
SENDLOG = rs("SENDLOG")
Flag = "RMS"

with adocmd
	.ActiveConnection = conn
	.CommandText = "dbo.UP_SendSMS"
	.CommandType = adCmdStoredProc
	.Parameters.Append .CreateParameter("@vcSndPhnId",advarwchar,adParamInput,15) 
	.Parameters.Append .CreateParameter("@vcRcvPhnId",advarwchar,adParamInput,15) 
	.Parameters.Append .CreateParameter("@vcSndMsg",advarwchar,adParamInput,200) 
	.Parameters.Append .CreateParameter("@vcMsgID",advarwchar,adParamInput,20) 
	.Parameters("@vcSndPhnId") = FromUserPhone
	.Parameters("@vcRcvPhnId") = ToUserPhone
	.Parameters("@vcSndMsg") = MSTMSG
	.Parameters("@vcMsgID") = Flag
	.Execute , , adExecuteNoRecords 
End with



strTxtSql = " INSERT INTO SC_SMS(USERNO,JOBNO,MANAGER,CONFIRMFLAG,fromusername,fromuserphone,touserphone,MSTMSG,SENDLOG) "
strTxtSql = strTxtSql & " VALUES('" & USENO & "','" & JOBNO & "','" & MANAGER & "','" & CONFIRMFLAG & "','" & fromusername & "','" & fromuserphone & "','" & touserphone & "','" & MSTMSG & "','" & SENDLOG & "') "

objConn.Execute strTxtSql

Set adocmd = Nothing
rs.movenext 
loop


strSql = "delete from pd_sms_temp" 
objConn.Execute strSql
	

objConn.Close
Set objConn = Nothing

%>
<html>
뭐냐 이거...
</html>