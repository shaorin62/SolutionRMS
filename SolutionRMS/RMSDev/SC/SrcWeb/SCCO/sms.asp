<!--METADATA TYPE= "typelib"  NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<%response.Charset = "euc-kr"%>
<%
'MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
Dim FromUserName , FromUserPhone, ToUserPhone
Dim Msg,Flag
Dim MSTMSG

Const conn ="provider=sqloledb; data source=10.110.10.55,1433; initial catalog=SMS; user id=rmsuser; password = rms12#$"

Set adocmd = Server.CreateObject("ADODB.Command")

MSTMSG = "" : FromUserName = "" : FromUserPhone = "" : ToUserPhone = "" : Msg = "" : Flag = "" : 

MSTMSG = Request("MSTMSG")

'보내는사람
FromUserName = Request("FromUserName")
FromUserPhone = Request("FromUserPhone")
RESPONSE.Write MSTMSG

'받을사람
ToUserPhone = Request("ToUserPhone")


'Msg = "RMS SMS발송 테스트"
Msg = FromUserName + "님으로부터 "+MSTMSG
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
			  .Parameters("@vcSndMsg") = Msg
			  .Parameters("@vcMsgID") = Flag
              .Execute , , adExecuteNoRecords 
       End with
Set adocmd = Nothing

%>
