<!--METADATA TYPE= "typelib"  NAME= "ADODB Type Library" FILE="C:\Program Files\Common Files\SYSTEM\ADO\msado15.dll"  -->
<%response.Charset = "euc-kr"%>
<%
Dim NAME, GBN , FromUserName, FromUserEmail
Dim ToUserEmail
Dim SUBJECT
Dim CONTENT
Dim Msg,Flag
dim objMail
Dim objConfig

if isnull(mail) then mail = ""

NAME = "" : GBN = "" : FromUserName = "" : FromUserEmail = "" : ToUserEmail = "" : SUBJECT = "" : CONTENT = ""


NAME = Request("NAME")
GBN = Request("GBN")

'�����»��
FromUserName = Request("FromUserName")
FromUserEmail = Request("FromUserEmail")


'�������
ToUserEmail = Request("ToUserEmail")

SUBJECT =  NAME + "  " + GBN + "  ���ν�û"
CONTENT = FromUserName + "�����κ��� " + NAME + "  " + GBN + " ���ο�û�� �ֽ��ϴ�."

Set objMail = Server.CreateObject("CDO.Message")

objMail.From = FromUserEmail
objMail.To = ToUserEmail
objMail.Subject = SUBJECT
objMail.TextBody = CONTENT
objMail.Send
Set objMail = Nothing

%>
<SCRIPT LANGUAGE="JavaScript">
<!--
	this.close();
//-->
</SCRIPT>

