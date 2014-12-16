<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOSAPBUSINO.aspx.vb" Inherits="SC.SCCOSAPBUSINO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SAP 거래처연결</title> 
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOSAPBUSINO.aspx
'기      능 : SAP거래처 연동
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By KTY
'*************************************************************************************
***
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub Set_CustInfo
	Dim strBUSINO
	Dim strBANKTYPE
	Dim strRETURNLIST
	Dim strRETURNBANK
	Dim Arraychk
	With frmThis
	
		strRETURNLIST = "" : strRETURNBANK = "" : strBUSINO = "" : strBANKTYPE = ""
		
		strRETURNLIST = .txtRETURNLIST.value
		strRETURNBANK = .txtRETURNBANK.value

		strBUSINO = strRETURNLIST
		strBANKTYPE = strRETURNBANK
		parent.Set_CustValue strBUSINO, strBANKTYPE
		'parent.Set_RFCValue strBUSINO, strBANKTYPE
		
		.txtRETURNLIST.value = ""
		.txtRETURNBANK.value = ""
		 
	End With
End Sub

-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM id="frmThis" action="SCCOSAPBUSINO.aspx" method="post" runat="server">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
				<tr>
					<td><asp:textbox id="txtSAPBUSINO" runat="server"></asp:textbox></td>
					<td><asp:textbox id="txtCNT" runat="server"></asp:textbox></td>
					<td><asp:textbox id="RT_RETURNLIST" runat="server"></asp:textbox></td>
					<td><asp:textbox id="RT_RETURNBANK" runat="server"></asp:textbox></td>
				</tr>
			</TABLE>
			<div id="rfcfield">
				<INPUT id=txtRETURNLIST style="WIDTH: 160px; HEIGHT: 21px" type=text size=2 value="<%# me.RT_RETURNLIST.text %>" name=txtRETURNLIST>
				<INPUT id=txtRETURNBANK style="WIDTH: 160px; HEIGHT: 21px" type=text size=2 value="<%# me.RT_RETURNBANK.text %>" name=txtRETURNBANK>
			</div>
		</FORM>
	</body>
</HTML>
