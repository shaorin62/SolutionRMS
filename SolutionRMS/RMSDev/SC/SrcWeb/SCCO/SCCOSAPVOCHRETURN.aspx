<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOSAPVOCHRETURN.aspx.vb" Inherits="SC.SCCOSAPVOCHRETURN" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SAP ��ǥ��� RETURN</title> 
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOSAPBUSINO.aspx
'��      �� : SAP�ŷ�ó ����
'Ư��  ���� : 
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
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit


'-----------------------------
' �̺�Ʈ ���ν��� 
'-----------------------------	
Sub Set_VochReturn
	Dim strValue
	strValue = .txtVOCHRETURN.value
	parent.RFC_EndMsg strValue
End Sub

Sub Test
	
End Sub
-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM id="frmThis" action="SCCOSAPVOCHRETURN.aspx" method="post" runat="server">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
				<tr>
					<td><asp:textbox id="txtVOCHRETURN" runat="server"></asp:textbox></td>
				</tr>
			</TABLE>
			<div id="rfcfield">
				<asp:textbox id="txtRETURN" runat="server"></asp:textbox>
				<!--<INPUT id=txtO_RETURN style="WIDTH: 48px; HEIGHT: 21px" type=text size=2 value='<%# DataBinder.Eval(SapProxy21, "O_RETURN") %>' name=txtO_RETURN>-->
			</div>
		</FORM>
	</body>
</HTML>
