<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTTEST.aspx.vb" Inherits="SC.SCRTTest" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>SCFUTest</title> 
		<!--<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

'��ưŬ�� �̺�Ʈ      

Sub imgPrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	
	If frmThis.txtModuleDir.value="" then
		ModuleDir = "SC"
	Else
		ModuleDir = frmThis.txtModuleDir.value
	End if
	
	If frmThis.txtModuleDir.value="" then
		ReportName = "SCMENU.rpt"
	Else
		ReportName = frmThis.txtReportName.value
	End if
	
	If frmThis.txtParams.value = "" then
		Params = "SJCC:����"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
	
	gShowReportWindow ModuleDir, ReportName, Params
	''gShowReportWindow "SC","SCMENU.rpt","SJCC:����"

	gFlowWait meWAIT_OFF
End Sub
//-->
		</script>
	</HEAD>
	<body>
		<form id="frmThis">
			<P><FONT face="����"></FONT>&nbsp;</P>
			<P><FONT face="����">*** ũ����Ż ����Ʈ ���� ***</FONT></P>
			<P>
				<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px" cellSpacing="1" cellPadding="1" width="75%" border="0">
					<TR>
						<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF" border="0" name="imgWaiting">
						</TD>
					</TR>
				</TABLE>
				<TABLE id="Table1" style="WIDTH: 619px; HEIGHT: 143px" cellSpacing="1" cellPadding="1" width="619" border="1">
					<TR class="EVENROW">
						<TD style="WIDTH: 370px"><FONT face="����">ModuleDir </FONT><INPUT id="txtModuleDir" type="text" value="SC"></TD>
						<TD><FONT face="����">����: ����Ʈ�� ���� ��ġ�� ��������
								<BR>
								��� ���丮 (��: SC, CO, PO, AP&nbsp;��...)</FONT></TD>
					</TR>
					<TR class="ODDROW">
						<TD style="WIDTH: 370px"><FONT face="����">ReportName</FONT> <INPUT id="txtReportName" type="text" value="SCMENU.rpt"></TD>
						<TD><FONT face="����">����Ʈ�� �̸�(��: SCMENU.rpt )</FONT></TD>
					</TR>
					<TR class="EVENROW">
						<TD style="WIDTH: 370px; HEIGHT: 16.571pt"><FONT face="����">Params</FONT> <INPUT id="txtParams" type="text" value="SJCC:����"></TD>
						<TD style="HEIGHT: 16.571pt"><FONT face="����">�Ķ���� ��(��: SJCC:����)</FONT></TD>
					</TR>
					<TR class="ODDROW">
						<TD style="WIDTH: 370px">
							<P align="left"><FONT face="����" color="#ff3366"><STRONG>%% ����: �Ķ���͸� �ѱ� ��� ������ String 
										Type���� �ѱ�� ���� �������̳� ��¥���� ��� String Type �����͸� ������ TO_NUMBER, TO_DATE �� ��ȯ �Ͽ� ����Ʈ 
										��� %%&nbsp;&nbsp; --&gt; </STRONG></FONT><STRONG><FONT face="����" color="#ff3366">
										ũ����Ż ����Ʈ����.doc ����</FONT></STRONG></P>
						</TD>
						<TD><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gif" width="54" border="0" name="imgPrint"><FONT face="����">&nbsp;</FONT></TD>
					</TR>
				</TABLE>
			</P>
			<P><FONT face="����"></FONT>&nbsp;</P>
			<P><FONT face="����"></FONT>&nbsp;</P>
		</form>
	</body>
</HTML>
