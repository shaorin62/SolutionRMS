<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTTEST.aspx.vb" Inherits="SC.SCRTTest" codePage="949" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>SCFUTest</TITLE>
		<META content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<META content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<META content="VBScript" name="vs_defaultClientScript">
		<META content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema"><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActivX COM ClassID -->
		<!--#INCLUDE VIRTUAL = "../../../Etc/SCUIClass.inc"-->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Sub Window_OnLoad
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitPageSetting mobjSCGLCtl,"SC"
	frmThis.txtParams.value = gStrUsrBU &":����"
End Sub

'��ưŬ�� �̺�Ʈ      

Sub imgOnePrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	Dim Opt
	
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
		Params = gStrUsrBU & ":����"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
 
	if frmThis.txtOpt.value <> "B" then
		Opt = "A"	
	Else
	    Opt = "B"	
	end if		  

 	gShowReportWindow ModuleDir, ReportName, Params, Opt

	gFlowWait meWAIT_OFF
End Sub

Sub imgTwoPrint_onclick
	gFlowWait meWAIT_ON

	Dim ModuleDir 	
	Dim ReportName 
	Dim Params 
	Dim Opt,i
	
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
		Params = gStrUsrBU & ":����"
	Else
		Params = frmThis.txtParams.value
	End if                                                     
 
	if frmThis.txtOpt.value <> "B" then
		Opt = "A"	
	Else
	    Opt = "B"	
	end if
	 
	For i=0 To 3
		Select Case i
			Case 0 : Params = gStrUsrBU &":����"
			Case 1 : Params = gStrUsrBU &":����"
			Case 2 : Params = gStrUsrBU &":����"
			Case 3 : Params = gStrUsrBU &":�繫"
		End Select
				
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	Next 
	gFlowWait meWAIT_OFF
End Sub


Sub btnIFRAME_Onclick 
	'gShowIFrameReport(���������ӳ���, ���,����Ʈ��,�Ķ����,�ɼ�)
    'gShowiFrameReport(iFrameName, Module,ReportName,Params, Opt)                                             
	 gShowiFrameReport ifrTest, "SC", "SCMENU.rpt", "HPC:����", "A"
End Sub

//-->
		</SCRIPT>
	</HEAD>
	<BODY class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)">
		<FORM id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="790">
				<TBODY>
					<TR>
						<TD style="WIDTH: 790px">
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gif" border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD align="left" width="49" rowSpan="2"><IMG id="imgTEST" height="28" src="../../images/TitleIcon.gif" width="49"></TD>
												<TD align="left" height="4"></TD></TR>
											<TR>
												<TD class="TITLE">ũ����Ż ����Ʈ �׽�Ʈ
													������</TD></TR></TABLE></TD></TR></TABLE><FONT face="����" size="2">
								&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								&nbsp;<BR>&nbsp;* ����Ʈ ����&nbsp;�μ� &gt;&gt;</FONT><IMG id="imgOnePrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gif" width="54" align="absMiddle" border="0" name="imgOnePrint"><FONT face="����" size="2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
								*
								����Ʈ ���� �μ� &gt;&gt;</FONT><IMG id="ImgTwoPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gif" width="54" align="absMiddle" border="0" name="ImgTwoPrint"> <FONT face="����" size="2">
								(�������� ��� �Ұ���!! �����߻�)</FONT> <FONT face="����" size="2"><BR></FONT><FONT face="����" size="2"><BR></FONT>
							<TABLE id="Table1" style="WIDTH: 704px; HEIGHT: 192px" cellSpacing="1" cellPadding="1" width="704" border="1">
								<TR class="EVENROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16px">
										<P align="left"><FONT face="����" size="2">ModuleDir</FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 7.066pt"><FONT face="����" size="2"></FONT>
										<P align="left"><INPUT id="txtModuleDir" style="WIDTH: 64px; HEIGHT: 22px" type="text" size="5" value="SC"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16px">
										<P align="left"><FONT face="����" size="2">&nbsp;����: ����Ʈ�� ���� ��ġ�� �������� ��� ���丮 <BR>&nbsp;(��:
												SC, CO, PO,
												AP&nbsp;��...)</FONT></P></TD></TR>
								<TR class="ODDROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 14px">
										<P align="left"><FONT face="����" size="2">ReportName </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 7.381pt">
										<P align="left"><INPUT id="txtReportName" style="WIDTH: 224px; HEIGHT: 22px" type="text" size="32" value="SCMENU.rpt"></P></TD>
									<TD class="LABEL" style="HEIGHT: 14px">
										<P align="left"><FONT face="����" size="2">&nbsp;����Ʈ�� �̸�(��: SCMENU.rpt )</FONT></P></TD></TR>
								<TR class="EVENROW">
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16.571pt">
										<P align="left"><FONT face="����" size="2">Params </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 16.571pt">
										<P align="left"><INPUT id="txtParams" style="WIDTH: 224px; HEIGHT: 22px" type="text" size="32"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16.571pt">
										<P align="left"><FONT face="����" size="2">&nbsp;�Ķ���� ��(��: SJCC:����)</FONT></P></TD></TR>
								<TR>
									<TD class="LABEL" style="WIDTH: 101px; HEIGHT: 16.571pt">
										<P align="left"><FONT face="����" size="2">Option </FONT></P></TD>
									<TD class="DATA" style="WIDTH: 231px; HEIGHT: 16.571pt">
										<P align="left"><INPUT id="txtOpt" style="WIDTH: 80px; HEIGHT: 22px" type="text" size="8" value="A" name="Text1"></P></TD>
									<TD class="LABEL" style="HEIGHT: 16.571pt">
										<P align="left"><FONT face="����" size="2">&nbsp;ȭ�����:A ���������:B</FONT></P></TD></TR>
								<TR class="ODDROW">
									<TD class="DATA" colSpan="3"><FONT face="����"></FONT>
										<P align="left"><FONT face="����"><FONT color="#ff0000"><STRONG><BR>*</STRONG>&nbsp;</FONT><STRONG><FONT color="#ff0000">���ǻ���<BR>
														&nbsp;1.&nbsp;�Ķ���ʹ� String Type ���� �Ѿ�
														���ϴ�.&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp;�������̳�
														��¥���� ��� ����Ʈ SQL������ TO_NUMBER, TO_DATE �� ��ȯ�Ͽ�&nbsp;����մϴ�.<BR>&nbsp;2. ����Ʈ����
														*&nbsp;�� % �� �νĵǵ���&nbsp;�Ǿ� �ֽ��ϴ�.
														LIKE % �� ������&nbsp;����ϽǶ��� * �� �Ѱ��ֽø�
														�˴ϴ�.<BR></FONT></STRONG></FONT></P></TD></TR></TABLE>
							<P></P><FONT face="����">
								<P></FONT><FONT face="����" size="2">&nbsp;&nbsp; * ũ����Ż ����Ʈ �� ��ġ �ȵɶ� �Ʒ���
								������
								�ٿ�����ż� �������� ��ġ�Ͻʽÿ�.</FONT><FONT face="����" size="2"><BR>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
								����Ʈ ��� �ٿ�ޱ� --&gt;
								Crystal Report Viewer&nbsp;<BR></FONT><FONT face="����"><FONT size="2"><STRONG>&nbsp; </STRONG><FONT face="����">
										*&nbsp; ������
										���ǻ���&nbsp;<BR>&nbsp;&nbsp;&nbsp;&nbsp; &nbsp;����â�� ��
										��쿡�� ���� ����Ʈ���� �ݵ�� Window â�� �̸��� �ٸ��� �Ͽ���
										�մϴ�.<BR></FONT></FONT></P></FONT></TD></TR>
					<TR><TD><INPUT type="button" value="iFrameTest ��ư" id="btnIFRAME" class="button"></TD></TR>
					<TR>
						<TD><IFRAME id="ifrtest" src="" frameborder="0" style="BORDER-RIGHT: #6699ff 1px solid; BORDER-TOP: #6699ff 1px solid; BORDER-LEFT: #6699ff 1px solid; WIDTH: 704px; BORDER-BOTTOM: #6699ff 1px solid; HEIGHT: 152px" scrolling="no"></IFRAME></TD>
					</TR>
				</TBODY></TABLE></TD></TR></TBODY></TABLE></FORM>
	</BODY>
</HTML>
