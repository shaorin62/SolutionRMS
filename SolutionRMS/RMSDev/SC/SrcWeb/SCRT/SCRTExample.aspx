<%@ Page CodeBehind="SCRTExample.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="SC.SCRTExample" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>���� ��ȸ ����</TITLE> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/SC/���� ��ȸ ����(��ǥ����)(SCRTExample)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCRTExample.aspx
'��      �� : ���� ��ȸ
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2007/10/25 By Kim Jung Hoon
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="JavaScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM Class ID -->
		<!-- #INCLUDE VIRTUAL=../../../Etc/SCUIClass.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mRowCnt

'=============================
' �̺�Ʈ ���ν��� 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgPrint_onclick	'��¹�ư Ŭ����
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	
	gFlowWait meWAIT_ON
		with frmThis
		          			  
			ModuleDir = "SC"

			ReportName = "SCRTEXAMPLE.rpt"
			
			Params = .txtTYY_MM.value & ":" & .txtFYY_MM.value
            Opt = "A"
		end with
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	gFlowWait meWAIT_OFF
End Sub
'-----------------------------------
' ��Ÿ change  �̺�Ʈ
'-----------------------------------
Sub txtFYY_MM_onchange
	gSetChange
End Sub

Sub txtTYY_MM_onchange
	gSetChange
End Sub

'=============================
' UI���� ���ν��� 
'=============================
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage
 '����������ü����
 '���Ѽ���/������ ����/ȭ������ ���� �⺻ �۾��� ����
	gInitPageSetting mobjSCGLCtl,"SC" 

	InitPageData
End Sub

Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis

	End With
End Sub

Sub EndPage
    gEndPage    
End Sub 

-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gIf)">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="420" border="0" style="WIDTH: 420px">
				<TR>
					<TD style="WIDTH: 538px"><FONT face="����"></FONT>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIf"
							border="0">
							<TR>
								<td style="WIDTH: 293px" align="left" width="293" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../images/TitleIcon.gIf" width="49"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</TR>
										<tr>
											<td class="TITLE"><FONT face="����">���� ��ȸ ����</FONT></td>
										</tr>
									</table>
								<TD vAlign="middle" align="center" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 180px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIf"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 100px; HEIGHT: 24px" cellSpacing="0" cellPadding="0"
										width="204" border="0">
										<TR>
											<TD></TD>
											<TD width="3"><FONT face="����"><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gIf'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gIf'" height="20" alt="�ڷḦ �μ��մϴ�."
														src="../../../images/imgPrint.gIf" width="54" border="0" name="imgPrint"></FONT></TD>
											<TD></TD>
											<TD style="WIDTH: 161px; HEIGHT: 24px"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIf'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIf'" height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIf"
													width="54" border="0" name="imgClose"></TD>
											<TD></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblform1" style="WIDTH: 419px" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 790px" align="center"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 790px; HEIGHT: 8px" vAlign="middle">
									<TABLE id="Table1" style="WIDTH: 418px" cellSpacing="1" cellPadding="0" width="418" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 20%"><FONT face="����"> JOBCUST</FONT></TD>
											<TD class="DATA" width="80%"><FONT face="����"></FONT><FONT face="����">&nbsp;<INPUT class="INPUT" id="txtFYY_MM" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="7">&nbsp;&nbsp;~&nbsp;<INPUT class="INPUT" id="txtTYY_MM" style="WIDTH: 100px; HEIGHT: 22px" type="text" size="7"></FONT>&nbsp;
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				<TR>
					<TD class="BOTTOMSPLIT" style="WIDTH: 790px; HEIGHT: 1px" width="790"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
		</form>
		</TD></TR></TABLE></SCRIPT>
	</body>
</HTML>
