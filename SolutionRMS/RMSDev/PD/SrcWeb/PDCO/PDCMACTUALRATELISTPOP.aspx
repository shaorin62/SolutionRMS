<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMACTUALRATELISTPOP.aspx.vb" Inherits="PD.PDCMACTUALRATELISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�����й��� ����Ʈ</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/���۰�����ȣ ��� ȭ��
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBNO.aspx
'��      �� : ���۰�����ȣ C/D/U/R
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/19 By Kim Tae Ho
'			 2) 
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt 
Dim mobjPDCOACTUALRATE '�����ڵ�, Ŭ����
Dim mobjPDCOGET
Dim mstrCheck
Const meTab = 9

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------
' ��� ��ư
'-----------------------------	
Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'****************************************************************************************
' SHEET���� ����
'****************************************************************************************
Sub sprSht_JOBNODEPT_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_JOBNODEPT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_JOBNODEPT.MaxRows
				sprSht_JOBNODEPT_Change 1, intcnt
			next
		end if
	end with
End Sub

sub sprSht_JOBNODEPT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_JOBNODEPT, ""
		end if
	end with
end sub


'=========================================================================================
' UI���� ���ν��� ����  - INIT,,  INITPAGEDATA ...
'=========================================================================================
'-----------------------------------------------------------------------------------------
Sub InitPage()
' ������ ȭ�� ������ �� �ʱ�ȭ 
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjPDCOACTUALRATE	= gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis		
		gSetSheetColor mobjSCGLSpr, .sprSht_JOBNODEPT
		mobjSCGLSpr.SpreadLayout .sprSht_JOBNODEPT, 7, 0
		mobjSCGLSpr.SpreadDataField .sprSht_JOBNODEPT, "SEQ | EMPNAME | EMPNO | DEPTNAME | DEPTCODE | JOBNOSEQ | ACTRATE"
		mobjSCGLSpr.SetHeader .sprSht_JOBNODEPT,        "����|�����|����ڻ��|���μ�|���μ��ڵ�|JOBSEQ|�μ������Է�"
		mobjSCGLSpr.SetColWidth .sprSht_JOBNODEPT, "-1","   5|	  10|        10|      28|          10|     6|          12" 
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_JOBNODEPT, "ACTRATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "EMPNAME | EMPNO | DEPTNAME | DEPTCODE | JOBNOSEQ",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT, true, "SEQ | JOBNOSEQ"
		mobjSCGLSpr.SetScrollBar .sprSht_JOBNODEPT,2,True,0,-1
		mobjSCGLSpr.colhidden .sprSht_JOBNODEPT, "JOBNOSEQ | SEQ",true
		
		.sprSht_JOBNODEPT.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)
				case 1 : .txtJOBNAME.value = vntInParam(i)
			end select
		next
		
		if .txtJOBNO.value <> "" then
			SelectRtn
		end IF
	end with
End Sub
'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.sprSht_JOBNODEPT.MaxRows = 0
		.txtJOBNAME.focus
	End with
End Sub

Sub EndPage()
	set mobjPDCOACTUALRATE = Nothing
	set mobjPDCOGET = Nothing
	gEndPage
End Sub


'****************************************************************************************
' UI ���� - ��ȸ ���� ���� ����  
'****************************************************************************************
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim strRow,strJOBNO , strJOBNOSEQ
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_JOBNODEPT.MaxRows = 0
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'��Ʈ1�� JOB��ȣ�� ��������µ� ���
		If .txtJOBNO.value = "" Then
			gErrorMsgBox "JOB�� ���� �Ͻʽÿ�.","��ȸ�ȳ�"
			Exit Sub
		Else
			strJOBNO = .txtJOBNO.value 
		End If
		'��Ʈ1�� JOBNOSEQ��ȣ�� ��������µ� ���
		
		strJOBNOSEQ = "1"
		
		vntData = mobjPDCOACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,strJOBNOSEQ)
		
		If not gDoErrorRtn ("SelectRtn_DTL_JOBNODEPT") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_JOBNODEPT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht_JOBNODEPT,meCLS_FLAG
			
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
		End If		
	END WITH
End Sub
'****************************************************************************************
' SHEET���� ��
'****************************************************************************************
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML><XML id="xmlBind1"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="80%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD colSpan="2">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" width="400" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
														border="0">
														<TR>
															<TD align="left" width="100%" height="2"></TD>
														</TR>
													</TABLE>
												</td>
											</tr>
											<tr>
												<td height="3"></td>
											</tr>
											<tr>
												<td class="TITLE">�����й��� ����Ʈ</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="800" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="40%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 5px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="top">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 47px; CURSOR: hand" onclick="vbscript:Call CleanField(txtJOBNAME, txtJOBNO)">JOB 
													��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 80px"><INPUT class="INPUT_L" id="txtJOBNAME" title="�ڵ��" style="WIDTH: 208px; HEIGHT: 22px" type="text"
														maxLength="255" align="left" size="29" name="txtJOBNAME"> 
												<TD class="SEARCHLABEL" style="WIDTH: 47px; CURSOR: hand" onclick="vbscript:Call CleanField(txtJOBNAME, txtJOBNO)">JOB 
													NO</TD>
												<td>
												<INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 56px; HEIGHT: 22px" accessKey=",M"
														type="text" maxLength="7" size="4" name="txtJOBNO">
												</td>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<tr>
						<td colSpan="2">
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left">
																<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
																	border="0">
																	<TR>
																		<TD align="left" width="100%" height="2"></TD>
																	</TR>
																</TABLE>
															</td>
														</tr>
														<tr>
															<td height="3"></td>
														</tr>
														<tr>
															<td class="TITLE">���μ�/�����</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start--></TD>
											</TR>
										</TABLE>
									</TD>
									<!--job ���� ��ư�ִ��ڸ�--></TR>
							</TABLE>
						</td>
					</tr>
					<!--Input End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 400px" vAlign="top" align="center" colSpan="2">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht_JOBNODEPT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31803">
									<PARAM NAME="_ExtentY" VALUE="10583">
									<PARAM NAME="_StockProps" VALUE="64">
									<PARAM NAME="Enabled" VALUE="-1">
									<PARAM NAME="AllowCellOverflow" VALUE="0">
									<PARAM NAME="AllowDragDrop" VALUE="0">
									<PARAM NAME="AllowMultiBlocks" VALUE="0">
									<PARAM NAME="AllowUserFormulas" VALUE="0">
									<PARAM NAME="ArrowsExitEditMode" VALUE="0">
									<PARAM NAME="AutoCalc" VALUE="-1">
									<PARAM NAME="AutoClipboard" VALUE="-1">
									<PARAM NAME="AutoSize" VALUE="0">
									<PARAM NAME="BackColorStyle" VALUE="0">
									<PARAM NAME="BorderStyle" VALUE="1">
									<PARAM NAME="ButtonDrawMode" VALUE="0">
									<PARAM NAME="ColHeaderDisplay" VALUE="2">
									<PARAM NAME="ColsFrozen" VALUE="0">
									<PARAM NAME="DAutoCellTypes" VALUE="1">
									<PARAM NAME="DAutoFill" VALUE="1">
									<PARAM NAME="DAutoHeadings" VALUE="1">
									<PARAM NAME="DAutoSave" VALUE="1">
									<PARAM NAME="DAutoSizeCols" VALUE="2">
									<PARAM NAME="DInformActiveRowChange" VALUE="1">
									<PARAM NAME="DisplayColHeaders" VALUE="1">
									<PARAM NAME="DisplayRowHeaders" VALUE="1">
									<PARAM NAME="EditEnterAction" VALUE="0">
									<PARAM NAME="EditModePermanent" VALUE="0">
									<PARAM NAME="EditModeReplace" VALUE="0">
									<PARAM NAME="FormulaSync" VALUE="-1">
									<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
									<PARAM NAME="GridColor" VALUE="12632256">
									<PARAM NAME="GridShowHoriz" VALUE="1">
									<PARAM NAME="GridShowVert" VALUE="1">
									<PARAM NAME="GridSolid" VALUE="1">
									<PARAM NAME="MaxCols" VALUE="500">
									<PARAM NAME="MaxRows" VALUE="500">
									<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
									<PARAM NAME="NoBeep" VALUE="0">
									<PARAM NAME="NoBorder" VALUE="0">
									<PARAM NAME="OperationMode" VALUE="0">
									<PARAM NAME="Position" VALUE="0">
									<PARAM NAME="ProcessTab" VALUE="0">
									<PARAM NAME="Protect" VALUE="-1">
									<PARAM NAME="ReDraw" VALUE="1">
									<PARAM NAME="RestrictCols" VALUE="0">
									<PARAM NAME="RestrictRows" VALUE="0">
									<PARAM NAME="RetainSelBlock" VALUE="-1">
									<PARAM NAME="RowHeaderDisplay" VALUE="1">
									<PARAM NAME="RowsFrozen" VALUE="0">
									<PARAM NAME="ScrollBarExtMode" VALUE="0">
									<PARAM NAME="ScrollBarMaxAlign" VALUE="-1">
									<PARAM NAME="ScrollBars" VALUE="3">
									<PARAM NAME="ScrollBarShowMax" VALUE="-1">
									<PARAM NAME="SelectBlockOptions" VALUE="15">
									<PARAM NAME="ShadowColor" VALUE="-2147483633">
									<PARAM NAME="ShadowDark" VALUE="-2147483632">
									<PARAM NAME="ShadowText" VALUE="-2147483630">
									<PARAM NAME="StartingColNumber" VALUE="1">
									<PARAM NAME="StartingRowNumber" VALUE="1">
									<PARAM NAME="UnitType" VALUE="1">
									<PARAM NAME="UserResize" VALUE="3">
									<PARAM NAME="VirtualMaxRows" VALUE="-1">
									<PARAM NAME="VirtualMode" VALUE="0">
									<PARAM NAME="VirtualOverlap" VALUE="0">
									<PARAM NAME="VirtualRows" VALUE="0">
									<PARAM NAME="VirtualScrollBuffer" VALUE="0">
									<PARAM NAME="VisibleCols" VALUE="0">
									<PARAM NAME="VisibleRows" VALUE="0">
									<PARAM NAME="VScrollSpecial" VALUE="0">
									<PARAM NAME="VScrollSpecialType" VALUE="0">
									<PARAM NAME="Appearance" VALUE="0">
									<PARAM NAME="TextTip" VALUE="0">
									<PARAM NAME="TextTipDelay" VALUE="500">
									<PARAM NAME="ScrollBarTrack" VALUE="0">
									<PARAM NAME="ClipboardOptions" VALUE="15">
									<PARAM NAME="CellNoteIndicator" VALUE="0">
									<PARAM NAME="ShowScrollTips" VALUE="0">
									<PARAM NAME="DataMember" VALUE="">
									<PARAM NAME="OLEDropMode" VALUE="0">
								</OBJECT>
							</DIV>
						</TD>
					</TR>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%" colSpan="2"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
