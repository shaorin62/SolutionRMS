<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCDDEPTMST.aspx.vb" Inherits="SC.SCCDDEPTMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�������� ����</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : �������� ����
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/27 By KIM TAE YUB
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
		<script language="vbscript" id="clientEventHandlersVBS">
'�������� ����
Dim mobjSCCODEPTMST
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mstrCheck
mstrCheck = True
CONST meTAB = 9
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	mobjSCGLSpr.ExportExcelFile frmThis.sprSht
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgE_HR_onclick
	Dim intRtn
	with frmThis
	
	
		intRtn = gYesNoMsgbox("���������� �ֽ������� ��ü �Ͻðڽ��ϱ�?","EHR-����Ȯ��")
		
		IF intRtn <> vbYes then exit Sub
		
		gFlowWait meWAIT_ON
		Call EHR_LOAD()
		window.setTimeout "SelectRtn", 3000	
		gOkMsgBox "����Ǿ����ϴ�..","����ȳ�!"
		gFlowWait meWAIT_OFF
	End With
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjSCCODEPTMST = gCreateRemoteObject("cSCCO.ccSCCODEPTMST")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub Grid_Layout()
	Dim intGBN
	Dim strComboList 
	
	strComboList =  "���" & vbTab & "�̻��"
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 6, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CC_CODE|CC_NAME|OC_NAME|USE_YN|SDATE|EDATE"
		mobjSCGLSpr.SetHeader .sprSht,		    "�μ��ڵ�|�μ���|�����μ���|�������|�μ�������|�μ�������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "12      |34    |34        |12      |15        |15"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USE_YN"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "SDATE|EDATE"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CC_CODE|USE_YN|SDATE|EDATE",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CC_NAME|OC_NAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CC_CODE|CC_NAME|OC_NAME"
	End with

	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
				
		vntData = mobjSCCODEPTMST.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTCODE.value,.txtDEPTNAME.value,.cmbYN.value)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
				
				for i = 1 to .sprSht.MaxRows
					if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"EDATE",i)) < cdate(mid(Now,1,10)) then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HAAE8EE, &H000000,false
					else
						If i Mod 2 = 0 Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HF4EDE3, &H000000,False
						Else
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HFFFFFF, &H000000,False
						End If
					end if 
				
				next
				
   				
   			Else
   				initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'������� ��Ʈ ��ư Ŭ��
Function DataValidation ()
	DataValidation = false	
	With frmThis
		'IF not gDataValidation(frmThis) then exit Function	
	End With
	DataValidation = True
End Function

'�������
Sub ProcessRtn()
	Dim intRtn
   	Dim vntData
	with frmThis
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CC_CODE|CC_NAME|OC_NAME|USE_YN|SDATE|EDATE")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		'ó�� ������ü ȣ��
		intRtn = mobjSCCODEPTMST.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 1 Then
			gErrorMsgBox intRtn & " �� �� ����Ǿ����ϴ�.","����ȳ�"
			End If
			SelectRtn
   		end if
   	end with
End Sub

Sub EndPage()
	set mobjSCCODEPTMST = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0
	End with
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	With frmThis
		If Row = 0 and Col = 4  Then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 4, 4,,, , , , , , mstrCheck
			if mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
		end if
	end with
End Sub

</script>
<script language="javascript">
function EHR_LOAD(){
	ifrm_EHR.location.href = "SCCDEHRDEPT.asp";		
}
</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF"
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
											<td class="TITLE">�������� ����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1">
								</TD>
							</TR>
						</TABLE>						
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 86px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')"
													width="86">&nbsp;�μ��ڵ�
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 95px"><INPUT class="INPUT_L" id="txtDEPTCODE" style="WIDTH: 96px; HEIGHT: 22px" type="text" maxLength="8"
														size="10" name="txtDEPTCODE" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 86px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTNAME,'')"
													width="86">�μ���
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 249px"><INPUT class="INPUT_L" id="txtDEPTNAME" style="WIDTH: 248px; HEIGHT: 22px" type="text" maxLength="255"
														size="36" name="txtDEPTNAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 86px" width="86">��뱸��
												</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbYN" title="��뱸��" style="WIDTH: 136px" name="cmbYN">
														<OPTION value="A">��ü</OPTION>
														<OPTION value="Y" selected >���</OPTION>
														<OPTION value="N">�̻��</OPTION>
													</SELECT>
												</TD>
												<td class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left">
																<TABLE cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF"
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
															<td class="TITLE">�������� ����&nbsp;</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgE_HR" onmouseover="JavaScript:this.src='../../../images/imgE_HROn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgE_HR.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																	src="../../../images/imgE_HR.gIF" border="0" name="imgE_HR"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgchangeSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgchangeSave.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgchangeSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End-->
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<TR vAlign="top" align="left">
									<!--����-->
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" >
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="19129">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
								<tr><td><iframe id="ifrm_EHR" width="0" height="0" frameborder="0"></iframe></td></tr>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
		
	</body>
</HTML>
