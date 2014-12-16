<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMEMPMST.aspx.vb" Inherits="MD.MDCMEMPMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��ü�ڵ� ���</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
Dim mobjMDCMEMPMST
Dim mobjMDCMGET
Dim mlngRowCnt,mlngColCnt

Dim mUploadFlag

CONST meTAB = 9
Dim mstrCheck
mstrCheck = True
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
	With frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrMsgBox "������ �ڷᰡ �����ϴ�.","���� �ȳ�"
			exit sub
		end if
	end with
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgpwdinit_onclick ()
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrMsgBox "��й�ȣ �ʱ�ȭ�� �ڷᰡ �����ϴ�.","���� �ȳ�"
			exit sub
		end if
	end with
	ProcessRtn_PWDCHANGE
End Sub
Sub imgE_HR_onclick
Dim intRtn

	with frmThis
	intRtn = gYesNoMsgbox("��������� �ֽ������� ��ü �Ͻðڽ��ϱ�?","EHR-����Ȯ��")
	IF intRtn <> vbYes then exit Sub
	gFlowWait meWAIT_ON
	Call EHR_LOAD()
	window.setTimeout "SelectRtn", 3000	
	gFlowWait meWAIT_OFF
	End With

End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'����������ü ����	
	Set mobjMDCMEMPMST = gCreateRemoteObject("cMDCO.ccMDCOEMPMST")
	set mobjMDCMGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")
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
	strComboList =  "����" & vbTab & "����" & vbTab & "����"
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		'CC_CODE,CC_NAME,OC_CODE,OC_NAME,USE_YN,STDATE,EDATE
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0
		
		mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "EMPNO|EMP_NAME|CC_CODE|BTN|CC_NAME|SC_EMP_STATUS|USE_YN|PASSWORD"
		mobjSCGLSpr.SetHeader .sprSht,		    "���|�����|�μ��ڵ�|�μ���|��������|��뿩��|��й�ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "11  |30    |15    |2|40    |12      |12      |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USE_YN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EMPNO|EMP_NAME|CC_CODE|SC_EMP_STATUS",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CC_NAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EMPNO|EMP_NAME|CC_CODE|CC_NAME"
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,6,6,,,strComboList
	End with
	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDCMEMPMST.GetEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value,.txtEMP_NAME.value,.cmbGUBUN.value,.txtDEPTCD.value,.txtDEPTNAME.value,.cmbUSEYN.value)
		if not gDoErrorRtn ("SelectRtn") then
			If mlngRowCnt > 0 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
			Else
			initpageData
			End If
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
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CC_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CC_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CC_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtEMPNO.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = 7  then 
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 7,7,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
				
			next
		end if
	end with
End Sub
'������� ��Ʈ ��ư Ŭ��

'Validation
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
   	dim vntData
   	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	with frmThis
   		'������ Validation
		'if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"EMPNO|EMP_NAME|CC_CODE|CC_NAME|SC_EMP_STATUS|USE_YN|PASSWORD")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		'ó�� ������ü ȣ��
		intRtn = mobjMDCMEMPMST.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			If intRtn > 0 Then
			gErrorMsgBox intRtn & " �� �� ����Ǿ����ϴ�.","����ȳ�"
			End If
			SelectRtn
   		end if
   	end with
End Sub

Sub ProcessRtn_PWDCHANGE()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ

	with frmThis
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("ProcessRtn_PWDCHANGE") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "��й�ȣ�� �ʱ�ȭ�� �ڷ�" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("��й�ȣ�� �ʱ�ȭ �Ͻðڽ��ϱ�?","���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		for i = intSelCnt-1 to 0 step -1
			strEMPNO = mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",vntData(i))
			
			intRtn = mobjMDCMEMPMST.ProcessRtn_PWDCHANGE(gstrConfigXml,strEMPNO)
		next
		IF not gDoErrorRtn ("ProcessRtn_PWDCHANGE") then
			gOkMsgBox "������� ��й�ȣ�� �ʱ�ȭ �Ǿ����ϴ�.","����ȳ�"
   		End IF
		SelectRtn
	End with
End Sub


'-----------------------------------------------------------------------------------------
' ����μ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------	
Sub ImgCC_onclick
	Call JOBREQU_DEPTCD_POP()
End Sub
Sub JOBREQU_DEPTCD_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtDEPTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = vntRet(0,0)	'Code�� ����
			.txtDEPTNAME.value = vntRet(1,0)	'�ڵ�� ǥ��
		gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub
Sub txtDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPTNAME.value))
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,0)
					.txtDEPTNAME.value = vntData(1,0)
				Else
					Call JOBREQU_DEPTCD_POP()
				End If
   			end if
   		end with
	window.event.returnValue = false
	window.event.cancelBubble = true
	End If
End Sub
Sub EndPage()
	set mobjMDCMEMPMST = Nothing
	set mobjMDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.MaxRows = 0
	End with
End Sub

sub DeleteRtn

	
End Sub



		</script>
		<script language="javascript">
		function EHR_LOAD(){
		ifrm_EHR.location.href = "MDCMEHREMP.asp";		
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;����� ����</td>
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
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 1040px; HEIGHT: 32px" cellSpacing="0" cellPadding="0"
							width="1040" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,'')">���
												</TD>
												<TD class="SEARCHDATA" width="81" style="WIDTH: 81px"><INPUT class="INPUTL" id="txtEMPNO" style="WIDTH: 80px; HEIGHT: 22px" type="text" maxLength="8"
														size="8" name="txtEMPNO" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMP_NAME,'')">�����
												</TD>
												<TD class="SEARCHDATA" width="89" style="WIDTH: 89px"><INPUT class="INPUTL" id="txtEMP_NAME" style="WIDTH: 88px; HEIGHT: 22px" type="text" maxLength="255"
														name="txtEMP_NAME" size="9"></TD>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">�μ���
												</TD>
												<TD class="SEARCHDATA" width="215" style="WIDTH: 215px"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPTNAME" title="���μ���" style="WIDTH: 136px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="100" size="17" name="txtDEPTNAME"><IMG id="imgCC" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="imgCC"><INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPTCD" title="���μ��ڵ�" style="WIDTH: 55px; HEIGHT: 22px"
														accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtDEPTCD">
												</TD>
												<TD class="SEARCHLABEL" width="70">��������
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 105px"><SELECT id="cmbGUBUN" title="��뱸��" style="WIDTH: 105px" name="cmbGUBUN">
														<OPTION value="A" selected>��ü</OPTION>
														<OPTION value="0">����</OPTION>
														<OPTION value="1">����</OPTION>
														<OPTION value="3">����</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" width="70">��뱸��
												</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbUSEYN" title="��뱸��" style="WIDTH: 104px" name="cmbUSEYN">
														<OPTION value="A" selected>��ü</OPTION>
														<OPTION value="Y">���</OPTION>
														<OPTION value="N">�̻��</OPTION>
													</SELECT>
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 1040px; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;��������� ����</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgE_HR" onmouseover="JavaScript:this.src='../../../images/imgE_HROn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgE_HR.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																	src="../../../images/imgE_HR.gIF" border="0" name="imgE_HR"></TD>
															<TD><IMG id="imgpwdinit" onmouseover="JavaScript:this.src='../../../images/imgpwdinitOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgpwdinit.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgpwdinit.gIF" border="0" name="imgpwdinit"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgchangeSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgchangeSave.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgchangeSave.gIF" border="0" name="imgchangeSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--���� �� �׸���-->
								<TR vAlign="top" align="left">
									<!--����-->
									<TD class="DATAFRAME">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 690px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 690px"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="18256">
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
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR> 
			<!--List End-->
			<!--BodySplit Start-->
			<!--Brench End-->
			<!--Bottom Split Start-->
			<TR>
				<TD class="BOTTOMSPLIT" style="WIDTH: 1040px" id="lblstatus"></TD>
			</TR>
			<!--Bottom Split End--> </TBODY></TABLE> 
			<!--Input Define Table End--> 
			</TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE><iframe id="ifrm_EHR" width="0" height="0" frameborder="0"></iframe>
	</body>
</HTML>
