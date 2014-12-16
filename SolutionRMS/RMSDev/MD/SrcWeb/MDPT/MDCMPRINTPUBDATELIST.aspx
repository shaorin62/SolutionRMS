<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTPUBDATELIST.aspx.vb" Inherits="MD.MDCMPRINTPUBDATELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�������೻��</title>
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMGET, mobjEXECUTE'�����ڵ�, Ŭ����

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
	
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	
	if frmThis.txtFPUB_DATE.value = "" AND frmThis.txtTPUB_DATE.value = "" then
		gErrorMsgBox "�������� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
'	if frmThis.txtCLIENTCODE.value = ""  then
'		gErrorMsgBox "�����ָ� �Է��Ͻÿ�","��ȸ�ȳ�"
'		exit Sub
'	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'sub imgPrint_onclick ()
'	gFlowWait meWAIT_ON
'	mobjSCGLSpr.SSPrint  frmThis.sprSht,window.document.title,"",0,0,0,0,true, true,false,2,32,true,-1,-1,-1,1
'	gFlowWait meWAIT_OFF
'end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgCalFrom_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtFPUB_DATE,frmThis.imgCalFrom,""
End Sub

Sub imgCalTo_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtTPUB_DATE,frmThis.imgCalTo,""
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtCLIENTNAME.value), trim(.txtMEDCODE.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMPRINTEXECUSTLISTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetPRINTCLIENT_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt,TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtCLIENTNAME.value), trim(.txtMEDCODE.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME.value = trim(vntRet(6,0))
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			
			if not gDoErrorRtn ("GetCUSTNO_HIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME.value = trim(vntData(6,0))
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' ��ü���ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgMEDCODE_onclick
	Call MED_CODE_POP()
End Sub

'���� ������List ��������
Sub MED_CODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtMEDNAME.value), trim(.txtCLIENTCODE.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMPRINTEXEMEDLISTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtMEDCODE.value = trim(vntRet(0,0))		' Code�� ����
			.txtMEDNAME.value = trim(vntRet(1,0))     ' �ڵ�� ǥ��
			gSetChangeFlag .txtMEDCODE                ' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMEDNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetPRINTMED_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt,TRIM(.txtFPUB_DATE.value), TRIM(.txtTPUB_DATE.value), trim(.txtMEDNAME.value), trim(.txtCLIENTCODE.value))
			if not gDoErrorRtn ("GetPRINTMED_LIST") then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,0))
					.txtMEDNAME.value = trim(vntData(1,0))
				Else
					Call MED_CODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub cmbMED_FLAG_onchange
	Dim strMED_FLAGNAME
	with frmThis
		if frmThis.cmbMED_FLAG.value = "MP01" Then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0, 0,0
			mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 4, 1
			mobjSCGLSpr.SpreadDataField .sprSht, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD_STEP | DAN | STD_CM | COL_DEG | AMOUNT"
			
			mobjSCGLSpr.SetHeader .sprSht,         "������|��ü��|�����|�԰�|�����"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "    15|    30|    30| 6|6|6|6|23"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMOUNT", -1, -1, 0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD_STEP | DAN | STD_CM | COL_DEG | AMOUNT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "PUB_DATE|STD_STEP | DAN | STD_CM | COL_DEG",-1,-1,2,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|PROGRAM_NAME",-1,-1,0,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "PUB_DATE"
			
		elseif frmThis.cmbMED_FLAG.value = "MP02" Then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0, 0,0
			mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 4, 1
			mobjSCGLSpr.SpreadDataField .sprSht, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD | STD_STEP | PAGE | COL_DEG | AMOUNT"
			
			mobjSCGLSpr.SetHeader .sprSht,         "������|��ü��|�����|�԰�|�����"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "    15|    30|    30| 6|6|6|6|23"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMOUNT", -1, -1, 0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD | STD_STEP | PAGE | COL_DEG | AMOUNT"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "PUB_DATE|STD | STD_STEP | PAGE | COL_DEG",-1,-1,2,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|PROGRAM_NAME",-1,-1,0,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "PUB_DATE"
		end if
		SelectRtn
	end with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjEXECUTE	= gCreateRemoteObject("cMDCO.ccMDCOEXECUTE")
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 4, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD_STEP | DAN | STD_CM | COL_DEG | AMOUNT"
		
		mobjSCGLSpr.SetHeader .sprSht,         "������|��ü��|�����|�԰�|�����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "    15|    30|    30| 6|6|6|6|23"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMOUNT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PUB_DATE | MEDNAME | PROGRAM_NAME | STD_STEP | DAN | STD_CM | COL_DEG | AMOUNT"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PUB_DATE|STD_STEP | DAN | STD_CM | COL_DEG",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|PROGRAM_NAME",-1,-1,0,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht, "PUB_DATE"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjEXECUTE = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtFPUB_DATE.value = gNowDate
		.txtTPUB_DATE.value = gNowDate
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtFPUB_DATE.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjEXECUTE.SelectRtn_PUBDATALIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtFPUB_DATE.value, .txtTPUB_DATE.value, .txtCLIENTCODE.value, .txtMEDCODE.value, .txtCLIENTSUBCODE.value, .cmbMED_FLAG.value)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	'For intCnt = 1 To .sprSht.MaxRows 
	'	mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
	'	If mobjSCGLSpr.GetTextBinding(.sprSht,"STD",intCnt) = "�Ұ�" Then
	'	mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
	'	End If
	'Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TBODY>
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
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;�Ź�/�������� �������೻��</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="110" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center"><FONT face="����">
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFPUB_DATE, txtTPUB_DATE)">������</TD>
													<TD class="SEARCHDATA" width="440"><INPUT dataFld="FPUB_DATE" class="INPUT" id="txtFPUB_DATE" title="������" style="WIDTH: 96px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtFPUB_DATE"><IMG id="imgCalFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalFrom">&nbsp;~
														<INPUT dataFld="TPUB_DATE" class="INPUT" id="txtTPUB_DATE" title="������" style="WIDTH: 96px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtTPUB_DATE"><IMG id="imgCalTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalTo">&nbsp; 
														����Ź �ŷ����� ���� ����
													</TD>
													<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">������</TD>
													<TD class="SEARCHDATA" width="440"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 168px; HEIGHT: 22px"
															type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
															border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
															type="text" maxLength="6" align="left" name="txtCLIENTCODE">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
														<SELECT id="cmbMED_FLAG" title="��ü����" style="WIDTH: 80px" name="cmbMED_FLAG">
															<OPTION value="MP01" selected>�Ź�</OPTION>
															<OPTION value="MP02">����</OPTION>
														</SELECT>
													</TD>
												</TR>
												<TR>
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDCODE, txtMEDNAME)">��ü��</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMEDNAME" title="�ڵ��" style="WIDTH: 173px; HEIGHT: 22px" type="text"
															maxLength="100" align="left" name="txtMEDNAME"><IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
															align="absMiddle" border="0" name="ImgMEDCODE"><INPUT class="INPUT_L" id="txtMEDCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px" type="text"
															maxLength="6" align="left" size="5" name="txtMEDCODE">
													</TD>
													<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)">�����</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="�ڵ��" style="WIDTH: 168px; HEIGHT: 22px"
															type="text" maxLength="100" align="left" size="22" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
															border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT_L" id="txtCLIENTSUBCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
															type="text" maxLength="6" align="left" name="txtCLIENTSUBCODE">
													</TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 740px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 738px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 1038px; HEIGHT: 738px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27464">
												<PARAM NAME="_ExtentY" VALUE="19526">
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
								<!--List End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
