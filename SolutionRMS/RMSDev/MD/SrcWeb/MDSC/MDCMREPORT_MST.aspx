<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMREPORT_MST.aspx.vb" Inherits="MD.MDCMREPORT_MST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��� ���γ���</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987"> <!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET"> <!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include--> <!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->  <!-- UI ���� ActiveX COM --> <!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->  <!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET, mobjEXECUTE, mobjMDSCREPORT_MST'�����ڵ�, Ŭ����
Dim mClientsubcode

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
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
    
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			
			if .txtYEARMON.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			end if
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					
					if .txtYEARMON.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					end if
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub rdA_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdB_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdDS_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdDO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdOS_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdOO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdRS_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdRO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdSS_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdSO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = ""  then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdPS_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdPO_onclick
	SetChangeLayout
	if frmThis.txtYEARMON.value = "" then
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjMDSCREPORT_MST	= gCreateRemoteObject("cMDSC.ccMDSCREPORT_MST")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
												'  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   												'1|
		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GUBUN", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "GUBUN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN",-1,-1,2,2,false
		
		.sprSht.style.visibility = "visible"
    End With
		
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSCREPORT_MST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.rdA.checked = TRUE
		.txtCLIENTNAME1.focus()
		
		SetChangeLayout
	End with	
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
												'  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
   												'1|
		
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GUBUN", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "GUBUN"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN",-1,-1,2,2,false
	End With
End Sub

Sub SetChangeLayout () 
	With frmThis
		gInitComParams mobjSCGLCtl,"MC"
		mobjSCGLCtl.DoEventQueue
		gSetSheetDefaultColor()
		
		Call Grid_init()
		
		if .rdA.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|��ü��|û���μ�|�귣���|�����|�ΰ���|�հ�|��������|���������|�հ�|����ڹ�ȣ|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|	  15|    9|      18|    15|      15|	  15|    10|    10|  10|       7|        10|  10|        12|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | BUSINO | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT | COMMISSION | SUMSUSUVAT", -1, -1,0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1,2
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdB.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|��ü��|û���μ�|�귣���|�����|�ΰ���|�հ�|��������|���������|�հ�|����ڹ�ȣ|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|    9|      18|    15|      15|	  15|    10|    10|  10|       7|        10|  10|        12|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | BUSINO | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT | COMMISSION | SUMSUSUVAT", -1, -1,0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1,2
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdDS.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | TITLE | TIMNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|����|û���μ�|��޾�|�ΰ���|�հ�|�����|������|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|    18|      15|    10|    10|  10|    10|     7|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | TITLE | TIMNAME | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT | COMMISSION", -1, -1,0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1,2
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | TITLE | TIMNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdDO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | TIMNAME | OUT_AMT | VAT | SUMAMTVAT | TITLE | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|û���μ�|���ֺ�|�ΰ���|�հ�|����|����ڹ�ȣ|���|��ǥ��ȣ|���޹��|������|������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|      15|    10|    10|  10|    20|        13|  20|      12|       8|     8|    8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | DOCUMENTDATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | TIMNAME | TITLE | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "OUT_AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | TIMNAME | OUT_AMT | VAT | SUMAMTVAT | TITLE | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN | TRU_VOCH_NO",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|��ü��|û���μ�|�귣���|�����|�ΰ���|�հ�|��������|���������|�հ�|����ڹ�ȣ|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|    15|      15|      15|    10|    10|  10|       7|        10|  10|        12|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | BUSINO | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT | COMMISSION | SUMSUSUVAT", -1, -1,0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1,2
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | REAL_MED_NAME | MEDNAME | TIMNAME | SUBSEQNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | SUMSUSUVAT | BUSINO | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdOS.checked then
		'������Ʈ��		û���μ�		�����	�ΰ���	��  ��					��  ��
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | TIMNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|������Ʈ��|û���μ�|�����|�ΰ���|�հ�|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|        18|      15|    10|    10|  10|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | TIMNAME | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | TIMNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdOO.checked then
		'�ŷ�ó��	ǰ��		û���μ�	�����	�ΰ���	��  ��				����ڹ�ȣ	��  ��
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | TIMNAME | AMT | VAT | SUMAMTVAT | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|ǰ��|û���μ�|�����|�ΰ���|�հ�|����ڹ�ȣ|���|��ǥ��ȣ|���޹��|������|������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|  15|      15|    10|    10|  10|        12|  20|      12|       8|     8|    8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | DOCUMENTDATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | TIMNAME | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | TIMNAME | AMT | VAT | SUMAMTVAT | BUSINO | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN | TRU_VOCH_NO",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdRS.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|������Ʈ��|�����|�ΰ���|�հ�|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|        18|    10|    10|  10|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdRO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|ǰ��|���ֺ�|������Ʈ��|���|��ǥ��ȣ|���޹��|������|������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|  15|    10|        20|  20|      12|       8|     8|    8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | DOCUMENTDATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "OUT_AMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN | TRU_VOCH_NO",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdSS.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|������Ʈ��|�����|�ΰ���|�հ�|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|        18|    10|    10|  10|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | PROJECTNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdSO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|ǰ��|���ֺ�|������Ʈ��|���|��ǥ��ȣ|���޹��|������|������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|   10|      15|  18|    10|        20|  20|      12|       8|     8|    8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | DOCUMENTDATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "OUT_AMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | PROJECTNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN | TRU_VOCH_NO",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdPS.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | JOBNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|���۰Ǹ�|�����|�ΰ���|�հ�|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|    10|    10|  10|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | JOBNAME | MEMO", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | JOBNAME | AMT | VAT | SUMAMTVAT | MEMO"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		elseif .rdPO.checked then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 2, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | JOBNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetHeader .sprSht,        "���|�����ָ�|�� ��|�ŷ�ó��|ǰ��|���ֺ�|���۰Ǹ�|���|��ǥ��ȣ|���޹��|������|������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  8|      13|	 9|      18|  15|    10|      20|  20|      12|       8|     8|    8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | DOCUMENTDATE", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | JOBNAME | MEMO | TRU_VOCH_NO | PAYCODENAME", -1, -1, 100
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "OUT_AMT", -1, -1,0
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | CLIENTNAME | GUBUN | EXCLIENTNAME | SUMM | OUT_AMT | JOBNAME | MEMO | TRU_VOCH_NO | PAYCODENAME | DEMANDDAY | DOCUMENTDATE"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | GUBUN | TRU_VOCH_NO",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "YEARMON | CLIENTNAME | GUBUN"
		else
			Call Grid_init()
		end if
		
   	End With
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE
   	Dim strGUBUN
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		
		IF .rdA.checked THEN
			strGUBUN = .rdA.value
		ELSEIF .rdB.checked THEN
			strGUBUN = .rdB.value
		ELSEIF .rdDS.checked THEN
			strGUBUN = .rdDS.value
		ELSEIF .rdDO.checked THEN
			strGUBUN = .rdDO.value
		ELSEIF .rdO.checked THEN
			strGUBUN = .rdO.value
		ELSEIF .rdOS.checked THEN
			strGUBUN = .rdOS.value
		ELSEIF .rdOO.checked THEN
			strGUBUN = .rdOO.value
		ELSEIF .rdRS.checked THEN
			strGUBUN = .rdRS.value
		ELSEIF .rdRO.checked THEN
			strGUBUN = .rdRO.value
		ELSEIF .rdSS.checked THEN
			strGUBUN = .rdSS.value
		ELSEIF .rdSO.checked THEN
			strGUBUN = .rdSO.value
		ELSEIF .rdPS.checked THEN
			strGUBUN = .rdPS.value
		ELSEIF .rdPO.checked THEN
			strGUBUN = .rdPO.value
		end if
		
		vntData = mobjMDSCREPORT_MST.SelectRtn_REPORT_MST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strGUBUN)

		if not gDoErrorRtn ("SelectRtn_CLIENTYEARCUSTTIMNAMELIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	
		For intCnt = 1 To .sprSht.MaxRows 
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
			If mobjSCGLSpr.GetTextBinding(.sprSht,3,intCnt) = "��  ��" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
			End If
			
			If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,3,intCnt),2) = "�հ�" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
			End If
		Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<FORM id="frmThis" method="post" runat="server"> <!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--Top TR Start-->
				<TR>
					<TD> <!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
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
											<td class="TITLE">��� ���γ���</td>
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
							</TR>
						</TABLE>
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="50">���</TD>
											<TD class="SEARCHDATA" width="150"><INPUT class="INPUT" id="txtYEARMON" title="������Է��ϼ���" style="WIDTH: 89px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="12" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="50">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 143px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="14" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
												<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHDATA" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<TABLE class="SEARCHDATA" id="tblKey" style="BORDER-TOP-STYLE: none" cellSpacing="1" cellPadding="0"
										width="100%" border="0">
										<tr>
											<TD class="SEARCHLABEL" width="50" rowSpan="2">����
											</TD>
											<TD class="SEARCHDATA"><INPUT id="rdA" type="radio" CHECKED value="A" name="chkGBN">&nbsp;����&nbsp;&nbsp;
												<INPUT id="rdB" type="radio" value="B" name="chkGBN">&nbsp;�μ�&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<INPUT id="rdDS" type="radio" value="DS" name="chkGBN">&nbsp;����(��޾�/�����)&nbsp; <INPUT id="rdDO" type="radio" value="DO" name="chkGBN">&nbsp;����(���ֺ�)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<INPUT id="rdO" type="radio" value="O" name="chkGBN">&nbsp;�¶���1&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<INPUT id="rdOS" type="radio" value="OS" name="chkGBN">&nbsp;�¶���2(�����)&nbsp;&nbsp;&nbsp;
												<INPUT id="rdOO" type="radio" value="OO" name="chkGBN">&nbsp;�¶���2(���ֺ�)
											</TD>
										</tr>
										<tr>
											<TD class="SEARCHDATA"><INPUT id="rdRS" type="radio" value="RS" name="chkGBN">&nbsp;���θ��(�����)&nbsp;&nbsp;
												<INPUT id="rdRO" type="radio" value="RO" name="chkGBN">&nbsp;���θ��(���ֺ�)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<INPUT id="rdSS" type="radio" value="SS" name="chkGBN">&nbsp;SP.Comm(�����)&nbsp; <INPUT id="rdSO" type="radio" value="SO" name="chkGBN">&nbsp;SP.Comm(���ֺ�)
												<INPUT id="rdPS" type="radio" value="PS" name="chkGBN">&nbsp;����(�����)&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												<INPUT id="rdPO" type="radio" value="PO" name="chkGBN">&nbsp;����(���ֺ�)
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31856">
										<PARAM NAME="_ExtentY" VALUE="16563">
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
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></FORM>
	</body>
</HTML>
