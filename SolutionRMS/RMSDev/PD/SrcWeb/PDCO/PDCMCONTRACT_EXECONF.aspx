<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT_EXECONF.aspx.vb" Inherits="PD.PDCMCONTRACT_EXECONF" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��༭ ��� �� Ȯ��</title>
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
'HISTORY    :1) 2009/11/21 By Ȳ����
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mcomecalender
Dim mobjSCCOCONTRACT, mobjSCCOGET
Dim mstrCheck
Dim mstrChk

CONST meTAB = 9
mcomecalender = FALSE

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
'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
'��ȸ��ư
Sub imgQuery_onclick
	If frmThis.txtFrom.value = ""  and frmThis.txtTO.value = "" Then
		gErrorMsgBox "���Ⱓ�� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	End If
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�ʱ�ȭ��ư
Sub imgCho_onclick
	InitPageData
End Sub

Sub imgNew_onclick
	InitPageData
End Sub

'������ư
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'���� ��ư
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'���ι�ư
Sub imgConf_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_ConfOK
	gFlowWait meWAIT_OFF
End Sub

Sub imgAgreeCanCel_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_ConfCAN
	gFlowWait meWAIT_OFF
End Sub



'������ư
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------------------------------------------------------------
' ������ ��ȸ ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub imgCUSTCODE_onclick
	Call CUST_POP()
End Sub

'���� ������List ��������
Sub CUST_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCUSTCODE.value), trim(.txtCUSTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMCONTRACT_EXE_POP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCUSTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCUSTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
     	
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetCONTRACT_EXE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCUSTCODE.value),trim(.txtCUSTNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCUSTCODE.value = trim(vntData(0,1))
					.txtCUSTNAME.value = trim(vntData(1,1))
				Else
					Call CUST_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'�̹�����ư Ŭ����
Sub imgCUSTCODE1_onclick
	Call CUST_POP1()
End Sub

'���� ������List ��������
Sub CUST_POP1
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCUSTCODE1.value), trim(.txtCUSTNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMCONTRACT_EXE_POP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCUSTCODE1.value = vntRet(0,0) and .txtCUSTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCUSTCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCUSTNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			selectrtn
     	end if
     	
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCUSTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetCONTRACT_EXE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCUSTCODE1.value),trim(.txtCUSTNAME1.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCUSTCODE1.value = trim(vntData(0,1))
					.txtCUSTNAME1.value = trim(vntData(1,1))
					selectrtn
				Else
					Call CUST_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' ������Ʈ�� �� �޷� /
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgTO,"txtTo_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgFROM2_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtSTDATE,frmThis.imgFROM,"txtSTDATE_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO2_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtEDDATE,frmThis.imgTO,"txtEDDATE_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
		gSetChange
	end with
End Sub

'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************

Sub cmbGBN_Onchange
	gSetChange
End Sub

Sub txtCONTRACTNAME_Onchange
	gSetChange
End Sub

Sub txtCUSTNAME_Onchange
	gSetChange
End Sub

Sub txtCUSTCODE_Onchange
	gSetChange
End Sub

Sub txtCONTRACTDAY_Onchange
	gSetChange
End Sub

Sub txtFROM_Onchange
	gSetChange
End SuB

Sub txtTO_Onchange
	gSetChange
End SuB

Sub txtSTDATE_Onchange
	gSetChange
End SuB

Sub txtEDDATE_Onchange
	gSetChange
End Sub

Sub txtAMT_Onchange
	gSetChange
End Sub

Sub txtMEMO_Onchange
	gSetChange
End Sub

'****************************************************************************************
' �����ʵ� �ĸ� ����
'****************************************************************************************
Sub txtAMT_onfocus
	with frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub

Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'****************************************************************************************
' �̺�Ʈ ó��
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		If Row > 0 and Col > 1 Then		
			sprShtToFieldBinding Col,Row
		elseIf Row = 0 and Col = 1  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End if
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	With frmThis
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	End With
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
end sub
	
	

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
			.cmbGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",Row)
			.txtCONTRACTNO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",Row)
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
			.txtCUSTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTNAME",Row)
			.txtCUSTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",Row)
			.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
			.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
			.txtEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			.txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
			.txtCONDITION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONDITION",Row)
			.txtSEQ.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
			.txtCONFIRMFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",Row)
			.txtCONFIRM_USER.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRM_USER",Row)
			.txtCONFIRMDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMDATE",Row)
			
		If .txtAMT.value <> "" Then
			call gFormatNumber(.txtAMT,0,true)
		End If
	End with
End Function


'=============================================
' UI���� ���ν��� 
'=============================================
'---------------------------------------------
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjSCCOCONTRACT	= gCreateRemoteObject("cSCCO.ccSCCOCONTRACT")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")

	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	'------------------------------------------
	'���� ������ ��Ʈ						
	'------------------------------------------
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht	
		mobjSCGLSpr.SpreadLayout .sprSht, 20, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, " CHK | CONFIRMFLAGNAME | GBNNAME | CONTRACTNO | CONTRACTNAME | CUSTNAME | CONTRACTDAY | STDATE | EDDATE | AMT | CONDITION | MEMO | GBN | CUSTCODE | SEQ | CONFIRMFLAG | CONFIRM_USER | CONFIRMDATE | CUSER | CDATE"
		mobjSCGLSpr.SetHeader .sprSht,		 " ����|����|����|����ȣ|����|�����|�����|������|������|�ݾ�|����(����)����|Ư�����|�����ڵ�|������ڵ�|����|�����÷���|������|���γ���|�Է���|�Է³���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  4|   6|  12|      10|    10|    15|    15|    15|    10|  10|            20|      20|       0|         0|   0|         0|    10|      15|    10|     15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONTRACTDAY | STDATE | EDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONFIRMFLAGNAME | GBNNAME | CONTRACTNO | CONTRACTNAME | CUSTNAME | MEMO | CONFIRM_USER | CONFIRMDATE | CUSER | CDATE", -1, -1, 100
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONDITION | MEMO", -1, -1, 1000
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMFLAGNAME | GBNNAME | CONTRACTNO",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprsht, true, "CONFIRMFLAGNAME | GBNNAME | CONTRACTNO | CONTRACTNAME | CUSTNAME | CONTRACTDAY | STDATE | EDDATE | AMT | CONDITION | MEMO | CONFIRM_USER | CONFIRMDATE | CUSER | CDATE"
		mobjSCGLSpr.ColHidden .sprSht, "GBN | CUSTCODE | SEQ | CONFIRMFLAG", true
		
		.sprSht.style.visibility = "visible"
    End With
    
	'ȭ�� �ʱⰪ ����
	InitPageData

End Sub

Sub EndPage()
	set mobjSCCOCONTRACT = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	with frmThis
		.sprSht.MaxRows = 0
	
		.txtCONTRACTDAY.value = gNowDate
		.txtFROM.value = Mid(gNowDate,1,4) & "-"  & Mid(gNowDate,6,2) & "-" & "01"
		.txtSTDATE.value = gNowDate
		.txtEDDATE.value = gNowDate
		
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)

		COMBO_TYPE
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'û���� ��ȸ���� ����
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtTo.value = date2
	end if
End Sub

'------------------------------------------
' select �ڽ� ������ ���ε��� ����
'------------------------------------------
sub COMBO_TYPE()
   	Dim vntData, vntData2
   	
    With frmThis
		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOCONTRACT.Get_COMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt)
		vntData2 = mobjSCCOCONTRACT.Get_COMBO_VALUE2(gstrConfigXml, mlngRowCnt, mlngColCnt)
     
		if not gDoErrorRtn ("Get_COMBO_VALUE") then
			 gLoadComboBox .cmbGBN1, vntData, False
			 gLoadComboBox .cmbGBN, vntData2, False
   		end if
   	end with
end sub

'-----------------------------------------------------------------------------------------
' ��������ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strGBN
	Dim strFROM
	Dim strTO
	Dim strCONTRACTNO
	Dim strCUSTSCODE
	Dim strCUSTSNAME
	Dim strCONFIRMFLAG
	Dim vntData
	
	Dim i, strCols
   	Dim strRows
	Dim intCnt, intCnt2
	Dim strtemp
	
	'On error resume next
	
	with frmThis
		
		.sprSht.MaxRows = 0
		
		strGBN = .cmbGBN1.value
		strFROM = .txtFROM.value
		strTO = .txtTo.value
		strCONTRACTNO = .txtCONTRACTNO1.value 
		strCUSTSCODE = TRIM(.txtCUSTCODE1.value)
		strCONFIRMFLAG = .cmbCONFIRM1.value
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOCONTRACT.SelectRtn_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt, strGBN, strFROM, strTO, strCUSTSCODE, strCONTRACTNO, strCONFIRMFLAG)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			
   			
   			If mlngRowCnt > 0 Then
   				sprShtToFieldBinding 1,1
   			Else
   				.sprSht.MaxRows = 0
   				'InitPageData
   			End If
   			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE			
		END IF
   	end with
End Sub

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn
	Dim intRtn, intRtn2
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strCONTRACTNAME
	
	with frmThis
		strCONTRACTNAME = ""
		
		if DataValidation =false then exit sub
		
		if .cmbGBN.value = "" then
			gErrorMsgBox " ��� ������ �����ϼ���.","����ȳ�" 
			.cmbGBN.focus()
			exit sub
		end if
		
		strCONTRACTNAME = .txtCONTRACTNAME.value
		intRtn2 = gYesNoMsgbox( strCONTRACTNAME & "  �ڷḦ ����/�����Ͻðڽ��ϱ�?","�ڷ����� Ȯ��")
		If intRtn2 <> vbYes Then exit Sub
		
		strMasterData = gXMLGetBindingData (xmlBind)
		
		
		intRtn = mobjSCCOCONTRACT.ProcessRtn(gstrConfigXml, strMasterData)
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			
			SelectRtn
		End If
	End with
End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	
	'On error resume next
	with frmThis
   		IF not gDataValidation(frmThis) then exit Function
   	End with
	DataValidation = true
End Function

'--------------------
'������ ����
'---------------------

Sub ProcessRtn_CONFOK ()
	Dim intRtn, intCnt2
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strchk 
	Dim strYEARMON
	Dim strSEQ
	Dim strGBN
	Dim i
	Dim strCONTRACTDAY
	
	with frmThis
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",intCnt2) <> "N" then
					gErrorMsgBox "üũ�� ������ �� " +  i + " ��° ���� ���´� �̹� ���λ����Դϴ�. �̽��λ����� �����͸� ������ �� �ֽ��ϴ�.","���ξȳ�!"
					Exit Sub
				end if 
				strchk = false
			end if
		Next
		
		if strchk then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���","���ξȳ�!"
			exit sub
		end if
		
		intRtn = gYesNoMsgbox("�ڷḦ ���� �Ͻðڽ��ϱ�?","����Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
	
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows  to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				
				strCONTRACTDAY = MID(REPLACE(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",i),"-",""),1,6)
				
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strGBN = mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",i)
				
				If strSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjSCCOCONTRACT.ProcessRtn_CONFOK(gstrConfigXml,strSEQ, strGBN, strCONTRACTDAY)
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		if not gDoErrorRtn ("ProcessRtn_CONFOK") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			
			SelectRtn
		End If
	End with
End Sub

'------------------------
'���� ���
'------------------------


Sub ProcessRtn_CONFCAN ()
	Dim intRtn, intCnt2
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strchk 
	Dim strSEQ
	Dim i
	Dim strCONTRACTDAY
	
	with frmThis
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",intCnt2) <> "Y" then
					gErrorMsgBox "üũ�� ������ �� " +  i + " ��° ���� ���´� �̹� �̽��λ����Դϴ�. ���λ����� �����͸� ���� ��� �� �� �ֽ��ϴ�.","���ξȳ�!"
					Exit Sub
				end if 
				strchk = false
			end if
		Next
		
		if strchk then
			gErrorMsgBox "���� ��� �� �����͸� üũ�� �ּ���","���ξȳ�!"
			exit sub
		end if
		
		intRtn = gYesNoMsgbox("�ڷḦ ���� ��� �Ͻðڽ��ϱ�?","����Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
	
		
		for i = .sprSht.MaxRows  to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				
				
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				
				If strSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjSCCOCONTRACT.ProcessRtn_CONFCAN(gstrConfigXml,strSEQ)
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		if not gDoErrorRtn ("ProcessRtn_CONFCAN") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ���� ��� " & mePROC_DONE,"����ȳ�" 
			
			SelectRtn
		End If
	End with
End Sub



'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCONTRACTNO
	
	with frmThis
		'���õ� �ڷḦ ������ ���� ����
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		
		IF intRtn <> vbYes then exit Sub
		
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
					strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
					intRtn = mobjSCCOCONTRACT.DeleteRtn(gstrConfigXml,strSEQ)
				End IF
   			End If
   			
   			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,i
   			End IF
		next
		
		gWriteText lblstatus, "�ڷᰡ �����Ǿ����ϴ�."

		SelectRtn
	End with
	err.clear
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="73" background="../../../images/back_p.gIF"
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
											<td class="TITLE">��༭ ��ȸ&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
						<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFrom,txtTO)">���Ⱓ</TD>
								<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 24px"><INPUT class="INPUT" id="txtFrom" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="9" name="txtFrom"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="9" name="txtTo"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										align="absMiddle" border="0" name="imgTo">
								</TD>
								<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME1, txtCUSTCODE1)">�����</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCUSTNAME1" title="����ڸ�" style="WIDTH: 184px; HEIGHT: 22px"
										type="text" maxLength="255" align="left" size="25" name="txtCUSTNAME1"> <IMG id="ImgCUSTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCUSTCODE1">
									<INPUT class="INPUT" id="txtCUSTCODE1" title="������ڵ�" style="WIDTH: 112px; HEIGHT: 22px"
										type="text" maxLength="20" align="left" size="13" name="txtCUSTCODE1"></TD>
								<TD style="WIDTH: 50px"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ ��ȸ�մϴ�."
										src="../../../images/imgQuery.gIF" name="imgQuery"></TD>
							</TR>
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO1, '')">��༭��ȣ</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCONTRACTNO1" title="������ȣ" style="WIDTH: 216px; HEIGHT: 22px"
										type="text" maxLength="255" align="left" size="30" name="txtCONTRACTNO1">
								</TD>
								<TD class="SEARCHLABEL">�������</TD>
								<TD class="SEARCHDATA" vAlign="middle" align="left" colSpan="2"><SELECT class="INPUT" id="cmbGBN1" title="�������" style="WIDTH: 120px" name="cmbGBN1"></SELECT>
									<SELECT id="cmbCONFIRM1" title="��������" style="WIDTH: 65px" name="cmbCONFIRM1">
										<OPTION value="" selected>��ü</OPTION>
										<OPTION value="Y">����</OPTION>
										<OPTION value="N">�̽���</OPTION>
									</SELECT>
								</TD>
							</TR>
						</TABLE>
						<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="350" height="20">
									<table id="TABLE1" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
											<td class="TITLE">��༭ ���</td>
										</tr>
									</table>
								</TD>
								<TD id="TD1" vAlign="middle" align="right" height="20" runat="server">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="ImgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" height="20" alt="�ڷḦ�ʱ�ȭ�մϴ�"
													src="../../../images/imgCho.gif" width="64" border="0" name="imgFind"></TD>
											<td><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
													src="../../../images/imgNew.gIF" border="0" name="imgNew"></td>
											<TD><IMG id="imgConf" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgAgree.gIF" border="0" name="imgConf"></TD>
											<TD><IMG id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelON.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'"
													height="20" alt="�ڷḦ ��������մϴ�." src="../../../images/imgAgreeCanCel.gIF" border="0"
													name="imgAgreeCanCel"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<!--TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD-->
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 11px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1"
										cellPadding="0" align="left" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCONTRACTTYPE, '')">�������</TD>
											<TD class="SEARCHDATA" style="WIDTH: 200px; HEIGHT: 25px"><SELECT dataFld="GBN" class="INPUT" id="cmbGBN" title="�������" style="WIDTH: 122px" dataSrc="#xmlBind"
													name="cmbGBN"></SELECT><INPUT dataFld="CONTRACTNO" class="NOINPUT_L" id="txtCONTRACTNO" title="��༭��ȣ" style="WIDTH: 70px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="6" size="10" name="txtCONTRACTNO">
											<TD class="SEARCHLABEL" style="WIDTH: 50px; CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')">����</TD>
											<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 25px"><INPUT dataFld="CONTRACTNAME" class="INPUT_L" id="txtCONTRACTNAME" title="����" style="WIDTH: 245px; HEIGHT: 21px"
													accessKey=",M" dataSrc="#xmlBind" type="text" size="33" name="txtCONTRACTNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 50px; CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCUSTNAME, txtCUSTCODE)">�����</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="CUSTNAME" class="INPUT_L" id="txtCUSTNAME" title="����ڸ�" style="WIDTH: 130px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="32" name="txtCUSTNAME">
												<IMG id="ImgCUSTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCUSTCODE">
												<INPUT dataFld="CUSTCODE" class="INPUT" id="txtCUSTCODE" title="������ڵ�" style="WIDTH: 109px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="20" align="left" size="12"
													name="txtCUSTCODE"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCONTRACTDAY, '')">�����</TD>
											<TD class="SEARCHDATA" style="HEIGHT: 25px"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="�����" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY">
												<IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
													height="15" alt="ImgCONTRACTDAY" src="../../../images/btnCalEndar.gIF" align="absMiddle"
													border="0" name="ImgCONTRACTDAY"> <INPUT dataFld="SEQ" id="txtSEQ" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind" type="hidden"
													name="txtSEQ"> <INPUT dataFld="CONFIRMFLAG" id="txtCONFIRMFLAG" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtCONFIRMFLAG"><INPUT dataFld="CONFIRM_USER" id="txtCONFIRM_USER" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtCONFIRM_USER"><INPUT dataFld="CONFIRMDATE" id="txtCONFIRMDATE" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtCONFIRMDATE">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSTDATE,txtEDDATE)">�Ⱓ</TD>
											<TD class="SEARCHDATA" style="HEIGHT: 25px"><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="���Ⱓ ������" style="WIDTH: 95px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtSTDATE">
												<IMG id="imgFROM2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
													height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgFROM2">&nbsp;~
												<INPUT dataFld="EDDATE" class="INPUT" id="txtEDDATE" title="���Ⱓ ������" style="WIDTH: 95px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtEDDATE">
												<IMG id="imgTO2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
													height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgTO2">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtAMT,'')">���ݾ�</TD>
											<TD class="SEARCHDATA" style="HEIGHT: 25px"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="���ݾ�" style="WIDTH: 130px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="18" name="txtAMT"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONDITION, '')">����(����)����</TD>
											<TD class="SEARCHDATA" colSpan="8"><TEXTAREA dataFld="CONDITION" id="txtCONDITION" title="����(����)����" style="WIDTH: 816px; HEIGHT: 36px"
													accessKey=",M" dataSrc="#xmlBind" name="txtCONDITION" wrap="hard" cols="99"></TEXTAREA></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEMO, '')">Ư�����</TD>
											<TD class="SEARCHDATA" colSpan="8"><TEXTAREA dataFld="MEMO" id="txtMEMO" title="Ư�̻���" style="WIDTH: 816px; HEIGHT: 36px" dataSrc="#xmlBind"
													name="txtMEMO" wrap="hard" cols="99"></TEXTAREA></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27490">
								<PARAM NAME="_ExtentY" VALUE="8784">
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
								<PARAM NAME="EditEnterAction" VALUE="5">
								<PARAM NAME="EditModePermanent" VALUE="0">
								<PARAM NAME="EditModeReplace" VALUE="0">
								<PARAM NAME="FormulaSync" VALUE="-1">
								<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
								<PARAM NAME="GridColor" VALUE="12632256">
								<PARAM NAME="GridShowHoriz" VALUE="1">
								<PARAM NAME="GridShowVert" VALUE="1">
								<PARAM NAME="GridSolid" VALUE="1">
								<PARAM NAME="MaxCols" VALUE="19">
								<PARAM NAME="MaxRows" VALUE="0">
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
				</tr>
				<!--BodySplit End-->
				<!--List Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE></TR></TABLE></TR></TABLE></TR></TABLE></FORM>
	</body>
</HTML>
