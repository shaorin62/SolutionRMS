<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVVOCH.aspx.vb" Inherits="MD.MDCMCATVVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���̺� ��ǥ����</title>
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
'HISTORY    :1) 2009/11/24 By Ȳ����
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
Dim mlngRowCnt,mlngColCnt
Dim mobjMDCMCATVVOCH
Dim mobjMDCOVOCH
Dim mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim vntData_ProcesssRtn_SUSU
Dim mstrPROCESS
Dim mstrSTAY

mstrSTAY = TRUE

mstrGUBUN = "T"
mstrPROCESS = ""
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

'�������� ��ư �����
Sub Set_delete(byVal strmode)
	With frmThis
		IF .rdT.checked = TRUE then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'��ȸ��ư
Sub imgQuery_onclick
	If frmThis.txtYEARMON.value = "" Then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	End If

	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub btnTab1_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
	frmThis.btnTab3.style.backgroundImage = meURL_TAB
	frmThis.btnTab4.style.backgroundImage = meURL_TAB
	
	pnlTab_amt.style.visibility = "visible" 
	pnlTab_gen.style.visibility = "hidden" 
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "hidden" 
	
	pnlGUBUN.style.visibility = "hidden" 
	pnlFLAG.style.visibility = "hidden" 
	
	frmThis.txtEDYEARMON.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "T"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

Sub btnTab2_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TABON
	frmThis.btnTab3.style.backgroundImage = meURL_TAB
	frmThis.btnTab4.style.backgroundImage = meURL_TAB
	
	pnlTab_amt.style.visibility = "hidden" 
	pnlTab_gen.style.visibility = "visible" 
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "hidden" 
	
	pnlGUBUN.style.visibility = "visible" 
	pnlFLAG.style.visibility = "hidden" 
	
	frmThis.txtEDYEARMON.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "D"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

Sub btnTab3_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
	frmThis.btnTab3.style.backgroundImage = meURL_TABON
	frmThis.btnTab4.style.backgroundImage = meURL_TAB
	
	pnlTab_amt.style.visibility = "hidden" 
	pnlTab_gen.style.visibility = "hidden" 
	pnlTab_susu.style.visibility = "visible" 
	pnlTab_out.style.visibility = "hidden" 
	
	pnlGUBUN.style.visibility = "hidden" 
	pnlFLAG.style.visibility = "hidden" 
	
	frmThis.txtEDYEARMON.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "S"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

Sub btnTab4_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
	frmThis.btnTab3.style.backgroundImage = meURL_TAB
	frmThis.btnTab4.style.backgroundImage = meURL_TABON
	
	pnlTab_amt.style.visibility = "hidden" 
	pnlTab_gen.style.visibility = "hidden" 
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "visible" 
	
	pnlGUBUN.style.visibility = "hidden" 
	pnlFLAG.style.visibility = "visible" 
	
	frmThis.txtEDYEARMON.style.visibility = "visible"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "O"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub


'������ư Ŭ��
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		
		if mstrGUBUN = "T"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_AMT
		elseif mstrGUBUN = "D"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_GEN
		elseif mstrGUBUN = "S"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_SUSU
		elseif mstrGUBUN = "O"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_OUT
		end if
		
	End With
	gFlowWait meWAIT_OFF
End Sub


'�ݱ��ư Ŭ��
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'��ǥ���� Ŭ��
Sub ImgvochCre_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� Ŭ��
Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'������ǥ����Ŭ��
Sub ImgErrVochDel_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� ���� Ŭ��
Sub imgVochDelco_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub


'--�����ư Ŭ��
Sub ImgSUMMApp_onclick()
	Dim intRtn
	
	with frmThis
		if .cmbSETTING.value = "" then
			gErrorMsgBox "�����Ͻ� �÷� ���� �����ϴ�. ","�������"
			exit sub
		end if 
		
		'��޾�
		if mstrGUBUN = "T"  then 
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_AMT)
		'�Ϲ�
		elseif mstrGUBUN = "D"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_GEN)
			settingRowChange (.sprSht_GENDTL)
		elseif mstrGUBUN = "S"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			'settingRowChange (.sprSht_SUSU)
			settingRowChange (.sprSht_SUSUDTL)
		elseif mstrGUBUN = "O"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUT)
		end if
	End With
End Sub

sub settingRowChange(sprsht)
	Dim strSETTINGDATA
	Dim intCnt 
	Dim i ,j

	with frmThis
		intCnt = 0
		
		for j = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) <> "DTL" Then
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",j) = "1" Then
					intCnt = intCnt + 1
				End if 
			END IF
		next
		
		if right(sprsht.ID,3) <> "DTL" Then
			if intCnt = 0 Then
				gErrorMsgBox "üũ�� �����Ͱ� �����ϴ�. �����Ͻ� �����͸� üũ�ϼ���. ","�������"
				EXIT SUB
			End if
		End if
		
		strSETTINGDATA = ""
		strSETTINGDATA = .txtSUMM.value
		
		for i = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) = "DTL" Then
				mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
			ELSE 
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",i) = "1" Then
					
					mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
				End if
			End if 
		next 
	End with
end sub

'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub



'��ü�� �˾� ��ư
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,1))
					.txtREAL_MED_NAME.value = trim(vntData(1,1))
					
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'�̹�����ư Ŭ����
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtEXCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			selectrtn
     	end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEXCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(0,0))
					.txtEXCLIENTNAME.value = trim(vntData(1,0))
					selectrtn
				Else
					Call EXCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'����üũ
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'����üũ
Sub rdSALE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'����üũ
Sub rdBUY_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'MC���� -> ����üũ
Sub rdMCBUY_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUMM_onchange
	Dim blnByteCHk
	Dim intRtn
	blnByteCHk =  checkBytes(frmThis.txtSUMM.value)
	
	If blnByteCHk  > 23 Then
		intRtn = gYesNoMsgbox("������ ũ��� 23Byte �� ������ �����ϴ�. �ʱ�ȭ �Ͻðڽ��ϱ�?","ó���ȳ�!")
		
		IF intRtn <> vbYes then exit Sub
		
		frmThis.txtSUMM.value = ""
	End If
End Sub

function checkBytes(expression)
	Dim VLength
	Dim temp
	Dim EscTemp
	Dim i
	
	VLength=0
	temp = expression
	
	if temp <> "" then
		for i=1 to len(temp) 
			if mid(temp,i,1) <> escape(mid(temp,i,1))  then
				EscTemp=escape(mid(temp,i,1))
				if (len(EscTemp)>=6) then
					VLength = VLength +2
				else
				VLength = VLength +1
				end if
			else
				VLength = VLength +1
			end if
		Next
	end if

	checkBytes = VLength
end function

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
'-----------------------------------
' SpreadSheet ü����
'-----------------------------------
Sub sprSht_AMT_Change(ByVal Col, ByVal Row)
	with frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_AMT,"PREPAYMENT") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,True,"TODATE",Row,Row,false
			End If
		End if
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_AMT, Col, Row
End Sub

Sub sprSht_GEN_Change(ByVal Col, ByVal Row)
	with frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_GEN,"PREPAYMENT") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,True,"TODATE",Row,Row,false
			End If
		End if
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_GEN, Col, Row
End Sub

Sub sprSht_GENDTL_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"PAYCODE") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE", Row) = "P" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE", Row) = "N" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE", Row) = "Z" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"DEBTOR",Row, "404100"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE", Row) = "I" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"DEBTOR",Row, "404103"
			end if 
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_GENDTL,"VENDOR",Row)
			
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_GENDTL)
			
		end if 
		
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_GENDTL, Col, Row
End Sub

Sub sprSht_SUSU_Change(ByVal Col, ByVal Row)
	Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
	Dim strFROMDATE, strTODATE, strSUMMTEXT
	
	With frmThis
		if .sprSht_SUSUDTL.MaxRows > 0 then
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)
			strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMM",Row)
			strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SEMU",Row)
			strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DEMANDDAY",Row)
			strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DUEDATE",Row)
			strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DOCUMENTDATE",Row)
			strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row)
			strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"FROMDATE",Row)
			strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TODATE",Row)
			strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMMTEXT",Row)
			
			
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
					mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) = strTAXNO Then
					
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMM",intCnt, strSUMM
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SEMU",intCnt, strSEMU
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DUEDATE",intCnt, strDUEDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"FROMDATE",intCnt, strFROMDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"TODATE",intCnt, strTODATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
				End If
			Next
		end if
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, Col, Row
End Sub

Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,false
			End If
		End if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PAYCODE") Then 
		
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUT,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUT)
		end if 
		
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

'-----------------------------------
' SpreadSheet Ŭ��
'-----------------------------------
Sub sprSht_AMT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_AMT, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_AMT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_AMT, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_GEN_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			.sprSht_GENDTL.MaxRows = 0
			for intCnt = 1 To .sprSht_GEN.MaxRows
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_GEN, 1, 1, intCnt, intCnt, "", , , , , mstrCheck
			Next    

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_GEN.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_GEN, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_GEN_ButtonClicked (Col,Row,ButtonDown)
	if Col = 1 and Row > 0 then 
		if mobjSCGLSpr.GetTextBinding( frmThis.sprSht_GEN,"CHK",Row) = 1 THEN
			SelectRtn_GENDTL Col,Row
		ELSE
			call DeleteRtn_GENDTL(Row)
		END IF
	end if
End Sub

Sub DeleteRtn_GENDTL (Row)
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strSEQ	

	With frmThis
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_GENDTL.MaxRows to 1 step -1
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXNO",Row)

			if mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",i) = strTAXYEARMON and _
			   mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",i) = strTAXNO then
				
				mobjSCGLSpr.DeleteRow .sprSht_GENDTL,i
				
			end if				
		next
	End With
	err.clear	
End Sub

Sub sprSht_SUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			.sprSht_SUSUDTL.MaxRows = 0
			for intCnt = 1 To .sprSht_SUSU.MaxRows
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1, 1, intCnt, intCnt, "", , , , , mstrCheck
			Next    

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_SUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_SUSU_ButtonClicked (Col,Row,ButtonDown)
	if Col = 1 and Row > 0 then 
		if mobjSCGLSpr.GetTextBinding( frmThis.sprSht_SUSU,"CHK",Row) = 1 THEN
			SelectRtn_SUSUDTL Col,Row
		ELSE
			call DeleteRtn_SUSUDTL(Row)
		END IF
	end if
End Sub

Sub DeleteRtn_SUSUDTL (Row)
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strSEQ	

	With frmThis
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_SUSUDTL.MaxRows to 1 step -1
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)

			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",i) = strTAXYEARMON and _
			   mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",i) = strTAXNO then
				
				mobjSCGLSpr.DeleteRow .sprSht_SUSUDTL,i
				
			end if				
		next
	End With
	err.clear	
End Sub

Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

'-----------------------------------
' SpreadSheet ���� Ŭ��
'-----------------------------------
sub sprSht_AMT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_AMT, ""
		end if
	end with
end sub


sub sprSht_GEN_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_GEN, ""
		end if
	end with
end sub

sub sprSht_SUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_SUSU, ""
		end if
	end with
end sub

sub sprSht_OUT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUT, ""
		end if
	end with
end sub

'----------------------------------------------------------
'��Ʈ �ڵ� ��� [��Ʈ Ű��]
'----------------------------------------------------------
'��޾�
Sub sprSht_AMT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_AMT
	End With
End Sub

'�Ϲ�
Sub sprSht_GEN_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_GEN
	End With
End Sub

'�Ϲݻ�
Sub sprSht_GENDTL_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_GENDTL
	End With
End Sub

'������
Sub sprSht_SUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_SUSU
	End With
End Sub

'������ ��
Sub sprSht_SUSUDTL_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_SUSUDTL
	End With
End Sub

'�����
Sub sprSht_OUT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_OUT
	End With
End Sub


SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"VAT") Then
		
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"VAT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

'---------------------------------------------
'��Ʈ ���콺 ��
'---------------------------------------------
'��޾�
Sub sprSht_AMT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_AMT
	end with
End Sub

'�Ϲ�
Sub sprSht_GEN_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_GEN
	end with
End Sub

'�Ϲ� ��
Sub sprSht_GENDTL_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_GENDTL
	end with
End Sub

'������
Sub sprSht_SUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSU
	end with
End Sub

'������ ��
Sub sprSht_SUSUDTL_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSUDTL
	end with
End Sub

'����� 
Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_OUT
	end with
End Sub

'-----------------------------------
'��Ʈ���� ���콺�� �ݾ��ջ� �̺�Ʈ
'-----------------------------------
sub MouseUp_SumAmt(sprSht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i,j
	Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"AMT") or SprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"VAT") then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

sub chkSUMMAE_Onclick
	Dim i,j
	Dim intRANK
	with frmThis 
	
		intRANK = 1
		for i = 1 to  .sprSht_GENDTL.MaxRows
			strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",i)
			if strBUSINO <> "" then
				exit for
			end if
		next
		
		for j = 1 to  .sprSht_GENDTL.MaxRows
			if strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j) then
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"TRANSRANK",j, intRANK
			else
				strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j)
				intRANK = intRANK + 1
				mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"TRANSRANK",j, intRANK		
			end if
		next
	
	end with
end sub
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intGBN
	Dim strComboPREPAYMENT
	Dim strBMORDER
	
	'����������ü ����	
	Set mobjMDCMCATVVOCH = gCreateRemoteObject("cMDCT.ccMDCTCATVVOCH")
	Set mobjMDCOGET		 = gCreateRemoteObject("cMDCO.ccMDCOGET")
	Set mobjMDCOVOCH	 = gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
	
    with frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strBMORDER = " " & vbTab & "AD0120" & vbTab & "AD0110"
		'**************************************************
		'��޾� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_AMT
		mobjSCGLSpr.SpreadLayout .sprSht_AMT, 33, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_AMT,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | DEPT_NAME | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | DEBTOR | ACCOUNT | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN"
		mobjSCGLSpr.SetHeader .sprSht_AMT,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�μ�|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|��ü���|����|��������|����|BMORDER|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG | AMTGBN"
		mobjSCGLSpr.SetColWidth .sprSht_AMT, "-1",  "   4|       8|        10|    15|  17|       5|         6|  13|  10|    10|       6| 5|       8|       8|        10|      15|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|       0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_AMT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_AMT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_AMT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_AMT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_AMT, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | DEPT_NAME | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | DEBTOR | ACCOUNT | BMORDER | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_AMT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_AMT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_AMT, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_AMT,mobjSCGLSpr.CnvtDataField(.sprSht_AMT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_AMT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_AMT,mobjSCGLSpr.CnvtDataField(.sprSht_AMT,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_AMT,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_AMT, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR | ACCOUNT | DEBTOR",-1,-1,2,2,false '���
		'mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,true,"CUSTOMERCODE | CUSTNAME | REAL_MED_NAME | DEPT_NAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | DEBTOR | ACCOUNT | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_AMT, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_AMT,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		
		'**************************************************
		'�Ϲ� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_GEN
		mobjSCGLSpr.SpreadLayout .sprSht_GEN, 28, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_GEN,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | DOCUMENTDATE | PAYCODE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN"
		mobjSCGLSpr.SetHeader .sprSht_GEN,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������������|���VENDOR|��ü���|����|������|���޹��|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG | AMTGBN"
		mobjSCGLSpr.SetColWidth .sprSht_GEN, "-1",  "   4|       8|        10|    15|  17|  10|    10|       6| 5|       8|           8|        10|      15|   0|     8|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|       0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_GEN, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_GEN, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_GEN, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_GEN, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GEN, "CUSTOMERCODE | CUSTNAME | AMT | VAT | SEMU | BP | VENDOR | REAL_MED_NAME | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GEN, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GEN, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_GEN, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GEN, "PAYCODE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_GEN,mobjSCGLSpr.CnvtDataField(.sprSht_GEN,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_GEN,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_GEN, "SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR",-1,-1,2,2,false '���
		'mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,true,"CUSTOMERCODE | CUSTNAME | REAL_MED_NAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_GEN, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_GEN,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		gSetSheetColor mobjSCGLSpr, .sprSht_GENDTL
		mobjSCGLSpr.SpreadLayout .sprSht_GENDTL, 34, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_GENDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_GENDTL,		    "��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|��ü���|����|��������|����|BMORDER|������|���޹��|BANKTYPE |�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG | AMTGBN | TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_GENDTL, "-1",  "        8|        10|    15|  17|       5|         8|  10|    10|       6| 5|       8|       8|        10|      15|   0|       7|   7|      7|     8|      20|       10|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|       0|       0|         5"
		mobjSCGLSpr.SetRowHeight .sprSht_GENDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_GENDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_GENDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | SEMU | BP | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | BMORDER | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_GENDTL, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_GENDTL, "PAYCODE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_GENDTL,mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_GENDTL,mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"BMORDER"),-1,-1,strBMORDER,,80
		
		mobjSCGLSpr.SetCellAlign2 .sprSht_GENDTL, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR",-1,-1,2,2,false '���
		'mobjSCGLSpr.SetCellsLock2 .sprSht_GENDTL,true,"CUSTOMERCODE | CUSTNAME | REAL_MED_NAME | SUMM | AMT | BP | VENDOR | GBN |  TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG| TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_GENDTL, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_GENDTL,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		'**************************************************
		'������ ��Ʈ ������ hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 26, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|����|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG|AMTGBN"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    15|  17|  10|    10|       6| 5|       8|       8|        10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      0|     0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "CUSTOMERCODE | CUSTNAME | SEMU | BP | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR",-1,-1,2,2,false '���
		'mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_SUSU,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		'**************************************************
		'������ ��Ʈ ������ dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_SUSUDTL, 31, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSUDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TAXNOSEQ"
		mobjSCGLSpr.SetHeader .sprSht_SUSUDTL,		    "��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|����|��������|����|BMORDER|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG |AMTGBN|����"
		mobjSCGLSpr.SetColWidth .sprSht_SUSUDTL, "-1",  "       8|        10|    15|  17|       5|         8|  10|    10|       6| 5|       8|       8|        10|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|       0|     0|  10"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | SEMU | BP | VENDOR | GBN | DEBTOR | ACCOUNT | BMORDER | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TAXNOSEQ", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSUDTL, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),-1,-1,strBMORDER,,80
		
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR | ACCOUNT | DEBTOR",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true,"CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | BP | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_SUSUDTL,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		'**************************************************
		'����� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 35, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER |DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "����|��ǥ����|������|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|����|��������|����|BMORDER|������|���޹��|BANKTYPE|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG|AMTGBN|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1",  "   4|       8|    15|        10|    15|  17|       5|         8|  10|    10|       6| 5|       8|       8|        10|   0|       7|   7|      7|     8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      0|     0|        10"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | SEMU | BP | VENDOR | GBN | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE|BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR | ACCOUNT | DEBTOR",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG|SUMMTEXT",-1,-1,0,2,false '����
		'mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | AMT | SEMU | BP | VENDOR | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN  | GFLAG | MEDFLAG | DUEDATE | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht_OUT,"VOCHNO | ERRCODE | ERRMSG"		
	End with

	pnlTab_amt.style.visibility = "visible" 
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtEDYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht_AMT.MaxRows = 0
		.sprSht_GEN.MaxRows = 0
		.sprSht_GENDTL.MaxRows = 0
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		
		.txtYEARMON.focus()
		
		Get_COMBO_VALUE	
		'ó���� ���� ���� ����
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		
		frmThis.txtEDYEARMON.style.visibility = "hidden"
	End with
End Sub

Sub EndPage()
	set mobjMDCMCATVVOCH = Nothing
	Set mobjMDCOGET = Nothing
	Set mobjMDCOVOCH = Nothing
	gEndPage	
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht_GEN.MaxRows = 0
		.sprSht_GENDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCMCATVVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_GEN, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_GENDTL, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUT, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	


'-----------------------------------------------------------------------------------------
' �׸��� ���� �޺� ����
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strCODE, row, sprsht)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCODE = replace(strCODE,"-","")
		
       	vntData = mobjMDCMCATVVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
		If not gDoErrorRtn ("Get_SUBCOMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 sprsht, "BANKTYPE",Row,Row,vntData,,160 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub SelectRtn (strVOCH_TYPE)
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN

	with frmThis
		.sprSht_AMT.MaxRows = 0
		.sprSht_GEN.MaxRows = 0
		.sprSht_GENDTL.MaxRows = 0
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		
		IF strVOCH_TYPE = "T" THEN
			CALL SelectRtn_AMT()
		ELSEIF strVOCH_TYPE = "D" THEN
			CALL SelectRtn_GEN()
		ELSEIF strVOCH_TYPE = "S" THEN
			CALL SelectRtn_SUSU()
		ELSEIF strVOCH_TYPE = "O" THEN
			CALL SelectRtn_OUT()
		END IF
		mstrSTAY = TRUE
   	end with
End Sub

Sub SelectRtn_AMT ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
	
	with frmThis
		.sprSht_AMT.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strREAL_MED_CODE = .txtREAL_MED_CODE.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
			
		vntData = mobjMDCMCATVVOCH.SelectRtn_AMT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
										  		 strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, _
										  		 strGBN)

		if not gDoErrorRtn ("SelectRtn_AMT") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_AMT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht_AMT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_AMT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_AMT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"DUEDATE",intCnt,intCnt,false
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_AMT,True,"TODATE",intCnt,intCnt,false
					End If
				Next
				AMT_SUM .sprSht_AMT
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_GEN ()
   	Dim vntData
    Dim intCnt, intCnt2
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
    Dim strSALEGBN
	
	with frmThis
		.sprSht_GEN.MaxRows = 0
		.sprSht_GENDTL.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strREAL_MED_CODE = .txtREAL_MED_CODE.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		if .rdSALE.checked then
			strSALEGBN = "Y"
		elseIF .rdBUY.checked THEN
			strSALEGBN = "N"
		Else
			strSALEGBN = "C"
		end if
		
		vntData = mobjMDCMCATVVOCH.SelectRtn_GEN(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
												 strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, _
												 strGBN, strSALEGBN)

		if not gDoErrorRtn ("SelectRtn_GEN") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_GEN, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				if .rdSALE.checked then
					mobjSCGLSpr.ColHidden .sprSht_GEN, "PAYCODE", true 
					mobjSCGLSpr.ColHidden .sprSht_GENDTL, "PAYCODE", true 
					mobjSCGLSpr.ColHidden .sprSht_GENDTL, "BANKTYPE", true 
				else
					mobjSCGLSpr.ColHidden .sprSht_GEN, "PAYCODE", false
					mobjSCGLSpr.ColHidden .sprSht_GENDTL, "PAYCODE", false
					mobjSCGLSpr.ColHidden .sprSht_GENDTL, "BANKTYPE", false  
				end if
				
				For intCnt = 1 To .sprSht_GEN.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_GEN, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_GEN, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_GEN,false,"DUEDATE",intCnt,intCnt,false
					End If
				Next
				
				if .rdSALE.checked then
					For intCnt2 = 1 To .sprSht_GEN.MaxRows
						mobjSCGLSpr.SetTextBinding .sprSht_GEN,"VENDOR",intCnt2, "1048636968"
					Next
				end if
				
				AMT_SUM .sprSht_GEN
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_GENDTL (Col, Row)
	Dim vntData
   	Dim i,j, strCols
    Dim intCnt
    Dim strTAXYEARMON
    Dim strTAXNO
    Dim strRow
    Dim strBUSINO
    Dim intRANK
    Dim strPAYCODE
    
	with frmThis
		'Sheet�ʱ�ȭ
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
	
		strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXYEARMON",Row)
		strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXNO",Row)
		strPAYCODE = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"PAYCODE",Row)
				
		if .rdSALE.checked then
			vntData = mobjMDCMCATVVOCH.SelectRtn_GENDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		elseif .rdBUY.checked then
			vntData = mobjMDCMCATVVOCH.SelectRtn_GENBUYDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		else
			vntData = mobjMDCMCATVVOCH.SelectRtn_GENMCBUYDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		end if
																							
		If not gDoErrorRtn ("SelectRtn_GENDTL") Then
			If mlngRowCnt >0 Then
				strRow = 0
				strRow = .sprSht_GENDTL.MaxRows + 1
				Call mobjSCGLSpr.SetClipBinding (.sprSht_GENDTL,vntData, 1, strRow, mlngColCnt, mlngRowCnt,True)
				
				For intCnt = 1 To .sprSht_GENDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) = strTAXNO Then
						if .rdBUY.checked = true and .rdF.checked = true then
							mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"PAYCODE",intCnt, strPAYCODE
							sprSht_GENDTL_Change mobjSCGLSpr.CnvtDataField(.sprSht_GENDTL,"PAYCODE"),intCnt
						end if
					End If
				Next
				
				'������ ��� �÷��� ����
				if .rdSALE.checked then
					For intCnt = 1 To .sprSht_GENDTL.MaxRows
						mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"VENDOR",intCnt, "1048636968"
					Next
				end if
				
				'����� ���� ����.
				if .chkSUMMAE.checked then
					intRANK = 1
					for i = 1 to  .sprSht_GENDTL.MaxRows
						strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",i)
						if strBUSINO <> "" then
							exit for
						end if
					next
					
					for j = 1 to  .sprSht_GENDTL.MaxRows
						if strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j) then
							mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"TRANSRANK",j, intRANK
						else
							strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j)
							intRANK = intRANK + 1
							mobjSCGLSpr.SetTextBinding .sprSht_GENDTL,"TRANSRANK",j, intRANK		
						end if
					next
				end if 
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_SUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
	
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strREAL_MED_CODE = .txtREAL_MED_CODE.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMCATVVOCH.SelectRtn_SUSU(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
												  strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, _
												  strGBN)

		if not gDoErrorRtn ("SelectRtn_SUSU") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_SUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht_SUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"DUEDATE",intCnt,intCnt,false
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"TODATE",intCnt,intCnt,false
					End If
				Next
				AMT_SUM .sprSht_SUSU
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_SUSUDTL (Col, Row)
	Dim vntData
   	Dim i, strCols
    Dim intCnt
    Dim strTAXYEARMON
    Dim strTAXNO
    Dim strRow
    
	with frmThis
		'Sheet�ʱ�ȭ
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
		Dim strFROMDATE, strTODATE, strSUMMTEXT
		
		strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
		strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)
		strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMM",Row)
		strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SEMU",Row)
		strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DEMANDDAY",Row)
		strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DUEDATE",Row)
		strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DOCUMENTDATE",Row)
		strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row)
		strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"FROMDATE",Row)
		strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TODATE",Row)
		strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMMTEXT",Row)
				
		vntData = mobjMDCMCATVVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
																							
		If not gDoErrorRtn ("SelectRtn_SUSUDTL") Then
			If mlngRowCnt >0 Then
				strRow = 0
				strRow = .sprSht_SUSUDTL.MaxRows + 1
				Call mobjSCGLSpr.SetClipBinding (.sprSht_SUSUDTL,vntData, 1, strRow, mlngColCnt, mlngRowCnt,True)
				
				For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) = strTAXNO Then
						
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMM",intCnt, strSUMM
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SEMU",intCnt, strSEMU
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DUEDATE",intCnt, strDUEDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"FROMDATE",intCnt, strFROMDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"TODATE",intCnt, strTODATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
					End If
				Next
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_OUT ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, strGBN
    Dim strEDYEARMON
	
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strEDYEARMON = .txtEDYEARMON.value
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strREAL_MED_CODE = .txtREAL_MED_CODE.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value
		strEXCLIENTCODE = .txtEXCLIENTCODE.value
		strEXCLIENTNAME = .txtEXCLIENTNAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMCATVVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, strEDYEARMON, _
												strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME, _
												strGBN, strEXCLIENTCODE, strEXCLIENTNAME)

		if not gDoErrorRtn ("SelectRtn_OUT") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DUEDATE",intCnt,intCnt,false
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,false
					End If
				Next
				
				AMT_SUM .sprSht_OUT
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

Function DataValidation_AMT ()
	DataValidation_AMT = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_AMT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_AMT = True
End Function

Function DataValidation_GEN ()
	DataValidation_GEN = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_GENDTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_GEN = True
End Function

Function DataValidation_SUSU ()
	DataValidation_SUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_SUSUDTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_SUSU = True
End Function

Function DataValidation_OUT ()
	DataValidation_OUT = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_OUT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUT = True
End Function

'�������
Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn
	
	with frmThis
		IF mstrPROCESS = "Create" THEN
			IF NOT .rdF.checked THEN
				gErrorMsgBox "�̿Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if 
		IF mstrPROCESS = "Delete" THEN
			IF NOT .rdT.checked THEN
				gErrorMsgBox "�Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if 
		
		IF mstrSTAY THEN 
			mstrSTAY = FALSE
			IF strVOCH_TYPE = "T" THEN
				if DataValidation_AMT =false then exit sub
				CALL ProcessRtn_AMT()
			ELSEIF strVOCH_TYPE = "D" THEN
				if DataValidation_GEN =false then exit sub
				CALL ProcessRtn_GEN()
			ELSEIF strVOCH_TYPE = "S" THEN
				if DataValidation_SUSU =false then exit sub
				CALL ProcessRtn_SUSU()
			ELSEIF strVOCH_TYPE = "O" THEN
				if DataValidation_OUT =false then exit sub
				CALL ProcessRtn_OUT()
			END IF
		ELSE
			gErrorMsgBox "��ǥó�� �������Դϴ�.","��ǥó�� �ȳ�"
		END IF
   	end with
End Sub

'�������
Sub ProcessRtn_AMT()
	Dim intRtn
	
	'ä���� �����ϱ� ���� ���� 
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_AMT,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | DEPT_NAME | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | DEBTOR | ACCOUNT | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0001"
		
		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_AMT.MaxRows
				if mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"chk",intCnt) = "1" then		
				
					'ä���� �����Ѵ�.
					'--------------------------------------------------------------------------------------

					strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

					strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"POSTINGDATE",intCnt),"-","")
					strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"MEDFLAG",intCnt)
					strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",intCnt)
					strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",intCnt)'
					strTYPE				= "7"

					if strGROUPSEQ = true then
						strGROUP = TRUE
					else 
						strGROUP = FALSE
					END IF 

					If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
						gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
						Exit Sub
					END IF 

					strGROUPSEQ = FALSE
					
					'���� ������ RMS ä�� ��������
					vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
					
					strVOCHNORMS =  vntData(0,1)
					'---------------------------------------------------------------------------------------
					
					strIF_CNT = strIF_CNT + 1
					strRMS_DOC_TYPE = "T"

					if strIF_CNT = "1" then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BMORDER",intCnt)
					else
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BMORDER",intCnt)
					end if

					strHSEQ = strHSEQ+1
				end if 
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_AMT.MaxRows
				if mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"chk",intCnt) = "1" then		
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "Z"
		
					if strIF_CNT = "1" then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BMORDER",intCnt)
					else
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"BMORDER",intCnt)
					end if
					
					strHSEQ = strHSEQ+1
				end if 
			Next
		
		end if 

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)

   	end with
End Sub

'�������
Sub ProcessRtn_GEN()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim intColFlag, bsdiv, intMaxCnt
	
	'��ǥ ä���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht_GENDTL, meINS_FLAG
		
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_GENDTL,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		if  .chkSUMMAE.checked = true and .rdSALE.checked= true then 
			gErrorMsgBox "�ջ��� ������ȸ�� �����մϴ�. ������ �����Ͻ÷��� �ջ꿡 üũ�� �������ּ���","����ȳ�"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF .rdSALE.checked THEN
			IF_GUBUN = "RMS_0002"
		ELSE
			IF_GUBUN = "RMS_0012"
		END IF
		
		'�ִ밪
		if .chkSUMMAE.checked = true then
			intColFlag = 0
			For intMaxCnt = 1 To .sprSht_GENDTL.MaxRows
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			Next
		end if		
		
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		if .chkSUMMAE.checked = true then
			if mstrPROCESS = "Create" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M" 
	                
					For i = 1 To .sprSht_GENDTL.MaxRows
						'û���հ�
						If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",i)) = intCnt Then
							lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",i))
							lngSUMAMT = lngSUMAMT + lngAMT
							lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",i))
							lngSUMVAT = lngSUMVAT + lngVAT
						End If
					Next

					For i = 1 To .sprSht_GENDTL.MaxRows
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",i)) = intCnt Then
								'û���հ�,�ΰ����հ�,û������ ����� ������ ����
							If intCnt2 = intCnt Then
							Else
								'ä���� �����Ѵ�.(�ջ���ǥ�� ä�� ����)
								'--------------------------------------------------------------------------------------
								strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

								strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",i),"-","")
								strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"MEDFLAG",i)
								strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",i)
								strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",i)
							
								'����
								IF .rdSALE.checked THEN
									strTYPE = "1"
								'����
								ELSE
									strTYPE = "4"
								END IF

								if strGROUPSEQ = true then
									strGROUP = TRUE
								else 
									strGROUP = FALSE
								END IF 

								If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
									gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
									Exit Sub
								END IF 

								strGROUPSEQ = FALSE
								
								'���� ������ RMS ä�� ��������
								vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
								
								strVOCHNORMS =  vntData(0,1)
								'---------------------------------------------------------------------------------------
									
								strIF_CNT = strIF_CNT + 1

								strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",i)
								strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",i)
								strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",i)
								strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",i)
								strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",i)
								strAMT			= lngSUMAMT
								strVAT			= lngSUMVAT
								strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",i)
								strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",i)
								strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",i)
								strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",i)
								strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",i)
								strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",i)
								strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",i)
								strRMS_DOC_TYPE = "M"
								strACCOUNT		= ""
								strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",i)
								strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",i)
								strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",i)
								strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",i)
								strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",i)
								strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",i)
								strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",i)
								strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",i)
								strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",i)
								strVOCHNO		= strVOCHNORMS
								strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",i)
								strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",i)
								
								if strIF_CNT = "1" then
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strVENDOR + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
								else
									strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strVENDOR + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
								end if
											
								For j = 1 To .sprSht_GENDTL.MaxRows

									If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",j)) = intCnt Then	
										strIF_CNT = strIF_CNT + 1
										
										strISEQ = strISEQ+1
										
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",j) + "|" + _
													cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",j)) + "|" + _
													cstr(0) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",j) + "|" + _ 
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",j) + "|" + _ 
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",j) + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",j) + "|" + _
													"" + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",j) + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",j),"-","") + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",j) + "|" + _  
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",j),"-","") + "|" + _
													strVOCHNORMS + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",j) + "|" + _  
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",j)
													
									end if
								Next
								strHSEQ = strHSEQ + 1
								strISEQ = 1
								intCnt2 = intCnt
								'���Ѿ�����Ʈ
							End If
						End If
					Next
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
	                
					For i = 1 To .sprSht_GENDTL.MaxRows
						'û���հ�
						If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",i)) = intCnt Then
							lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",i))
							lngSUMAMT = lngSUMAMT + lngAMT
							lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",i))
							lngSUMVAT = lngSUMVAT + lngVAT
						End If
					Next

					For i = 1 To .sprSht_GENDTL.MaxRows
						
						If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",i)) = intCnt Then
							'û���հ�,�ΰ����հ�,û������ ����� ������ ����
							If intCnt2 = intCnt Then
							Else
								strIF_CNT = strIF_CNT + 1
								
								strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",i)
								strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",i)
								strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",i)
								strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",i)
								strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",i)
								strAMT			= lngSUMAMT
								strVAT			= lngSUMVAT
								strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",i)
								strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",i)
								strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",i)
								strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",i)
								strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",i)
								strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",i)
								strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",i)
								strRMS_DOC_TYPE = "Z"
								strACCOUNT		= ""
								strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",i)
								strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",i)
								strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",i)
								strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",i)
								strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",i)
								strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",i)
								strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",i)
								strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",i)
								strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",i)
								strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VOCHNO",i)
								strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",i)
								strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",i)
								
								strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
											cstr(strISEQ) + "|" + _
											replace(strPOSTINGDATE,"-","") + "|" + _
											strCUSTOMERCODE + "|" + _
											strSUMM + "|" + _
											strBA + "|" + _
											strCOSTCENTER + "|" + _
											cstr(strAMT) + "|" + _
											cstr(strVAT) + "|" + _
											strSEMU + "|" + _ 
											strBP + "|" + _ 
											replace(strDEMANDDAY,"-","") + "|" + _
											strVENDOR + "|" + _
											strTAXYEARMON + "|" + _
											strTAXNO + "|" + _
											strGFLAG + "|" + _
											strRMS_DOC_TYPE + "|" + _ 
											strACCOUNT + "|" + _
											strDEBTOR + "|" + _
											replace(strDOCUMENTDATE,"-","") + "|" + _
											strPREPAYMENT + "|" + _
											replace(strFROMDATE,"-","") + "|" + _
											replace(strTODATE,"-","") + "|" + _
											strSUMMTEXT + "|" + _
											strAMTGBN + "|" + _
											strPAYCODE + "|" + _  
											replace(strDUEDATE,"-","") + "|" + _
											strVOCHNO + "|" + _
											strBANKTYPE + "|" + _
											strBMORDER
											
								For j = 1 To .sprSht_GENDTL.MaxRows
									If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TRANSRANK",j)) = intCnt Then	
										strIF_CNT = strIF_CNT + 1
										
										strISEQ = strISEQ+1
										
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",j) + "|" + _
													cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",j)) + "|" + _
													cstr(0) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",j) + "|" + _ 
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",j) + "|" + _ 
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",j) + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",j) + "|" + _
													"" + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",j) + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",j),"-","") + "|" + _
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",j) + "|" + _  
													replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",j),"-","") + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VOCHNO",j) + "|" + _
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",j) + "|" + _  
													mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",j)
													
									end if
								Next
								strHSEQ = strHSEQ + 1
								strISEQ = 1
							End If
							'���Ѿ�����Ʈ.
							intCnt2 = intCnt
						End If
						
					Next
				Next
			end if 
			
		else
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_GENDTL.MaxRows
					strIF_CNT = strIF_CNT + 1
			
					IF .rdSALE.checked THEN
						strRMS_DOC_TYPE = "M"
					ELSEIF .rdMCBUY.checked THEN
						strRMS_DOC_TYPE = "M"
					ELSE
						strRMS_DOC_TYPE = "O"
					END IF
					
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------
							
						'DTL ��Ʈ�� ���� �ο� ������ �ϳ��� ��ǥ��ȣ�� ä���ȴ�.
						If strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) and _
								strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) Then
						ELSE

							strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

							strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",intCnt),"-","")
							strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"MEDFLAG",intCnt)
							strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt)
							strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt)'
							
							'����
							IF .rdSALE.checked THEN
								strTYPE = "1"
							'����
							ELSE
								strTYPE = "4"
							END IF

							if strGROUPSEQ = true then
								strGROUP = TRUE
							else 
								strGROUP = FALSE
							END IF 

							If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
								gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
								Exit Sub
							END IF 

							strGROUPSEQ = FALSE
							
							'���� ������ RMS ä�� ��������
							vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
							
							strVOCHNORMS =  vntData(0,1)

										
						END IF
						'---------------------------------------------------------------------------------------
					If strIF_CNT = "1" Then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",intCnt) + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",intCnt)
				else
						
						if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) and _
							strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) THEN
							
							strHSEQ = strHSEQ
							strISEQ = strISEQ+1
						else 
							strHSEQ = strHSEQ + 1
							strISEQ = 1
						end if
					
					
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",intCnt) + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",intCnt)
					end if
					
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt)

					
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_GENDTL.MaxRows
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "Z"
		
					if strIF_CNT = "1" then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VOCHNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",intCnt) + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",intCnt)
					else
						if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) and _
							strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) THEN
							
							strHSEQ = strHSEQ
							strISEQ = strISEQ+1
						else 
							strHSEQ = strHSEQ + 1
							strISEQ = 1
						end if
						
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"VOCHNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BANKTYPE",intCnt) + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"BMORDER",intCnt)
					end if
					
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXYEARMON",intCnt)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GENDTL,"TAXNO",intCnt)
				Next
			
			end if 
		end if

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)

   	end with
End Sub


Sub ProcessRtn_SUSU()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	
	'ä���� �����ϱ� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht_SUSUDTL, meINS_FLAG
		
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_SUSUDTL,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TAXNOSEQ")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		strTAXYEARMON = "" : strTAXNO = ""
		
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0002"
		
		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1
		
				strRMS_DOC_TYPE = "D"
				
				'ä���� �����Ѵ�.
				'--------------------------------------------------------------------------------------
					
				'DTL ��Ʈ�� ���� �ο� ������ �ϳ��� ��ǥ��ȣ�� ä���ȴ�.
				If strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) Then
				ELSE

					strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

					strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","")
					strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"MEDFLAG",intCnt)
					strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
					strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)'
					strTYPE				= "1"

					if strGROUPSEQ = true then
						strGROUP = TRUE
					else 
						strGROUP = FALSE
					END IF 

					If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
						gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
						Exit Sub
					END IF 

					strGROUPSEQ = FALSE
					
					'���� ������ RMS ä�� ��������
					vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
					
					strVOCHNORMS =  vntData(0,1)
				END IF
				'---------------------------------------------------------------------------------------
				

				if strIF_CNT = "1" then

					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
								cstr(strISEQ) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
								strRMS_DOC_TYPE + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
								"" + "|" + _  
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
								strVOCHNORMS + "|" + _
								"" + "|" + _  
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				else
					
					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if
				
				
					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
								cstr(strISEQ) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
								strRMS_DOC_TYPE + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
								"" + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
								strVOCHNORMS + "|" + _
								"" + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				end if
				
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)

				
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1
		
				strRMS_DOC_TYPE = "Z"
	
				if strIF_CNT = "1" then

					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
								cstr(strISEQ) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
								strRMS_DOC_TYPE + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
								"" + "|" + _  
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt) + "|" + _   
								"" + "|" + _  
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				else
					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if
					
					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
								cstr(strISEQ) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
								cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
								strRMS_DOC_TYPE + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
								"" + "|" + _ 
								replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt) + "|" + _
								"" + "|" + _ 
								mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				end if
				
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)
			Next
		
		end if 

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)

   	end with
End Sub

'����� ���� ����
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim strCUSTOMERCODE
	Dim intColFlag, bsdiv, intMaxCnt
	
	'��ǥ ä���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CLIENTNAME | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0012"
		
		
		'�ִ밪
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		IF .rdDIV.checked THEN
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
					
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------

						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)'
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'���� ������ RMS ä�� ��������
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)
						'---------------------------------------------------------------------------------------
						
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "O"

						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						else
							
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
			
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						end if
					end if 
				Next
			end if 
		ELSE
			if mstrPROCESS = "Create" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M" 
	                
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								'û���հ�,�ΰ����հ�,û������ ����� ������ ����
								If intCnt2 = intCnt Then
								Else
									'ä���� �����Ѵ�.(�ջ���ǥ�� ä�� ����)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'���� ������ RMS ä�� ��������
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1

									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE = "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
									
									if strIF_CNT = "1" then
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
													
									else
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									
									end if
												
												
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",j)  
			
											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								End If
								'���Ѿ�����Ʈ.
							End If
						End If
					Next
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
	                
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								'û���հ�,�ΰ����հ�,û������ ����� ������ ����
								If intCnt2 = intCnt Then
								Else
									strIF_CNT = strIF_CNT + 1
									
									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE = "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
									
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
												
												
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",j)

											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								End If
								'���Ѿ�����Ʈ.
								intCnt2 = intCnt
							End If
						End If
					Next
				Next
			end if 
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'****************************************************************************************
' ä�� ����ó��
'****************************************************************************************
Function InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
	InsertRtn_VOCHNO = false
   	Dim strVOCHNO
	With frmThis
		
		'ä���� ����& �����Ѵ� (������ �ߺ��� ���� SAP �ʿ��� ������ �� ��쿡�� ���� ��ȣ�� �����Ǵ� ���� ���´�.).
		intRtn = mobjMDCOVOCH.InsertRtn_VOCHNO(gstrConfigXml,strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
		If not gDoErrorRtn ("InsertRtn_VOCHNO") Then
		
			If intRtn = 0 Then
				Exit Function
			End If		
   		End If
   	end With
   	InsertRtn_VOCHNO = true
End Function

'---------------------------------------------------
' ��ǥ���� �� ��ǥ��ȣ �޾ƿ��� �� ���� RMS������Ʈ
'---------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
	
		if mstrPROCESS ="Create" then
			IF mstrGUBUN = "T" THEN
				intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A2")
			ELSEIF mstrGUBUN = "D" THEN
				if .rdSALE.checked then
					intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A2")
				ELSEif .rdMCBUY.checked then
					intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A2AA")
				else
					intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A2D")
				end if
			ELSEIF mstrGUBUN = "S" THEN
				intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A2")
			ELSEIF mstrGUBUN = "O" THEN
				intRtn = mobjMDCMCATVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "A22")
			END IF

			
			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "T" THEN
					mobjSCGLSpr.SetFlag  .sprSht_AMT,meCLS_FLAG
				ELSEIF mstrGUBUN = "D" THEN
					mobjSCGLSpr.SetFlag  .sprSht_GEN,meCLS_FLAG
				ELSEIF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "�����Դϴ�..","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   			
   		elseif mstrPROCESS ="Delete" then
   			IF mstrGUBUN = "D" THEN
				if .rdSALE.checked then
					intRtn = mobjMDCMCATVVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CATV" )
				ELSEif .rdMCBUY.checked then
					intRtn = mobjMDCMCATVVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CATV" )
				else
					intRtn = mobjMDCMCATVVOCH.VOCHDELL_BUY(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CATV", "A2D")
				end if
			elseIF mstrGUBUN = "O" THEN
				intRtn = mobjMDCMCATVVOCH.VOCHDELL_BUY(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CATV", "A22")
			ELSE
				intRtn = mobjMDCMCATVVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CATV" )
			END IF
   			
   			if not gDoErrorRtn ("VOCHDELL") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "T" THEN
					mobjSCGLSpr.SetFlag  .sprSht_AMT,meCLS_FLAG
				ELSEIF mstrGUBUN = "D" THEN
					mobjSCGLSpr.SetFlag  .sprSht_GEN,meCLS_FLAG
				ELSEIF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF
				
				gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				
				SelectRtn(mstrGUBUN)
   			end if
   		end if 
   		IF mstrGUBUN = "T" THEN
			.sprSht_AMT.focus()
		ELSEIF mstrGUBUN = "D" THEN
			.sprSht_GEN.focus()
		ELSEIF mstrGUBUN = "S" THEN
			.sprSht_SUSU.focus()
		ELSEIF mstrGUBUN = "O" THEN
			.sprSht_OUT.focus()
		END IF
	End With
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   	
		IF NOT .rdE.checked THEN
			gErrorMsgBox "������ȸ�� �����մϴ�.","�����׻���"
			exit sub
		end if 
		
		IF mstrGUBUN = "T" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_AMT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "D" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_GEN,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "S" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_SUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "O" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		intRtn = mobjMDCMCATVVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'��� �÷��� Ŭ����
			IF mstrGUBUN = "T" THEN
				mobjSCGLSpr.SetFlag  .sprSht_AMT,meCLS_FLAG
			ELSEIF mstrGUBUN = "D" THEN
				mobjSCGLSpr.SetFlag  .sprSht_GEN,meCLS_FLAG
			ELSEIF mstrGUBUN = "S" THEN
				mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "O" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
			END IF
			
			if intRtn > 0 Then
				gErrorMsgBox "���� ��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
			End If
			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

'-----------------------------------------
'��ǥ ���� ����
'-----------------------------------------
Sub DeleteRtn (strGUBUN)
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strVOCHNO
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
	
		If mstrGUBUN = "D"  then  
			If .sprSht_GEN.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_GEN.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "T"  then
			If .sprSht_AMT.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_AMT.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "S"  then
			If .sprSht_SUSU.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_SUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "O"  then
			If .sprSht_OUT.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUT.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		END IF
	
		intRtn = gYesNoMsgbox("���������� SAP���� ���ε� ��ǥ�� SAP���� ����Ͽ� RMS�ʿ��� ������ �� ������ RMS�� ��ǥ�� ������ �����Ҷ� ����մϴ�. " & vbCrlf & "  " & vbCrlf & " ��ǥ�� ������ �����Ͻðڽ��ϱ�?","�������� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		If mstrGUBUN = "T"  then  
			for i = .sprSht_AMT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_AMT,"VOCHNO",i)
					
					intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_AMT,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "S"  then
			for i = .sprSht_SUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"VOCHNO",i)
					
					intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "D"  then
			for i = .sprSht_GEN.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_GEN,"VOCHNO",i)
					
					if .rdSALE.checked then
						intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV" )
					ELSEIF .rdMCBUY.checked then
						intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV" )
					else
						intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV", "A2D")
					end if
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_GEN,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "O"  then
			for i = .sprSht_OUT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
					
					intRtn = mobjMDCMCATVVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CATV", "A22")
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		END IF
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
			SelectRtn (strGUBUN)
	End With
	err.clear	
End Sub


		</script>
		<script language="javascript">
		//##########################################################################################################################################
		//******************************************��1) frmSapCon ���� ������ �� �̿��Ͽ� Submit �ϴ� �Լ�
		//##########################################################################################################################################

		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {
			//���
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			//dtl 
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			
			window.frames[0].document.forms[0].submit();
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
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���̺� ��ǥ����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
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
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%" height="93%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')"
												width="60">&nbsp;���
											</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtEDYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" size="9" name="txtEDYEARMON"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="60">&nbsp;������
											</TD>
											<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 142px; HEIGHT: 22px"
													maxLength="100" size="16" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME,txtREAL_MED_CODE)"
												width="60">&nbsp;��ü��
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 142px; HEIGHT: 22px"
													maxLength="100" size="16" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" maxLength="6" name="txtREAL_MED_CODE">
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
										<TR>
											<TD class="SEARCHLABEL">����
											</TD>
											<TD class="SEARCHDATA" colSpan="4"><INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;�Ϸ�&nbsp;
												<INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" value="rdF" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')"
													CHECKED>&nbsp;�̿Ϸ�&nbsp; <INPUT id="rdE" title="������ǥ ������ȸ" type="radio" value="rdE" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;����&nbsp;
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="60">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="�ڵ��" style="WIDTH: 142px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="18" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">
												<INPUT class="INPUT" id="txtEXCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtEXCLIENTCODE"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<td class="DATA">
									�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
									<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
													type="button" value="��޾�" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
													type="button" size="20" value="�Ϲ�" name="btnTab2"> <INPUT class="BTNTAB" id="btnTab3" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
													type="button" size="20" value="������" name="btnTab3"> <INPUT class="BTNTAB" id="btnTab4" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
													type="button" size="20" value="�����" name="btnTab4">
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
													<TR>
														<td><IMG id="ImgvochCre" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/ImgvochCre.gIF" border="0" name="ImgvochCre"></td>
														<td><IMG id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/imgVochDel.gIF" border="0" name="imgVochDel"></td>
														<td><IMG id="ImgErrVochDel" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																height="20" alt="������ǥ �� �����մϴ�." src="../../../images/ImgErrVochDel.gIF" border="0"
																name="ImgErrVochDel"></td>
														<td><IMG id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																height="20" alt="��ǥ�� ������ �����մϴ�." src="../../../images/imgVochDelco.gIF" border="0"
																name="imgVochDelco" title="SAP���� ���������Ͽ� RMS���� ������ �� ������ RMS��ǥ�� ������ �����Ѵ�."></td>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
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
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHDATA" width="90" onclick="vbscript:Call gCleanField(txtSUMM,'')">
												<select id="cmbSETTING" style="WIDTH: 90px">
													<OPTION value="" selected>����</OPTION>
													<OPTION value="POSTINGDATE">��ǥ����</OPTION>
													<OPTION value="SUMM">����</OPTION>
													<OPTION value="BA">�������</OPTION>
													<OPTION value="COSTCENTER">�ڽ�Ʈ����</OPTION>
													<OPTION value="SEMU">�����ڵ�</OPTION>
													<OPTION value="DEMANDDAY">���ޱ���</OPTION>
													<OPTION value="DUEDATE">�Աݱ���</OPTION>
													<OPTION value="DEBTOR">��������</OPTION>
													<OPTION value="ACCOUNT">����</OPTION>
													<OPTION value="PREPAYMENT">�����ݱ���</OPTION>
													<OPTION value="SUMMTEXT">����TEXT</OPTION>
												</select></TD>
											<TD class="DATA"><INPUT class="INPUT_L" id="txtSUMM" title="��������" style="WIDTH: 368px; HEIGHT: 21px" size="56"
													name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													title="���並 �ϰ� �����մϴ�" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
													height="20" alt="���並 �ϰ� �����մϴ�" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0"
													name="ImgSUMMApp">
												<DIV id="pnlGUBUN" align="center" style="VISIBILITY: hidden; WIDTH: 450px; POSITION: absolute; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdSALE" title="����" type="radio" CHECKED value="SALE" name="rdGROUP">&nbsp;����&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdBUY" title="����" type="radio" value="BUY" name="rdGROUP">&nbsp;����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="chkSUMMAE" title="�����ջ�üũ" type="checkbox" name="chkSUMMAE">&nbsp;�ջ�&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdMCBUY" title="MC�Ϲݸ������ �������� ����" type="radio" value="MCBUY" name="rdGROUP">&nbsp;MC�Ϲݸ���-&gt;����</DIV>
												<DIV id="pnlFLAG" align="center" style="VISIBILITY: hidden; WIDTH: 250px; POSITION: absolute; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdDIV" title="����" type="radio" CHECKED value="rdDIV" name="rdDIVGUBUN">&nbsp;����&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdSUM" title="�ջ�" type="radio" value="rdSUM" name="rdDIVGUBUN">&nbsp;�ջ�</DIV>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 80%" vAlign="top" align="center">
									<DIV id="pnlTab_amt" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_AMT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32464">
											<PARAM NAME="_ExtentY" VALUE="13097">
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
									<DIV id="pnlTab_gen" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_GEN" style="WIDTH: 100%; HEIGHT: 50%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32464">
											<PARAM NAME="_ExtentY" VALUE="9181">
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
										<OBJECT id="sprSht_GENDTL" style="WIDTH: 100%; HEIGHT: 50%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32464">
											<PARAM NAME="_ExtentY" VALUE="3942">
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
									<DIV id="pnlTab_susu" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_SUSU" style="WIDTH: 100%; HEIGHT: 70%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32464">
											<PARAM NAME="_ExtentY" VALUE="9181">
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
										<OBJECT id="sprSht_SUSUDTL" style="WIDTH: 100%; HEIGHT: 30%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32464">
											<PARAM NAME="_ExtentY" VALUE="3942">
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
									<DIV id="pnlTab_out" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_OUT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="12965">
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
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
				</TR>
			</TABLE>
			<P>
				<!--Input Define Table End--> </TD></TR> 
				<!--Top TR End--> </TABLE> 
				<!--Main End--></P>
		</FORM>
		</TR></TABLE><iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 100%; HEIGHT: 300px" name="frmSapCon"
			src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
