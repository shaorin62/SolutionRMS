<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORVOCH.aspx.vb" Inherits="MD.MDCMOUTDOORVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� ��ǥ����</title>
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
'HISTORY    :1) 2011/01/22 By kty
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
Dim mobjMDCMOUTDOORVOCH
Dim mobjMDCOVOCH
Dim mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim mstrPROCESS
Dim mstrSTAY

mstrSTAY = TRUE

mstrGUBUN = "S"
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

'����
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TABON
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "visible" 
	pnlTab_gen.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "hidden" 
	
	frmThis.txtEDYEARMON.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "S"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'����
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TABON
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_gen.style.visibility = "visible" 
	
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
		
		if mstrGUBUN = "S"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_SUSU
		elseif mstrGUBUN = "O"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_OUT
		end if
		
	End With
	gFlowWait meWAIT_OFF
End Sub

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
		if mstrGUBUN = "S"  then 
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_SUSU)
			settingRowChange (.sprSht_SUSUDTL)
		'�Ϲ�
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

'�� �˾� ��ư
Sub ImgTIMCODE1_onclick
	Call TIMCODE_POP()
End Sub

'���� ������List ��������
Sub TIMCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value), _
							trim(.txtTIMCODE.value), trim(.txtTIMNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' �ڵ�� ǥ��
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
											trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					
					SelectRtn
				Else
					Call TIMCODE_POP()
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
Sub sprSht_SUSU_Change(ByVal Col, ByVal Row)
	Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
	Dim strFROMDATE, strTODATE, strSUMMTEXT
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT") Then
			DeleteRtn_SUSUDTL (Row)
			
			SelectRtn_SUSUDTL Col,Row
			If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"TODATE",Row,Row,false
			End If
		End If
		
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
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"paycode") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "P" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "N" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "Z" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404100"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "I" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404103"
			end if
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUT,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUT)
			 
		end if 

		If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,false
		End If
	
		'�������� ��û���� ��ǥ���ڿ� �������� 90 �� �ʰ� �Ұ�� �޼��� �ڽ� ����ֱ�
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"DEMANDDAY") then
			IF datediff("D",mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE", Row),mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY", Row)) > 90 THEN
				gErrorMsgBox "��ǥ���ڿ� �������� ���̰� 90���� �ʰ� �Ͽ����ϴ�.","Ȯ�� ����."
			End if
		end if
		
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

'-----------------------------------
' SpreadSheet Ŭ��
'-----------------------------------
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

'-----------------------------------
' SpreadSheet ���� Ŭ��
'-----------------------------------
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

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDCMOUTDOORVOCH = gCreateRemoteObject("cMDOT.ccMDOTOUTDOORVOCH")
	Set mobjMDCOVOCH		= gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	Set mobjMDCOGET		    = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
    mobjSCGLCtl.DoEventQueue
    
    Dim strComboPREPAYMENT
	Dim strSemuComboListB, strSemuComboListA
	Dim strBMORDER
	
	gSetSheetDefaultColor
    with frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strSemuComboListB =  "B0" & vbTab & "D0" & vbTab & "B7"
		strSemuComboListA =  "I0" & vbTab & "A0" & vbTab & "I0" & vbTab & "V0"
		strBMORDER = "AD0160" & vbTab & " "
		'**************************************************
		'���� ��Ʈ ������ hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 26, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG | VOCH_TYPE"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������������|����|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|AMTGBN|MEDFLAG|VOCH_TYPE"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    15|  20|  10|    10|       7| 5|       8|          10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|      0|		0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "CUSTOMERCODE | CUSTNAME | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CUSTOMERCODE | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | BP | DUEDATE  | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | VOCH_TYPE"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "GBN | GFLAG | DEMANDDAY | AMTGBN | MEDFLAG | VOCH_TYPE", true 
		
		'**************************************************
		'���� ��Ʈ ������ dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_SUSUDTL, 30, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSUDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSUDTL,		    "��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������������|����|��������|����|BMORDER|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|AMTGBN|VENDOR|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSUDTL, "-1",  "       8|        10|    15|  20|       5|         8|  10|    10|       7| 5|       8|          10|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSUDTL, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "CUSTOMERCODE | BA | SEMU | BP | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | GBN | VOCHNO",-1,-1,2,2,false '���
		'mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true,"CUSTOMERCODE | CUSTNAME | AMT | DUEDATE | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN | GFLAG | DEMANDDAY | AMTGBN | VENDOR | MEDFLAG", true 

		'**************************************************
		'���� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 36, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|������������|����|��������|����|BMORDER|������|���ޱ���|���޹��|BANKTYPE|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|����|GFLAG|���ⱸ��|AMTGBN|VENDOR|RMSNO|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|   8|    0|      10|     0|     0|    0|		 0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '����
		'mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"CUSTOMERCODE | CUSTNAME  | AMT | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO", true 

	End with
	pnlTab_susu.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMOUTDOORVOCH = Nothing
	Set mobjMDCOVOCH = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtEDYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.txtYEARMON.focus	
		
		Get_COMBO_VALUE	
		'ó���� ���� ���� ����
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		
		frmThis.txtEDYEARMON.style.visibility = "hidden"
	End with
End Sub


Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_OUT.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
	
		vntData = mobjMDCMOUTDOORVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
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
		
       	vntData = mobjMDCMOUTDOORVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
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
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		
		IF strVOCH_TYPE = "S" THEN
			CALL SelectRtn_SUSU()
		ELSEIF strVOCH_TYPE = "O" THEN
			CALL SelectRtn_OUT()
		END IF
		mstrSTAY = TRUE
   	end with
End Sub

Sub SelectRtn_SUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strGBN
    Dim strREAL_MED_CODE, strREAL_MED_NAME
	
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strTIMCODE = .txtTIMCODE.value
		strTIMNAME = .txtTIMNAME.value
		strREAL_MED_CODE = .txtREAL_MED_CODE.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
	
		vntData = mobjMDCMOUTDOORVOCH.SelectRtn_SUSU(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, _
													 strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, _
													 strREAL_MED_CODE, strREAL_MED_NAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_SUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "S"
								
				mobjSCGLSpr.SetClipbinding .sprSht_SUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht_SUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"DUEDATE",intCnt,intCnt,false
					End If
				
					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"TODATE",intCnt,intCnt,false
					End If		
				Next
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
		
		if .rdF.checked then
			vntData = mobjMDCMOUTDOORVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO, strPREPAYMENT)
		else
		
			vntData = mobjMDCMOUTDOORVOCH.SelectRtn_SUSUDTL_COMMIT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		end if
																							
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
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strGBN
    Dim strREAL_MED_CODE, strREAL_MED_NAME
	Dim strEDYEARMON
	Dim strAORGUBUN
	
	with frmThis
		.sprSht_OUT.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strEDYEARMON	= .txtEDYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		strTIMCODE		= .txtTIMCODE.value
		strTIMNAME		= .txtTIMNAME.value
		strREAL_MED_CODE= .txtREAL_MED_CODE.value
		strREAL_MED_NAME= .txtREAL_MED_NAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMOUTDOORVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, strEDYEARMON, _
													strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, _
													strREAL_MED_CODE, strREAL_MED_NAME,  _
													strGBN)

		if not gDoErrorRtn ("SelectRtn_OUT") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "O"

				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"ACCOUNT",-1,-1,false
				
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				'ACCOUNT
				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"POSTINGDATE",intCnt,intCnt,false
					End If
				
					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,false
					End If	
				Next
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

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

Function DataValidation_GEN ()
	DataValidation_GEN = false	
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
	DataValidation_GEN = True
End Function

'�������
Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn
	gFlowWait meWAIT_ON
	
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
			IF strVOCH_TYPE = "S" THEN
				if DataValidation_SUSU =false then exit sub
				CALL ProcessRtn_SUSU()
			ELSEIF strVOCH_TYPE = "O" THEN
				if DataValidation_GEN =false then exit sub
				CALL ProcessRtn_OUT()
			END IF
		ELSE
			gErrorMsgBox "��ǥó�� �������Դϴ�.","��ǥó�� �ȳ�"
		END IF
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
		
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_SUSUDTL,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG")
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
		
		IF_GUBUN = "RMS_0006"
		
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

					strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" :  strTYPE = ""

					strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","")
					strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"MEDFLAG",intCnt)
					strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
					strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
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
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
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

'�������[������ǥ����]
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'��ǥ ä���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK")
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
	
		IF_GUBUN = "RMS_0007"
		
		'��Ʈ ��ü�� ���鼭 üũ�� ���� TRANSRANK �� �ִ� ���� ��´�.
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		'�ջ꿡 ����� ����
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		'����üũ ������ ���
		IF .rdDIV.checked THEN
		
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" :  strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)
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
							strHSEQ = strHSEQ+1
							
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
		'�ջ� üũ�� ���
		ELSE
			if mstrPROCESS = "Create" then
				'1���� TRANRANK �� �ְ�����ŭ ���鼭 ��ǥó��
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'��ü ��Ʈ�� ���鼭 üũ�� �� �������� TRANSRANK �� ������ �ݾ��� �ջ��Ѵ�.
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
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = "1" then	
						'üũ�� �Ǿ��ִ� �����Ͱ� �ջ��Ҽ��ִ� �����Ͷ�� 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'���� �ջ� �����ʹ� �ѵ����͸� �����Ѵ�.
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
									'������ ���� ������ ��� �ѵ����ͷ� ����� �ݾ��� �ջ��� �ݾ��� �Է��Ѵ�.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
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
									
									IF strIF_CNT = "1" THEN
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
									ELSE
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
									END IF 
									
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then
												'�ջ� �����Ͷ�� 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i) + "|" + _
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
											END IF 
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								END IF
							end if
						end if 
					Next
				 Next
			'�ջ��� ��� ��ǥ����
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
							'TRANSRANK �� ������ �ݾװ� �ΰ����� �ջ��Ѵ�.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUT.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'�ݾ��հ� �ΰ��� �հ�� ����� ������ ����
								Else
									strIF_CNT = strIF_CNT + 1
									'������ ���� ������ ��� �ѵ����ͷ� ����� �ݾ��� �ջ��� �ݾ��� �Է��Ѵ�.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
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
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
											END IF
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								END IF
								intCnt2 = intCnt
							END IF
						end if 
					Next
				NEXT
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

'--------------------------------------------------
' ��ǥ���� �� ��ǥ��ȣ �޾ƿ��� �� ���� RMS������Ʈ
'--------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	gFlowWait meWAIT_ON
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
	
		if mstrPROCESS ="Create" then
			if mstrGUBUN = "S" then
				intRtn = mobjMDCMOUTDOORVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "D")
			elseif mstrGUBUN = "O" then
				intRtn = mobjMDCMOUTDOORVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "D2")
			end if 

			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT, meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   			
   		elseif mstrPROCESS ="Delete" then
   		
   			intRtn = mobjMDCMOUTDOORVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "OUTDOOR")
   			
   			if not gDoErrorRtn ("VOCHDELL") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "O" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   		end if 
   		IF mstrGUBUN = "S" THEN
			.sprSht_SUSU.focus()
		ELSEIF mstrGUBUN = "O" THEN
			.sprSht_OUT.focus()
		END IF
	End With
	gFlowWait meWAIT_OFF
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   		
   		IF NOT .rdE.checked THEN
			gErrorMsgBox "������ȸ�� �����մϴ�.","�����׻���"
			exit sub
		end if 
		
		IF mstrGUBUN = "S" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_SUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "O" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF
		
		'ó�� ������ü ȣ��
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		intRtn = mobjMDCMOUTDOORVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'��� �÷��� Ŭ����
			IF mstrGUBUN = "S" THEN
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
	
		If mstrGUBUN = "S"  then  
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
		If mstrGUBUN = "S"  then
			for i = .sprSht_SUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"VOCHNO",i)
					
					intRtn = mobjMDCMOUTDOORVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "OUTDOOR" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
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
					
					intRtn = mobjMDCMOUTDOORVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "OUTDOOR" )
					
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
		//***************************************��1) frmSapCon ���� ������ �� �̿��Ͽ� Submit �ϴ� �Լ�********************************************
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
												<TABLE cellSpacing="0" cellPadding="0" width="82" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���� ��ǥ����&nbsp;</td>
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
							<TBODY>
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
													width="50">&nbsp;���
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 80px"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 80px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="9" name="txtYEARMON"></TD>
												<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtEDYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
														maxLength="6" size="9" name="txtEDYEARMON"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
													width="50">&nbsp;������
												</TD>
												<TD class="SEARCHDATA" width="190"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 112px; HEIGHT: 22px"
														type="text" maxLength="100" size="13" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
														accessKey=",M" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 28px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME, txtTIMCODE)"
													width="28">��</TD>
												<TD class="SEARCHDATA" width="190"><INPUT class="INPUT_L" id="txtTIMNAME" title="����" style="WIDTH: 112px; HEIGHT: 22px" type="text"
														maxLength="100" size="22" name="txtTIMNAME"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTIMCODE"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME,txtREAL_MED_CODE)"
													width="50">&nbsp;��ü��
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 112px; HEIGHT: 22px"
														type="text" maxLength="100" size="16" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgREAL_MED_CODE"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
														accessKey=",M" type="text" maxLength="6" name="txtREAL_MED_CODE">
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
												<TD class="SEARCHDATA" colspan="5">
													<INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;�Ϸ�&nbsp;
													<INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" value="rdF" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')"
														CHECKED>&nbsp;�̿Ϸ�&nbsp; <INPUT id="rdE" title="������ǥ ������ȸ" type="radio" value="rdE" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;����&nbsp;
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 30px"></TD>
								</TR>
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE height="15" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
														type="button" value="����" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" size="20" value="����" name="btnTab2">
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
												<TD class="DATA"><INPUT class="INPUT_L" id="txtSUMM" title="��������" style="WIDTH: 368px; HEIGHT: 21px" type="text"
														size="56" name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
														title="���並 �ϰ� �����մϴ�" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20"
														alt="���並 �ϰ� �����մϴ�" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0" name="ImgSUMMApp">
													<DIV id="pnlFLAG" style="VISIBILITY: hidden; WIDTH: 250px; POSITION: absolute; HEIGHT: 24px"
														align="center" ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdDIV" title="����" type="radio" CHECKED value="rdDIV" name="rdDIVGUBUN">&nbsp;����&nbsp;&nbsp;&nbsp; 
														&nbsp; <INPUT id="rdSUM" title="�ջ�" type="radio" value="rdSUM" name="rdDIVGUBUN">&nbsp;�ջ�&nbsp;
													</DIV>
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
										<DIV id="pnlTab_susu" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_SUSU" style="WIDTH: 100%; HEIGHT: 70%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="8811">
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
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="3784">
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
											<OBJECT id="sprSht_OUT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="12568">
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
							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
				</TR>
			</TABLE>
		</FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 100%; HEIGHT: 300px" src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx"
			name="frmSapCon"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>