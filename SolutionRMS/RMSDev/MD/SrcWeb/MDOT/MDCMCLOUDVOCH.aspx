<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUDVOCH.aspx.vb" Inherits="MD.MDCMCLOUDVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CGVŬ���� ��ǥ����</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : CGVŬ���� ��ǥó�� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : CGVŬ���� ��ǥó��
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
Dim mobjMDCMCLOUDVOCH
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
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "visible" 
	pnlTab_outsusu.style.visibility = "hidden" 
	pnlTab_cgv.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "hidden" 
	
	frmThis.txtYEARMON1.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "S"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'���������
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TABON
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_outsusu.style.visibility = "visible" 
	pnlTab_cgv.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "visible" 
	
	frmThis.txtYEARMON1.style.visibility = "hidden"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "GO"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'CGV����
Sub btnTab3_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	frmThis.btnTab3.style.backgrounDimage = meURL_TABON
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_outsusu.style.visibility = "hidden" 
	pnlTab_cgv.style.visibility = "visible" 
	
	pnlFLAG.style.visibility = "visible" 
	
	frmThis.txtYEARMON1.style.visibility = "visible"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "GO2"
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
		elseif mstrGUBUN = "GO"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_OUTSUSU
		elseif mstrGUBUN = "GO2"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_CGV
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
		elseif mstrGUBUN = "GO"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUTSUSU)
		elseif mstrGUBUN = "GO2"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_CGV)
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

Sub sprSht_OUTSUSU_Change(ByVal Col, ByVal Row)
	Dim strCODE 
	with frmThis
	
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"PAYCODE") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404103"
			end if 
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUTSUSU,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUTSUSU)
			
		end if 
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"TODATE",Row,Row,false
		End If
		
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUTSUSU, Col, Row
End Sub

Sub sprSht_CGV_Change(ByVal Col, ByVal Row)
	Dim strCODE 
	with frmThis
	
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"paycode") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DEBTOR",Row, "404103"
			end if 
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUTSUSU,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUTSUSU)
			
		end if 
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,True,"TODATE",Row,Row,false
		End If
		
		'ACCOUNT | DEBTOR | DOCUMENTDATE | DEMANDDAY
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"DOCUMENTDATE") and Row = 1 then
			for i=1 to .sprSht_CGV.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DOCUMENTDATE",i, mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"ACCOUNT") and Row = 1 then
			for i=1 to .sprSht_CGV.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"ACCOUNT",i, mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"DEBTOR") and Row = 1 then
			for i=1 to .sprSht_CGV.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DEBTOR",i, mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"DEMANDDAY") and Row = 1 then
			for i=1 to .sprSht_CGV.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"DEMANDDAY",i, mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY", Row) 
			Next
		end if
		
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_CGV, Col, Row
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

Sub sprSht_OUTSUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUTSUSU, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_CGV_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_CGV, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_CGV.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_CGV, 1, intCnt
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

sub sprSht_OUTSUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUTSUSU, ""
		end if
	end with
end sub

sub sprSht_CGV_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CGV, ""
		end if
	end with
end sub


'----------------------------------------------------------
'��Ʈ �ڵ� ��� [��Ʈ Ű��]
'----------------------------------------------------------
'��޾�
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

'��޾� ->����
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

'��޾�->���Ի�
Sub sprSht_OUTSUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_OUTSUSU
	End With
End Sub

'�Ϲ�
Sub sprSht_CGV_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_CGV
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
Sub sprSht_SUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSU
	end with
End Sub

'��޾� ->����
Sub sprSht_SUSUDTL_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSUDTL
	end with
End Sub

'��޾� ->���Ի�
Sub sprSht_OUTSUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_OUTSUSU
	end with
End Sub

'�Ϲ�
Sub sprSht_CGV_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_CGV
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

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDCMCLOUDVOCH = gCreateRemoteObject("cMDOT.ccMDOTCLOUDVOCH")
	Set mobjMDCOVOCH	  = gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	Set mobjMDCOGET		  = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
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
		strSemuComboListB =  "B0" & vbTab & "B5" & vbTab & "BR"
		strSemuComboListA =  "I0" & vbTab & "A0" & vbTab & "AI" & vbTab & "A8" & vbTab & "AZ"
		strBMORDER = "AD0430" & vbTab & " "
		
		'**************************************************
		'���� ��Ʈ ������ hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 25, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������������|����|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|AMTGBN|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    15|  20|  10|    10|       7| 5|       8|          10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|      0"
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
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | BP | DUEDATE  | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "GBN | GFLAG | DEMANDDAY | AMTGBN | MEDFLAG", true 
		
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
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true," CUSTOMERCODE | CUSTNAME | AMT | BP | DUEDATE | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN | GFLAG | DEMANDDAY | AMTGBN | VENDOR | MEDFLAG", true 
		
		'**************************************************
		'��������� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUTSUSU
		mobjSCGLSpr.SpreadLayout .sprSht_OUTSUSU, 36, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUTSUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUTSUSU,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|������������|����|��������|����|BMORDER|������|���ޱ���|���޹��|BANKTYPE|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|����|GFLAG|���ⱸ��|AMTGBN|VENDOR|RMSNO|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUTSUSU, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|   8|    0|      10|     0|     0|    0|		 0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUTSUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUTSUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUTSUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUTSUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUTSUSU, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUTSUSU, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUTSUSU, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_OUTSUSU, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO", true 
		
		'**************************************************
		'CGV���� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_CGV
		mobjSCGLSpr.SpreadLayout .sprSht_CGV, 38, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_CGV,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_CGV,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|������������|����|��������|����|BMORDER|������|���ޱ���|���޹��|BANKTYPE|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|���ⱸ��|AMTGBN|VENDOR|RMSNO|�������ڹ�ȣ|�����|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_CGV, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      10|     0|     0|    0|		      0|     0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht_CGV, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_CGV, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CGV, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_CGV, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_CGV, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_CGV,mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_CGV,mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_CGV,mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_CGV, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_CGV, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,true,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_CGV, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME", true 
		mobjSCGLSpr.CellGroupingEach .sprSht_CGV,"VOCHNO | ERRCODE | ERRMSG"
		
	End with
	pnlTab_susu.style.visibility = "visible" 
    
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMCLOUDVOCH = Nothing
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
		.txtYEARMON1.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUTSUSU.MaxRows = 0
		.sprSht_CGV.MaxRows = 0
		.txtYEARMON.focus	
		
		frmThis.txtYEARMON1.style.visibility = "hidden"
		
		Get_COMBO_VALUE	
		'ó���� ���� ���� ����
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
	End with
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_OUTSUSU.MaxRows = 0
		.sprSht_CGV.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
	
		vntData = mobjMDCMCLOUDVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUTSUSU, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_CGV, "PAYCODE",,,vntData,,160
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
		
       	vntData = mobjMDCMCLOUDVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
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
		.sprSht_OUTSUSU.MaxRows = 0
		.sprSht_CGV.MaxRows = 0
		
		IF strVOCH_TYPE = "S" THEN
			CALL SelectRtn_SUSU()
		ELSEIF strVOCH_TYPE = "GO" THEN
			CALL SelectRtn_OUTSUSU()
		ELSEIF strVOCH_TYPE = "GO2" THEN
			CALL SelectRtn_CGV()
		END IF
		mstrSTAY = TRUE
   	end with
End Sub

Sub SelectRtn_SUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
    
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
			
		vntData = mobjMDCMCLOUDVOCH.SelectRtn_SUSU(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, _
													 strCLIENTCODE, strCLIENTNAME, strGBN)

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
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		
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
		
		
		'if .rdF.checked then
			vntData = mobjMDCMCLOUDVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO, strPREPAYMENT)
		'else
		'	vntData = mobjMDCMCLOUDVOCH.SelectRtn_SUSUDTL_COMMIT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		'end if
																							
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

Sub SelectRtn_OUTSUSU ()
   	Dim vntData
    Dim intCnt
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN

	with frmThis
		.sprSht_OUTSUSU.MaxRows = 0
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMCLOUDVOCH.SelectRtn_OUTSUSU(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
													  strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_OUTSUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "GO"

				mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"ACCOUNT",-1,-1,false
				
				mobjSCGLSpr.SetClipbinding .sprSht_OUTSUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				'ACCOUNT
				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"POSTINGDATE",intCnt,intCnt,false
					End If
				
					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"TODATE",intCnt,intCnt,false
					End If	
				Next
				AMT_SUM .sprSht_OUTSUSU
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_CGV ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON,strYEARMON1, strCLIENTCODE, strCLIENTNAME, strGBN

	with frmThis
		.sprSht_CGV.MaxRows = 0
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strYEARMON1		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMCLOUDVOCH.SelectRtn_CGV(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, strYEARMON1, _
												  strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_CGV") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "GO2"

				mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"ACCOUNT",-1,-1,false
				
				mobjSCGLSpr.SetClipbinding .sprSht_CGV, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				'ACCOUNT
				For intCnt = 1 To .sprSht_CGV.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_CGV, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,true,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_CGV, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"POSTINGDATE",intCnt,intCnt,false
					End If
				
					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_CGV,True,"TODATE",intCnt,intCnt,false
					End If	
				Next
				AMT_SUM .sprSht_CGV
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

Function DataValidation_OUTSUSU ()
	DataValidation_OUTSUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_OUTSUSU.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUTSUSU = True
End Function

Function DataValidation_CGV ()
	DataValidation_CGV = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_CGV.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_CGV = True
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
			ELSEIF strVOCH_TYPE = "GO" THEN
				if DataValidation_OUTSUSU =false then exit sub
				CALL ProcessRtn_OUTSUSU()
			ELSEIF strVOCH_TYPE = "GO2" THEN
				if DataValidation_CGV =false then exit sub
				CALL ProcessRtn_CGV()
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
Sub ProcessRtn_OUTSUSU()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'��ǥ ä���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUTSUSU,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK")
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
		For intMaxCnt = 1 To .sprSht_OUTSUSU.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",intMaxCnt))
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
				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",intCnt) = "1" then	
					
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt)'
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
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
			
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ+1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						end if												
					end if 
				Next
			end if
		'�ջ� üũ�� ���
		ELSE
			if mstrPROCESS = "Create" then
				'1���� TRANRANK �� �ְ���ŭ ���鼭 ��ǥó��
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'��ü ��Ʈ�� ���鼭 üũ�� �� �������� TRANSRANK �� ������ �ݾ��� �ջ��Ѵ�.
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",i) = "1" then	
						'üũ�� �Ǿ��ִ� �����Ͱ� �ջ��Ҽ��ִ� �����Ͷ�� 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'���� �ջ� �����ʹ� �ѵ����͸� �����Ѵ�.
								Else
									'ä���� �����Ѵ�.(�ջ���ǥ�� ä�� ����)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)'
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
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
									
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
									
									For j = 1 To .sprSht_OUTSUSU.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",j) = 1 Then
												'�ջ� �����Ͷ�� 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",j)
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
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
							'TRANSRANK �� ������ �ݾװ� �ΰ����� �ջ��Ѵ�.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'�ݾ��հ� �ΰ��� �հ�� ����� ������ ����
								Else
									strIF_CNT = strIF_CNT + 1
									'������ ���� ������ ��� �ѵ����ͷ� ����� �ݾ��� �ջ��� �ݾ��� �Է��Ѵ�.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i)	
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
									
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
												
																			
									For j = 1 To .sprSht_OUTSUSU.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",j) = 1 Then
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
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

'�������[������ǥ����]
Sub ProcessRtn_CGV()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'��ǥ ä���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
		
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_CGV,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK")
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
		For intMaxCnt = 1 To .sprSht_CGV.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",intMaxCnt))
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
				For intCnt = 1 To .sprSht_CGV.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"chk",intCnt) = "1" then		
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",intCnt)'
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
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",intCnt)
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_CGV.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
			
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ+1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",intCnt)
						end if												
					end if 
				Next
			end if
		'�ջ� üũ�� ���
		ELSE
			if mstrPROCESS = "Create" then
				'1���� TRANRANK �� �ְ���ŭ ���鼭 ��ǥó��
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'��ü ��Ʈ�� ���鼭 üũ�� �� �������� TRANSRANK �� ������ �ݾ��� �ջ��Ѵ�.
					For i = 1 To .sprSht_CGV.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_CGV.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"chk",i) = "1" then	
						'üũ�� �Ǿ��ִ� �����Ͱ� �ջ��Ҽ��ִ� �����Ͷ�� 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'���� �ջ� �����ʹ� �ѵ����͸� �����Ѵ�.
								Else
									'ä���� �����Ѵ�.(�ջ���ǥ�� ä�� ����)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",i)'
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
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",i)
									
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
									
									For j = 1 To .sprSht_CGV.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",j) = 1 Then
												'�ջ� �����Ͷ�� 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",j)
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
					
					For i = 1 To .sprSht_CGV.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 Then
							'TRANSRANK �� ������ �ݾװ� �ΰ����� �ջ��Ѵ�.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_CGV.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'�ݾ��հ� �ΰ��� �հ�� ����� ������ ����
								Else
									strIF_CNT = strIF_CNT + 1
									'������ ���� ������ ��� �ѵ����ͷ� ����� �ݾ��� �ջ��� �ݾ��� �Է��Ѵ�.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VOCHNO",i)	
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",i)
									
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
												
																			
									For j = 1 To .sprSht_CGV.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",j) = 1 Then
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"BMORDER",i) 
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
				intRtn = mobjMDCMCLOUDVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "G")
			elseif mstrGUBUN = "GO" then
				intRtn = mobjMDCMCLOUDVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "GO")
			elseif mstrGUBUN = "GO2" then
				intRtn = mobjMDCMCLOUDVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "GO2")
			end if 

			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "GO" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "GO2" THEN
					mobjSCGLSpr.SetFlag  .sprSht_CGV, meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			End if

   		elseif mstrPROCESS ="Delete" then
   			if mstrGUBUN = "S" then
				intRtn = mobjMDCMCLOUDVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CGV")
			else
				intRtn = mobjMDCMCLOUDVOCH.VOCHDELL_BUY(gstrConfigXml, strRETURNLIST, mstrGUBUN, "CGV", mstrGUBUN)
			end if 
   			
   			if not gDoErrorRtn ("VOCHDELL") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "GO" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "GO2" THEN
					mobjSCGLSpr.SetFlag  .sprSht_CGV,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			End if
   		End if

   		IF mstrGUBUN = "S" THEN
			.sprSht_SUSU.focus()
		ELSEIF mstrGUBUN = "GO" THEN
			.sprSht_OUTSUSU.focus()
		ELSEIF mstrGUBUN = "GO2" THEN
			.sprSht_CGV.focus()
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
		ELSEIF mstrGUBUN = "GO" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUTSUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "GO2" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_CGV,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF
		
		'ó�� ������ü ȣ��
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		intRtn = mobjMDCMCLOUDVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'��� �÷��� Ŭ����
			IF mstrGUBUN = "S" THEN
				mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "GO" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "GO2" THEN
				mobjSCGLSpr.SetFlag  .sprSht_CGV,meCLS_FLAG
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
		ELSEIf mstrGUBUN = "GO"  then
			If .sprSht_OUTSUSU.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUTSUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "GO2"  then
			If .sprSht_CGV.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_CGV.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 THEN
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
					
					intRtn = mobjMDCMCLOUDVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CGV" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "GO"  then
			for i = .sprSht_OUTSUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i)
					
					intRtn = mobjMDCMCLOUDVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CGV", "GO" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUTSUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "GO2"  then
			for i = .sprSht_CGV.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"VOCHNO",i)
					
					intRtn = mobjMDCMCLOUDVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "CGV", "GO2")
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_CGV,i
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
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="131" background="../../../images/back_p.gIF"
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
											<td class="TITLE">CGVŬ���� ��ǥ����&nbsp;</td>
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
												<TD class="SEARCHDATA" style="WIDTH: 164px"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 60px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="9" name="txtYEARMON">
														<INPUT class="INPUT" id="txtYEARMON1" style="WIDTH: 60px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="9" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
													width="50">&nbsp;������
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 112px; HEIGHT: 22px"
														type="text" maxLength="100" size="13" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
														accessKey=",M" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
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
									<td class="DATA">
										�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
											accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
										<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
											readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
									</td>
								</TR>
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE height="15" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
														type="button" value="����" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" size="20" value="���������" name="btnTab2"> <INPUT class="BTNTAB" id="btnTab3" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" size="20" value="CGV����" name="btnTab3">
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
										<DIV id="pnlTab_susu" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_SUSU" style="WIDTH: 100%; HEIGHT: 70%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31882">
												<PARAM NAME="_ExtentY" VALUE="9340">
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
												<PARAM NAME="_ExtentX" VALUE="31882">
												<PARAM NAME="_ExtentY" VALUE="3995">
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
										<DIV id="pnlTab_outsusu" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_OUTSUSU" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31882">
												<PARAM NAME="_ExtentY" VALUE="13335">
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
										<DIV id="pnlTab_cgv" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_CGV" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31882">
												<PARAM NAME="_ExtentY" VALUE="13335">
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
