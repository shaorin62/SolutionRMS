<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMAOR_MEDIUM.aspx.vb" Inherits="MD.MDCMAOR_MEDIUM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>AOR ��ü���� ���/��ȸ</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'�ý��۱��� : MD/��ŷ ȭ��(MDCMBOOKING)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMAOR_MEDIUM.aspx
'��      �� : AOR ���� ���� �ݾ� ��ȸ ����
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2012.05.15 OH SE HOON
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT id="clientEventHandlersVBS" language="vbscript">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOAORMEDIUM, mobjMDCOGET
Dim mstrCheck
Dim mstrHIDDEN
Dim mcomecalender

CONST meTAB = 9
mcomecalender = FALSE
mstrHIDDEN = 0
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody").style.display = "inline"
			document.getElementById("tblSheet").style.height = "65%"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody").style.display = "none"
			document.getElementById("tblSheet").style.height = "82%"
		End If

		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

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
	If frmThis.txtYEARMON1.value = "" Then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
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

'�űԹ�ư
Sub imgREG_onclick ()
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		'mobjSCGLSpr.ExportComboType = "2"
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	On error resume next
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData

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
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'��ü�� �˾� ��ư
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData

		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "B")
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
				Else
					Call REAL_MED_CODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' �˾� ��ư[�Է¿�]
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
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")

			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
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
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
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
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'���ۻ�/����� �˾� 
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtEXCLIENTCODE
		End If
	end With
End Sub

Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)

			vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code�� ����
					.txtEXCLIENTNAME.value = trim(vntData(2,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'���μ� �˾� 
Sub imgDEPT_CD_onclick
	Call DEPT_CD_POP()
End Sub

Sub DEPT_CD_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtDEPT_NAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtDEPT_CD.value = trim(vntRet(0,0))	'Code�� ����
			.txtDEPT_NAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtDEPT_CD
		End If
	end With
End Sub

'���μ� �˾�
Sub txtDEPT_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPT_NAME.value))
			
			If not gDoErrorRtn ("GetCC") Then
				If mlngRowCnt = 1 Then
					.txtDEPT_CD.value = trim(vntData(0,1))
					.txtDEPT_NAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call DEPT_CD_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'��ü �˾� ��ư
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'���� ������List ��������
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value), _
							trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "")
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtMEDCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtMEDNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtREAL_MED_CODE.value = trim(vntRet(3,0))       ' �ڵ�� ǥ��
			.txtREAL_MED_NAME.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMEDNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
											trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "")
			
			If not gDoErrorRtn ("GetMEDGUBNCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtMEDNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtREAL_MED_CODE.value = trim(vntData(3,1))
					.txtREAL_MED_NAME.value = trim(vntData(4,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call MEDCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' ������ �޷�
'****************************************************************************************
Sub imgCalEndar_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar,"txtDEMANDDAY_onchange()"
	Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DEMANDDAY"), frmThis.sprSht.ActiveRow)
	mcomecalender = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCARD_AMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbMED_FLAG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtREAL_MED_CODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMISSION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_CARD.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEX_CARD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEMANDDAY_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEXCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEXCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtOUT_AMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtOUT_AMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_AMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************
Sub txtYEARMON_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		
		AMT_CAL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCARD_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, frmThis.txtCARD_AMT.value
		EXCARD_CAL frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub cmbMED_FLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		IF frmThis.cmbMED_FLAG.value =  "A" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "������"
		ELSEIF frmThis.cmbMED_FLAG.value =  "A2" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���̺�"
		ELSEIF frmThis.cmbMED_FLAG.value =  "T" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���������"
		ELSEIF frmThis.cmbMED_FLAG.value =  "B" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "�Ź�"
		ELSEIF frmThis.cmbMED_FLAG.value =  "C" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "����"
		ELSEIF frmThis.cmbMED_FLAG.value =  "O" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���ͳ�"
		ELSEIF frmThis.cmbMED_FLAG.value =  "D" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "����"			
		END IF
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCOMMISSION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
		EXCARD_CAL frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCOMMI_RATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, frmThis.txtCOMMI_RATE.value
		COMMISSION_CAL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtEX_CARD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, frmThis.txtEX_CARD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEX_CARD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, frmThis.txtEX_CARD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtDEMANDDAY_onchange
	Dim strdate 
	Dim strDEMANDDAY
	strdate = "" : strDEMANDDAY = ""
	
	With frmThis
		strdate=.txtDEMANDDAY.value
		If mcomecalender Then
			strDEMANDDAY = strdate
		else
			If len(strdate) = 4 Then
				strDEMANDDAY = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strDEMANDDAY = strdate
			elseif len(strdate) = 3 Then
				strDEMANDDAY = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strDEMANDDAY = strdate
			End If
		End If

		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtOUT_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, frmThis.txtOUT_AMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtEX_AMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, frmThis.txtEX_AMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

Sub txtMEMO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End SuB

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'���ް���
Sub txtAMT_onblur
	With frmThis
		Call gFormatNumber(.txtAMT,0,True)
	end With
End Sub
'���ް��� �ΰ���
Sub txtVAT_onblur
	With frmThis
		Call gFormatNumber(.txtVAT,0,True)
	end With
End Sub

'�ΰ��� ����
Sub txtSUMAMTVAT_onblur
	With frmThis
		Call gFormatNumber(.txtSUMAMTVAT,0,True)
	end With
End Sub

'������
Sub txtCOMMISSION_onblur
	With frmThis
		Call gFormatNumber(.txtCOMMISSION,0,True)
	end With
End Sub

'ī�� ������
Sub txtCARD_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtCARD_AMT,0,True)
	end With
End Sub

'ī�� ���� �ݾ�
Sub txtEX_CARD_onblur
	With frmThis
		Call gFormatNumber(.txtEX_CARD,0,True)
	end With
End Sub

'��ü��Ȯ���ݾ�
Sub txtOUT_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtOUT_AMT,0,True)
	end With
End Sub

'���۴���� Ȯ�� �ݾ�
Sub EX_AMT_onblur
	With frmThis
		Call gFormatNumber(.txtEX_AMT,0,True)
	end With
End Sub

'-----------------------------------------------------------------------------------------   
' õ���� ������ ���ֱ� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'���ް���
Sub txtAMT_onfocus
	With frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end With
End Sub

'���ް��� �ΰ���
Sub txtVAT_onfocus
	With frmThis
		.txtVAT.value = Replace(.txtVAT.value,",","")
	end With
End Sub

'�ΰ��� ����
Sub txtSUMAMTVAT_onfocus
	With frmThis
		.txtSUMAMTVAT.value = Replace(.txtSUMAMTVAT.value,",","")
	end With
End Sub

'������
Sub txtCOMMISSION_onfocus
	With frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end With
End Sub

'ī�� ������
Sub txtCARD_AMT_onfocus
	With frmThis
		.txtCARD_AMT.value = Replace(.txtCARD_AMT.value,",","")
	end With
End Sub

'ī�� ���� �ݾ�
Sub txtEX_CARD_onfocus
	With frmThis
		.txtEX_CARD.value = Replace(.txtEX_CARD.value,",","")
	end With
End Sub

'��ü��Ȯ���ݾ�
Sub txtOUT_AMT_onfocus
	With frmThis
		.txtOUT_AMT.value = Replace(.txtOUT_AMT.value,",","")
	end With
End Sub

'���۴���� Ȯ�� �ݾ�
Sub txtEX_AMT_onfocus
	With frmThis
		.txtEX_AMT.value = Replace(.txtEX_AMT.value,",","")
	end With
End Sub

'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		
		frmThis.txtSELECTAMT.value = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON1.value
		
		IF frmThis.cmbMED_FLAG.value = "A" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "������"
		ELSEIF frmThis.cmbMED_FLAG.value = "A2" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���̺�"
		ELSEIF frmThis.cmbMED_FLAG.value = "T" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���������"
		ELSEIF frmThis.cmbMED_FLAG.value = "B" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "�Ź�"
		ELSEIF frmThis.cmbMED_FLAG.value = "C" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "����"
		ELSEIF frmThis.cmbMED_FLAG.value = "O" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "���ͳ�"
		ELSEIF frmThis.cmbMED_FLAG.value = "D" THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "����"
		END IF 
		'AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | OUT_AMT | EX_AMT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEMANDDAY",frmThis.sprSht.ActiveRow, gNowDate2
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VAT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, 15
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, 0
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, "G00076"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, "�������� �÷���(��)"
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub sprSht_Keyup(KeyCode, Shift) 
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
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
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or _ 
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") or _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT") Then
			
			strSUM = 0 : intSelCnt = 0 : intSelCnt1 = 0
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i,j
	Dim vntData_col, vntData_row
	
	With frmThis
		strSUM = 0 : intColCnt = 0 : intRowCnt = 0
		
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or _ 
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") or _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT") Then
			
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intColCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intRowCnt)
					
					for i = 0 to intColCnt -1
						if vntData_col(i) <> "" then
							FOR j = 0 TO intRowCnt -1
								If vntData_row(j) <> "" Then
									if typename(mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))) = "String" then
										exit sub
									end if 
									strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
								End If
							Next
						end if 
					next
					.txtSELECTAMT.value = strSUM
				End If
			else
				.txtSELECTAMT.value = 0
			End If
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim strCode, strCodeName
   	Dim intCnt
	With frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strCode = "" : strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON") Then .txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDDAY") Then .txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MED_FLAG") Then 
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "������" THEN
				.cmbMED_FLAG.value = "A"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���̺�" THEN
				.cmbMED_FLAG.value = "A2"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���������" THEN
				.cmbMED_FLAG.value = "T"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "�Ź�" THEN
				.cmbMED_FLAG.value = "B"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "����" THEN
				.cmbMED_FLAG.value = "C"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���ͳ�" THEN
				.cmbMED_FLAG.value = "O"
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "����" THEN
				.cmbMED_FLAG.value = "D"
			END IF 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then 
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			AMT_CAL Col,Row
		end if 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then 
			.txtCOMMI_RATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
			COMMISSION_CAL Col,Row	'��������� ���
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then 
			.txtCOMMISSION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
			EXCARD_CAL  Col, Row	'ī����������ܱݾ� ���
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CARD_AMT") Then 
			.txtCARD_AMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
			EXCARD_CAL  Col, Row	'ī����������ܱݾ� ���
		end if

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CARD") Then .txtEX_CARD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then .txtOUT_AMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_AMT")  Then .txtEX_AMT.value  = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_AMT",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE") Then	.txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						.txtCLIENTCODE.value = vntData(0,1)
						.txtCLIENTNAME.value = vntData(1,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
	
		'��ü �����
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDCODE") Then .txtMEDCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", _
													strCode, strCodeName, "MED_PRINT")

				If not gDoErrorRtn ("GetMEDGUBNCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntData(4,1)
						.txtMEDCODE.value = vntData(0,1)
						.txtMEDNAME.value = vntData(1,1)
						.txtREAL_MED_CODE.value = vntData(3,1)
						.txtREAL_MED_NAME.value = vntData(4,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_CODE") Then .txtREAL_MED_CODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row) = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntData(1,1))
						.txtREAL_MED_CODE.value = trim(vntData(0,1))	    ' Code�� ����
						.txtREAL_MED_NAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��

						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_CD") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.txtDEPT_CD.value = trim(vntData(0,1))
						.txtDEPT_NAME.value = trim(vntData(1,1))
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTCODE") Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code�� ����
						.txtEXCLIENTNAME.value = trim(vntData(2,1))	'�ڵ�� ǥ��
						
						.txtEXCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME"), Row
						.txtEXCLIENTNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO") Then .txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))

			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				.txtCLIENTCODE.value = vntRet(0,0)		
				.txtCLIENTNAME.value = vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then		
			vntInParams = array("","" , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)), "MED_PRINT")

			vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntRet(4,0)
				.txtMEDCODE.value = vntRet(0,0)
				.txtMEDNAME.value = vntRet(1,0)
				.txtREAL_MED_CODE.value = vntRet(3,0)
				.txtREAL_MED_NAME.value = vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then		
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntRet(1,0))
				.txtREAL_MED_CODE.value = vntRet(0,0)
				.txtREAL_MED_NAME.value = vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)

			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				
				.txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code�� ����
				.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'�ڵ�� ǥ��

				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(1,0)
				
				.txtDEPT_CD.value = trim(vntRet(0,0))	'Code�� ����
				.txtDEPT_NAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If

		sprShtToFieldBinding Col, Row
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�.
		.txtCLIENTNAME1.focus()
		.sprSht.Focus
	End With
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			sprShtToFieldBinding Col,Row
		elseif Row = 0 and Col = 1 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	end With
End Sub

Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

'��Ʈ�� �������ѷο��� ������ ��� �ʴ��� ���ε�
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '�׸��� �����Ͱ� ������ ������.
		
		.txtYEARMON.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtCARD_AMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "������" THEN
			.cmbMED_FLAG.value	= "A"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���̺�" THEN
			.cmbMED_FLAG.value	= "A2"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���������" THEN
			.cmbMED_FLAG.value	= "T"
			
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "�Ź�" THEN
			.cmbMED_FLAG.value	= "B"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "����" THEN
			.cmbMED_FLAG.value	= "C"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "���ͳ�" THEN
			.cmbMED_FLAG.value	= "O"
		
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "����" THEN
			.cmbMED_FLAG.value	= "D"
		END IF 
		
		.txtREAL_MED_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtREAL_MED_CODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtMEDNAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		.txtMEDCODE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		
		.txtCOMMISSION.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		.txtCOMMI_RATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtEX_CARD.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		.txtDEMANDDAY.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		.txtEXCLIENTNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		.txtEXCLIENTCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		
		.txtDEPT_NAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtDEPT_CD.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		
		.txtOUT_AMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		.txtEX_AMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_AMT",Row)
		
		.txtMEMO.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
   	end With
	
	Call gFormatNumber(frmThis.txtAMT,0,True)
	Call gFormatNumber(frmThis.txtCARD_AMT,0,True)
	Call gFormatNumber(frmThis.txtCOMMISSION,0,True)
	Call gFormatNumber(frmThis.txtEX_CARD,0,True)
	Call gFormatNumber(frmThis.txtOUT_AMT,0,True)
	Call gFormatNumber(frmThis.txtEX_AMT,0,True)
	
	Call Field_Lock ()
End Function

'------------------------------------------------------------------------------------------
' Field_Lock  �ŷ�������ȣ�� ���ݰ�꼭 ��ȣ�� ������ �����Ҽ� ������ �ʵ带 ReadOnlyó��
'------------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			'�ŷ������� �����Ǹ� �ʵ带 ��ٴ�.
			If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> "" Then
				.txtYEARMON.className		= "NOINPUT" : .txtYEARMON.readOnly		= True
				'������
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly			= True
				.txtCARD_AMT.className		= "NOINPUT_R" : .txtCARD_AMT.readOnly		= True
				.cmbMED_FLAG.disabled = True
				
				'��ü��
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				
				'��ü
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True

				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True
				.txtEX_CARD.className		= "NOINPUT_R" : .txtEX_CARD.readOnly		= True
				.txtDEMANDDAY.className		= "NOINPUT" : .txtDEMANDDAY.readOnly		= True : .imgCalEndar.disabled = True
				
				'���۴����
				.txtEXCLIENTNAME.className = "NOINPUT_L" : .txtEXCLIENTNAME.readOnly	= True : .ImgEXCLIENTCODE.disabled = True
				.txtEXCLIENTCODE.className = "NOINPUT_L" : .txtEXCLIENTCODE.readOnly	= True
				
				'���μ�
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .ImgDEPT_CD.disabled = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				
				.txtOUT_AMT.className		= "NOINPUT_R" : .txtOUT_AMT.readOnly		= True
				.txtEX_AMT.className		= "NOINPUT_R" : .txtEX_AMT.readOnly			= True
				
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True

			else 
				.txtYEARMON.className		= "INPUT" : .txtYEARMON.readOnly			= False
				'������
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly		= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly		= False
				
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly				= False
				.txtCARD_AMT.className		= "INPUT_R" : .txtCARD_AMT.readOnly			= False
				.cmbMED_FLAG.disabled = False
				
				'��ü��
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly	= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly	= False
				
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly			= False : .ImgMEDCODE.disabled = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly			= False
				
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly		= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly		= False
				.txtEX_CARD.className		= "INPUT_R" : .txtEX_CARD.readOnly			= False
				.txtDEMANDDAY.className		= "INPUT" : .txtDEMANDDAY.readOnly		= False	: .imgCalEndar.disabled = false

				'���۴����
				.txtEXCLIENTNAME.className = "INPUT_L" : .txtEXCLIENTNAME.readOnly		= False : .ImgEXCLIENTCODE.disabled = False
				.txtEXCLIENTCODE.className = "INPUT_L" : .txtEXCLIENTCODE.readOnly		= False
				
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly		= False : .ImgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly			= False
				
				.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly			= False
				.txtEX_AMT.className		= "INPUT_R" : .txtEX_AMT.readOnly			= False
				
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly				= False
			End If
		else
			.txtYEARMON.className		= "INPUT" : .txtYEARMON.readOnly				= False
			'������
			.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly			= False : .ImgCLIENTCODE.disabled = False
			.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly			= False
			
			.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly					= False
			.txtCARD_AMT.className		= "INPUT_R" : .txtCARD_AMT.readOnly				= False
			.cmbMED_FLAG.disabled = False
			
			'��ü��
			.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly		= False : .ImgREAL_MED_CODE.disabled = False
			.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly		= False
			
			'��ü
			.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly				= False : .ImgMEDCODE.disabled = False
			.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly				= False
			
			.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly			= False
			.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly			= False
			.txtEX_CARD.className		= "INPUT_R" : .txtEX_CARD.readOnly				= False
			.txtDEMANDDAY.className		= "INPUT" : .txtDEMANDDAY.readOnly			= False : .imgCalEndar.disabled = false
							
			'���۴����
			.txtEXCLIENTNAME.className = "INPUT_L" : .txtEXCLIENTNAME.readOnly			= False : .ImgEXCLIENTCODE.disabled = False
			.txtEXCLIENTCODE.className = "INPUT_L" : .txtEXCLIENTCODE.readOnly			= False
		
			'���μ�
			.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly			= False : .ImgDEPT_CD.disabled = False
			.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly				= False
							
			.txtOUT_AMT.className		= "INPUT_R" : .txtOUT_AMT.readOnly				= False
			.txtEX_AMT.className		= "INPUT_R" : .txtEX_AMT.readOnly				= False
			
			.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly					= False

		End If
	End With
End Sub

'���ް��� ���
sub	AMT_CAL (ByVal Col, ByVal Row)
	Dim intAMT
	Dim intVAT
	Dim intSUMAMTVAT
	with frmThis

		intAMT = 0 : intVAT = 0 : intSUMAMTVAT = 0

		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intVAT = intAMT * 0.1
		intSUMAMTVAT = intAMT + intVAT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VAT",frmThis.sprSht.ActiveRow, intVAT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",frmThis.sprSht.ActiveRow, intSUMAMTVAT
		
		COMMISSION_CAL Col, Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'���� ������ ���
SUB COMMISSION_CAL (ByVal Col, ByVal Row)
	Dim intAMT
	Dim intCOMMI_RATE
	Dim intCOMMISSION
	with frmThis

		intAMT = 0 : intCOMMI_RATE = 0 : intCOMMISSION = 0
		
		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		intCOMMISSION = round(intAMT * (intCOMMI_RATE / 100),0)
		
		.txtCOMMISSION.value = intCOMMISSION
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, intCOMMISSION

		CARD_CAL Col, Row

	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
END SUB

'ī�� ������ ���(1.32%)
sub CARD_CAL (ByVal Col, ByVal Row)
	Dim intAMT			'���ް���
	Dim intCARD_AMT		'ī�� ������
	
	with frmThis
		intAMT = 0 : intCARD_AMT = 0 
		
		intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		intCARD_AMT = round(intAMT * 0.0132,0)
		
		.txtCARD_AMT.value = intCARD_AMT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CARD_AMT",frmThis.sprSht.ActiveRow, intCARD_AMT
		
		EXCARD_CAL Col,Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'ī������� ���� ���
sub EXCARD_CAL (ByVal Col, ByVal Row)
	Dim intCARD_AMT		'ī�� ������
	Dim intCOMMISSION	'���������
	Dim intEX_CARD		'ī�� ���������ܱݾ�
	
	with frmthis
		intCARD_AMT = 0 : intCOMMISSION = 0 : intEX_CARD = 0

		intCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		intCARD_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CARD_AMT",Row)
		
		intEX_CARD = intCOMMISSION - intCARD_AMT
		.txtEX_CARD.value = intEX_CARD
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CARD",frmThis.sprSht.ActiveRow, intEX_CARD
		
		EX_CAL Col,Row
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'��ü ����� ���� ���
sub EX_CAL (ByVal Col, ByVal Row)
	Dim intEX_CARD		'ī�� ���������ܱݾ�
	Dim intOUT_AMT		'��ü�� Ȯ���ݾ�
	Dim intEX_AMT		'��ü ����� Ȯ���ݾ�
	
	with frmthis
		intEX_CARD = 0 : intOUT_AMT = 0 : intEX_AMT = 0

		
		intEX_CARD = mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CARD",Row)
		
		intOUT_AMT = clng(intEX_CARD) * 0.3
		intEX_AMT = clng(intEX_CARD) * 0.7
		
		.txtOUT_AMT.value = intOUT_AMT
		.txtEX_AMT.value = intEX_AMT
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUT_AMT",frmThis.sprSht.ActiveRow, intOUT_AMT
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_AMT",frmThis.sprSht.ActiveRow, intEX_AMT
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub


'========================================================================================
' UI���� ���ν��� 
'========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjMDCOAORMEDIUM	= gCreateRemoteObject("cMDSC.ccMDSCAORMEDIUM")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	mobjSCGLCtl.DoEventQueue
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 28, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | MED_FLAG | CLIENTCODE | CLIENTNAME| DEMANDDAY | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | EX_AMT | MEMO | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|����|��ü����|�������ڵ�|�����ָ�|û����|���ް���|VAT|VAT���Աݾ�|��������|������|ī�������(1.32%)|ī�����������|��ü�ڵ�|��ü��|��ü���ڵ�|��ü���|��ü��Ȯ���ݾ�|���۴�����ڵ�|���۴�����|���μ��ڵ�|���μ���|���۴����Ȯ���ݾ�|���|�ŷ�������ȣ|���ݰ�꼭��ȣ|��ǥ��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   8|   4|      10|         0|      14|    10|      14| 12|         14|       6|    14|				 14|            14|		  0|    10|     	0|      12|			   14|			   0|          14|           0|        12|                14|  15|             0|             0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "MED_FLAG", -1, -1, "������" & vbTab & "���̺�" & vbTab & "���������" & vbTab & "�Ź�" & vbTab & "����" & vbTab & "���ͳ�" & vbTab & "����" , 10, 90, False, False
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME| REAL_MED_CODE | REAL_MED_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT | VAT | SUMAMTVAT | COMMISSION | CARD_AMT | EX_CARD | OUT_AMT | EX_AMT ", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "SEQ | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE | MEDCODE | REAL_MED_CODE | EXCLIENTCODE | DEPT_CD | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | YEARMON | CLIENTCODE | CLIENTNAME| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | DEMANDDAY | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD ",-1,-1,2,2,False
		.sprSht.style.visibility = "visible"
    End With
	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCOAORMEDIUM = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	'�ʱ� ������ ����
	With frmThis
		.sprSht.MaxRows = 0
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		
		.txtDEMANDDAY.value  = gNowDate2
		.cmbMED_FLAG.value = "A"
		
	End With
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim vntData2
	Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
   	Dim strRows
	Dim intCnt, intCnt2

	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		intCnt2 = 1

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strYEARMON = "" : strCLIENTCODE = "" : strCLIENTNAME = "" :	strREAL_MED_CODE = "" :	strREAL_MED_NAME = "" 

		strYEARMON		 = .txtYEARMON1.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCLIENTNAME	 = .txtCLIENTNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value

		vntData = mobjMDCOAORMEDIUM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, _
											  strCLIENTCODE, strCLIENTNAME, _
											  strREAL_MED_CODE, strREAL_MED_NAME)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)

   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
	   			For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> "" Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next

				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,1,27,True
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				sprShtToFieldBinding 2, 1
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   			End If
   			Field_Lock
   		End If
   	end With
End Sub

Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strDataCHK
	Dim lngCol, lngRow

	With frmThis
   		'������ Validation
		'If DataValidation =False Then exit Sub
		'On error resume Next
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "YEARMON | MED_FLAG | MEDCODE | REAL_MED_CODE | EXCLIENTCODE | DEPT_CD",lngCol, lngRow, False) 
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ���/��ü����/��ü/��ü��/���۴����/���μ� (��)�� �ʼ� �Է� ���� �Դϴ�..","����ȳ�"
			Exit Sub
		End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | MED_FLAG | CLIENTCODE | CLIENTNAME| DEMANDDAY | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | CARD_AMT | EX_CARD | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | EX_AMT | MEMO | COMMI_TRANS_NO | COMMI_TAX_NO | COMMI_VOCH_NO")
		intRtn = mobjMDCOAORMEDIUM.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox intRtn &" ���� �ڷᰡ ����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   	end With
End Sub

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim lngchkCnt

	lngchkCnt = 0
	With frmThis
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
			End If
		Next

		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If

		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)

				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDCOAORMEDIUM.DeleteRtn(gstrConfigXml,strYEARMON,dblSEQ)
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
				End If
   				intCnt = intCnt + 1
   			End If
		Next

		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox intCnt & "���� �ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
	End With
	err.clear
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">AOR �������</td>
										</tr>
									</table>
								</TD>
								<TD height="20" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table Start-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblKey" class="SEARCHDATA" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
									width="50">���</TD>
								<TD class="SEARCHDATA" width="100"><INPUT accessKey="NUM" style="WIDTH: 96px; HEIGHT: 22px" id="txtYEARMON1" class="INPUT"
										title="�����ȸ" maxLength="6" size="10" name="txtYEARMON1"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">������</TD>
								<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 173px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="�ڵ��"
										maxLength="100" align="left" size="22" name="txtCLIENTNAME1"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
									<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT_L" title="�ڵ���ȸ"
										maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD style="WIDTH: 45px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)">��ü��</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 173px; HEIGHT: 22px" id="txtREAL_MED_NAME1" class="INPUT_L" title="��ü���"
										maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG style="CURSOR: hand" id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgREAL_MED_CODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
									<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtREAL_MED_CODE1" class="INPUT_L" title="��ü���ڵ�"
										maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" align="right"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF"
										height="20"></TD>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 20px" class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="500" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
										<tr>
											<td class="TITLE" vAlign="middle"><span style="CURSOR: hand" id="spnHIDDEN" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG style="CURSOR: hand" id="imgTableUp" border="0" name="imgTableUp" alt="�ڷḦ �˻��մϴ�."
														align="absMiddle" src="../../../images/imgTableUp.gif"></span> &nbsp;&nbsp;&nbsp;&nbsp;���� 
												�հ� : <INPUT style="WIDTH: 120px; HEIGHT: 22px" id="txtSELECTAMT" class="NOINPUTB_R" title="���ñݾ�"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="top" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
										<TR>
											<TD><IMG style="CURSOR: hand" id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgCho"
													alt="�ڷḦ �ʱ�ȭ." src="../../../images/imgCho.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" border="0" name="imgREG"
													alt="�ű��ڷḦ �����մϴ�.." src="../../../images/imgNew.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" border="0" name="imgSave"
													alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
													alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 120px" vAlign="top" align="center">
									<TABLE id="tblHidden" class="DATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD class="LABEL" width="70">���</TD>
											<TD style="WIDTH: 150px" class="DATA"><INPUT accessKey="NUM" style="WIDTH: 118px; HEIGHT: 22px" id="txtYEARMON" dataSrc="#xmlBind"
													class="INPUT" title="���" dataFld="YEARMON" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" maxLength="6" size="13"
													name="txtYEARMON"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="60">������</TD>
											<TD style="WIDTH: 200px" class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="�����ָ�" dataFld="CLIENTNAME" maxLength="100" size="33" name="txtCLIENTNAME">&nbsp;<IMG style="CURSOR: hand" id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">&nbsp;<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="�������ڵ�" dataFld="CLIENTCODE" maxLength="10" size="4" name="txtCLIENTCODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtAMT, '')"
												width="70">���ް���</TD>
											<TD style="WIDTH: 200px" class="DATA"><INPUT accessKey="NUM" style="WIDTH: 196px; HEIGHT: 22px" id="txtAMT" dataSrc="#xmlBind"
													class="INPUT_R" title="���ް���" dataFld="AMT" maxLength="13" size="17" name="txtAMT">
											</TD>
											<TD class="LABEL" width="70">ī�������</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtCARD_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="ī�������ݾ�" dataFld="CARD_AMT" maxLength="13" size="17" name="txtCARD_AMT">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL">��ü����</TD>
											<TD style="WIDTH: 148px" class="DATA"><SELECT style="WIDTH: 112px" id="cmbMED_FLAG" dataSrc="#xmlBind" title="��ü����" dataFld="MED_FLAG"
													name="cmbMED_FLAG">
													<OPTION selected value="A">������</OPTION>
													<OPTION value="A2">���̺�</OPTION>
													<OPTION value="T">���������</OPTION>
													<OPTION value="B">�Ź�</OPTION>
													<OPTION value="C">����</OPTION>
													<OPTION value="O">���ͳ�</OPTION>
													<OPTION value="D">����</OPTION>
												</SELECT>
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)">��ü��</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtREAL_MED_NAME" dataSrc="#xmlBind" class="INPUT_L"
													title="��ü���" dataFld="REAL_MED_NAME" maxLength="100" size="32" name="txtREAL_MED_NAME">&nbsp;<IMG style="CURSOR: hand" id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgREAL_MED_CODE" align="absMiddle" src="../../../images/imgPopup.gIF">&nbsp;<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtREAL_MED_CODE" dataSrc="#xmlBind" class="INPUT_L"
													title="��ü���ڵ�" dataFld="REAL_MED_CODE" maxLength="10" size="4" name="txtREAL_MED_CODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtCOMMISSION, '')">������</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 123px; HEIGHT: 22px" id="txtCOMMISSION" dataSrc="#xmlBind"
													class="INPUT_R" title="������" dataFld="COMMISSION" maxLength="13" size="17" name="txtCOMMISSION">
												<INPUT style="WIDTH: 60px; HEIGHT: 22px" id="txtCOMMI_RATE" dataSrc="#xmlBind" class="INPUT_R"
													title="��������" dataFld="COMMI_RATE" maxLength="6" size="5" name="txtCOMMI_RATE">%
											</TD>
											<TD class="LABEL">ī������</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtEX_CARD" dataSrc="#xmlBind"
													class="INPUT_R" title="ī����������ܱݾ�" dataFld="EX_CARD" maxLength="13" size="17" name="txtEX_CARD">
											</TD>
										</TR>
										<tr>
											<TD class="LABEL">û����</TD>
											<TD class="DATA"><INPUT accessKey="DATE,M" style="WIDTH: 120px; HEIGHT: 22px" id="txtDEMANDDAY" dataSrc="#xmlBind"
													class="INPUT" title="û����" dataFld="DEMANDDAY" maxLength="10" size="14" name="txtDEMANDDAY">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16">
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtMEDNAME, txtMEDCODE)">��ü��</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtMEDNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="��ü��" dataFld="MEDNAME" maxLength="100" size="13" name="txtMEDNAME"> <IMG style="CURSOR: hand" id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgMEDCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtMEDCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="��ü���ڵ�" dataFld="MEDCODE" maxLength="6" size="2" name="txtMEDCODE"></TD>
											<TD class="LABEL">��üȮ����</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 196px; HEIGHT: 22px" id="txtOUT_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="��ü��Ȯ���ݾ�" dataFld="OUT_AMT" maxLength="13" size="17" name="txtOUT_AMT">
											</TD>
											<TD class="LABEL">����Ȯ����</TD>
											<TD class="DATA"><INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtEX_AMT" dataSrc="#xmlBind"
													class="INPUT_R" title="���۴���Ȯ���ݾ�" dataFld="EX_AMT" maxLength="13" size="17" name="txtEX_AMT">
											</TD>
										</tr>
										<tr>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtDEPT_NAME, txtDEPT_CD)">���μ�</TD>
											<TD class="DATA"><INPUT style="WIDTH: 75px; HEIGHT: 22px" id="txtDEPT_NAME" dataSrc="#xmlBind" class="INPUT_L"
													title="���μ���" dataFld="DEPT_NAME" maxLength="100" size="6" name="txtDEPT_NAME">
												<IMG style="CURSOR: hand" id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="imgDEPT_CD"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtDEPT_CD" dataSrc="#xmlBind"
													class="INPUT_L" title="���μ��ڵ�" dataFld="DEPT_CD" maxLength="6" size="3" name="txtDEPT_CD"></TD>
											<TD style="CURSOR: hand; HEIGHT: 22px" class="LABEL" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)">���۴���</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtEXCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="���ۻ��" dataFld="EXCLIENTNAME" maxLength="100" size="30" name="txtEXCLIENTNAME">
												<IMG style="CURSOR: hand" id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgEXCLIENTCODE"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtEXCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="���ۻ��ڵ�" dataFld="EXCLIENTCODE" maxLength="10" size="4" name="txtEXCLIENTCODE"></TD>
											<TD class="LABEL">�޸�</TD>
											<TD class="DATA" colSpan="4"><INPUT style="WIDTH: 397px; HEIGHT: 22px" id="txtMEMO" dataSrc="#xmlBind" class="INPUT_R"
													title="���" dataFld="MEMO" maxLength="255" size="17" name="txtMEMO"></TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="BODYSPLIT"></TD>
							</TR>
							<!--BodySplit End--></TABLE>
						<TABLE id="tblSheet" border="0" cellSpacing="0" cellPadding="0" width="100%" height="65%">
							<TR>
								<td style="WIDTH: 100%; HEIGHT: 100%" class="DATA" vAlign="top" align="center">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31882">
										<PARAM NAME="_ExtentY" VALUE="13520">
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
										<PARAM NAME="MaxCols" VALUE="44">
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
								</td>
							</TR>
							<TR>
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
