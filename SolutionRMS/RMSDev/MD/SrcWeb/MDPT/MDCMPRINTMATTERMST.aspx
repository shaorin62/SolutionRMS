<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTMATTERMST.aspx.vb" Inherits="MD.MDCMPRINTMATTERMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�����ڵ� ����</title>
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
Dim mobjMDCOMATTERMST
Dim mobjMDCMGET
Dim mlngRowCnt,mlngColCnt
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
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
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

Sub imgDelete_onclick ()
	Dim i
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'����ó��
'-----------------------------
Sub imgConf_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn_ConfOK
	gFlowWait meWAIT_OFF
End Sub

sub imgNewReg_onclick ()
	With frmThis
		Call sprSht_Keydown(meINS_ROW, 0)
	End With 
End sub


Sub chkMC_onclick()
	With frmThis
		.txtAPPNAME.value = "M&C (���� ����)"
		.txtAPPCODE.value = "K00006"
	End With 
End SUb
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
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			imgQuery_onclick
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					imgQuery_onclick
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ����ȸ�˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgTIMCODE_onclick
	Call TIMCODE_POP()
End Sub

'���� ������List ��������
Sub TIMCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value), _
							trim(.txtTIMCODE.value), trim(.txtTIMNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,445)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' �ڵ�� ǥ��
			imgQuery_onclick
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
											trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					imgQuery_onclick
				Else
					Call TIMCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' �귣���ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,445)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtTIMNAME.value = trim(vntRet(5,0))	' �����ָ� ǥ��

			imgQuery_onclick
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
												trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))		' ������ ǥ��
					.txtCLIENTNAME.value = trim(vntData(3,1))	' ������
					.txtTIMCODE.value = trim(vntData(4,1))	' ������
					.txtTIMNAME.value = trim(vntData(5,1))	' ������
					imgQuery_onclick
				Else
					Call SUBSEQCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ����� �ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value), "") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtEXCLIENTCODE.value = trim(vntRet(1,0))  ' Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))  ' �ڵ�� ǥ��
			imgQuery_onclick
     	End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value), "")
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))
					.txtEXCLIENTNAME.value = trim(vntData(2,1))
					imgQuery_onclick
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' �ϰ������� ���ۻ� �˾�
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgAPP_onclick
	Call APP_POP()
End Sub

'���� ������List ��������
Sub APP_POP
	Dim vntRet
	Dim vntInParams
	Dim strFLAG

	With frmThis
	
		If .chkAOR.checked = True Then
			strFLAG = "G"
		ElseIf .chkCRE.checked = True Then
			strFLAG = "K"
		ElseIf .chkMC.checked = True Then
			strFLAG = "C"
		ELSE
			strFLAG = ""
		End If
		
		vntInParams = array(trim(.txtAPPCODE.value), trim(.txtAPPNAME.value), strFLAG) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtAPPCODE.value = vntRet(1,0) and .txtAPPNAME.value = vntRet(2,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtAPPCODE.value = trim(vntRet(1,0))  ' Code�� ����
			.txtAPPNAME.value = trim(vntRet(2,0))  ' �ڵ�� ǥ��
     	End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtAPPNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
   		Dim strFLAG
		On error resume next
		With frmThis
		
			If .chkAOR.checked = True Then
				strFLAG = "G"
			ElseIf .chkCRE.checked = True Then
				strFLAG = "K"
			ElseIf .chkMC.checked = True Then
				strFLAG = "C"
			ELSE
				strFLAG = ""
			End If
			
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtAPPCODE.value),trim(.txtAPPNAME.value), strFLAG)
			
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtAPPCODE.value = trim(vntData(1,1))
					.txtAPPNAME.value = trim(vntData(2,1))
					'.txtMEDNAME.focus()
				Else
					Call APP_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub


'�ϰ����� ��ư Ŭ��
Sub ImgBundleApp_onclick
	Dim i
	Dim strCNT
	
	With frmThis
		strCNT = 0
		If .txtAPPCODE.value = "" Or .txtAPPNAME.value = "" Then
			gErrorMsgBox "�ϰ������� ���ۻ縦 �Է��ϼ���.","�ϰ�����ȳ�"
			Exit Sub
		End If
		
		For i=1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",i, .txtAPPCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",i, .txtAPPNAME.value
				mobjSCGLSpr.CellChanged frmThis.sprSht, 13, i
				strCNT = strCNT + 1
			End If
		Next
		
		If strCNT = 0 Then
			gErrorMsgBox "�ϰ������� ���� �����ϼ���.","�ϰ�����ȳ�"
			Exit Sub
		End If
	End With
End Sub

'-----------------------------------
' SpreadSheet sprSht �̺�Ʈ
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,false,frmThis.sprSht.ActiveRow,3,3,true
		
		frmThis.txtCLIENTNAME.focus
		frmThis.sprSht.focus
		
	End If
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	With frmThis
		If Row = 0 and Col = 1 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			ElseIf mstrCheck = False Then 
				mstrCheck = True
			End If
			For intcnt = 1 To .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' �������� ��ưŬ�� �̺�Ʈ
'-----------------------------------------------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUBSEQ") Then '�귣��
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_TIMCODE.aspx",vntInParams , 640,445)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(3,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntData(4,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntData(5,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntData(6,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(7,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(8,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(9,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntData(10,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(11,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col+1,Row
			End If
			.txtMATTERNAME.focus
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		ElseIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNEX") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			.txtMATTERNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	Dim intColor
   	intColor = ""
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQ") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If strCode = "" AND strCodeName = "" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, ""
   			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCMGET.Get_BrandInfo_TIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)),  _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))

				If not gDoErrorRtn ("Get_BrandInfo_TIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntData(3,1)
						'mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntData(4,0)
						'mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntData(5,0)
						'mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntData(6,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntData(7,1)
						'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(8,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntData(9,1)
						'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntData(10,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntData(11,1)
						
						.txtMATTERNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME"), Row
						.txtMATTERNAME.focus
						.sprSht.focus 
					End If
   				End If
   			ElseIf strCode = "" AND strCodeName = "" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, ""
   			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			strCode		= mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCMGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(2,1)			
						.txtMATTERNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME"), Row
						.txtMATTERNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet, vntInParams
	Dim strGUBUN
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_TIMCODE.aspx",vntInParams , 640,445)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(3,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntData(4,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntData(5,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntData(6,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(7,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(8,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(9,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntData(10,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(11,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col+1,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtMATTERNAME.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
		
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtMATTERNAME.focus
		.sprSht.Focus
	End With
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	'����������ü ����	
	set mobjMDCOMATTERMST	= gCreateRemoteObject("cMDCO.ccMDCOMATTERMST")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "137px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'��ȸ �� ���� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 18, 0, 0, 2
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | ATTR02 | MATTERCODE | MATTER | SUBSEQ | BTNSUBSEQ | SUBSEQNAME | DEPT_NAME | TIMNAME | CLIENTSUBNAME | CUSTCODE | CUSTNAME | EXCLIENTCODE | BTNEX | EXCLIENTNAME | MEMO | UUSER | UDATE"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|��뱸��|�ڵ�|�����|�귣���ڵ�|�귣���|�귣����μ�|��|�����|�������ڵ�|������|���ۻ��ڵ�|���ۻ��|���|����������|����������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|      10|   7|    15|         9|2|    13|            10|13|	   0|         9|    18|         9|2|    13|  15|         7|       12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "ATTR02", -1, -1, "���" & vbTab & "�̻��" & vbTab & "���" & vbTab & "���ο�û" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTNSUBSEQ | BTNEX"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUBSEQNAME | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | MEMO | UUSER | UDATE", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MATTERCODE | DEPT_NAME | TIMNAME | CLIENTSUBNAME | CUSTCODE | CUSTNAME | UUSER | UDATE" 
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SUBSEQ | MATTERCODE | CUSTCODE | EXCLIENTCODE",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MATTER | SUBSEQNAME | TIMNAME | CLIENTSUBNAME | CUSTNAME | EXCLIENTNAME | MEMO",-1,-1,0,2, false
		
    End With    
	pnlTab1.style.visibility = "visible"
	

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub


'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	With frmThis
		.sprSht.MaxRows = 0	
		
	End With
End Sub


Sub EndPage()
	set mobjMDCMMATTERMST = Nothing
	Set mobjMDCMGET = Nothing
	gEndPage	
End Sub



Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTNAME, strCLIENTCODE
   	Dim strTIMNAME, strTIMCODE
   	Dim strSUBSEQNAME, strSUBSEQ
   	Dim strEXCLIENTNAME, strEXCLIENTCODE
   	Dim strMATTERNAME, strMATTERCODE
   	Dim strUSE_YN
   	Dim intCnt, intCnt2, strRows
   	Dim dblcnt
    
	'On error resume next
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		dblcnt = true

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strTIMNAME		= .txtTIMNAME.value
		strTIMCODE		= .txtTIMCODE.value
		strSUBSEQNAME	= .txtSUBSEQNAME.value
		strSUBSEQ		= .txtSUBSEQ.value
		strEXCLIENTNAME	= .txtEXCLIENTNAME.value
		strEXCLIENTCODE	= .txtEXCLIENTCODE.value
		strMATTERNAME	= .txtMATTERNAME.value
		strUSE_YN		= .cmbUSE_YNSEARCH.value
		
		vntData = mobjMDCOMATTERMST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												strCLIENTNAME, strCLIENTCODE, _
												strTIMNAME, strTIMCODE, _
												strSUBSEQNAME, strSUBSEQ, _
												strEXCLIENTNAME, strEXCLIENTCODE, _
												strMATTERNAME, _
												"B", strUSE_YN)
			
		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt > 0 Then
				Call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
				
				FOR i =1 TO .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR02",i) = "���" or mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR02",i) = "�̻��" Then
						If dblcnt Then
							strRows = i
							dblcnt = false
						Else
							strRows = strRows & "|" & i
						End If
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,3,18,True
				mobjSCGLSpr.SetCellsLock2 .sprSht, false, "EXCLIENTCODE | BTNEX | EXCLIENTNAME"
				
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			End If
   		End If
   	End With
End Sub

Function DataValidation ()
	DataValidation = false	
	With frmThis
		'If not gDataValidation(frmThis) Then exit Function	
	End With
	DataValidation = True
End Function

'�������
Sub ProcessRtn()
	Dim intRtn
	Dim lngCol, lngRow
   	Dim vntData, vntData_Src
   	Dim strYEAR
   	Dim strMEDFLAG
   	Dim strDataCHK
   	
	With frmThis
		For intCnt = 1 to .sprSht.MaxRows
			If Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"MATTER",intCnt)) = ""  Then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End If
		Next
	
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "MATTER | SUBSEQ | SUBSEQNAME | CUSTCODE | CUSTNAME | EXCLIENTCODE | EXCLIENTNAME",lngCol, lngRow, False) 
		
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� �����/�귣��/������/���ۻ�� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | MATTERCODE | MATTER | SUBSEQ | BTNSUBSEQ | SUBSEQNAME | TIMNAME | CLIENTSUBNAME | CUSTCODE | CUSTNAME | EXCLIENTCODE | BTNEX | EXCLIENTNAME | MEMO | ATTR02")
		
		If Not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"	
			Exit Sub 
		End If
		
		strYEAR = Mid(gNowDate,3,2)
		strMEDFLAG = "B"
		
		'ó�� ������ü ȣ��
		intRtn = mobjMDCOMATTERMST.ProcessRtn(gstrConfigXml,vntData, strYEAR, strMEDFLAG)
		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox intRtn & " �� �� ����Ǿ����ϴ�.","����ȳ�"
			SelectRtn
   		End If
   	End With
End Sub

Sub ProcessRtn_CONFOK ()
	Dim vntData, vntData2
	Dim intCnt, intRtn, i
	Dim lngCol, lngRow
	Dim strMATTERCODE
	Dim intCnt2
	Dim strDataCHK
	Dim strYEAR, strMEDFLAG
	Dim strchk
	
	strchk = true
	
	With frmThis
		
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR02",intCnt2) <> "���ο�û" then
					gErrorMsgBox "üũ�� ������ �� " +  i + " ��° ���� ���´� ���ο�û���°� �ƴմϴ�. ���ο�û������ �����͸� ������ �� �ֽ��ϴ�.","���ξȳ�!"
					Exit Sub
				end if 
				strchk = false
			end if
		Next
		
		if strchk then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���","���ξȳ�!"
			exit sub
		end if
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "MATTER | SUBSEQ | SUBSEQNAME | CUSTCODE | CUSTNAME | EXCLIENTCODE | EXCLIENTNAME",lngCol, lngRow, False) 
		
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� �����/�귣��/������/���ۻ�� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If
		
		intRtn = gYesNoMsgbox("�ڷḦ ���� �Ͻðڽ��ϱ�?","����Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData2 = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | MATTERCODE | MATTER | SUBSEQ | BTNSUBSEQ | SUBSEQNAME | TIMNAME | CLIENTSUBNAME | CUSTCODE | CUSTNAME | EXCLIENTCODE | BTNEX | EXCLIENTNAME | MEMO | ATTR02")
		
		If Not IsArray(vntData2) Then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"	
			Exit Sub 
		End If
		
		strYEAR = Mid(gNowDate,3,2)
		strMEDFLAG = "B"
		
		'ó�� ������ü ȣ��
		intRtn = mobjMDCOMATTERMST.ProcessRtn(gstrConfigXml,vntData2, strYEAR, strMEDFLAG)
		MsgBox .sprSht.MaxRows 
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows  to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strMATTERCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",i)
				
				If strMATTERCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDCOMATTERMST.ProcessRtn_CONFOK(gstrConfigXml,strMATTERCODE)
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("ProcessRtn_CONF") Then
			gErrorMsgBox "�ڷᰡ ���� �Ǿ����ϴ�.","���ξȳ�!"
   		End If

		SelectRtn
	End With
	err.clear	
End Sub

'------------------------------------------
'������ ����
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strMATTERCODE
	Dim	strMATTERCODE2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strMATTERCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"MATTERCODE",i)
				If strMATTERCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht, i
				Else
					vntData = mobjMDCOMATTERMST.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strMATTERCODE) 
					If mlngRowCnt > 0 Then
						strMSG = ""
						For intCnt = 0 To mlngRowCnt-1
							If vntData(0,intCnt) = "B" Then
								strMSG = strMSG & " �μ�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A2" Then
								strMSG = strMSG & " ���̺�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A" Then
								strMSG = strMSG & " ������: " & vntData(1,intCnt) & "��" 
							End If
						Next
						gErrorMsgBox i & "���� �ڵ�� " & strMSG & " �� ����� ����Ǿ��ֽ��ϴ�.","�����ȳ�!"
						Exit Sub
					End If
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
	
		'���õ� �ڷḦ ������ ���� ����
		For i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strMATTERCODE2 = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",i)
				If strMATTERCODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					intRtn = mobjMDCOMATTERMST.DeleteRtn(gstrConfigXml, strMATTERCODE2)
					
					IF not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn") Then
   			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		SelectRtn
	End With
	err.clear
End Sub
-->

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
											<td class="TITLE">�μ�-�������� ����</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
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
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" height="95%" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="70">������</TD>
											<TD class="SEARCHDATA" width="270"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 192px; HEIGHT: 22px"
													type="text" maxLength="100" size="26" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME,txtTIMCODE)"
												width="70">��</TD>
											<TD class="SEARCHDATA" width="270"><INPUT class="INPUT_L" id="txtTIMNAME" title="����" style="WIDTH: 192px; HEIGHT: 22px" type="text"
													maxLength="100" size="20" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgTIMCODE"> <INPUT class="INPUT_L" id="txtTIMCODE" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" size="6" name="txtTIMCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME,txtSUBSEQ)"
												width="70">�귣��</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 184px; HEIGHT: 22px"
													type="text" maxLength="100" size="25" name="txtSUBSEQNAME"> <IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgSUBSEQCODE"> <INPUT class="INPUT_L" id="txtSUBSEQ" title="�귣���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" name="txtSUBSEQ"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="70">���۴����</TD>
											<TD class="SEARCHDATA" width="270"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="�������" style="WIDTH: 192px; HEIGHT: 22px"
													type="text" maxLength="100" size="20" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgEXCLIENTCODE"> <INPUT class="INPUT_L" id="txtEXCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" size="6" name="txtEXCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME,'')"
												width="70">�����</TD>
											<TD class="SEARCHDATA" width="270"><INPUT class="INPUT_L" id="txtMATTERNAME" title="�����" style="WIDTH: 264px; HEIGHT: 22px"
													type="text" maxLength="200" size="38" name="txtMATTERNAME"></TD>
											<TD class="SEARCHLABEL">��뱸��</TD>
											<TD class="SEARCHDATA">
												<TABLE id="tblButton" width="100%" cellSpacing="0" cellPadding="2" align="left" border="0">
													<TR>
														<TD class="SEARCHDATA" align="left"><SELECT id="cmbUSE_YNSEARCH" title="��뱸��" style="WIDTH: 100px" name="cmbUSE_YNSEARCH">
																<OPTION value="">��ü</OPTION>
																<OPTION value="Y">���</OPTION>
																<OPTION value="N">�̻��</OPTION>
																<OPTION value="R">���</OPTION>
																<OPTION value="S" selected>���ο�û</OPTION>
															</SELECT>
														</TD>
														<TD align="right" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ ��ȸ�մϴ�."
																src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 15px"></TD>
							</TR>
							<!--Input Start-->
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAPPNAME,txtAPPCODE)"
												width="70">����</TD>
											<TD class="DATA" width="270">&nbsp;&nbsp; <INPUT id="chkAOR" type="radio" CHECKED value="AOR" name="chkGUBUN">AOR&nbsp;&nbsp; 
												&nbsp; <INPUT id="chkCRE" type="radio" value="CRE" name="chkGUBUN"> ũ������&nbsp;&nbsp;
												<INPUT id="chkMC" type="radio" value="MC" name="chkGUBUN">��ü��(���� ��)
											</TD>
											<TD class="DATA" width="350"><INPUT class="INPUT_L" id="txtAPPNAME" title="�����" style="WIDTH: 192px; HEIGHT: 22px" type="text"
													maxLength="100" size="20" name="txtAPPNAME"> <IMG id="ImgAPP" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgAPP">
												<INPUT class="INPUT_L" id="txtAPPCODE" title="�����ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" size="6" name="txtAPPCODE">&nbsp;<IMG id="ImgBundleApp" onmouseover="JavaScript:this.src='../../../images/ImgBundleAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgBundleApp.gIF'" height="20" alt="�ϰ������մϴ�" src="../../../images/ImgBundleApp.gif"
													align="absMiddle" border="0" name="ImgBundleApp"></TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<!--td><IMG id="imgNewReg" onmouseover="JavaScript:this.src='../../../images/imgNewRegOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgNewReg.gif'"
																height="20" alt="�ű��ڷḦ ����մϴ�." src="../../../images/imgNewReg.gIF" border="0" name="imgNewReg"></td-->
														<TD><IMG id="imgConf" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgAgree.gIF" border="0" name="imgConf"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
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
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--���� �� �׸���-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="15372">
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
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></FORM>
		</TR></TABLE>
	</body>
</HTML>
