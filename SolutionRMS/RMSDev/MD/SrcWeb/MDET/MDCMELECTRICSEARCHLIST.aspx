<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRICSEARCHLIST.aspx.vb" Inherits="MD.MDCMELECTRICSEARCHLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>û�೻�� ����</title>
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
Dim mlngRowCnt1, mlngColCnt1
Dim mobjMDCOGET, mobjMDETELEC_TRAN'�����ڵ�, Ŭ����
Dim mstrCheck
Dim mobjMDCMCODETR
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

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtYEARMON1.value = ""  then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		'mobjSCGLSpr.ExportMerge = true
		'mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSetting_onclick
	Call ProcessRtn_ConfirmOK()
End Sub

Sub ImgConfirmCancel_onclick
	ProcessRtn_ConfirmCancel
End Sub

Sub ImgCodeMapping_onclick
	Dim intRtn
	gFlowWait meWAIT_ON
	intRtn = gYesNoMsgbox("�ϰ��ڵ� ��Ī �۾��� �Ͻðڽ��ϱ�?" & vbcrlf & "�ϰ��۾��� ���Ͽ� �귣�� �ڵ� �� �����ڵ� �۾��� �Ǿ��־�� �մϴ�!","�ڵ��Ī Ȯ��")
	IF intRtn <> vbYes then exit Sub
	
	ProcessRtn_CodeBatch
	
	gFlowWait meWAIT_OFF
End Sub

Sub imgSUBSEQApp_onclick
	gFlowWait meWAIT_ON
	SUBSEQApp
	gFlowWait meWAIT_OFF
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
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub 
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
   		Dim i
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
   		Dim i
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

'�� �˾� ��ư
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'���� ������List ��������
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE1.value = vntRet(0,0) and .txtTIMNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTIMCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtTIMNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(5,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
											trim(.txtTIMCODE1.value),trim(.txtTIMNAME1.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME1.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(5,1))
				Else
					Call TIMCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'��ü �˾� ��ư
Sub ImgMEDCODE1_onclick
	Call MEDCODE1_POP()
End Sub

'���� ������List ��������
Sub MEDCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtMEDCODE1.value), trim(.txtMEDNAME1.value), "MED_PRINT")
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE1.value = vntRet(0,0) and .txtMEDNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtMEDCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtMEDNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtREAL_MED_CODE1.value = trim(vntRet(3,0))       ' �ڵ�� ǥ��
			.txtREAL_MED_NAME1.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			SELECTRTN
			
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMEDNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), _
												trim(.txtMEDCODE1.value),trim(.txtMEDNAME1.value), "MED_PRINT")
			
			If not gDoErrorRtn ("GetMEDGUBNCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE1.value = trim(vntData(0,1))	    ' Code�� ����
					.txtMEDNAME1.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtREAL_MED_CODE1.value = trim(vntData(3,1))
					.txtREAL_MED_NAME1.value = trim(vntData(4,1))
					SELECTRTN
				Else
					Call MEDCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'�귣��
Sub ImgSUBSEQ1_onclick
	Call SUBSEQCODE1_POP()
End Sub

Sub SUBSEQCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ1.value), trim(.txtSUBSEQNAME1.value), trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ1.value = vntRet(0,0) and .txtSUBSEQNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ1.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME1.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE1.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtTIMNAME1.value = trim(vntRet(5,0))	' �����ָ� ǥ��
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ1.value),trim(.txtSUBSEQNAME1.value),  _
												trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ1.value = trim(vntData(0,1))
					.txtSUBSEQNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(2,1))		' ������ ǥ��
					.txtCLIENTNAME1.value = trim(vntData(3,1))	' ������
					.txtTIMCODE1.value = trim(vntData(4,1))	' ������
					.txtTIMNAME1.value = trim(vntData(5,1))	' ������
				Else
					Call SUBSEQCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'����� ��ư �˾� ��ȸ��
Sub ImgMATTERCODE1_onclick
	Call MATTERCODE1_POP()
End Sub

Sub MATTERCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis													
		vntInParams = array(trim(.txtCLIENTNAME1.value), trim(.txtTIMNAME1.value), "" ,"", _
							trim(.txtMATTERNAME1.value), "", "A") '<< �޾ƿ��°�� 
		
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
		
		If isArray(vntRet) Then
			If .txtMATTERCODE1.value = vntRet(0,0) and .txtMATTERNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtMATTERCODE1.value = trim(vntRet(0,0))	' �����ڵ� ǥ��
			.txtMATTERNAME1.value = trim(vntRet(1,0))	' ����� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' �������ڵ� ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE1.value = trim(vntRet(4,0))		' ���ڵ� ǥ��
			.txtTIMNAME1.value = trim(vntRet(5,0))		' ���� ǥ��
			.txtSUBSEQ1.value	  = trim(vntRet(6,0))
			.txtSUBSEQNAME1.value  = trim(vntRet(7,0))
			SELECTRTN
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME1_onkeydown
	Dim vntData
   	Dim i
	
	If window.event.keyCode = meEnter Then
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME1.value),trim(.txtTIMNAME1.value), "","", _
											trim(.txtMATTERNAME1.value),"", "A")
			If not gDoErrorRtn ("GetMATTER") Then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE1.value = trim(vntData(0,1))	' �����ڵ� ǥ��
					.txtMATTERNAME1.value = trim(vntData(1,1))	' ����� ǥ��
					.txtCLIENTCODE1.value = trim(vntData(2,1))	' �������ڵ� ǥ��
					.txtCLIENTNAME1.value = trim(vntData(3,1))	' �����ָ� ǥ��
					.txtTIMCODE1.value	  = trim(vntData(4,1))	' ���ڵ� ǥ��
					.txtTIMNAME1.value	  = trim(vntData(5,1))	' ���� ǥ��
					.txtSUBSEQ1.value	  = trim(vntData(6,1))
					.txtSUBSEQNAME1.value  = trim(vntData(7,1))
				Else
					Call MATTERCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'��Ʈ �̺�Ʈ
sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	Dim vntRet
	with frmThis
		if Row = 0 and Col > 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
		end if
	end with
end sub
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ Ŭ���� 
'-----------------------------------------------------------------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
		if Col = 3 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 3, 3, , , "", , , , , mstrCheck
			
			if mstrCheck = True then 
				mstrCheck = False	
			elseif mstrCheck = False then 
				mstrCheck = True
			end if	
		end if 
	End With
End Sub  



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
		'SelectRtn_DTL frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CNT") OR _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CNT") Then
				strCOLUMN = "CNT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CNT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		Else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	Dim strColFlag
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
		strCOLUMN = ""
		strColFlag = 0
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CNT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
						End If
					Next
					
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
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjMDETELEC_TRAN	= gCreateRemoteObject("cMDET.ccMDETELEC_TRAN")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 4, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | SEQ | CHK | CLIENTNAME | SUBSEQ | SUBSEQNAME | DEPTNAME | CLIENTSUBNAME | TIMNAME | GUBUN | MEDNAME | REAL_MED_NAME | PROGRAM | MATTERCODE | MATTERNAME | EXCLIENTNAME | ROLLSTDATE | ROLLEDDATE | PRICE | CNT | AMT | TRU_TRANS_NO "
		mobjSCGLSpr.SetHeader .sprSht,		     "���|����|����|�����ָ�|�귣���ڵ�|�귣���|�μ���|����θ�|����|����|��ü��|��ü���|����|�����ڵ�|�����|���۴�����|�������|��������|�ܰ�|Ƚ��|�ݾ�|�ŷ�������ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   10|    5|   4|      15|         8|      10|     8|       9|   9|   5|    10|      15|    15|       8|    20|          15|       8|       8|  14|   7|  15|            12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | PRICE | CNT | AMT ", -1, -1, 0'
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ROLLSTDATE | ROLLEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | CLIENTNAME | SUBSEQ | SUBSEQNAME | DEPTNAME | CLIENTSUBNAME | TIMNAME | MEDNAME | REAL_MED_NAME | PROGRAM | MATTERCODE | MATTERNAME | EXCLIENTNAME | TRU_TRANS_NO ", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | SEQ | CLIENTNAME | SUBSEQ | SUBSEQNAME | DEPTNAME | CLIENTSUBNAME | TIMNAME | MEDNAME | REAL_MED_NAME | PROGRAM | MATTERCODE | MATTERNAME | EXCLIENTNAME | ROLLSTDATE | ROLLEDDATE | PRICE | CNT | AMT | TRU_TRANS_NO "
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN | TRU_TRANS_NO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ", true
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
	'�ڵ� �ϰ� ��Ī PorceDure
	Batch_Code
	
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	set mobjMDETELEC_TRAN = Nothing
	set mobjMDCMCODETR = Nothing
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
		.txtYEARMON1.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtCLIENTNAME1.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	
	with frmThis
		if .txtYEARMON1.value = "" Then
			gErrorMsgBox "����������Ͻʽÿ�.","��ȸ�ȳ�!"
			Exit Sub
		End If
		
		.sprSht.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDETELEC_TRAN.SelectRtn_ConfirmList(gstrConfigXml, mlngRowCnt, mlngColCnt,.cmbGBN1.value, _
														.txtYEARMON1.value, _
														.txtCLIENTCODE1.value,    .txtCLIENTNAME1.value, _
														.txtSUBSEQ1.value,        .txtSUBSEQNAME1.value, _
														.txtMATTERCODE1.value,    .txtMATTERNAME1.value, _
														.txtTIMCODE1.value,		.txtTIMNAME1.value, _
														.txtMEDCODE1.value,       .txtMEDNAME1.value, _
														.txtREAL_MED_CODE1.value, .txtREAL_MED_NAME1.value)
			
		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG	
   		end if
		
		gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
		AMT_SUM
   	end with
End Sub

Sub AMT_SUM
	Dim intCnt
	Dim lngAMT
	Dim lngAMTSUM
	
	with frmThis
   			lngAMT = 0
   			lngAMTSUM = 0
	   		
   			For intCnt = 1 to .sprSht.MaxRows
				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
				lngAMTSUM = lngAMTSUM + lngAMT
			Next
			.txtSUMAMT.value = lngAMTSUM
			call gFormatNumber(.txtSUMAMT,0,true)
	End With		
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ε��� �ڵ��Ī�� �����ٸ� �ڵ� ��Ī ���� ���� �� ó��
'-----------------------------------------------------------------------------------------
Sub Batch_Code
	Dim vntData
   	Dim intCnt
   	Dim intRtn
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDETELEC_TRAN.SelectRtn_CodeProc(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value)

		if not gDoErrorRtn ("SelectRtn_CodeProc") then
   			If mlngRowCnt > 0  Then
   				If vntData(0,1) = "N" Then
   					intRtn = gYesNoMsgbox("�ϰ��ڵ� ��Ī �۾��� �Ͻðڽ��ϱ�?" & vbcrlf & "�ϰ��۾��� ���Ͽ� �귣�� �ڵ� �� �����ڵ� �۾��� �Ǿ��־�� �մϴ�!","�ڵ��Ī Ȯ��")
					IF intRtn <> vbYes then exit Sub
					ProcessRtn_CodeBatch
				End IF
   			End If
   		end if
   	end with
End SUb

'------------------------------------------
' �ڵ� �ϰ�ó�� �������
'------------------------------------------
Sub ProcessRtn_CodeBatch
	Dim intRtn
	Dim strYEARMON
	With frmThis
		If .txtYEARMON1.value <> "" Then
			If  Len(.txtYEARMON1.value) = 6  Then
				strYEARMON = .txtYEARMON1.value
			Else 
				gErrorMsgBox "�ش����� �ùٸ��� �����Ͻʽÿ�.��: 200810",""
				Exit Sub
			End IF
		Else
			gErrorMsgBox "��� �� ���� �Ͻʽÿ�",""
			Exit Sub
		End If

		On error resume next
		intRtn = mobjMDETELEC_TRAN.ProcessRtn_CodeBatch(gstrConfigXml,strYEARMON)
		
		if not gDoErrorRtn ("ProcessRtn_CodeBatch") then 
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgbox "������ ����" & mePROC_DONE,"ó���ȳ�!"
			SelectRtn
   		end if
   	end with
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="98%" border="0">
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
											<td class="TITLE">û�೻�� ����</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
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
									<TABLE id="tblButton" style="WIDTH: 80px; HEIGHT: 20px" cellSpacing="0" cellPadding="2"
										width="80" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgClose.gIF" border="0" name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
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
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 66px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')">��&nbsp;&nbsp; 
												��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 78px"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 78px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="7" name="txtYEARMON1">
											</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)">������</TD>
											<TD class="SEARCHDATA" style="WIDTH: 200px"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 123px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="14" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
												<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1">
											</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME1, txtMEDCODE1)">��ü</TD>
											<TD class="SEARCHDATA" style="WIDTH: 200px"><INPUT class="INPUT_L" id="txtMEDNAME1" title="��ü��" style="WIDTH: 123px; HEIGHT: 22px"
													type="text" size="15" name="txtMEDNAME1"> <IMG id="ImgMEDCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgMEDCODE1"> <INPUT class="INPUT_L" id="txtMEDCODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" size="2" name="txtMEDCODE1"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1, txtSUBSEQ1)">�귣��</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="�귣���" style="WIDTH: 140px; HEIGHT: 22px"
													type="text" maxLength="100" size="18" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgSUBSEQ1"> <INPUT class="INPUT_L" id="txtSUBSEQ1" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="8" size="3" name="txtSUBSEQ1"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL">�ŷ�����</TD>
											<TD class="SEARCHDATA">&nbsp;<SELECT id="cmbGBN1" title="���α��м���" style="WIDTH: 72px; HEIGHT: 22px" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
													name="cmbGBN1">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="Y">����</OPTION>
													<OPTION value="N">�̻���</OPTION>
												</SELECT>
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)">����</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 123px; HEIGHT: 22px" type="text"
													maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" size="6" name="txtTIMCODE1">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)">��ü��</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="��ü���" style="WIDTH: 123px; HEIGHT: 22px"
													type="text" maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" name="txtREAL_MED_CODE1">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERCODE1, txtMATTERNAME1)">����</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMATTERNAME1" title="�����" style="WIDTH: 140px; HEIGHT: 22px"
													type="text" maxLength="100" size="30" name="txtMATTERNAME1"> <IMG id="ImgMATTERCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgMATTERCODE1"> <INPUT class="INPUT_L" id="txtMATTERCODE1" title="�����ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="10" size="4" name="txtMATTERCODE1"></TD>
										</TR>
									</TABLE>
									<table class="SEARCHDATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgCodeMapping" onmouseover="JavaScript:this.src='../../../images/ImgCodeMappingOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgCodeMapping.gIF'"
																height="20" alt="����ó���� ����մϴ�." src="../../../images/ImgCodeMapping.gif" border="0"
																name="ImgCodeMapping"></TD>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start-->
				<TR vAlign="top" align="left">
					<!--����-->
					<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31829">
								<PARAM NAME="_ExtentY" VALUE="13176">
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
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TABLE></TR></TABLE></FORM>
	</body>
</HTML>
