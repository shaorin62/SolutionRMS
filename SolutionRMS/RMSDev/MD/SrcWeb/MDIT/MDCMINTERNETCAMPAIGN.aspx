<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNETCAMPAIGN.aspx.vb" Inherits="MD.MDCMINTERNETCAMPAIGN" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>ķ���� ����</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOCUSTEXELIST.aspx
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By KTY
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
Dim mobjMDITINTERNETCAMPAIGN '�����ڵ�, Ŭ����
Dim mobjMDCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'====================================================
' �̺�Ʈ ���ν��� 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'---------------------------------------------------
'-----------------------------------
'��ȸ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'���߰�
'-----------------------------
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtCLIENTNAME1.focus
		.sprSht.focus
	End With 
end sub

'-----------------------------------
' ����   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'����
'-----------------------------------
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
' �ݱ�
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'****************************************************************************************
' �˾� �̺�Ʈ, ķ����, ������
'****************************************************************************************
Sub ImgCAMPAIGN_CODE1_onclick
	Call CAMPAIGN_POP()
End Sub
'���� ������List ��������
Sub CAMPAIGN_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtYEARMON.value, .txtCAMPAIGN_CODE1.value, .txtCAMPAIGN_NAME1.value)
			
		vntRet = gShowModalWindow("MDCMINTERNETCAMPAIGNPOP.aspx",vntInParams , 520,525)
			
		if isArray(vntRet) then
			.txtYEARMON.value = vntRet(0,0)	
			.txtCAMPAIGN_CODE1.value = vntRet(1,0)
			.txtCAMPAIGN_NAME1.value = vntRet(2,0)
			SelectRtn
		end if
		
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCAMPAIGN_NAME1_onkeydown
	if window.event.keyCode = meEnter Or window.event.keyCode = meTab then
		Dim vntData
   		Dim i, strCols
		
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCAMPAIGN_INFO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, .txtCAMPAIGN_CODE1.value,.txtCAMPAIGN_NAME1.value)

			if not gDoErrorRtn ("GetCAMPAIGN_INFO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON.value = vntData(0,1)
					.txtCAMPAIGN_CODE1.value = vntData(1,1)
					.txtCAMPAIGN_NAME1.value = vntData(2,1)
					selectRtn           
				Else
					Call CAMPAIGN_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
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
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'--------------------------------------------------
' SpreadSheet �̺�Ʈ
'--------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt

	With frmThis
		'�� �̺�Ʈ
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME", Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, trim(vntData(5,1))
						
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		'����� �̺�Ʈ
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, trim(vntData(4,1))
						
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		'�������̺�Ʈ
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			vntRet = gShowModalWindow("SCCOTIMPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(5,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row)))
								
			vntRet = gShowModalWindow("SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			
		End If	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtCLIENTNAME1.focus
		.sprSht.Focus
	End With
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End if
	End With
End Sub

'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		CALL DateClean (Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2))
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MCCOMMI_RATE",frmThis.sprSht.ActiveRow, "35"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, "00006132"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, "Interactive Comm.��"
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus()
	End If
End Sub

Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, date1
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",.sprSht.ActiveRow, date2
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	Dim strSTD_STEP, strSTD_CM, strSTD_FACE, strSTD_PAGE, strPRICE
   	Dim strAMT
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
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
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						.txtCLIENTNAME1.focus
						.sprSht.focus ()
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus ()
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", strCodeName, _
													  "", "")

				If not gDoErrorRtn ("Get_BrandInfo") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntData(4,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntData(5,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntData(8,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntData(9,1)
											
						.txtCLIENTNAME1.focus ()
						.sprSht.focus ()
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME"), Row
						.txtCLIENTNAME1.focus ()
						.sprSht.focus ()
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, trim(vntData(5,1))
						
						.txtCLIENTNAME1.focus ()
						.sprSht.focus()
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtCLIENTNAME1.focus ()
						.sprSht.focus ()
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then		
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "G")

				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(2,1)			
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
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
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)) , "", "")
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(5,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(8,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(9,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then			
			vntInParams = array("", "" , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(5,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then		
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)),"G")
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtCLIENTNAME1.focus
		.sprSht.Focus
	End With
End Sub

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNTIM") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(5,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUB") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)) , _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(5,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(8,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(9,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
				
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNEX") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)),"G")
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
				
		.txtCLIENTNAME1.focus()
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	End With
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' ������ ȭ�� ������ �� �ʱ�ȭ 
'----------------------------------------------------------------------
	'����������ü ����	
	set mobjMDITINTERNETCAMPAIGN = gCreateRemoteObject("cMDIT.ccMDITINTERNETCAMPAIGN")
	set mobjMDCOGET				 = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 10, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 16, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CAMPAIGN_CODE | CAMPAIGN_NAME | SUBSEQ | BTNSUB | SUBSEQNAME | TIMCODE | BTNTIM | TIMNAME | CLIENTCODE | BTN | CLIENTNAME | TBRDSTDATE | TBRDEDDATE | MCCOMMI_RATE | EXCLIENTCODE | BTNEX | EXCLIENTNAME | DEPT_CD | DEPT_NAME | MEMO"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|�ڵ�|ķ���θ�|�귣���ڵ�|�귣��|���ڵ�|����|�������ڵ�|������|������|������|��������|������ڵ�|������|�μ��ڵ�|���μ�|���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|   5|      20|         7|2|  10|	    5|2|10|         7|2|  10|     9|     9|       7|         7|2|    10|       0|      19|  15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTNTIM | BTNSUB | BTN | BTNEX"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "MCCOMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CAMPAIGN_CODE | CAMPAIGN_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | CLIENTCODE | CLIENTNAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CAMPAIGN_CODE | DEPT_CD | DEPT_NAME"
		mobjSCGLSpr.ColHidden .sprSht, "DEPT_CD", true
	
		.sprSht.style.visibility = "visible"

    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDITINTERNETCAMPAIGN = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
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
		
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	End With
End Sub

'------------------------------------------
' HDR ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE
   	Dim strCAMPAIGN_CODE
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'���� �ʱ�ȭ
		strYEARMON		 = ""
		strCLIENTCODE	 = ""
		strCAMPAIGN_CODE = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		strYEARMON		 = .txtYEARMON.value 
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCAMPAIGN_CODE = .txtCAMPAIGN_CODE1.value
		
		vntData = mobjMDITINTERNETCAMPAIGN.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strCAMPAIGN_CODE)

		If not gDoErrorRtn ("SelectRtn") Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   		End if
   	End With
End Sub

'------------------------------------------
' HDR ����/���� ó�� 
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	Dim strYEAR
	
	With frmThis
   		'������ Validation
		'if DataValidation =false then exit sub
		'On error resume next
		
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "CAMPAIGN_NAME | CLIENTCODE | SUBSEQ",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ķ���θ�/������/�귣��� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CAMPAIGN_CODE | CAMPAIGN_NAME | SUBSEQ | BTNSUB | SUBSEQNAME | TIMCODE | BTNTIM | TIMNAME | CLIENTCODE | BTN | CLIENTNAME | TBRDSTDATE | TBRDEDDATE | MCCOMMI_RATE | EXCLIENTCODE | BTNEX | EXCLIENTNAME | DEPT_CD | DEPT_NAME | MEMO")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
				
		strYEAR = Mid(gNowDate2,3,2)
		
		intRtn = mobjMDITINTERNETCAMPAIGN.ProcessRtn(gstrConfigXml,vntData, strYEAR)
	
		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		End If
   	End With
End Sub

'------------------------------------------
'������ ����
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strCAMPAIGN_CODE
	Dim strCAMPAIGN_CODE2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strCAMPAIGN_CODE = mobjSCGLSpr.GetTextBinding( .sprSht,"CAMPAIGN_CODE",i)
				If strCAMPAIGN_CODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					vntData = mobjMDITINTERNETCAMPAIGN.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strCAMPAIGN_CODE) 
					If mlngRowCnt > 0 Then
						gErrorMsgBox i & "���� " & mlngRowCnt & "���� û�൥���ͷ� ����Ǿ��ֽ��ϴ�.","�����ȳ�!"
						Exit Sub
					End If
					lngchkCnt = lngchkCnt + 1
				End If
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
				strCAMPAIGN_CODE2 = mobjSCGLSpr.GetTextBinding(.sprSht,"CAMPAIGN_CODE",i)
			
				If strCAMPAIGN_CODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					intRtn = mobjMDITINTERNETCAMPAIGN.DeleteRtn(gstrConfigXml, strCAMPAIGN_CODE2)
					
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
				<TBODY>
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
												<td class="TITLE">���ͳ�-ķ����</td>
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
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1">
									</TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
													width="70">���</TD>
												<TD class="SEARCHDATA" width="180"><INPUT class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 80px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="10" size="6" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCAMPAIGN_CODE1, txtCAMPAIGN_NAME1)"
													width="70">ķ����</TD>
												<TD class="SEARCHDATA" width="255"><INPUT class="INPUT_L" id="txtCAMPAIGN_NAME1" title="�ڵ��" style="WIDTH: 174px; HEIGHT: 22px"
														type="text" maxLength="100" size="23" name="txtCAMPAIGN_NAME1"> <IMG id="ImgCAMPAIGN_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCAMPAIGN_CODE1"> <INPUT class="INPUT" id="txtCAMPAIGN_CODE1" title="�ڵ��Է�" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtCAMPAIGN_CODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
													width="70">������</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 174px; HEIGHT: 22px"
														type="text" maxLength="100" size="23" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="�ڵ��Է�" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="16113">
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
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
