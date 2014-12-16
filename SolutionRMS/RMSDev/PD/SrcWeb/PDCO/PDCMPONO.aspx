<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPONO.aspx.vb" Inherits="PD.PDCMPONO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>PROJECT ���� ���α׷�</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : PROJECT ��� ȭ��(PDCMPONO)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPONO.aspx
'��      �� : ������Ʈ ��� �� ����
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/10/27 By Tae Ho Kim
'			 2) 
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPONO '�����ڵ�, Ŭ����
Dim mstrPROCESS
Dim mstrPROCESS2 '��ȸ�����̸� true �űԻ����̸� false
Dim mstrCheck
Dim mobjMDLOGIN
Dim mobjMDCMEMP
Dim mobjPDCMGET
Dim mobjSCCOGet
Dim mstrHIDDEN
CONST meTAB = 9
mstrPROCESS = TRUE
mstrPROCESS2 = TRUE
mstrCheck = True
mstrHIDDEN = 0
'=============================
' �̺�Ʈ ���ν��� 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

Sub Set_SELECTTBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblSelectBody").style.display = "inline"
		Else
			document.getElementById("tblSelectBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgNew_onclick
	DataClean
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

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
	with frmThis
	mobjSCGLSpr.ExcelExportOption = true 
	mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub



'****************************************************************************************
' ��ü�� �Է½ÿ� �ڵ� ��ü������ ��������
'****************************************************************************************
Sub GetRealMedCode (strMEDCODE, strMEDNAME)
	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMGET.GetREALMEDNO1(gstrConfigXml,mlngRowCnt,mlngColCnt,strMEDCODE,strMEDNAME)

		if isArray(vntData) then
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, vntData(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, vntData(1,0)
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
   		end if
   	end with
END Sub


'****************************************************************************************
' �˾� �̺�Ʈ, ������, ��ü��, ��ü��
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))       ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE1                  ' gSetChangeFlag objectID	 Flag ���� �˸�
		end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value),"A")
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strGBN

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.cmbGROUPGBN.value = trim(vntRet(4,0)) 
				
     		'GetBrandDefaultFind	
     			
			if .sprSht.ActiveRow >0 Then
				If .cmbGROUPGBN.value = "2" Then
					strGBN = "�׷�"
				ElseIf .cmbGROUPGBN.value = "1" Then
					strGBN = "��׷�"
				End If
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
	
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtSUBSEQNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
     	
	End with

	'GetBrandAndDept '������ �������� �������� ���μ��� �����´�.
	
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),"A")
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					.cmbGROUPGBN.value = trim(vntData(4,1))
					if .sprSht.ActiveRow >0 Then
						If .cmbGROUPGBN.value = "2" Then
							strGBN = "�׷�"
						ElseIf .cmbGROUPGBN.value = "1" Then
							strGBN = "��׷�"
						End If
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
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
' �귣���ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgSUBSEQ_onclick
	Call BRANDCODE_POP()
End Sub

Sub BRANDCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strGBN

	with frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			.txtSUBSEQ.value = vntRet(0,0)			' �귣���ڵ�
			.txtSUBSEQNAME.value = vntRet(1,0)		' �귣���
			.txtCLIENTCODE.value = vntRet(2,0)		' �������ڵ�
			.txtCLIENTNAME.value = vntRet(3,0)		' �����ָ�
			.txtTIMCODE.value = vntRet(4,0)	' ���ڵ�
			.txtCLIENTTEAMNAME.value = vntRet(5,0)	' ����
			.cmbGROUPGBN.value = vntRet(10,0)	
			.txtCPDEPTCD.value = vntRet(8,0)		' �μ��ڵ�
			.txtCPDEPTNAME.value = vntRet(9,0)		' �μ���
			
			
			
			.txtCPEMPNAME.focus()					' ��Ŀ�� �̵�
			'msgbox vntRet(3,0)
			if .sprSht.ActiveRow >0 Then
						If .cmbGROUPGBN.value = "2" Then
							strGBN = "�׷�"
						ElseIf .cmbGROUPGBN.value = "1" Then
							strGBN = "��׷�"
						End If
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, .txtSUBSEQ.value
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, .txtSUBSEQNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, .txtTIMCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTTEAMNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			gSetChangeFlag .txtSUBSEQ		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtCLIENTCODE
			gSetChangeFlag .txtCPDEPTCD
     	end if
	End with
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown

	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = vntData(0,1)			' �귣���ڵ�
					.txtSUBSEQNAME.value = vntData(1,1)		' �귣���
					.txtCLIENTCODE.value = vntData(2,1)		' �������ڵ�
					.txtCLIENTNAME.value = vntData(3,1)		' �����ָ�
					.txtTIMCODE.value = vntData(4,1)	' ���ڵ�
					.txtCLIENTTEAMNAME.value = vntData(5,1)	' ����
					.txtCPDEPTCD.value = vntData(8,1)		' �μ��ڵ�
					.txtCPDEPTNAME.value = vntData(9,1)		' �μ���
					.cmbGROUPGBN.value = vntData(10,1)
					
					.txtCPEMPNAME.focus()
					if .sprSht.ActiveRow >0 Then
						If .cmbGROUPGBN.value = "2" Then
							strGBN = "�׷�"
						ElseIf .cmbGROUPGBN.value = "1" Then
							strGBN = "��׷�"
						End If
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, .txtSUBSEQ.value
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, .txtSUBSEQNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, .txtTIMCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTTEAMNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
				Else
					Call BRANDCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------
' ���μ� ��ȸ 
'-----------------------------
Sub ImgCPDEPTCD_onclick
	Call imgCC_POP()
End Sub

Sub imgCC_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtCPDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCPDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
			.txtCPDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtCPEMPNAME.focus()
			gSetChangeFlag .txtCPDEPTCD
		end if
	end with
End Sub

Sub txtCPDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCPDEPTNAME.value)
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtCPDEPTCD.value = trim(vntData(0,0))
					.txtCPDEPTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtCPEMPNAME.focus()
				Else
					Call imgCC_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ����ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCPEMPNO_onclick
	Call EMP_POP()
End Sub

'���� ������List ��������
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCPDEPTCD.value), trim(.txtCPDEPTNAME.value), trim(.txtCPEMPNO.value), trim(.txtCPEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCPEMPNO.value = vntRet(0,0) and .txtCPEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCPDEPTCD.value = trim(vntRet(2,0))  ' Code�� ����
			.txtCPDEPTNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
			.txtCPEMPNO.value = trim(vntRet(0,0))
			.txtCPEMPNAME.value = trim(vntRet(1,0))
			
			if .sprSht.ActiveRow >0 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"CPEMPNO",.sprSht.ActiveRow, .txtCPEMPNO.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CPEMPNAME",.sprSht.ActiveRow, .txtCPEMPNAME.value
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtMEMO.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCPEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtCPEMPNAME
			gSetChangeFlag .txtCPDEPTCD
			gSetChangeFlag .txtCPDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCPEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCPEMPNO.value, .txtCPEMPNAME.value,"A",.txtCPDEPTCD.value,.txtCPDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCPEMPNO.value = trim(vntData(0,1))
					.txtCPEMPNAME.value = trim(vntData(1,1))
					.txtCPDEPTCD.value = trim(vntData(2,1))
					.txtCPDEPTNAME.value = trim(vntData(3,1))
					
					if .sprSht.ActiveRow >0 Then
						
						mobjSCGLSpr.SetTextBinding .sprSht,"CPEMPNO",.sprSht.ActiveRow, .txtCPEMPNO.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPEMPNAME",.sprSht.ActiveRow, .txtCPEMPNAME.value
						
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTCD",.sprSht.ActiveRow, .txtCPDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CPDEPTNAME",.sprSht.ActiveRow, .txtCPDEPTNAME.value
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtMEMO.focus()
					gSetChangeFlag .txtCPEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ��/����� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTTEAM_onclick
	Call ImgCLIENTTEAM_POP()
End Sub

Sub ImgCLIENTTEAM_POP
	Dim vntRet, vntInParams
	Dim strGBN
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtCLIENTCODE.value) , trim(.txtCLIENTNAME.value),trim(.txtTIMCODE.value) , trim(.txtCLIENTTEAMNAME.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
			.txtTIMCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtCLIENTTEAMNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(4,0))
			.txtCLIENTNAME.value = trim(vntRet(5,0))
			.cmbGROUPGBN.value = trim(vntRet(6,0))
		 
			if .sprSht.ActiveRow >0 Then	
				If .cmbGROUPGBN.value = "2" Then
					strGBN = "�׷�"
				ElseIf .cmbGROUPGBN.value = "1" Then
					strGBN = "��׷�"
				End If
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtCPEMPNAME.focus()
			gSetChangeFlag .txtCPDEPTCD
		end if
	end with
End Sub


Sub txtCLIENTTEAMNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		Dim strGBN
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtTIMCODE.value) , trim(.txtCLIENTTEAMNAME.value))
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	'Code�� ����
					.txtCLIENTTEAMNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					.cmbGROUPGBN.value = trim(vntData(6,1))
					
					
					if .sprSht.ActiveRow >0 Then	
						If .cmbGROUPGBN.value = "2" Then
							strGBN = "�׷�"
						ElseIf .cmbGROUPGBN.value = "1" Then
							strGBN = "��׷�"
						End If
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGBN
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtCPEMPNAME.focus()
				Else
					Call imgCC_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub





'****************************************************************************************
' ������� �޷�
'****************************************************************************************
'��ȸ��
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================

Sub txtPROJECTNM_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PROJECTNM",frmThis.sprSht.ActiveRow, frmThis.txtPROJECTNM.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCLIENTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCLIENTCODE_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCPDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CPDEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCPDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCPDEPTCD_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CPDEPTCD",frmThis.sprSht.ActiveRow, frmThis.txtCPDEPTCD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCREDAY_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREDAY",frmThis.sprSht.ActiveRow, frmThis.txtCREDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCPEMPNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CPEMPNAME",frmThis.sprSht.ActiveRow, frmThis.txtCPEMPNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCPEMPNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CPEMPNO",frmThis.sprSht.ActiveRow, frmThis.txtCPEMPNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbGROUPGBN_onchange

	WITH frmThis
		Dim strGROUPGBN
		If .cmbGROUPGBN.value  = "1" Then
			strGROUPGBN = "�׷�"
			Else
			strGROUPGBN = "��׷�"
		End If	
	End with
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GROUPGBN",frmThis.sprSht.ActiveRow, strGROUPGBN
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtSUBSEQNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtSUBSEQ_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQ.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtMEMO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub



Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub



'�Է¿�
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtCREDAY,frmThis.imgCalEndar,"txtCREDAY_onchange()"
		gSetChange
	end with
End Sub
Sub txtCREDAY_onchange
	gSetChange
End Sub

'ķ���θ� Ű�ٿ�
Sub txtCAMPAIGN_NAME_onkeydown
	if window.event.keyCode = meEnter or window.event.keyCode = meTAB then
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'��������� Ű�ٿ�
Sub txtTBRDEDDATE_onkeydown
	if window.event.keyCode = meEnter or window.event.keyCode = meTAB then
		frmThis.txtCLIENTSUBNAME.focus()
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'�������� Ű�ٿ�
Sub txtMCCOMMI_RATE_onkeydown
	if window.event.keyCode = meEnter or window.event.keyCode = meTAB then
		frmThis.txtNOTE.focus()
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 and Col > 1 then		
			sprShtToFieldBinding Col,Row
		End If
	end with

End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'��Ʈ�� �������ѷο��� ������ ��� �ʴ��� ���ε�
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		'PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPEMPNO|MEMO
		.txtPROJECTNO.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNO",Row)
		.txtPROJECTNM.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNM",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtSUBSEQ.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		.txtSUBSEQNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		If mobjSCGLSpr.GetTextBinding(.sprSht,"GROUPGBN",Row) ="�׷�"  Then
		.cmbGROUPGBN.value	=	"2"
		Else
		.cmbGROUPGBN.value	=	"1"
		End IF
		.txtCREDAY.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CREDAY",Row)
		.txtCPDEPTCD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CPDEPTCD",Row)
		.txtCPDEPTNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CPDEPTNAME",Row)
		.txtCPEMPNO.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CPEMPNO",Row)
		.txtCPEMPNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CPEMPNAME",Row)
		.txtCLIENTTEAMNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTTEAMNAME",Row)
		.txtTIMCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		.txtMEMO.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
	
		
		
   	end with
	'CALL Field_Lock ()
End Function

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
	
	'Ű�� �����϶� ���ε�
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
End Sub
'=============================
' UI���� ���ν��� 
'=============================
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	set mobjPONO	= gCreateRemoteObject("cPDCO.ccPDCOPONO")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjMDLOGIN	= gCreateRemoteObject("cSCCO.ccSCCOLOGIN") 
	set mobjMDCMEMP = gCreateRemoteObject("cSCCO.ccSCCOEMPMST")
	set mobjSCCOGET	= gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|CLIENTTEAMNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPDEPTNAME|CPEMPNO|CPEMPNAME|MEMO|TIMCODE"
		mobjSCGLSpr.SetHeader .sprSht,		"������Ʈ�ڵ�|������Ʈ��|�������ڵ�|������|���̸�        |�귣���ڵ�|�귣��|�׷챸��|�����|�μ��ڵ�|���μ�|���|�����|���|���ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|        29|         0|    25|20            |         0|    18|      10|     8|       0|      19|   0|    10| 10 |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CPDEPTCD|CPEMPNO|MEMO"
		mobjSCGLSpr.ColHidden .sprSht, "PROJECTNO|CLIENTCODE|SUBSEQ|TIMCODE", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPDEPTNAME|CPEMPNO|CPEMPNAME|MEMO|CLIENTTEAMNAME",-1,-1,0,2,false
	
		.sprSht.MaxRows = 50
	End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	
End Sub



Sub EndPage()
	set mobjPONO = Nothing
	set mobjMDLOGIN = Nothing
	set mobjPDCMGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		
		.txtCREDAY.value = gNowDate
		
		
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""
	End with
	DataNewClean
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub DataNewClean
	with frmThis
	.txtCREDAY.value = ""
	.cmbGROUPGBN.selectedIndex  = -1
	End with
End Sub
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub
'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn ()
   Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strDELCODE
	Dim intRtnSave
	Dim vntData2
	with frmThis
	'On error resume next
	
		
		
  		'������ Validation
		if DataValidation =false then exit sub
		strCODE = .txtPROJECTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPEMPNO|MEMO|TIMCODE")
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		
		
		'ó�� ������ü ȣ��
		strMasterData = gXMLGetBindingData (xmlBind)
		
		if .txtPROJECTNO.value = "" then
			strSEQFlag = "new"
			intRtn = mobjPONO.ProcessRtn(gstrConfigXml,strMasterData, strSEQFlag)
		else
			'���⼭ ���� JOBNO �� ��ϵǾ�������, ���� �Ұ� ����
			strDELCODE = Trim(.txtPROJECTNO.value)
			'vntData2 = mobjPONO.GetPONODELSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strDELCODE)
			
		
			If mlngRowCnt > 0 Then
				'intRtnSave = gYesNoMsgbox("JOBNO�� ��ϵǾ��ִ� ������Ʈ �Դϴ�.�����Ͻðڽ��ϱ�?","�ڷắ�� Ȯ��")
				'IF intRtnSave <> vbYes then exit Sub
				gErrorMsgBox "��ϵǾ��ִ� JOBNO �� ������°� �Ƿڰ� �ƴѰ��� �ֽ��ϴ�." & vbcrlf & "������ �Ұ��� �մϴ�.","����ȳ�!"
				SelectRtn
				Exit Sub
			End IF
			
			
			
			intRtn = mobjPONO.ProcessRtnSheet(gstrConfigXml,vntData)
		end if
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if strSEQFlag = "new" then
				gErrorMsgBox " �ڷᰡ �ű�����" & mePROC_DONE,"����ȳ�" 
			else
				gErrorMsgBox " �ڷᰡ" & intRtn & " �� ��������" & mePROC_DONE,"����ȳ�" 
			end if
			SelectRtn
  		end if
 	end with
End Sub


'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		
   	
   	End with
	DataValidation = true
End Function

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		
		
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		'���ݰ�꼭 �Ϸ���ȸ
		vntData = mobjPONO.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtPROJECTNM1.value),Trim(.txtPROJECTNO1.value),Trim(.txtCLIENTNAME1.value),Trim(.txtCLIENTCODE1.value),"ST")
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht.MaxRows = 0	
			DATACLEAN		
			DataNewClean
			End If
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
			sprShtToFieldBinding 1,1
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub



'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strTBRDSTDATE,strTBRDEDDATE, strCAMPAIGN_CODE, strCAMPAIGN_NAME, strCLIENTCODE, strCLIENTNAME)
	With frmThis
		.txtTBRDSTDATE1.value = strTBRDSTDATE
		.txtTBRDEDDATE1.value = strTBRDEDDATE
		.txtCAMPAIGN_CODE1.value = strCAMPAIGN_CODE
		.txtCAMPAIGN_NAME1.value = strCAMPAIGN_NAME
		.txtCLIENTCODE1.value = strCLIENTCODE
		.txtCLIENTNAME1.value = strCLIENTNAME
	End With
End Sub


'****************************************************************************************
' ������ �Է½� �귣��� ���μ� ��������
'****************************************************************************************
Sub GetBrandAndDept ()
	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMGET.GetBRANDANDDEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value)

		if isArray(vntData) then
			.txtSUBSEQ.value = vntData(0,0)		             ' Code�� ����
			.txtSUBSEQNAME.value = vntData(1,0) 
			.txtDEPT_CD.value = vntData(2,0) 
			.txtDEPT_NAME.value = vntData(3,0) 
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, .txtSUBSEQ.value
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, .txtSUBSEQNAME.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",.sprSht.ActiveRow, .txtDEPT_CD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",.sprSht.ActiveRow, .txtDEPT_NAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
   		end if
   		
   		if .txtSUBSEQ.value = "" then
   			.txtSUBSEQNAME.focus()
   		elseif .txtDEPT_CD.value = "" then
   			.txtDEPT_NAME.focus()
   		end if
   	end with
END Sub



Function SelectRtn_Detail (ByVal strCAMPAIGN_CODE)
	dim vntData
	Dim intCnt
	Dim strRows
	on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjCAMPAIGN.SelectRtn_Detail (gstrConfigXml,mlngRowCnt,mlngColCnt,strCAMPAIGN_CODE)
	with frmThis
		IF mlngRowCnt = 0 THEN
			SelectRtn_Detail = True
		end if
	END WITH
End Function

Sub DataClean
	with frmThis
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = ""
		.txtCLIENTCODE.value = ""
		.txtCLIENTNAME.value = ""
		.txtSUBSEQ.value = ""
		.txtSUBSEQNAME.value = ""
		.txtCPDEPTCD.value = ""
		.txtCPDEPTNAME.value = ""
		.txtCPEMPNO.value = ""
		.txtCPEMPNAME.value = ""
		.txtMEMO.value = ""
		.txtCLIENTTEAMNAME.value = ""
		.txtTIMCODE.value = ""
		.cmbGROUPGBN.value = "1"
		.txtCREDAY.value = gNowDate
		.sprSht.MaxRows = 0
	End With
End Sub
Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub
Sub DefaultValue
Dim strGROUPGBN
	
	with frmThis
		If .cmbGROUPGBN.value  = "2" Then
		strGROUPGBN = "�׷�"
		Else
		strGROUPGBN = "��׷�"
		End If	
		mobjSCGLSpr.SetTextBinding .sprSht,"GROUPGBN",.sprSht.ActiveRow, strGROUPGBN
		mobjSCGLSpr.SetTextBinding .sprSht,"CREDAY",.sprSht.ActiveRow, .txtCREDAY.value 
	End with
End Sub
'ProjectNO ��ȸ�˾�
Sub ImgPROJECTNO1_onclick
	Call PONO_POP()
End Sub
'���� ������List ��������
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.focus()					' ��Ŀ�� �̵�
     	end if
	End with
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCODE
	
	with frmThis
		If .txtPROJECTNO.value = "" Or .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ �ڷᰡ �����ϴ�.","�����ȳ�"
			Exit Sub
		End If
		intSelCnt = 0
		strCODE = Trim(.txtPROJECTNO.value)
		vntData = mobjPONO.GetPONODELSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
		
	
		If mlngRowCnt > 0 Then
			gErrorMsgBox "JOBNO �� ��ϵǾ��ִ� PROJECT �Դϴ�.�����ɼ� �����ϴ�.","�����ȳ�"
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'�ڷ� ����
		intRtn = mobjPONO.DeleteRtn(gstrConfigXml,strCODE)
			
		IF not gDoErrorRtn ("DeleteRtn") then
			'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			msgbox "[" & strCODE & "] PROJECT �� �����Ǿ����ϴ�."
   		End IF
		'���� ���� ����
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
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;���۰���</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 90%" cellSpacing="0" cellPadding="0" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()" width="80">����� 
												�˻�</TD>
											<TD class="SEARCHDATA" width="230"><INPUT class="INPUT" id="txtFROM" title="�Ⱓ�˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
													type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPROJECTNO1, txtPROJECTNM1)"
												width="80">������Ʈ��</TD>
											<TD class="SEARCHDATA" width="260"><INPUT class="INPUT_L" id="txtPROJECTNM1" title="�ڵ��" style="WIDTH: 176px; HEIGHT: 22px"
													type="text" maxLength="100" size="24" name="txtPROJECTNM1"><IMG id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgPROJECTNO1"><INPUT class="INPUT" id="txtPROJECTNO1" title="�ڵ�" style="WIDTH: 56px; HEIGHT: 22px" type="text"
													maxLength="6" size="4" name="txtPROJECTNO1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80">������</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 144px; HEIGHT: 22px"
													type="text" maxLength="100" size="18" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE1"><INPUT class="INPUT" id="txtCLIENTCODE1" title="�ڵ��Է�" style="WIDTH: 56px; HEIGHT: 22px"
													type="text" maxLength="6" size="4" name="txtCLIENTCODE1"></TD>
											<td class="SEARCHDATA" width="53"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;&nbsp;������Ʈ ����<span id="spnSELECTHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_SELECTTBL_HIDDEN ()">(�����)</span></td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
														<!--<td><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gIF'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imglistcopy.gIF" width="77" border="0"
																name="Imgcopy"></td>-->
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></td>
														<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblSelectBody" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtPROJECTNM, txtPROJECTNO)"
												width="80">������Ʈ��</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="230"><INPUT dataFld="PROJECTNM" class="INPUT_L" id="txtPROJECTNM" title="������Ʈ��" style="WIDTH: 160px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="200" size="21" name="txtPROJECTNM" dataSrc="#xmlBind"><INPUT dataFld="PROJECTNO" class="NOINPUT" id="txtPROJECTNO" title="������Ʈ�ڵ�" style="WIDTH: 65px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="7" size="5" name="txtPROJECTNO">
											</TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSUBSEQNAME,txtSUBSEQ)"
												align="right" width="80">�귣��</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="260"><INPUT dataFld="SUBSEQNAME" class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 160px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQ" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
													name="ImgSUBSEQ"><INPUT dataFld="SUBSEQ" class="INPUT" id="txtSUBSEQ" title="�귣���ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtSUBSEQ"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCPDEPTNAME,txtCPDEPTCD)"
												align="right" width="80">���μ�</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="CPDEPTNAME" class="INPUT_L" id="txtCPDEPTNAME" title="���μ�(CP)��" style="WIDTH: 192px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="26" name="txtCPDEPTNAME"><IMG id="ImgCPDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
													name="ImgCPDEPTCD"><INPUT dataFld="CPDEPTCD" class="INPUT" id="txtCPDEPTCD" title="���μ�(CP)�ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCPDEPTCD"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCREDAY, '')">�����</TD>
											<TD class="DATA"><INPUT dataFld="CREDAY" class="INPUT" id="txtCREDAY" title="�����" style="WIDTH: 88px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCREDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
													name="imgCalEndar">
											</TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCLIENTTEAMNAME, txtTIMCODE)">��</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTTEAMNAME" class="INPUT_L" id="txtCLIENTTEAMNAME" title="�����(CP)��"
													style="WIDTH: 160px; HEIGHT: 22px" dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtCLIENTTEAMNAME"><IMG id="ImgCLIENTTEAM" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgCLIENTTEAM"><INPUT dataFld="TIMCODE" class="INPUT" id="txtTIMCODE" title="CLIENT��" style="WIDTH: 70px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtTIMCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtCPEMPNAME,txtCPEMPNO)"
												width="80">�����</TD>
											<TD class="DATA"><INPUT dataFld="CPEMPNAME" class="INPUT_L" id="txtCPEMPNAME" title="�����(CP)��" style="WIDTH: 192px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="26" name="txtCPEMPNAME"><IMG id="ImgCPEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
													name="ImgCPEMPNO"><INPUT dataFld="CPEMPNO" class="INPUT" id="txtCPEMPNO" title="�����(CP)���" style="WIDTH: 70px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCPEMPNO"></TD>
										</TR>
										<TR>
											<TD class="LABEL">�׷챸��</TD>
											<TD class="DATA"><SELECT dataFld="GROUPGBN" id="cmbGROUPGBN" title="�׷챸��" style="WIDTH: 112px" dataSrc="#xmlBind"
													name="cmbGROUPGBN">
													<OPTION value="2" selected>�׷�</OPTION>
													<OPTION value="1">��׷�</OPTION>
												</SELECT></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">������</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 160px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
													name="ImgCLIENTCODE"><INPUT dataFld="CLIENTCODE" class="INPUT" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCLIENTCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEMO, '')">�޸�</TD>
											<TD class="DATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="���" style="WIDTH: 285px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="1000" size="41" name="txtMEMO"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: visible; POSITION: relative" ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="42466">
											<PARAM NAME="_ExtentY" VALUE="13467">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
