<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPREESTLIST.aspx.vb" Inherits="PD.PDCMPREESTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����</title>
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMPREESTLIST, mobjPDCMGET,mobjPDCMJOBNO
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
Dim mstrCHKROW
mstrCHKROW = false
Const meTab = 9
mALLCHECK = TRUE
mstrCheck=True
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgFind_onclick()
Dim vntRet
	vntRet = gShowModalWindow("PDCMCHARGELISTPOP.aspx","" , 1060,730)
End Sub
'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
mstrCHKROW = false
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
	
End Sub
Sub imgQuery1_onclick
	gFlowWait meWAIT_ON
	SelectRtn_HDR2
	gFlowWait meWAIT_OFF
	
End Sub

Sub imgNew_onclick
	Dim vntInParams
	Dim vntRet
	Dim strRow
	with frmThis
	vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow))
	vntRet = gShowModalWindow("PDCMPREESTDTLNEW.aspx",vntInParams , 1060,780)
	strRow = .sprSht.ActiveRow
	selectRtn
	mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
	Call sprSht_click(1,strRow)
	End with
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
	mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel1_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht1
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgRowAdd_onclick ()
	
	With frmThis
		
		call sprSht_Keydown(meINS_ROW, 0)
		intiSprValue
	End With 
End Sub

Sub imgRowDel_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDetail_onclick()
Dim vntInParams
Dim vntRet
Dim strJOBNO
Dim strRow
	with frmThis
		if mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow) = "" then
				gErrorMsgBox "������ List �� �켱�����Ͻð� �Է� �Ͻʽÿ�.","ó���ȳ�" 
				Exit Sub
			End if
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow))
			vntRet = gShowModalWindow("PDCMPREESTDTL.aspx",vntInParams , 1060,780)
			
			
			.txtCLIENTSUBNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht1.Focus
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strRow = .sprSht.ActiveRow
			SelectRtn
			SelectRtn_DBLHDR(strJOBNO)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow	
	End with
End Sub

Sub Imgcopy_onclick()
Dim intRtn
Dim intRtnCopy
Dim strOLDCODE
Dim strJOBNO
Dim strCREDAY
Dim strCLIENTSUBCODE
Dim strCOMMITION
Dim strCLIENTCODE
Dim strSUBSEQ
Dim vntRet
Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			strJOBNO         = trim(vntRet(0,0))  ' Code�� ����
			strCREDAY        = trim(vntRet(6,0))  ' �ڵ�� ǥ��
			strCLIENTSUBCODE = trim(vntRet(2,0)) 
			strCOMMITION     = trim(vntRet(3,0)) 
			strCLIENTCODE    = trim(vntRet(4,0)) 
			strSUBSEQ        = trim(vntRet(5,0)) 
		Else
			Exit Sub
     	end if

		strOLDCODE = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",.sprSht1.ActiveRow)
		'strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
		'strCREDAY = Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",.sprSht.ActiveRow),"-","")
		'strCLIENTSUBCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow)
		'strCOMMITION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMITION",.sprSht.ActiveRow)
		'strCOMMITION = CDBL(strCOMMITION)
		'strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		'strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
		
		'strCREDAY,strCLIENTSUBCODE,strCOMMITION,strCLIENTCODE
		intRtn = gYesNoMsgbox("[" & strOLDCODE & "] �ڷḦ �����Ͻðڽ��ϱ�?","�������� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intRtnCopy = mobjPDCMPREESTLIST.ProcessRtn_COPY(gstrConfigXml,strOLDCODE,strJOBNO,strCREDAY,strCLIENTSUBCODE,strCOMMITION,strCLIENTCODE,strSUBSEQ)
		if not gDoErrorRtn ("ProcessRtn_COPY") then
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			.txtJOBNO.value = trim(vntRet(0,0)) 
			.txtJOBNAME.value = trim(vntRet(1,0)) 
			SelectRtn
  		end if
	End with 
End Sub
Sub intiSprValue
	Dim strJOBNAME
	Dim strJOBNO
	Dim strCREDAY
	Dim strCLIENTSUBCODE
	Dim strCOMMITION
	Dim strCLIENTCODE
	Dim strSUBSEQ
	with frmThis
		If .sprSht.MaxRows <> 0 Then
			strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow)
			strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
			strCREDAY = Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",.sprSht.ActiveRow),"-","")
			strCLIENTSUBCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow)
			strCOMMITION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMITION",.sprSht.ActiveRow)
			strCOMMITION = CDBL(strCOMMITION)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
			strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"JOBNO",.sprSht1.ActiveRow, strJOBNO
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"JOBNAME",.sprSht1.ActiveRow, strJOBNAME
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CREDAY",.sprSht1.ActiveRow, strCREDAY
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"SUSURATE",.sprSht1.ActiveRow, strCOMMITION
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CLIENTSUBCODE",.sprSht1.ActiveRow, strCLIENTSUBCODE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, strCLIENTCODE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht1,"SUBSEQ",.sprSht1.ActiveRow, strSUBSEQ
		End If
	End with
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
	

End Sub

Sub sprSht1_Keydown(KeyCode, Shift)
Dim intRtn
if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
	
	
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 8 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(13), cint(Shift), -1, 1)
		'intiSprValue
		end if
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if

End sub
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
		
				
     		'GetBrandDefaultFind	
     			
			
			.txtCLIENTSUBNAME.focus()					' ��Ŀ�� �̵�
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
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
				
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
' ������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME.value = trim(vntRet(6,0))
			
		
			
			.txtJOBNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME.value = trim(vntData(6,0))
					
				
					.txtCLIENTSUBNAME.focus()
					gSetChangeFlag .txtCLIENTSUBCODE
					gSetChangeFlag .CLIENTCODE
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'���� ������List ��������
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtJOBNO.value = trim(vntRet(0,0))  ' Code�� ����
			.txtJOBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO.value = trim(vntData(0,0))
					.txtJOBNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'�ڡڡڡڡڡڡڡڡڡڡڡڡ� �ϴ���ȸ���� �ڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡ�
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP1()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP1
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
		
				
     		'GetBrandDefaultFind	
     			
			
			.txtJOBNAME1.focus()					' ��Ŀ�� �̵�
			'gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
     	
	End with

	'GetBrandAndDept '������ �������� �������� ���μ��� �����´�.
	
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
			
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,0))
					.txtCLIENTNAME1.value = trim(vntData(1,0))
				
				Else
					Call CLIENTCODE_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgJOBNO1_onclick
	Call SEARCHJOB_POP1()
End Sub

'���� ������List ��������
Sub SEARCHJOB_POP1
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO1.value), trim(.txtJOBNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO1.value = vntRet(0,0) and .txtJOBNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtJOBNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtJOBNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO1.value),trim(.txtJOBNAME1.value))
			if not gDoErrorRtn ("txtJOBNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO1.value = trim(vntData(0,0))
					.txtJOBNAME1.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'�ڡڡڡڡڡڡڡڡڡڡڡڡ� �ϴ���ȸ�� �ڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡڡ�
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	Dim intCnt2
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht1.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
'	For intCnt2 = 1 To frmThis.sprSht1.MaxRows
'		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXNO",intCnt2) <> "" THEN
'			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "���ݰ�꼭��ȣ�� �����ϴ� ������ ������� �� �����ϴ�.","�μ�ȳ�!"
'			Exit Sub
'		End If
'	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'�μ��ư�� Ŭ���ϱ� ���� PD_CHARGE_TEMP���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�.
		'PD_CHARGE_TEMP���� ����
		intRtn = mobjPDCMPREESTLIST.DeleteRtn_temp(gstrConfigXml)
		'PD_CHARGE_TEMP���� ��
		
		ModuleDir = "PD"
		ReportName = "PDCMCHARGE.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",.sprSht1.activeRow)
		
		vntData = mobjPDCMPREESTLIST.Get_PREEST_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strPREESTNO)
	
		strcntsum = 0
		IF not gDoErrorRtn ("Get_PREEST_CNT") then
			datacnt = mlngRowCnt
			
			for i=1 to 3
				strUSERID = ""
				vntDataTemp = mobjPDCMPREESTLIST.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
				
		window.setTimeout "printSetTimeout", 10000
	
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMPREESTLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub



Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub






'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------




'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************

sub sprSht_DblClick (ByVal Col, ByVal Row)
Dim strJOBNO	
Dim vntInParams
Dim vntRet
Dim strRow
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row))
			vntRet = gShowModalWindow("PDCMESTDTL.aspx",vntInParams , 1060,780)
			strRow = Row
			'���⼭ ���� ���� ���� ȭ�� ȣ��
			.txtCLIENTSUBNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			
			SelectRtn
			SelectRtn_DBLHDR(strJOBNO)
			mobjSCGLSpr.ActiveCell .sprSht, Col, strRow			
		end if
	end with
end sub
Sub sprSht_Click(ByVal Col, ByVal Row)
Dim strJOBNO	
Dim vntInParams
Dim vntRet
Dim strRow
with frmThis
	mstrCHKROW = True
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			SelectRtn_DBLHDR(strJOBNO)
				
End with
End Sub
sub sprSht1_DblClick (ByVal Col, ByVal Row)
Dim strJOBNO	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		Else
			Dim vntInParams
			Dim vntRet
			Dim strRow
				with frmThis
					if mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow) = "" then
							gErrorMsgBox "������ List �� �켱�����Ͻð� �Է� �Ͻʽÿ�.","ó���ȳ�" 
							Exit Sub
						End if
						vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow))
						vntRet = gShowModalWindow("PDCMPREESTDTL.aspx",vntInParams , 1060,780)
						
						
						.txtCLIENTSUBNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
						.sprSht1.Focus
						strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
						strRow = .sprSht.ActiveRow
						SelectRtn
						SelectRtn_DBLHDR(strJOBNO)
						mobjSCGLSpr.ActiveCell .sprSht, 1, strRow	
				End with
		end if
	end with
end sub

Sub sprSht1_Change(ByVal Col, ByVal Row)
Dim vntData
Dim i, strCols
Dim strCode, strCodeName

	with frmThis
				'Long Type�� ByRef ������ �ʱ�ȭ
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				strCode = ""
				strCodeName = ""
				
				IF Col = 3 Then
						
					strCode = ""
					strCode		= mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow)
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNAME",.sprSht1.ActiveRow)
					
					vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)
					
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",.sprSht1.ActiveRow, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",.sprSht1.ActiveRow, vntData(1,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",.sprSht1.ActiveRow, vntData(2,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",.sprSht1.ActiveRow, vntData(3,0)	
						mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, vntData(4,0)
						mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",.sprSht1.ActiveRow, vntData(5,0)			
						mobjSCGLSpr.CellChanged .sprSht1, .sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					Else
						mobjSCGLSpr_ClickProc .sprSht1, Col, .sprSht1.ActiveRow
					End If
				END IF
				
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col, Row
End Sub
Sub mobjSCGLSpr_ClickProc(sprSht1, Col, Row)
	dim vntRet, vntInParams
	With frmThis
	
		IF Col = 3 Then
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( sprSht1,"JOBNAME",Row))
			vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(5,0)	
				mobjSCGLSpr.CellChanged sprSht1, Col,Row
				
				
			End IF
			.txtCLIENTSUBNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht1.Focus	
		
		end if
	End With
End Sub


Sub sprSht1_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strJOBNO
	Dim strRow
	with frmThis
		IF Col = 2 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") then exit Sub
			if mobjSCGLSpr.GetTextBinding( .sprSht1,"CONF",Row) = "Y" Then exit Sub	
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",Row), mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNAME",Row))
			vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUSURATE",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(5,0)				
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtCLIENTSUBNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht1.Focus
			mobjSCGLSpr.ActiveCell .sprSht1, Col+3, Row
		
		
		end if
	end with
end SUB
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjPDCMPREESTLIST	= gCreateRemoteObject("cPDCO.ccPDCOPREESTLIST")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
    set mobjPDCMJOBNO       = gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "160px"
	'pnlTab1.style.left= "7px"
	
	'pnlTab2.style.position = "absolute"
	'pnlTab2.style.top = "693px"
	'pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'��������û����Ʈ
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 18, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "REQDAY|PROJECTNO|JOBNO|JOBNAME|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|ENDFLAG|JOBGUBN|CREPART|CREGUBN|COMMITION|CLIENTCODE|PREESTNO|AMT|DEMANDYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		   "�Ƿ���|������Ʈ��ȣ|JOBNO|JOB��|������|�����|����θ�|�귣��|�귣���|����|��ü�κ�|��ü�з�|�ű�|��������|�������ڵ�|Ȯ�������ڵ�|�����ݾ�|û������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "10    | 0          |7    |   19|13    |6     |12      |6     |13      |   6|12      |12      |6   |0       |0         |0           |11		|10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|DEMANDYEARMON", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PROJECTNO|JOBNO|JOBNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|JOBGUBN|CREPART|CREGUBN|REQDAY|ENDFLAG|CLIENTNAME|PREESTNO|AMT|DEMANDYEARMON"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME|MEDNAME|PROGRAM_NAME|PUB_FACENAME|COL_DEG ", -1, -1, 100
		mobjSCGLSpr.ColHidden .sprSht, "PROJECTNO|COMMITION|CLIENTCODE|PREESTNO|CLIENTSUBCODE|SUBSEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTSUBNAME|SUBSEQNAME|CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTSUBCODE|SUBSEQ|JOBGUBN|CREPART|CREGUBN|JOBNO|ENDFLAG|DEMANDYEARMON",-1,-1,2,2,false
		
		
	    
	    '******************************************************************
		'����������Ʈ
		'******************************************************************
	    gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 13, 0, 0, 0,2
		mobjSCGLSpr.AddCellSpan  .sprSht1, 1, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1, "JOBNO|BTN|JOBNAME|PREESTNO|PREESTNAME|AMT|MEMO|CONF|CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ"
		mobjSCGLSpr.SetHeader .sprSht1,		"JOBNO|JOB��|�������ڵ�|��������|�ݾ�|���|Ȯ������|�Ƿ���|�����|Ŀ�̼�|�������ڵ�|�귣���ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "6|2|20|9|28|12|35|10|0|0|0|10|0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "SUSURATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "JOBNO|JOBNAME|PREESTNAME|MEMO", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht1, true, "PREESTNO|CONF|BTN|SUSURATE|JOBNO|JOBNAME|PREESTNAME|MEMO|CONF|AMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "PREESTNO|CONF",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht1, "CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ", true
		mobjSCGLSpr.SetScrollBar .sprSht1,2,False,0,-1
	
	    		
    End With    
	'pnlTab1.style.visibility = "visible"
	'pnlTab2.style.visibility = "visible"
	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'WITH frmThis
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
	'			case 1 : .txtCLIENTCODE.value = vntInParam(i)
	'			case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'��ȸ�߰��ʵ�
	'			case 3 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
	'			case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
	'			case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
	'		end select
	'	next
	'end with
	'SelectRtn
	Call SEARCHCOMBO_TYPE()
End Sub
'-----------------------------------------------------------------------------------------
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
    With frmThis   

		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB���� ȣ��
		vntENDFLAG = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '���ۻ��� ȣ��
		if not gDoErrorRtn ("COMBO_TYPE") then 
			 gLoadComboBox .cmbSEARCHENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbSEARCHJOBGUBN, vntJOBGUBN, False
   		end if    				   		
   	end with     	
End Sub

Sub EndPage()
	set mobjPDCMPREESTLIST = Nothing
	set mobjPDCMGET = Nothing
	set mobjPDCMJOBNO = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		
		'.txtCREDAY.value = gNowDate
		
		
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
	End with
	'DataNewClean
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
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
Sub DateClean2
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM1.value = date1
		.txtTO1.value = date2
	End With
End Sub
'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
  	Dim vntData
  	Dim strRow
	with frmThis
	'On error resume next
  		'������ Validation
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"JOBNO|JOBNAME|PREESTNO|PREESTNAME|AMT|MEMO|CREDAY|CLIENTSUBCODE|SUSURATE|CLIENTCODE|SUBSEQ")
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		if DataValidation =false then exit sub
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .sprSht1.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
		
		intRtn = mobjPDCMPREESTLIST.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ" & intRtn & " �� ����" & mePROC_DONE,"����ȳ�" 
			strRow = .sprSht.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
  		end if
 	end with
End Sub
'****************************************************************************************
' ������ȸ �޷�
'****************************************************************************************
'��ȸ��
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM1,frmThis.imgCalEndarFROM1,"txtFROM1_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo1,frmThis.imgCalEndarTO1,"txtTo1_onchange()"
		gSetChange
	end with
End Sub
'��ȸ��
Sub imgCalEndarFROM_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub

Sub txtFROM1_onchange
	gSetChange
End Sub


Sub txtTo1_onchange
	gSetChange
End Sub
'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	'On error resume next
	with frmThis
		for intCnt = 1 to .sprSht1.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNO",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNAME",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ���۹�ȣ �� �������� �� Ȯ���Ͻʽÿ�","�Է¿���"
				Exit Function
			End if
		next	
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	
	On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		
		
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		'���ݰ�꼭 �Ϸ���ȸ
		vntData = mobjPDCMPREESTLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),Trim(.txtJOBNO.value),Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
			ELSE
				
			End If
			Call SelectRtn_HDR()
			'gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
			'sprShtToFieldBinding 1,1
			Call sprSht_Click(1,1)
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub
',Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value
Sub SelectRtn_HDR ()
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
	strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),Trim(.txtJOBNO.value),Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					Else
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB
Sub SelectRtn_HDR2 ()
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM1.value,1,4) &  MID(.txtFROM1.value,6,2) &  MID(.txtFROM1.value,9,2)
	strTO =  MID(.txtTO1.value,1,4) &  MID(.txtTO1.value,6,2) &  MID(.txtTO1.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR2(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME1.value),Trim(.txtJOBNO1.value),Trim(.txtCLIENTCODE1.value),Trim(.txtCLIENTNAME1.value))
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					
					Else
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB

Sub SelectRtn_DBLHDR (ByVal strJOBNO)
	Dim vntData1
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
	strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
	strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMPREESTLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNAME.value),strJOBNO,Trim(.txtCLIENTSUBNAME.value),Trim(.txtCLIENTSUBCODE.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value)
	
	If not gDoErrorRtn ("SelectRtn_HDR") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht1.MaxRows = 0	
			Else
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",intCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,3,3,true
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,1,1,true
					Else
					'mobjSCGLSpr.SetCellsLock2 .sprSht1, False, "JOBNO|JOBNAME"
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
			End If
	End If	
	End with
End SUB





'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
	Dim strRow
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		
		'PREESTNO,ITEMCODESEQ
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"CONF",vntData(i)) = "Y" Then
				gErrorMsgBox "Ȯ�������� �����ϽǼ� ������, �󼼳������� Ȯ���� ����� �����Ͻʽÿ�.","�����ȳ�"
				Exit Sub
			End if
			intRtn = gYesNoMsgbox("�ڷ�� �󼼳��� �� �Բ� ���� �˴ϴ�. " & vbcrlf & "�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
			IF intRtn <> vbYes then exit Sub
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i)) <> "" Then
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i))
				intRtn = mobjPDCMPREESTLIST.DeleteRtn(gstrConfigXml,strPREESTNO)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				gWriteText "", "[" & strPREESTNO & "] �ڷᰡ �����Ǿ����ϴ�."
   			End IF
		next
		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		strRow = .sprSht.ActiveRow
		SelectRtn
		mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		Call sprSht_Click(1,strRow)
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
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;��������</td>
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
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
										width="85">������</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" size="20" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" size="5" name="txtCLIENTCODE"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
										width="85">JOB��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><INPUT class="INPUT_L" id="txtJOBNAME" title="���۰����� ��ȸ" style="WIDTH: 144px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="18" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="���۰����ڵ� ��ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="7" align="left" size="3" name="txtJOBNO">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="85">�Ƿ�����</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtFROM" title="�Ⱓ�˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
											accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
											border="0" name="imgCalEndarFROM">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgCalEndarTO"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)"
										width="85">�����</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="����θ� ��ȸ" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="20" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT" id="txtCLIENTSUBCODE" title="������ڵ� ��ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtCLIENTSUBCODE"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
										width="85">��ü�ι�</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><SELECT id="cmbSEARCHJOBGUBN" title="��ü�ι���ȸ" style="WIDTH: 232px" name="cmbSEARCHJOBGUBN"></SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="85">�����������</TD>
									<TD class="SEARCHDATA" colSpan="2"><SELECT id="cmbSEARCHENDFLAG" title="�Ϸᱸ��" style="WIDTH: 216px" name="cmbSEARCHENDFLAG"></SELECT></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left"  height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;JOB ����Ʈ</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
						
							<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="11721">
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
					</TR>
					<TR>
						<TD>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left"  height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;û�� ���� ����Ʈ</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgDetail" onmouseover="JavaScript:this.src='../../../images/imgDetailOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDetail.gif'"
														height="20" alt="�ڷ��� �󼼳����� �����մϴ�." src="../../../images/imgDetail.gIF" border="0"
														name="imgDetail"></TD>
												<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
														src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
												<td><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gIF'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imglistcopy.gIF" width="77" border="0"
														name="Imgcopy"></td>
												<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="������ ���������մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgRowDel"></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
												<TD><IMG id="imgExcel1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel1"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody1" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
										width="85">������</TD>
									<TD class="SEARCHDATA" style="WIDTH: 243px"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�����ָ�" style="WIDTH: 152px; HEIGHT: 22px"
											type="text" maxLength="100" size="20" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE1"><INPUT class="INPUT" id="txtCLIENTCODE1" title="�������ڵ�" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" size="5" name="txtCLIENTCODE1"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME1, txtJOBNO1)"
										width="85">JOB��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 234px"><INPUT class="INPUT_L" id="txtJOBNAME1" title="���۰����� ��ȸ" style="WIDTH: 144px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="18" name="txtJOBNAME1"><IMG id="ImgJOBNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgJOBNO1"><INPUT class="INPUT" id="txtJOBNO1" title="���۰����ڵ� ��ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtJOBNO1">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call DateClean2()"
										width="85">�Ƿ�����</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtFROM1" title="�Ⱓ�˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
											accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM1"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
											border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO1" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="7" name="txtTO1"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery1" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery1"></td>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
							<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									 VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="3889">
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
									<PARAM NAME="ReDraw" VALUE="-1">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� 
							�� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="�ݾ�" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
