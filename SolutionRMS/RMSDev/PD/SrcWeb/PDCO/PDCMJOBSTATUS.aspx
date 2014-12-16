<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBSTATUS.aspx.vb" Inherits="PD.PDCMJOBSTATUS" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB ������Ȳ��ȸ</title>
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
Dim mobjPDCOJOB
Dim mobjPDCMGET
Dim mobjSCCOGET
Dim mobjPDCODEMAND
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mstrDEPTCD
Dim mstrMANAGER

Dim mstrSelectCHK
mstrSelectCHK = "SELECT"
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
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub imgDivReDemand_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub ImgMine_onclick()
	gFlowWait meWAIT_ON
	SelectRtn_Mine
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------





sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'����������ü ����	
	Set mobjPDCOJOB = gCreateRemoteObject("cPDCO.ccPDCOJOB")
	Set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET") 
	set mobjPDCODEMAND	= gCreateRemoteObject("cPDCO.ccPDCODEMAND")
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 19, 0, 5
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK|REQDAY|JOBNO|SEQ|JOBNAME|CLIENTCODE|CLIENTNAME|JOBGUBN|CONFIRM|MEMO|TRANS|TAX|VOCH|EMPNAME|DEPTNAME|MANAGER|EMPNO|DEPTCD|RANKJOB"
		mobjSCGLSpr.SetHeader .sprSht,		    "����|��Ͽ�|JOBNO|SUBNO|JOB��|CLIENTCODE|������|���ۺι�|����|���|�ŷ�����|���ݰ�꼭|��ǥ|�����|���μ�|���δ��|���|�μ��ڵ�|GROUP"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "4   |10    |7    |6    |25   |0         |20    |8       |4   |6   |10        |10        |10  |8     |18      |8       |0   |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY"
		'mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT|VAT|SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|SEQ|JOBGUBN|CONFIRM|EMPNAME|MANAGER",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|MEMO|TRANS|TAX|VOCH|DEPTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"REQDAY|JOBNO|SEQ|JOBNAME|CLIENTCODE|CLIENTNAME|JOBGUBN|CONFIRM|MEMO|TRANS|TAX|VOCH|EMPNAME|DEPTNAME|MANAGER|EMPNO|DEPTCD|RANKJOB|CHK"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|EMPNO|DEPTCD", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"JOBNO"
	.cmbGBN.selectedIndex = 0
	End with
	DateClean
	pnlTab1.style.visibility = "visible" 

End Sub
'-----------------------------------------------------------------------------------------
' �˾�(��ȸ)
'-----------------------------------------------------------------------------------------
'-----------------------------------
' ������ �� JOBNO �˾� ��ư[��ȸ��]
'------------------------------------
Sub ImgCLIENTCODE1_onclick
	with frmThis
		IF .cmbSEARCH.value = "1" then
			Call CLIENTCODE1_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	End With
End Sub

'������ - ���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��		
     	end if
	End with
	SelectRtn
	gSetChange
End Sub

'JOBNO - ���� ������List ��������
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array( trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'������ �Ǵ� JOBNO �Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			if .cmbSEARCH.value = "1" Then '������Ʈ �ڵ� ���
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
				if not gDoErrorRtn ("GetHIGHCUSTCODE") then
					If mlngRowCnt = 1 Then
						.txtCLIENTCODE1.value = trim(vntData(0,1))
						.txtCLIENTNAME1.value = trim(vntData(1,1))
					Else
						Call CLIENTCODE1_POP()
					End If
   				end if
   			Else
   				vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value))
				
				if not gDoErrorRtn ("GetJOBNO") then
					If mlngRowCnt = 1 Then
						.txtCLIENTCODE1.value = trim(vntData(0,0))
						.txtCLIENTNAME1.value = trim(vntData(1,0))
					Else
						Call SEARCHJOB_POP()
					End If
   				end if
   			
   			End If
   		end with
   		SelectRtn
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' ����ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'���� ������List ��������
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../../../PD/SrcWeb/PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()				' ��Ŀ�� �̵�
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtEMPNAME
			
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetPDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					'.txtMEMO.focus()
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,4) & Mid(strDATE,6,2)
	date2 = Mid(strDATE,1,4) & Mid(strDATE,6,2) 

	with frmThis
		.txtFrom.value = date1
		.txtTo.value = date2
	End With
End Sub
Sub EndPage()
	set mobjPDCMSEARCH = Nothing
	set mobjSCCOGET = Nothing
	Set mobjPDCMGET = Nothing
	Set mobjPDCODEMAND = Nothing
	
	gEndPage	
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub
'=========================================================================================
'������ ó��
'=========================================================================================
Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strFROM
    Dim strTO
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = Trim(.txtFROM.value) 
		strTO =  Trim(.txtTO.value) 
		
		vntData = mobjPDCOJOB.SelectRtn_Status(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value,.cmbGBN.value,.cmbSEARCH.value)
		
		if not gDoErrorRtn ("SelectRtn_Status") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows '��ȸ�� ������ ó������ ������ ���鼭
						'JOB�� �÷� ����
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKJOB",intCnt) Mod 2 = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRM",intCnt) = "����" Then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),intCnt,intCnt,,,,,,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"CHK",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK",intCnt,intCnt,4,,,,,False
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CHK",intCnt,intCnt,false
					End If
				Next
   			Else
   				.sprSht.MaxRows = 0	
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
   	mstrSelectCHK = "SELECT"
End Sub

Sub SelectRtn_Mine ()
   	Dim vntData
   	Dim i, strCols
    Dim strFROM
    Dim strTO
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = Trim(.txtFROM.value) 
		strTO =  Trim(.txtTO.value) 
		
		vntData = mobjPDCOJOB.SelectRtn_Mine(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrDEPTCD,mstrMANAGER)
		
		if not gDoErrorRtn ("SelectRtn_Status") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows '��ȸ�� ������ ó������ ������ ���鼭
						'JOB�� �÷� ����
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKJOB",intCnt) Mod 2 = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRM",intCnt) = "����" Then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),intCnt,intCnt,,,,,,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"CHK",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK",intCnt,intCnt,4,,,,,False
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CHK",intCnt,intCnt,false
					End If
				Next
   			Else
   				.sprSht.MaxRows = 0	
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
   	mstrSelectCHK = "MINE"
End Sub

Sub ProcessRtn
	Dim vntData
	Dim intRtn
	Dim strSAVEGBN
	Dim intCnt,intCnt2,intCnt3,intMsgCnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg
	
	Dim dblChk
	
	'SMS ����
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		
		
		
		'��Ʈ�� ����� �����͸� �����´�.
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "���α��ڸ� ���� �Ͻʽÿ�.","���û�ȳ�!"
			Exit Sub
		End If
		
		dblChk = 0
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				
				dblChk = dblChk + 1
				mobjSCGLSpr.CellChanged .sprSht, 1, intCnt	
			End If
		Next
		
		if dblChk = 0 then
			gErrorMsgBox "û�� ���û�� �����͸� üũ �Ͻʽÿ�","���û�ȳ�!"
			exit sub
		end if
		
		
		'���α��� �� �׸��忡 ž��
		intMsgCnt = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			'�׸����� ���۰Ǹ� �� �����´�
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1" Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt2)
				 intMsgCnt = intMsgCnt +1
			End If
		Next
	
	
		If intMsgCnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] ���ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ���ο�û���ֽ��ϴ�"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] ��" & intMsgCnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ��" & intMsgCnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			End If
		End If
		
		intSaveRtn = gYesNoMsgbox("�ش絥���͸� û�����û SMS�߼� �Ͻðڽ��ϱ�?","���û�ȳ�!")
		IF intSaveRtn <> vbYes then 
			'���
		Else
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|JOBNO|SEQ")
		
			intRtn = mobjPDCODEMAND.ProcessRtn_ReDemand(gstrConfigXml,vntData,Trim(.txtEMPNO.value))
			If not gDoErrorRtn ("ProcessRtn_ReDemand") Then
				
				'������ �����Ͽ����Ƿ� SMS �߼�
				'������ ����� ���� ��������
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				
				vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
				
				'�����»������
				strFromUserName		= vntData_info(0,2)
				strFromUserEmail	= vntData_info(1,2)
				strFromUserPhone	= vntData_info(2,2)
				
				'�޴»�� ����
				strToUserName		=  vntData_info(0,1)
				strToUserEmail		=  vntData_info(1,1)
				strToUserPhone		=  vntData_info(2,1)
			
				
				
				
				call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				
				gOkMsgBox "���û�Ǿ����ϴ�.","���û�ȳ�!"
				.txtEMPNAME.value = ""
				.txtEMPNO.value = ""
				If mstrSelectCHK = "MINE" Then
					SelectRtn_Mine
				Else
					SelectRtn
				End IF
			End If
			
		End If
		
		
		
		
		
	End with

End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
	.sprSht.maxrows = 0
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	vntData = mobjPDCODEMAND.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
	
		if mlngRowCnt > 0 Then
		mstrDEPTCD = vntData(0,1)
		mstrMANAGER = vntData(1,1)
		end if
   	end if	
	End with
End Sub


		</script>
		<script language="javascript">
		//SMS �߼�
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
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
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="105" background="../../../images/back_p.gIF"
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
											<td class="TITLE">JOB ������Ȳ��ȸ</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 535px; HEIGHT: 28px" cellSpacing="0" cellPadding="0"
										width="535" border="0">
										<TR>
											<td class="TITLE" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)"
												vAlign="middle" align="center" width="50">������:</td>
											<td><INPUT class="NOINPUTB_L" id="txtEMPNAME" title="���α���" style="WIDTH: 96px; HEIGHT: 20px"
													type="text" maxLength="100" size="10" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													title="���α��ڼ���" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgEMPNO"> <INPUT class="NOINPUTB" id="txtEMPNO" title="���α��ڻ��" style="WIDTH: 58px; HEIGHT: 20px"
													type="text" maxLength="100" size="4" name="txtEMPNO"></td>
											<TD><IMG id="imgDivReDemand" onmouseover="JavaScript:this.src='../../../images/imgDivReDemandOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDivReDemand.gIF'"
													height="20" alt="���������� ���Ͽ� SMS �� ������ �մϴ�." src="../../../images/imgDivReDemand.gif"
													width="87" align="absMiddle" border="0" name="imgDivReDemand">
											</TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="ImgMine" onmouseover="JavaScript:this.src='../../../images/ImgMineOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/ImgMine.gIF'" height="20" alt="���μ� �� ����� �ڷḦ �˻��մϴ�."
													src="../../../images/ImgMine.gIF" width="100" border="0" name="ImgMine"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" width="1040"
							border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2"><FONT face="����">
									<TABLE class="SEARCHDATA"  id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()" width="90">��Ͽ�
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 163px"><INPUT class="INPUT" id="txtFROM" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtFROM)" size="6" name="txtFROM">&nbsp;~
												<INPUT class="INPUT" id="txtTO" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtTO)" size="6" name="txtTO"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1,txtCLIENTCODE1)"
												width="84"><SELECT id="cmbSEARCH" title="�����,������" style="WIDTH: 88px" name="cmbSEARCH">
													<OPTION value="1" selected>������</OPTION>
													<OPTION value="2">JOBNO</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA" style="WIDTH: 282px" width="282"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="��ȸ�뱤���ָ�" style="WIDTH: 200px; HEIGHT: 22px"
													type="text" maxLength="100" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="��ȸ�뱤�����ڵ�" style="WIDTH: 57px; HEIGHT: 22px"
													type="text" maxLength="7" size="4" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" width="90">����</TD>
											<TD class="SEARCHDATA"><SELECT id="cmbGBN" title="���±���" style="WIDTH: 104px; HEIGHT: 22px" name="cmbGBN">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="�Ƿ�">�Ƿ�</OPTION>
													<OPTION value="����">����</OPTION>
													<OPTION value="����">����</OPTION>
													<OPTION value="����">����</OPTION>
													<OPTION value="û��">û��</OPTION>
												</SELECT></TD>
										</TR>
									</TABLE>
									</FONT>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="HEIGHT: 2px"></TD>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 98%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 98%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="42466">
											<PARAM NAME="_ExtentY" VALUE="15849">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
		<iframe id="frmSMS" style="DISPLAY: none; WIDTH: 0px; HEIGHT: 0px" name="frmSMS"></iframe>
	</body>
</HTML>
