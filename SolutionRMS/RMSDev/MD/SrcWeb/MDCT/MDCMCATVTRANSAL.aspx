<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVTRANSAL.aspx.vb" Inherits="MD.MDCMCATVTRANSAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ����� ����</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ����� ��� ȭ��(MDCMCATVTRANSAL.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMCATVTRANSAL.aspx
'��      �� : ����Ź�ŷ����� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/28 By Kim Tae Yub
'			 2) 2009/09/11 By HWANG DUCK SU
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCTCATVTRANS
Dim mobjMDCOGET
Dim mobjSCCOGET
Dim mstrCheck, mstrCheck1, mstrGrid

CONST meTAB = 9

mstrCheck=True
mstrCheck1=True
mstrGrid = FALSE

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN(byVal strmode)
	With frmThis
		If  strmode = "EXTENTION"  Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "60%"
			document.getElementById("tblSheet2").style.height = "30%"
		ELSEIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ELSEIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "60%"
		END IF
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
Sub imgQuery_onclick
	IF frmThis.txtYEARMON1.value = "" and frmThis.txtCLIENTCODE1.value = "" then
		gErrorMsgBox "��ȸ������ �Է��Ͻÿ�.","��ȸ�ȳ�"
		Exit Sub
	end if
	
	mstrGrid = FALSE
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�ʱ�ȭ��ư
Sub imgCho_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
Sub ImgCRE_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmRequest_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Confirm_User
	gFlowWait meWAIT_OFF
End Sub


Sub imgALLPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData, vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCTCATVTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
     				 
		ReportName = "MDCMCATVTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = "" 
		vntDataTemp = mobjMDCTCATVTRANS.ProcessRtn_TEMP_ALL(gstrConfigXml, vntData, strUSERID)

		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout_All()", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout_All()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCTCATVTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strCNT
	Dim vntData
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCTCATVTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMCATVTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjMDCTCATVTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout('" & strTRANSYEARMON & "', '" & strTRANSNO & "')", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout(strTRANSYEARMON, strTRANSNO)
	Dim intRtn, intRtn2
	with frmThis
		intRtn = mobjMDCTCATVTRANS.DeleteRtn_temp(gstrConfigXml)
		intRtn2 = mobjMDCTCATVTRANS.DeleteRtnUpdate_PRINTSEQ(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ProcessRtn_Confirm_User ()
   	Dim vntData
   	Dim intRtn
	Dim strTRANSYEARMON, intTRANSNO
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	
	Dim intCnt,intCnt2,intCnt3, chkcnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg
	
	'SMS ����
	Dim vntData_info
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		IF .txtEMPNO.value = "" THEN
			gErrorMsgBox "���ο�û�ڸ� �Է� �Ͻʽÿ�",""
			exit sub
		END IF 
		
		chkcnt=0
		For intCnt = 1 To .sprSht_HDR.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "���ο�û�� �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if
		
		For intCnt2 = 1 To .sprSht_HDR.MaxRows
			'�׸����� ���۰Ǹ� �� �����´�
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt2) = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTNAME",intCnt2)
				 Exit For
			End If
		Next
	
		If chkcnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] ���ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ���ο�û���ֽ��ϴ�"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] ��" & chkcnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ��" & chkcnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			End If
		End If
		
		intSaveRtn = gYesNoMsgbox("�ش絥���͸� ���ο�û �Ͻðڽ��ϱ�?","���ο�û Ȯ��")
		IF intSaveRtn = vbYes then 
		
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
		

			'�����÷��� ����
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meINS_TRANS
			gXMLSetFlag xmlBind, meINS_TRANS

			'��Ʈ�� ����� �����͸� �����´�.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO")
			
			strTRANSYEARMON = .txtYEARMON1.value
			strCLIENTCODE	= .txtCLIENTCODE1.value
			strCLIENTNAME	= .txtCLIENTNAME1.value
			strTIMCODE		= .txtTIMCODE1.value
			strTIMNAME		= .txtTIMNAME1.value
			
			intRtn = mobjMDCTCATVTRANS.ProcessRtn_Confirm_User(gstrConfigXml,vntData,.txtEMPNO.value)
			
   			if not gDoErrorRtn ("ProcessRtn_Confirm_USER") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
				initpagedata	
				gOkMsgBox "���ο�û�� �Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					.txtTIMCODE1.value = strTIMCODE
					.txtTIMNAME1.value = strTIMNAME
					selectRtn
				Else
					initpagedata
				End If
				DateClean
   			end if
   		End If
   	end with
End Sub


'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP ()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(.txtYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "trans", "CATV") 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			IF vntRet(3,0) = "�Ϸ�" THEN
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(4,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			ELSE
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			END IF
			selectRtn
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag ���� �˸�
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
			
			vntData = mobjMDCOGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"","trans", "CATV")
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					selectRtn
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
' ������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'���� ������List ��������
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtYEARMON1.value), trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "trans", "CATV") 
							
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSTIMPOP.aspx",vntInParams , 413,465)
		'TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE
		if isArray(vntRet) then
			.txtYEARMON1.value = trim(vntRet(0,0))
			.txtTIMCODE1.value = trim(vntRet(5,0))
			.txtTIMNAME1.value = trim(vntRet(1,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			.txtCLIENTNAME1.value = trim(vntRet(2,0))
			selectrtn
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTRANSTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												  trim(.txtYEARMON1.value),trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
												  trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "", "trans", "CATV")
			
			if not gDoErrorRtn ("GetTRANSTIMCODE") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = trim(vntData(0,1))
					.txtTIMCODE1.value = trim(vntData(5,1))
					.txtTIMNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(2,1))
					selectrtn
				Else
					Call TIMCODE1_POP()
				End If
   			end if
   		end with
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
		
		vntRet = gShowModalWindow("../MDCO/MDCMEMPPOP.aspx",vntInParams , 413,435)
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
			vntData = mobjMDCOGET.GetMDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			
			if not gDoErrorRtn ("GetMDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
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

'****************************************************************************************
'��ȸ�ʵ� onclick onkeydown 
'****************************************************************************************
Sub chkVOCH_TYPE0_onclick 
	selectrtn
End Sub

Sub chkVOCH_TYPE1_onclick 
	selectrtn
End Sub

Sub chkVOCH_TYPE2_onclick 
	selectrtn
End Sub

Sub chkVOCH_TYPE3_onclick 
	selectrtn
End Sub

Sub txtYEARMON1_onkeydown
	'or window.event.keyCode = meTAB ���϶��� �ƴ� �����϶��� ��ȸ
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' ������ �޷�
'****************************************************************************************
Sub imgCalDemandday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'-----------------------------------------------------------------------------------------
' û������ ����
'-----------------------------------------------------------------------------------------
Sub txtYEARMON1_onblur
	With frmThis
		If .txtYEARMON1.value <> "" AND Len(.txtYEARMON1.value) = 6 Then DateClean
	End With
End Sub

'û�����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************

Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		elseif Row > 0 AND Col > 1 then
			mstrGrid = TRUE
			CALL Grid_Setting ()
			SelectRtn_DTL Col, Row
			'mstrGrid = false
		end if
	end with
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		IF mstrGrid = FALSE THEN
			if Row = 0 and Col = 1 then
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_DTL, 1, 1, , , "", , , , , mstrCheck1
				if mstrCheck1 = True then 
					mstrCheck1 = False
				elseif mstrCheck1 = False then 
					mstrCheck1 = True
				end if
				for intcnt = 1 to .sprSht_DTL.MaxRows
					sprSht_DTL_Change 1, intcnt
				next
			end if
		END IF
	end with
End Sub  


sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub


Sub sprSht_HDR_Keyup(KeyCode, Shift)
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
		mstrGrid = TRUE
		CALL Grid_Setting ()
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
			.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") Then
				strCOLUMN = "VAT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
				strCOLUMN = "SUMAMTVAT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))
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

Sub sprSht_DTL_Keyup(KeyCode, Shift)
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

	With frmThis
		If mstrGrid Then
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
				.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
					strCOLUMN = "SUMAMTVAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT")) Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
					
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			End If
		else
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") OR _
			   .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") THEN

				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
					strCOLUMN = "SUMAMTVAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") Then
					strCOLUMN = "COMMISSION"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION")) Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
							End If
						Next
					End If
				Next
					
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			End If
		end if
	End With
End Sub


Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_HDR.MaxRows >0 Then
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
				.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
				If .sprSht_HDR.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)
					
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
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,strCol,vntData_row(j))
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


Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
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
		If mstrGrid Then
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION")    Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		ELSE
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or _
			       .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") THEN
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		END IF
		
	End With
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim i
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strSEQ
	Dim strPRINT_SEQ
	Dim intRtn

	with frmThis
		If mstrGrid Then
			'���� �׸��尡 ���ݰ�꼭 �󼼳����϶� ��¼����� ���Ҷ� �߻�
			if Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRINT_SEQ") then
				for i=1 to .sprSht_DTL.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) <> "" then
						if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) = 0 then
							mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
						else
							
							if Row <> i then
								'�Է��Ѽ��ڿ� ���� ��¼����� ���ִٸ� ����
								if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row) = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) then
									gErrorMsgBox "��¼����� �ߺ��ԷµǾ����ϴ�.",""
									mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
									.txtCLIENTNAME1.focus() 
									.sprSht_DTL.focus()
									EXIT SUB
								end if
							end if
						end if
					end if
				next
				
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSYEARMON",Row)
				strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSNO",Row)
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",Row)
				strPRINT_SEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row)
				
				intRtn = mobjMDCTCATVTRANS.UPDATE_PRINTSEQ(gstrConfigXml,strTRANSYEARMON,strTRANSNO, strSEQ, strPRINT_SEQ)
			end if
		END IF
	end with	
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	'����������ü ����	
	set mobjMDCTCATVTRANS	= gCreateRemoteObject("cMDCT.ccMDCTCATVTRANS")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'******************************************************************
		'�ŷ����� ��� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 16, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | CONFIRM_USER | MEMO | CNT | MED_FLAG"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "����|����|��꼭|������|��ü����|���ް���|�ΰ�����|�հ�ݾ�|û����|������|�ŷ����|��ȣ|������|���|�����|��ü�����ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "  4|   4|     6|    15|       8|      12|      11|      12|     9|     9|       8|   5|    10|  14|       0|          0"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY | PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TRANSNO | AMT | VAT | SUMAMTVAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | TRANSYEARMON | CONFIRM_USER | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | CONFIRM_USER | MEMO | MED_FLAG"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "MED_FLAG|CNT", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | MED_FLAGNAME | CONFIRM_USER" ,-1,-1,2,2,false
		.sprSht_HDR.style.visibility = "visible"
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCTCATVTRANS = Nothing
	set mobjMDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


Sub Grid_Setting ()
	With frmThis
		mobjSCGLCtl.DoEventQueue
		If mstrGrid Then
		
			'Sheet �⺻Color ����
			gSetSheetDefaultColor() 
			'******************************************************************
			''�ŷ����� ������
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 29, 0, 0, 2
			mobjSCGLSpr.SpreadDataField .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | SUBSEQNAME | TIMNAME | MATTERNAME | PROGRAM | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON | REAL_MED_BUSINO"
			mobjSCGLSpr.SetHeader .sprSht_DTL,		"�ŷ������|�ŷ�����ȣ|����|����|��Ź����|������|��ü��|��ü��|�귣��|��|�����|���α׷�|������|������|�ݾ�|�ΰ���|��|������|��������|��ü����|���|�μ���|û����|������|��꼭���|��꼭��ȣ|Ȯ������|��Ź���|��ü�����ڹ�ȣ"
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "        0|	         0|	  0|   4|       4|	  15|    15|    13|    10|10|    15|       8|     9|    9|   10|    10|11|    10|       6|       7|  15|     0|     0|     0|	       9|         9|       0|      0|              10"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "TBRDSTDATE | TBRDEDDATE | DEMANDDAY | PRINTDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "PRINT_SEQ | TRUST_SEQ | AMT | VAT | SUMAMTVAT | COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | SUBSEQNAME | TIMNAME | MATTERNAME | MED_FLAGNAME | MEMO | DEPT_NAME | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON | REAL_MED_BUSINO", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, false, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | SUBSEQNAME | TIMNAME | MATTERNAME | PROGRAM | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON | REAL_MED_BUSINO " 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "TRANSYEARMON | TRANSNO | SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | SUBSEQNAME | TIMNAME | MATTERNAME | PROGRAM | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON | REAL_MED_BUSINO" 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | SUBSEQNAME | TIMNAME | MATTERNAME | PROGRAM | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | CONFIRMFLAG | TRUST_YEARMON | REAL_MED_BUSINO", FALSE
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, " TRUST_YEARMON|TRUST_SEQ | MED_FLAGNAME ",-1,-1,2,2,false
		
		Else
		
			'Sheet �⺻Color ����
			gSetSheetDefaultColor() 
			'******************************************************************
			'û�೻�� �׸���
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 40, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM | MEMO | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |TRANSCUSTRANK "
			mobjSCGLSpr.SetHeader .sprSht_DTL,		 "����|���|û����|G|����|�������ڵ�|�����ָ�|����ó�ڵ�|����ó��|���ڵ�|����|�귣���ڵ�|�귣��|ä���ڵ�|ä��|��ü���ڵ�|��ü���|�����ڵ�|�����|�ݾ�|�ΰ���|��|������|��������|������|������|MPP�ڵ�|MPP��|���α׷�|���|���μ��ڵ�|���μ�|���ۻ��ڵ�|���ۻ��|CIC/������ڵ�|CIC/�����|�ʼ�|��ǥ����|VAT|TRANSCUSTRANK"		
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 4|   6|     8|0|   0|         8|      15|         0|      15|     0|  12|         0|    12|       0|  12|         0|      12|       0|    12|  10|    10|10|    10|       8|    9|      9|      0|    0|       0| 20|           0|     0|          0|      0|             0|        0|   0|       0|  0|  0|            0"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK|TRU_TAX_FLAG "
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE ", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "AMT | VAT | SUMAMTVAT | COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "YEARMON | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM  | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |   TRANSCUSTRANK", -1, -1, 100   
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, false, "CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM | MEMO | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |  TRANSCUSTRANK" 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM | MEMO | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |  TRANSCUSTRANK" 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM | MEMO | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |TRANSCUSTRANK  ", FALSE
			mobjSCGLSpr.ColHidden .sprSht_DTL, " SEQ | CLIENTCODE | GREATCODE | MEDCODE | REAL_MED_CODE | SUBSEQ | TIMCODE | MATTERCODE | DEPT_CD |MPP_CODE | MPP_NAME | PROGRAM| GFLAG | EXCLIENTCODE  | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |  TRANSCUSTRANK", True
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "YEARMON | CNT | TBRDSTDATE | TBRDEDDATE| GFLAG",-1,-1,2,2,false '���
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CLIENTNAME | GREATNAME | TIMNAME | SUBSEQNAME | MEDNAME | REAL_MED_NAME | CLIENTSUBNAME | MATTERNAME | DEPT_NAME | MPP_NAME | EXCLIENTNAME | MEMO | PROGRAM",-1,-1,0,2,false '����
		End If
		
		.sprSht_DTL.style.visibility = "visible"
	End With
End Sub


'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		DateClean

		.txtPRINTDAY.value  = gNowDate
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		.chkVOCH_TYPE0.checked = TRUE
		.chkVOCH_TYPE1.checked = TRUE
		.chkVOCH_TYPE2.checked = TRUE
		.chkVOCH_TYPE3.checked = TRUE
		
		CALL Grid_Setting ()
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON1.value,1,4) & "-" & MID(frmThis.txtYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON, strDEMANDYEARMON
	Dim strCLIENTCODE, strTIMCODE
   	Dim i, strCols
   	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
   	Dim strVOCH_TYPE
    
	'On error resume next
	with frmThis
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strTIMCODE		= .txtTIMCODE1.value
		
		CALL Grid_Setting()
		
		vntData = mobjMDCTCATVTRANS.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strCLIENTCODE, strTIMCODE)
		
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   		
   		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF .chkVOCH_TYPE0.checked = TRUE THEN
			strVOCH_TYPE = strVOCH_TYPE & "0"
		END IF
		
		IF .chkVOCH_TYPE1.checked = TRUE THEN
			IF strVOCH_TYPE = "" THEN
				strVOCH_TYPE = strVOCH_TYPE & "1"
			ELSE 
				strVOCH_TYPE = strVOCH_TYPE & ",1"
			END IF
		END IF
		
		IF .chkVOCH_TYPE2.checked = TRUE THEN
			IF strVOCH_TYPE = "" THEN
				strVOCH_TYPE = strVOCH_TYPE & "2"
			ELSE 
				strVOCH_TYPE = strVOCH_TYPE & ",2"
			END IF
		END IF
		
		IF .chkVOCH_TYPE3.checked = TRUE THEN
			IF strVOCH_TYPE = "" THEN
				strVOCH_TYPE = strVOCH_TYPE & "3"
			ELSE 
				strVOCH_TYPE = strVOCH_TYPE & ",3"
			END IF
		END IF
		
		vntData2 = mobjMDCTCATVTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												strYEARMON, strCLIENTCODE, strTIMCODE, strVOCH_TYPE)
   		
		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData2,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				
   				For i = 1 to .sprSht_DTL.MaxRows
					IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"VOCH_TYPE",i) = "3" THEN
						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, i, i,&HCCFFFF, &H000000,False
						'mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,TRUE,i,1,-1,true
					END IF 
				next
   			else
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   		End If

   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)
				
		vntData = mobjMDCTCATVTRANS.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strTRANSYEARMON, strTRANSNO)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = false
   		End If
   	end with
End Sub


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht_DTL.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	chkcnt = 0
	
	with frmThis
		if mstrGrid then exit sub
		
		intColFlag = 0
		
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'�׷��ִ밪 ����
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSCUSTRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if

		'�����÷��� ����

		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME |  REAL_MED_CODE | REAL_MED_NAME | MATTERCODE | MATTERNAME | AMT | VAT | SUMAMTVAT | COMMISSION | COMMI_RATE | TBRDSTDATE | TBRDEDDATE | MPP_CODE | MPP_NAME | PROGRAM | MEMO | DEPT_CD | DEPT_NAME |  EXCLIENTCODE | EXCLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME  | CNT | VOCH_TYPE| TRU_TAX_FLAG |  TRANSCUSTRANK")
	
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		
		intRtn = mobjMDCTCATVTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON,intColFlag)
   		
   		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				.txtTIMCODE1.value = strTIMCODE
				.txtTIMNAME1.value = strTIMNAME
				selectRtn
			Else
				initpagedata
			End If
			DateClean
   		end if

   	end with
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
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "�������� �ʼ� �Է� ���� �Դϴ�.",""
			Exit Function
		End If
  	End with
	DataValidation = true
End Function


'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	Dim strPRINTDAY
   	Dim strMED_FLAG
   	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
   	Dim strCLIENTCODE, strCLIENTNAME
   	Dim lngchkCnt
   	
	with frmThis
		strDESCRIPTION = ""

		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "������ ������ �����ϴ�.","�����ȳ�!"
			Exit Sub
		END IF
		
		For i = 1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSNO",i)
				
				'DB�� ������ Ȯ�����ؾ� �������� ���������� �ߺ��������� ����ġ���� ������ ���´�.(����)
				'����� TRANSYEARMON�� TRANSNO�� DTL�� �󼼳���(�ŷ����� �߻�����)�� �����ϴ��� ��ȸ
				vntData = mobjMDCTCATVTRANS.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO) 
				If mlngRowCnt > 0 Then
					gErrorMsgBox i & "���� �ŷ������� ���ݰ�꼭�� �߻��� �󼼳����� �����մϴ�.","�����ȳ�!"
					Exit Sub
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
				
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0

		'�׸����� �÷��׸� INSERT�� �ٲ۴�.  �ʱⰪ�� UPDATE...
		mobjSCGLSpr.SetFlag  .sprSht_HDR, meINS_TRANS
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO ")
		
		intRtn = mobjMDCTCATVTRANS.DeleteRtn(gstrConfigXml,vntData)

		IF not gDoErrorRtn ("DeleteRtn") then
			'���õ� �ڷḦ ������ ���� ����
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "�ŷ������� �����Ǿ����ϴ�.","�����ȳ�!"
			if .sprSht_HDR.MaxRows > 0 then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, 1,1
				mstrGrid = true
				CALL Grid_Setting ()
				SelectRtn_DTL 1,1
			else
				mstrGrid = FALSE
				SelectRtn
			end if
   		End IF
	End with
	err.clear	
End Sub



-->
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
													<TABLE cellSpacing="0" cellPadding="0" width="220" background="../../../images/back_p.gIF"
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
												<td class="TITLE">û����� - �ŷ����� ����/��ȸ/����</td>
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
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 12px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1,'')"
													width="60">û�����</TD>
												<TD class="SEARCHDATA" width="130"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">������</TD>
												<TD class="SEARCHDATA" width="280"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 203px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="27" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
													<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1,txtTIMCODE1)"
													width="60">��</TD>
												<TD class="SEARCHDATA" width="280"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 203px; HEIGHT: 22px" type="text"
														maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTIMCODE1"></TD>
												<TD class="SEARCHDATA">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
												</TD>
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
												<TD class="TITLE" style="WIDTH: 100%; HEIGHT: 8px" vAlign="absmiddle"></TD>
											</TR>
											<TR>
												<TD class="TITLE" width="210" vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_normal.gif'
															align='absMiddle' border='0' name='btn_normal'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')">
														<IMG id='btn_multi' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_multi.gif'
															align='absMiddle' border='0' name='btn_multi'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')">
														<IMG id='btn_hide' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_hide.gif'
															align='absMiddle' border='0' name='btn_hide'></span>
												</TD>
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
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD style="FONT-WEIGHT: bold; FONT-SIZE: 12px"><span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">������:</span>
																&nbsp;<INPUT class="INPUT_L" id="txtEMPNAME" title="�����ȸ" style="WIDTH: 62px; HEIGHT: 22px" type="text"
																	maxLength="255" align="left" size="5" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
																	name="ImgEMPNO" title="���α��ڼ���"> <INPUT class="INPUT" id="txtEMPNO" title="�����ȸ" style="WIDTH: 46px; HEIGHT: 22px" readOnly
																	type="text" maxLength="8" align="left" size="2" name="txtEMPNO">&nbsp; <IMG id="ImgConfirmRequest" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'" height="20" alt="���õ� �ŷ������� ���ο�û�մϴ�." src="../../../images/ImgConfirmRequest.gIF"
																	align="absMiddle" border="0" name="ImgConfirmRequest"></TD>
															<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="ȭ���� �ʱ�ȭ �մϴ�."
																	src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="���õ� �ŷ������� �����մϴ�." src="../../../images/imgDelete.gIF" border="0"
																	name="imgDelete"></TD>
															<TD><IMG id="imgALLPrint" onmouseover="JavaScript:this.src='../../../images/imgALLPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgALLPrint.gif'"
																	height="20" alt="��ȸ�� �ŷ������� ��ü�μ��մϴ�.." src="../../../images/imgALLPrint.gIF" border="0"
																	name="imgALLPrint"></TD>
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
								<TR id="tblBody1">
									<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="4233">
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
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD class="TITLE" vAlign="absmiddle" align="left" width="400" height="22">&nbsp; 
													����Ź <INPUT id="chkVOCH_TYPE0" title="����Ź" type="checkbox" name="chkVOCH_TYPE0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
													���� <INPUT id="chkVOCH_TYPE1" title="����" type="checkbox" name="chkVOCH_TYPE1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
													�Ϲ� <INPUT id="chkVOCH_TYPE2" title="�Ϲ�" type="checkbox" name="chkVOCH_TYPE2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
													AOR  <INPUT id="chkVOCH_TYPE3" title="AOR"  type="checkbox" name="chkVOCH_TYPE3"></TD>
												<TD class="TITLE" align="left" width="400" height="22">û������ : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 100px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="32" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">&nbsp;&nbsp;&nbsp;&nbsp; 
													�������� : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 94px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalPrintday"></TD>
												<TD vAlign="middle" align="right" height="22">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgCRE" onmouseover="JavaScript:this.src='../../../images/ImgCREOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/ImgCRE.gif'" alt="CIC���� �ŷ������� �����մϴ�."
																	src="../../../images/ImgCRE.gif" border="0" name="ImgCRE"></TD>
															<!--<TD><IMG id="ImgCustSave" onmouseover="JavaScript:this.src='../../../images/ImgCustSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgCustSave.gIF'"
																	height="20" alt="�����ֺ��� �ŷ������� �����մϴ�.." src="../../../images/ImgCustSave.gIF" border="0"
																	name="ImgCustSave"></TD>-->
															<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="���� �ŷ������� ����մϴ�.." src="../../../images/imgPrint.gIF" border="0"
																	name="imgPrint"></TD>
															<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR id="tblBody2">
									<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="8625">
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
								<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE>
		</FORM>
		<iframe id="frmSMS" style="WIDTH: 500px;HEIGHT: 500px; DISPLAY: none;" name="frmSMS"></iframe> <!-- -->
	</body>
</HTML>
