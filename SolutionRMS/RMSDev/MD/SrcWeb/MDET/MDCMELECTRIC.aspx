<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRIC.aspx.vb" Inherits="MD.MDCMELECTRIC" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�����ı����Ź ���α׷�����</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : �����ı��� ��Ź
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMELECTRIC.aspx
'��      �� : �����Ľ�Ź����Ÿ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/18 By Ȳ����
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
Dim mobjMDETELEC_TRAN, mobjMDCMGET '�����ڵ�, Ŭ����
Dim mobjMDCMCODETR
Dim mstrHIDDEN
Dim mstrPROCESS	'�ű��̸� True ��ȸ�� False


CONST meTAB = 9
mstrHIDDEN = 0
mstrPROCESS = False

'=============================
' �̺�Ʈ ���ν��� 
'=============================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody1").style.display = "inline"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody1").style.display = "none"
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
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�űԹ�ư
Sub imgNEW_onclick ()
	Call sprSht_Keydown(meINS_ROW, 0)	
	mstrPROCESS = False
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

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		'mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' ���������Ѵ�.
'-----------------------------------------------------------------------------------------
Sub Imgcopy_onclick ()
	Dim intRtn
   	Dim vntData
	Dim intSelCnt,  i
	
	Dim strCHK, strYEARMON, strSEQ, strMATTERNAME
	Dim strMEDCODE, strMEDNAME,strMATTERCODE, strTBRDSTDATE, strTBRDEDDATE
	Dim intCNT, intPRICE, intAMT, strCLIENTCODE, strCLIENTNAME, strSUBSEQ, strSUBSEQNAME, strCLIENTSUBCODE
	Dim strCLIENTSUBNAME, strTIMCODE,  strTIMNAME, strDEPT_CD,  strDEPT_NAME,  strREAL_MED_CODE,  strREAL_MED_NAME
	Dim strPROGRAM, strSTD, strEXCLIENTCODE, strEXCLIENTNAME, strINPUT_MEDFLAG, strSPONSOR,  strROLLSTDATE, strROLLEDDATE
	Dim strBRDSTTIME, strBRDEDTIME, strTYPHOUR, strCMLAN, strBRDMON, strBRDTUE, strBRDWED, strBRDTHU, strBRDFRI, strBRDSAT
	Dim strBRDSUN, strADLOCALFLAG, strBRDDIV, strADSTOCFLAG, strINPUT_AREAFLAGNAME, intCOMMI_RATE, intCOMMISSION,  strVOCH_TYPE
	
	With frmThis
		intSelCnt = 0
		
		Dim strCNT, strCNT2
		strCNT2 = 0
		For i=1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				strCNT = i
				strCNT2 = strCNT2 +1
			End If
		Next
		If strCNT2 >1 Then
			gErrorMsgBox "��������� �ѰǸ� �����մϴ�.",""
			Exit Sub
		elseif strCNT2 =0 Then
			gErrorMsgBox "���������� �ο츦 �����Ͻÿ�.",""
			Exit Sub
		elseif strCNT2 = 1 Then
			If mstrPROCESS Then
				for i = .sprSht.MaxRows to 1 step -1
					If strCNT = i Then
					else 
						mobjSCGLSpr.DeleteRow .sprSht,i
					End If
				Next
			End If
		End If
		
		
		strYEARMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON ",.sprSht.ActiveRow)
		strMEDNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME ",.sprSht.ActiveRow)
		strMATTERNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME ",.sprSht.ActiveRow)
		strMEDCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE ",.sprSht.ActiveRow)
		strMATTERCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE ",.sprSht.ActiveRow)
		strTBRDSTDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE ",.sprSht.ActiveRow)
		strTBRDEDDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE ",.sprSht.ActiveRow)
		intCNT				=	mobjSCGLSpr.GetTextBinding(.sprSht,"CNT ",.sprSht.ActiveRow)
		intPRICE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE ",.sprSht.ActiveRow)
		intAMT				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT ",.sprSht.ActiveRow)
		strCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE ",.sprSht.ActiveRow)
		strCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME ",.sprSht.ActiveRow)
		strSUBSEQ			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ ",.sprSht.ActiveRow)
		strSUBSEQNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME ",.sprSht.ActiveRow)
		strCLIENTSUBCODE	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE ",.sprSht.ActiveRow)
		strCLIENTSUBNAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME ",.sprSht.ActiveRow)
		strTIMCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE ",.sprSht.ActiveRow)
		strTIMNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME ",.sprSht.ActiveRow)
		strDEPT_CD			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD ",.sprSht.ActiveRow)
		strDEPT_NAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME ",.sprSht.ActiveRow)
		strREAL_MED_CODE	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE ",.sprSht.ActiveRow)
		strREAL_MED_NAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME ",.sprSht.ActiveRow)
		strPROGRAM			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM ",.sprSht.ActiveRow)
		strSTD				=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD ",.sprSht.ActiveRow)
		strEXCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE ",.sprSht.ActiveRow)
		strEXCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME ",.sprSht.ActiveRow)
		strINPUT_MEDFLAG	=	mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_MEDFLAG ",.sprSht.ActiveRow)
		strSPONSOR			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR ",.sprSht.ActiveRow)
		strROLLSTDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLSTDATE ",.sprSht.ActiveRow)
		strROLLEDDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLEDDATE ",.sprSht.ActiveRow)
		strBRDSTTIME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSTTIME ",.sprSht.ActiveRow)
		strBRDEDTIME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDEDTIME ",.sprSht.ActiveRow)
		strTYPHOUR			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TYPHOUR ",.sprSht.ActiveRow)
		strCMLAN			=	mobjSCGLSpr.GetTextBinding(.sprSht,"CMLAN ",.sprSht.ActiveRow)
		strBRDMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDMON ",.sprSht.ActiveRow)
		strBRDTUE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTUE ",.sprSht.ActiveRow)
		strBRDWED			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDWED ",.sprSht.ActiveRow)
		strBRDTHU			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTHU ",.sprSht.ActiveRow)
		strBRDFRI			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDFRI ",.sprSht.ActiveRow)
		strBRDSAT			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSAT ",.sprSht.ActiveRow)
		strBRDSUN			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSUN ",.sprSht.ActiveRow)
		strADLOCALFLAG		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ADLOCALFLAG ",.sprSht.ActiveRow)
		strBRDDIV			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDDIV ",.sprSht.ActiveRow)
		strADSTOCFLAG		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ADSTOCFLAG ",.sprSht.ActiveRow)
		strINPUT_AREAFLAGNAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_AREAFLAGNAME ",.sprSht.ActiveRow)
		intCOMMI_RATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE ",.sprSht.ActiveRow)
		intCOMMISSION		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION ",.sprSht.ActiveRow)
		strVOCH_TYPE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE ",.sprSht.ActiveRow)
	
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, meINS_ROW, 0, -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strYEARMON
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",.sprSht.ActiveRow, strMEDNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",.sprSht.ActiveRow, strMATTERNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",.sprSht.ActiveRow, strMEDCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",.sprSht.ActiveRow, strMATTERCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",.sprSht.ActiveRow, strTBRDEDDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"CNT",.sprSht.ActiveRow, intCNT
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",.sprSht.ActiveRow, intPRICE
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, intAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, strCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, strCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, strSUBSEQ
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, strSUBSEQNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow, strCLIENTSUBCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",.sprSht.ActiveRow, strCLIENTSUBNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, strTIMCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",.sprSht.ActiveRow, strTIMNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",.sprSht.ActiveRow, strDEPT_CD
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",.sprSht.ActiveRow, strDEPT_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, strREAL_MED_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, strREAL_MED_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM",.sprSht.ActiveRow, strPROGRAM
		mobjSCGLSpr.SetTextBinding .sprSht,"STD",.sprSht.ActiveRow, strSTD
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, strEXCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, strEXCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"INPUT_MEDFLAG",.sprSht.ActiveRow, strINPUT_MEDFLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"SPONSOR",.sprSht.ActiveRow, strSPONSOR
		mobjSCGLSpr.SetTextBinding .sprSht,"ROLLSTDATE",.sprSht.ActiveRow, strROLLSTDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"ROLLEDDATE",.sprSht.ActiveRow, strROLLEDDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDSTTIME",.sprSht.ActiveRow, strBRDSTTIME
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDEDTIME",.sprSht.ActiveRow, strBRDEDTIME
		mobjSCGLSpr.SetTextBinding .sprSht,"TYPHOUR",.sprSht.ActiveRow, strTYPHOUR
		mobjSCGLSpr.SetTextBinding .sprSht,"CMLAN",.sprSht.ActiveRow, strCMLAN
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDMON",.sprSht.ActiveRow, strBRDMON
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDTUE",.sprSht.ActiveRow, strBRDTUE
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDWED",.sprSht.ActiveRow, strBRDWED
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDTHU",.sprSht.ActiveRow, strBRDTHU
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDFRI",.sprSht.ActiveRow, strBRDFRI
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDSAT",.sprSht.ActiveRow, strBRDSAT
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDSUN",.sprSht.ActiveRow, strBRDSUN
		mobjSCGLSpr.SetTextBinding .sprSht,"ADLOCALFLAG",.sprSht.ActiveRow, strADLOCALFLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"BRDDIV",.sprSht.ActiveRow, strBRDDIV
		mobjSCGLSpr.SetTextBinding .sprSht,"ADSTOCFLAG",.sprSht.ActiveRow, strADSTOCFLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"INPUT_AREAFLAGNAME",.sprSht.ActiveRow, strINPUT_AREAFLAGNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, intCOMMI_RATE
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",.sprSht.ActiveRow, intCOMMISSION
		mobjSCGLSpr.SetTextBinding .sprSht,"VOCH_TYPE",.sprSht.ActiveRow, strVOCH_TYPE

		gXMLSetFlag xmlBind, meUPD_TRANS
		mstrPROCESS = False
   	end With
end Sub

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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SELECTRTN
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
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SELECTRTN
				Else
					Call CLIENTCODE1_POP()
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
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
											trim(.txtTIMCODE1.value),trim(.txtTIMNAME1.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME1.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(5,1))
					SELECTRTN
				Else
					Call TIMCODE1_POP()
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
			SELECTRTN
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME1_onkeydown
	Dim vntData
   	Dim i, strCols
	
	If window.event.keyCode = meEnter Then
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCMGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
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
					.txtSUBSEQ.value	  = trim(vntData(6,1))
					.txtSUBSEQNAME.value  = trim(vntData(7,1))
				Else
					Call MATTERCODE1_POP()
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
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_NAME1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_CODE1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_NAME1.value= trim(vntData(0,1))
					.txtREAL_MED_CODE1.value = trim(vntData(1,1))
					SELECTRTN
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
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			'.txtGREATCODE.value = trim(vntRet(4,0))
			'.txtGREATNAME.value = trim(vntRet(5,0))
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE_ALL") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					'.txtGREATCODE.value = trim(vntData(4,1))
					'.txtGREATNAME.value = trim(vntData(5,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
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


'�귣��
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_ALL.aspx",vntInParams , 640,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))	
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	
			.txtCLIENTCODE.value = trim(vntRet(2,0))	
			.txtCLIENTNAME.value = trim(vntRet(3,0))	
			'.txtGREATCODE.value = trim(vntRet(4,0))	
			'.txtGREATNAME.value = trim(vntRet(5,0))	
			.txtTIMCODE.value = trim(vntRet(6,0))
			.txtTIMNAME.value = trim(vntRet(7,0))
			.txtCLIENTSUBCODE.value = trim(vntRet(8,0))	
			.TXTCLIENTSUBNAME.value = trim(vntRet(9,0))	
			.txtDEPT_CD.value = trim(vntRet(10,0))	
			.txtDEPT_NAME.value = trim(vntRet(11,0))	
			
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(7,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(8,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(9,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(10,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(11,0))
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_BrandInfo_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
													trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo_ALL") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	
					.txtCLIENTNAME.value = trim(vntData(3,1))	
					'.txtGREATCODE.value = trim(vntData(4,1))
					'.txtGREATNAME.value = trim(vntData(5,1))
					.txtTIMCODE.value = trim(vntData(6,1))		
					.txtTIMNAME.value = trim(vntData(7,1))		
					.txtCLIENTSUBCODE.value = trim(vntData(8,1))	
					.txtCLIENTSUBNAME.value = trim(vntData(9,1))	
					.txtDEPT_CD.value = trim(vntData(10,1))		
					.txtDEPT_NAME.value = trim(vntData(11,1))	
						
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(7,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(8,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(9,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(10,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(11,1))
						
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call SUBSEQCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'����� �˾� 
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

Sub CLIENTSUBCODE_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCLIENTSUBPOP_ALL.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtCLIENTSUBCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(3,0))	'Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(4,0))	'�ڵ�� ǥ��
			'.txtGREATCODE.value = trim(vntRet(5,0))	'�ڵ�� ǥ��
			'.txtGREATNAME.value = trim(vntRet(6,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtCLIENTCODE
		End If
	end With
End Sub

Sub txtCLIENTSUBNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCLIENTSUBCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
		
			If not gDoErrorRtn ("GetCLIENTSUBCODE_ALL") Then
			
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,1))	'Code�� ����
					.txtCLIENTSUBNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(3,1))	'Code�� ����
					.txtCLIENTNAME.value = trim(vntData(4,1))	'�ڵ�� ǥ��
					'.txtGREATCODE.value = trim(vntData(5,1))	'�ڵ�� ǥ��
					'.txtGREATNAME.value = trim(vntData(6,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
'�� �˾� ��ư
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
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP_ALL.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��.
			.txtCLIENTSUBCODE.value = trim(vntRet(2,0))       ' �ڵ�� ǥ��
			.txtCLIENTSUBNAME.value = trim(vntRet(3,0))       ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' �ڵ�� ǥ��
			'.txtGREATCODE.value = trim(vntRet(6,0))
			'.txtGREATNAME.value = trim(vntRet(7,0))
					
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(7,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
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
			vntData = mobjMDCMGET.GetTIMCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
									 		trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE_ALL") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTSUBCODE.value = trim(vntData(2,1))
					.txtCLIENTSUBNAME.value = trim(vntData(3,1))
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					'.txtGREATCODE.value = trim(vntData(6,1))
					'.txtGREATNAME.value = trim(vntData(7,1))
					
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(2,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))	
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(7,1))
						
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call TIMCODE_POP()
				End If
   			End If
   		End With
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

Sub txtDEPT_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPT_NAME.value))
			
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


'��ü��-ä�� �˾� ��ư-------
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'���� ������List ��������
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis   
	
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value), trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "MED_ELECTRIC")
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
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
												trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "MED_ELECTRIC")
			
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
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_CODE",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_NAME",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
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
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "B")
			
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



'����� ��ư �˾�
Sub ImgMATTERCODE_onclick
	Call MATTERCODE_POP()
End Sub

Sub MATTERCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTNAME.value), trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
							trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "A") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP_ALL.aspx",vntInParams , 780,630)
		
		If isArray(vntRet) Then
			If .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtMATTERCODE.value = trim(vntRet(0,0))	' �����ڵ� ǥ��
			.txtMATTERNAME.value = trim(vntRet(1,0))	' ����� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' �������ڵ� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE.value = trim(vntRet(4,0))		' ���ڵ� ǥ��
			.txtTIMNAME.value = trim(vntRet(5,0))		' ���� ǥ��
			.txtSUBSEQ.value = trim(vntRet(6,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(7,0))	' �귣��� ǥ��
			.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' ���ۻ��ڵ� ǥ��
			.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' ���ۻ��ڵ� ǥ��
			.txtDEPT_CD.value = trim(vntRet(10,0))		' �μ��ڵ� ǥ��
			.txtDEPT_NAME.value = trim(vntRet(11,0))	' �μ��� ǥ��
			.txtCLIENTSUBCODE.value = trim(vntRet(12,0))	' ������ڵ� ǥ��
			.txtCLIENTSUBNAME.value = trim(vntRet(13,0))	' ����θ� ǥ��
			'.txtGREATCODE.value = trim(vntRet(14,0))	' ����ó�ڵ� ǥ��
			'.txtGREATNAME.value = trim(vntRet(15,0))	' ����ó�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntRet(7,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(8,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(9,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(10,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(11,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(12,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(13,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(14,0))
				'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(15,0))
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME_onkeydown
	Dim vntData
   	Dim i, strCols
	
	If window.event.keyCode = meEnter Then
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
            
			vntData = mobjMDCMGET.GetMATTER_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtCLIENTNAME.value),trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
												trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "A")
			If not gDoErrorRtn ("GetMATTER_ALL") Then
				If mlngRowCnt = 1 Then
				
					.txtMATTERCODE.value = trim(vntData(0,1))	' �����ڵ� ǥ��
					.txtMATTERNAME.value = trim(vntData(1,1))	' ����� ǥ��
					.txtCLIENTCODE.value = trim(vntData(2,1))	' �������ڵ� ǥ��
					.txtCLIENTNAME.value = trim(vntData(3,1))	' �����ָ� ǥ��
					.txtTIMCODE.value	 = trim(vntData(4,1))	' ���ڵ� ǥ��
					.txtTIMNAME.value	 = trim(vntData(5,1))	' ���� ǥ��
					.txtSUBSEQ.value	 = trim(vntData(6,1))	' �귣�� ǥ��
					.txtSUBSEQNAME.value = trim(vntData(7,1))	' �귣��� ǥ��
					.txtEXCLIENTCODE.value = trim(vntData(8,1))	' ���ۻ��ڵ� ǥ��
					.txtEXCLIENTNAME.value = trim(vntData(9,1))	' ���ۻ�� ǥ��
					.txtDEPT_CD.value	 = trim(vntData(10,1))	' �μ��ڵ� ǥ��
					.txtDEPT_NAME.value	 = trim(vntData(11,1))	' �μ��� ǥ��
					.txtCLIENTSUBCODE.value	 = trim(vntData(12,1))	' ������ڵ� ǥ��
					.txtCLIENTSUBNAME.value	 = trim(vntData(13,1))	' ����θ� ǥ��
					'.txtGREATCODE.value	 = trim(vntData(14,1))	' ����ó�ڵ� ǥ��
					'.txtGREATNAME.value	 = trim(vntData(15,1))	' ����ó�� ǥ��
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntData(7,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(8,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(9,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(10,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(11,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(12,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(13,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(14,1))
						'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(15,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call MATTERCODE_POP()
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

			vntData = mobjMDCMGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
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


'�ܰ����� ���ͽÿ� �ݾ� �ڵ����
Sub txtPRICE_onkeydown
	if window.event.keyCode = meEnter OR window.event.keyCode = meTAB then
		AMT_CAL
		txtAMT_onchange
	end if
End Sub

Sub AMT_CAL
	Dim strCNT
	Dim strPRICE
	Dim strAMOUNT
	On error resume next
	with frmThis
		strCNT		= .txtCNT.value
		strPRICE	= .txtPRICE.value
		
		IF strCNT <> "" AND  strPRICE <> "" THEN
			strAMOUNT	= CDBL(strCNT) * CDBL(strPRICE)
		END IF
		.txtAMT.value = strAMOUNT
		call gFormatNumber(.txtAMT,0,true)
		COMMI_RATE_Cal
   	end with
END Sub

'****************************************************************************************
' ������ ���
'****************************************************************************************
Sub COMMI_RATE_Cal
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,dblCOMMI_RATE
	
	With frmThis
		intAMT		= .txtAMT.value
		
		if intAMT= "" then  Exit Sub

		if .txtCOMMI_RATE.value ="" Then
			.txtCOMMI_RATE.value = 15
			dblCOMMI_RATE	= .txtCOMMI_RATE.value
		else
			dblCOMMI_RATE	= .txtCOMMI_RATE.value
		end if
			
		.txtCOMMISSION.value = intAMT * dblCOMMI_RATE /100
		
		txtAMT_onchange
		txtCOMMI_RATE_onchange
		txtCOMMISSION_onchange
	End With
	txtCOMMISSION_onblur
End Sub


'-----------------------------
' �޷� 
'-----------------------------
Sub imgCalEndar_onclick
	gShowPopupCalEndar frmThis.txtTBRDSTDATE,frmThis.imgCalEndar,"txtTBRDSTDATE_onchange()"
	gXMLDataChanged xmlBind
End Sub

Sub imgCalEndar1_onclick
	gShowPopupCalEndar frmThis.txtTBRDEDDATE,frmThis.imgCalEndar1,"txtTBRDEDDATE_onchange()"
	gXMLDataChanged xmlBind
End Sub

Sub imgCalEndar4_onclick
	gShowPopupCalEndar frmThis.txtROLLSTDATE,frmThis.imgCalEndar4,"txtROLLSTDATE_onchange()"
	gXMLDataChanged xmlBind
End Sub

Sub imgCalEndar5_onclick
	gShowPopupCalEndar frmThis.txtROLLEDDATE,frmThis.imgCalEndar5,"txtROLLEDDATE_onchange()"
	gXMLDataChanged xmlBind
End Sub



'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
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
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************
Sub txtYEARMON_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'������
Sub txtCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'�귣��
Sub txtSUBSEQNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtSUBSEQ_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQ.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'�����
Sub txtCLIENTSUBNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtCLIENTSUBCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'��
Sub txtTIMNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, frmThis.txtTIMNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtTIMCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, frmThis.txtTIMCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'���μ�
Sub txtDEPT_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME_onchange",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_NAME_onchange.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtDEPT_CD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_CD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'��ü��
Sub txtMEDNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, frmThis.txtMEDNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtMEDCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, frmThis.txtMEDCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'��ü��
Sub txtREAL_MED_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtREAL_MED_CODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_CODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'����
Sub txtPROGRAM_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PROGRAM",frmThis.sprSht.ActiveRow, frmThis.txtPROGRAM.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'ǰ��
Sub txtSTD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD",frmThis.sprSht.ActiveRow, frmThis.txtSTD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'�����
Sub txtMATTERNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERNAME",frmThis.sprSht.ActiveRow, frmThis.txtMATTERNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtMATTERCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERCODE",frmThis.sprSht.ActiveRow, frmThis.txtMATTERCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'���ۻ�
Sub txtEXCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtEXCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'��ü����
Sub cmbINPUT_MEDFLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"INPUT_MEDFLAG",frmThis.sprSht.ActiveRow, frmThis.cmbINPUT_MEDFLAG.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'��������
Sub chkSPONSOR_onClick
	with frmThis
		IF .chkSPONSOR.checked = true THEN
			.txtCOMMI_RATE.readOnly = "FALSE"
			.txtCOMMI_RATE.className = "INPUT_R"
			.txtCOMMISSION.readOnly = "FALSE"
			.txtCOMMISSION.className = "INPUT_R"
			.txtCOMMI_RATE.value = "10"
			.chkVOCH_TYPE.disabled = TRUE
		ELSE
			.txtCOMMI_RATE.readOnly = "true"
			.txtCOMMI_RATE.className = "NOINPUT_R"
			.txtCOMMISSION.readOnly = "TRUE"
			.txtCOMMISSION.className = "NOINPUT_R"
			.txtCOMMI_RATE.value = ""
			.txtCOMMISSION.value = ""
			.chkVOCH_TYPE.disabled = FALSE
		END IF
		
		If .chkSPONSOR.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SPONSOR",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SPONSOR",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
		
		txtCOMMISSION_onchange  
		txtCOMMI_RATE_onchange
	end with
	gSetChange
End Sub

'����Ⱓ
Sub txtTBRDSTDATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDSTDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtTBRDEDDATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDEDDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'����Ⱓ
Sub txtROLLSTDATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ROLLSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtROLLSTDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtROLLEDDATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ROLLEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtROLLEDDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'��۽ð�
Sub txtBRDSTTIME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BRDSTTIME",frmThis.sprSht.ActiveRow, frmThis.txtBRDSTTIME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub
Sub txtBRDEDTIME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BRDEDTIME",frmThis.sprSht.ActiveRow, frmThis.txtBRDEDTIME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'�ñ�
Sub cmbTYPHOUR_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TYPHOUR",frmThis.sprSht.ActiveRow, frmThis.cmbTYPHOUR.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'�ʼ�
Sub txtCMLAN_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CMLAN",frmThis.sprSht.ActiveRow, frmThis.txtCMLAN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'��
Sub chkBRDMON_onClick
	with frmThis
		If .chkBRDMON.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDMON",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDMON",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub

'ȭ
Sub chkBRDTUE_onClick
	with frmThis
		If .chkBRDTUE.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDTUE",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDTUE",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub

'��
Sub chkBRDWED_onClick
	with frmThis
		If .chkBRDWED.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDWED",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDWED",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub

'��
Sub chkBRDTHU_onClick
	with frmThis
		If .chkBRDTHU.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDTHU",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDTHU",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	end with	
	gSetChange
End Sub

'��
Sub chkBRDFRI_onClick
	with frmThis
		If .chkBRDFRI.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDFRI",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDFRI",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub

'��
Sub chkBRDSAT_onClick
	with frmThis
		If .chkBRDSAT.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDSAT",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDSAT",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub

'��
Sub chkBRDSUN_onClick
	with frmThis
		If .chkBRDSUN.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDSUN",.sprSht.ActiveRow, 0
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BRDSUN",.sprSht.ActiveRow, 1
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If	
	end with
	gSetChange
End Sub


'��������
Sub txtADLOCALFLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ADLOCALFLAG",frmThis.sprSht.ActiveRow, frmThis.txtADLOCALFLAG.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'���౸��
Sub txtBRDDIV_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BRDDIV",frmThis.sprSht.ActiveRow, frmThis.txtBRDDIV.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'û�౸��
Sub txtADSTOCFLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ADSTOCFLAG",frmThis.sprSht.ActiveRow, frmThis.txtADSTOCFLAG.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'������
Sub txtINPUT_AREAFLAGNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"INPUT_AREAFLAGNAME",frmThis.sprSht.ActiveRow, frmThis.txtINPUT_AREAFLAGNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'�ܰ�
Sub txtPRICE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PRICE",frmThis.sprSht.ActiveRow, frmThis.txtPRICE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'ȸ��
Sub txtCNT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CNT",frmThis.sprSht.ActiveRow, frmThis.txtCNT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub


'�ݾ�
Sub txtAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'��������
Sub txtCOMMI_RATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, frmThis.txtCOMMI_RATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub

'������
Sub txtCOMMISSION_onchange  
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	gSetChange
End Sub




Sub chkVOCH_TYPE_onchange
	Dim strVOCH_TYPE
	WITH frmThis
		If .sprSht.ActiveRow >0 Then
			if .chkVOCH_TYPE.checked = true then
				strVOCH_TYPE = 3
			else
				strVOCH_TYPE = 2	
			end if 
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, strVOCH_TYPE
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end With
	gSetChange
End Sub

'-----------------------------------------------
' õ���� ������ ǥ�� ( Ƚ��, �ܰ�, �ݾ�, ������)
'-----------------------------------------------
'Ƚ��
Sub txtCNT_onblur
	with frmThis
		call gFormatNumber(.txtCNT,0,true)
	end with
	AMT_CAL
End Sub

'�ܰ�
Sub txtPRICE_onblur
	with frmThis
		call gFormatNumber(.txtPRICE,0,true)
	end with
	AMT_CAL
End Sub

'�ݾ�
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
	AMT_CAL
End Sub

'������
Sub txtCOMMISSION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMISSION,0,true)
	end with
End Sub

'
Sub txtCOMMI_RATE_onblur
	COMMI_RATE_Cal
End Sub




'-----------------------------------------------------------------------------------------
' õ���� ������ ���ֱ� ( Ƚ��, �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'Ƚ��
Sub txtCNT_onfocus
	with frmThis
		.txtCNT.value = Replace(.txtCNT.value,",","")
	end with
End Sub

'�ܰ�
Sub txtPRICE_onfocus
	with frmThis
		.txtPRICE.value = Replace(.txtPRICE.value,",","")
	end with
End Sub

'�ݾ�
Sub txtAMT_onfocus
	with frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub

'������
Sub txtCOMMISSION_onfocus
	with frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end with
End Sub


'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	with frmThis
		if Row > 0 and Col > 1 then
			sprShtToFieldBinding Col,Row
		end if
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"INPUT_MEDFLAG",frmThis.sprSht.ActiveRow, frmThis.cmbINPUT_MEDFLAG.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDSTDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDEDDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ROLLSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtROLLSTDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ROLLEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtROLLEDDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TYPHOUR",frmThis.sprSht.ActiveRow, frmThis.cmbTYPHOUR.value
		
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		'mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus
	End If
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
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
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
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON") Then .txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SEQ") Then .txtSEQ.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDCODE") Then .txtMEDCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then .txtMEDNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERCODE") Then .txtMATTERCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then .txtMATTERNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TBRDSTDATE") Then .txtTBRDSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TBRDEDDATE") Then .txtTBRDEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CNT") Then .txtCNT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CNT",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then .txtPRICE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then .txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE") Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then .txtCLIENTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQ") Then .txtSUBSEQ.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then .txtSUBSEQNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBCODE") Then .txtCLIENTSUBCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then .txtCLIENTSUBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMCODE") Then .txtTIMCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then .txtTIMNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_CD") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then .txtDEPT_NAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_CODE") Then .txtREAL_MED_CODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then .txtREAL_MED_NAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PROGRAM") Then .txtPROGRAM.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTCODE") Then .txtEXCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then .txtEXCLIENTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"INPUT_MEDFLAG") Then .cmbINPUT_MEDFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_MEDFLAG",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SPONSOR") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",Row) = "1" Then
				.chkSPONSOR.checked = True
			Else
				.chkSPONSOR.checked = False
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ROLLSTDATE") Then .txtROLLSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLSTDATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ROLLEDDATE") Then .txtROLLEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLEDDATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BRDSTTIME") Then .txtBRDSTTIME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSTTIME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BRDEDTIME") Then .txtBRDEDTIME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BRDEDTIME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TYPHOUR") Then .cmbTYPHOUR.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TYPHOUR",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CMLAN") Then .txtCMLAN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CMLAN",Row)
		
		
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDMON",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDMON",Row) = "1" Then
				.chkBRDMON.checked = True
			Else
				.chkBRDMON.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTUE",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTUE",Row) = "1" Then
				.chkBRDTUE.checked = True
			Else
				.chkBRDTUE.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDWED",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDWED",Row) = "1" Then
				.chkBRDWED.checked = True
			Else
				.chkBRDWED.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTHU",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTHU",Row) = "1" Then
				.chkBRDTHU.checked = True
			Else
				.chkBRDTHU.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDFRI",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDFRI",Row) = "1" Then
				.chkBRDFRI.checked = True
			Else
				.chkBRDFRI.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSAT",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSAT",Row) = "1" Then
				.chkBRDSAT.checked = True
			Else
				.chkBRDSAT.checked = False
			End If
		end if
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSUN",Row) then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSUN",Row) = "1" Then
				.chkBRDSUN.checked = True
			Else
				.chkBRDSUN.checked = False
			End If
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ADLOCALFLAG") Then .txtADLOCALFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ADLOCALFLAG",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BRDDIV") Then .txtBRDDIV.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BRDDIV",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ADSTOCFLAG") Then .txtADSTOCFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ADSTOCFLAG",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"INPUT_AREAFLAGNAME") Then .txtINPUT_AREAFLAGNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_AREAFLAGNAME",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then .txtCOMMI_RATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then .txtCOMMISSION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'��Ʈ�� �ִ� ������ �ؽ�Ʈ �ʵ忡 ���ε� 
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis 
		If .sprSht.MaxRows = 0 Then exit function '�׸��� �����Ͱ� ������ ������.
		.txtYEARMON.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtSEQ.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		.txtMEDCODE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		.txtMEDNAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		.txtMATTERCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		.txtMATTERNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		.txtTBRDSTDATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
		.txtTBRDEDDATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		.txtCNT.value				=   mobjSCGLSpr.GetTextBinding(.sprSht,"CNT",Row)
		.txtPRICE.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
		.txtAMT.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtCLIENTCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtCLIENTNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtSUBSEQ.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		.txtSUBSEQNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtCLIENTSUBCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",Row)
		.txtCLIENTSUBNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",Row)
		.txtTIMCODE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		.txtTIMNAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row)
		.txtDEPT_CD.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		.txtDEPT_NAME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtREAL_MED_CODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtREAL_MED_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtPROGRAM.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",Row)
		.txtSTD.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row)
		.txtEXCLIENTCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		.txtEXCLIENTNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		.cmbINPUT_MEDFLAG.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_MEDFLAG",Row)
	
		if mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",Row) then
			.chkSPONSOR.checked =true
		else
			.chkSPONSOR.checked =false
		end if
		
		.txtROLLSTDATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLSTDATE",Row)
		.txtROLLEDDATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"ROLLEDDATE",Row)
		.txtBRDSTTIME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSTTIME",Row)
		.txtBRDEDTIME.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDEDTIME",Row)
		.cmbTYPHOUR.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TYPHOUR",Row)
		.txtCMLAN.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"CMLAN",Row)
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDMON",Row) then
			.chkBRDMON.checked =true
		else
			.chkBRDMON.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTUE",Row) then
			.chkBRDTUE.checked =true
		else
			.chkBRDTUE.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDWED",Row) then
			.chkBRDWED.checked =true
		else
			.chkBRDWED.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDTHU",Row) then
			.chkBRDTHU.checked =true
		else
			.chkBRDTHU.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDFRI",Row) then
			.chkBRDFRI.checked =true
		else
			.chkBRDFRI.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSAT",Row) then
			.chkBRDSAT.checked =true
		else
			.chkBRDSAT.checked =false
		end if
		if mobjSCGLSpr.GetTextBinding(.sprSht,"BRDSUN",Row) then
			.chkBRDSUN.checked =true
		else
			.chkBRDSUN.checked =false
		end if
		
		.txtADLOCALFLAG.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"ADLOCALFLAG",Row)
		.txtBRDDIV.value				=	mobjSCGLSpr.GetTextBinding(.sprSht,"BRDDIV",Row)
		.txtADSTOCFLAG.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"ADSTOCFLAG",Row)
		.txtINPUT_AREAFLAGNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_AREAFLAGNAME",Row)
		.txtCOMMI_RATE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtCOMMISSION.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		.txtTRU_TRANS_NO.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",Row)
		.txtCOMMI_TRANS_NO.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row) = "3" THEN
			.chkVOCH_TYPE.checked = TRUE
		ELSE
			.chkVOCH_TYPE.checked = FALSE
		END IF 
		
		Field_Lock
	END WITH
End Function


'=============================
' UI���� ���ν��� 
'=============================
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()

	'����������ü ����	
	set mobjMDETELEC_TRAN	= gCreateRemoteObject("cMDET.ccMDETELEC_TRAN")
	set mobjMDCMCODETR		= gCreateRemoteObject("cMDCO.ccMDCOCODETR")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 52, 0, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | MEDNAME | MATTERNAME | MEDCODE | MATTERCODE | TBRDSTDATE | TBRDEDDATE | CNT | PRICE | AMT | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | CLIENTSUBCODE | CLIENTSUBNAME | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | REAL_MED_CODE | REAL_MED_NAME | PROGRAM | STD | EXCLIENTCODE | EXCLIENTNAME | INPUT_MEDFLAG | SPONSOR | ROLLSTDATE | ROLLEDDATE | BRDSTTIME | BRDEDTIME | TYPHOUR | CMLAN | BRDMON | BRDTUE | BRDWED | BRDTHU | BRDFRI | BRDSAT | BRDSUN | ADLOCALFLAG | BRDDIV | ADSTOCFLAG | INPUT_AREAFLAGNAME | COMMI_RATE | COMMISSION | VOCH_TYPE | TRU_TRANS_NO | COMMI_TRANS_NO "
		mobjSCGLSpr.SetHeader .sprSht, "����|�⵵|����|��ü��|�����|��ü�ڵ�|�����ڵ�|������|������|Ƚ��|�ܰ�|�ݾ�|�������ڵ�|������|�귣���ڵ�|�귣��|������ڵ�|�����|���ڵ�|��|�μ��ڵ�|���μ�|��ü���ڵ�|��ü��|����|ǰ��|���۴����ڵ�|���۴���|��ü����|��������|�������|��������|��۽���|�������|�ñ�|�ʼ�|������|ȭ����|������|�����|�ݿ���|�����|�Ͽ���|��������|���౸��|������|��������|������|û������|����Ź��ȣ|�������ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   0|    25|    25|       0|       0|     8|     8|   6|  11|  12|         0|     0|         0|     0|         0|     0|     0| 0|       0|      0|          0|     0|     0|   0|           0|       0|       0|       0|       0|       0|       0|       0|   0|   0|     0|     0|     0|     0|     0|     0|     0|       0|       0|     0|       0|     0|       0|         0|         0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDSTDATE | TBRDEDDATE | ROLLSTDATE | ROLLEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "SPONSOR | BRDMON | BRDTUE | BRDWED | BRDTHU | BRDFRI | BRDSAT | BRDSUN "
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | MEDNAME | MATTERNAME", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | CNT | PRICE | AMT | CMLAN | COMMISSION ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE ", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | SEQ |  MEDNAME |  MATTERNAME | MEDCODE | MATTERCODE | TBRDSTDATE | TBRDEDDATE | CNT | PRICE | AMT | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | CLIENTSUBCODE | CLIENTSUBNAME | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | REAL_MED_CODE | REAL_MED_NAME | PROGRAM | STD | EXCLIENTCODE | EXCLIENTNAME | INPUT_MEDFLAG | SPONSOR | ROLLSTDATE | ROLLEDDATE | BRDSTTIME | BRDEDTIME | TYPHOUR | CMLAN | BRDMON | BRDTUE | BRDWED | BRDTHU | BRDFRI | BRDSAT | BRDSUN | ADLOCALFLAG | BRDDIV | ADSTOCFLAG | INPUT_AREAFLAGNAME | COMMI_RATE | COMMISSION | TRU_TRANS_NO | COMMI_TRANS_NO "
		mobjSCGLSpr.ColHidden .sprSht, " SEQ |   MEDCODE | MATTERCODE | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | CLIENTSUBCODE | CLIENTSUBNAME | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | REAL_MED_CODE | REAL_MED_NAME | PROGRAM | STD | EXCLIENTCODE | EXCLIENTNAME | INPUT_MEDFLAG | SPONSOR | ROLLSTDATE | ROLLEDDATE | BRDSTTIME | BRDEDTIME | TYPHOUR | CMLAN | BRDMON | BRDTUE | BRDWED | BRDTHU | BRDFRI | BRDSAT | BRDSUN | ADLOCALFLAG | BRDDIV | ADSTOCFLAG | INPUT_AREAFLAGNAME | COMMI_RATE | COMMISSION | VOCH_TYPE | TRU_TRANS_NO | COMMI_TRANS_NO", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | TBRDSTDATE | TBRDEDDATE",-1,-1,2,2,false
		
		.sprSht.style.visibility = "visible"

    End With
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDETELEC_TRAN	= Nothing
	set mobjMDCMGET			= Nothing
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
	
		.txtYEARMON1.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtTBRDSTDATE.value = gNowDate2
		.txtTBRDEDDATE.value = gNowDate2
		.txtROLLSTDATE.value = gNowDate2
		.txtROLLEDDATE.value = gNowDate2
		.cmbINPUT_MEDFLAG.value = "01"
		

		COMBO_TYPE
		Get_COMBO_VALUE
		
		
		Field_Lock
		.txtCOMMI_RATE.readOnly = "TRUE"
		.txtCOMMI_RATE.className = "NOINPUT_R"
		.txtCOMMISSION.readOnly = "TRUE"
		.txtCOMMISSION.className = "NOINPUT_R"
		.txtCOMMI_RATE.value = ""
		.txtCOMMISSION.value = ""
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		Call sprSht_Keydown(meINS_ROW, 0)
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'û���� ��ȸ���� ����
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	msgbox date2
End Sub

'------------------------------------------
' select �ڽ� ������ ���ε��� ����
'------------------------------------------
sub COMBO_TYPE()
	Dim vntME_KIND
   	Dim vntMD_SAF
   	
	  
    With frmThis

		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntME_KIND = mobjMDETELEC_TRAN.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt, "ME_KIND")
		vntMD_SAF = mobjMDETELEC_TRAN.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt, "MD_SAF")
     
		if not gDoErrorRtn ("COMBO_TYPE") then
			 gLoadComboBox .cmbINPUT_MEDFLAG, vntME_KIND, False
			 gLoadComboBox .cmbTYPHOUR, vntMD_SAF, False
   		end if
   	end with
end sub

'-----------------------------------------------------------------------------------------
' �׸��� �޺��ڽ� ����
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntME_KIND, vntMD_SAF
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntME_KIND = mobjMDETELEC_TRAN.Get_COMBOMEDFLAG_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		vntMD_SAF = mobjMDETELEC_TRAN.Get_COMBOTYPHOUR_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBOMEDFLAG_VALUE") AND not gDoErrorRtn ("Get_COMBOTYPHOUR_VALUE")  Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "INPUT_MEDFLAG",,,vntME_KIND,,60 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TYPHOUR",,,vntMD_SAF,,60 
			mobjSCGLSpr.TypeComboBox = True 
   		End If
   	End With
End Sub


'-----------------------------------------------------------------------------------------
' Field_Lock  �ŷ�������ȣ�� ���ݰ�꼭 ��ȣ�� ������ �����Ҽ� ������ �ʵ带 ReadOnlyó��
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",.sprSht.ActiveRow) <> "" THEN
				.txtYEARMON.className		= "NOINPUT"   : .txtYEARMON.readOnly	= True
			else
				.txtYEARMON.className		= "INPUT"   : .txtYEARMON.readOnly		= FALSE 
			End If
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",.sprSht.ActiveRow) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> "" Then		
	
				'������
				.txtCLIENTNAME.className	 = "NOINPUT_L" : .txtCLIENTNAME.readOnly	= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	 = "NOINPUT_L" : .txtCLIENTCODE.readOnly	= True
				'�귣��
				.txtSUBSEQNAME.className	 = "NOINPUT_L" : .txtSUBSEQNAME.readOnly	= True : .ImgSUBSEQCODE.disabled = True
				.txtSUBSEQ.className		 = "NOINPUT_L" : .txtSUBSEQ.readOnly		= True 
				'�����
				.txtCLIENTSUBNAME.className	= "NOINPUT_L" : .txtCLIENTSUBNAME.readOnly	= True : .imgCLIENTSUBCODE.disabled	 = True
				.txtCLIENTSUBCODE.className	= "NOINPUT_L" : .txtCLIENTSUBCODE.readOnly	= True 
				'��
				.txtTIMNAME.className		= "NOINPUT_L" : .txtTIMNAME.readOnly		= True : .ImgTIMCODE.disabled	 = True
				.txtTIMCODE.className		= "NOINPUT_L" : .txtTIMCODE.readOnly		= True
				'���μ�
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .imgDEPT_CD.disabled	 = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				'��ü
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled	 = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True
				'��ü��
				.txtREAL_MED_NAME.className	= "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled	 = True
				.txtREAL_MED_CODE.className	= "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				'����
				.txtPROGRAM.className		= "NOINPUT_L" : .txtPROGRAM.readOnly		= True
				'ǰ��
				.txtSTD.className			= "NOINPUT_L" : .txtSTD.readOnly		= True
				'����
				.txtMATTERNAME.className	= "NOINPUT_L" : .txtMATTERNAME.readOnly		= True : .ImgMATTERCODE.disabled = True
				.txtMATTERCODE.className	= "NOINPUT_L" : .txtMATTERCODE.readOnly		= True
				'���۴���
				.txtEXCLIENTNAME.className	= "NOINPUT_L" : .txtEXCLIENTNAME.readOnly	= True : .ImgEXCLIENTCODE.disabled = True
				.txtEXCLIENTCODE.className	= "NOINPUT_L" : .txtEXCLIENTCODE.readOnly	= True
				'��ü����
				.cmbINPUT_MEDFLAG.disabled	= True 
				'����
				.chkSPONSOR.disabled		= True 
				'����Ⱓ
				.txtTBRDSTDATE.className	= "NOINPUT"   : .txtTBRDSTDATE.readOnly		= True : .imgCalEndar.disabled  = True 
				.txtTBRDEDDATE.className	= "NOINPUT"   : .txtTBRDEDDATE.readOnly		= True : .imgCalEndar1.disabled  = True 
				'����Ⱓ
				.txtROLLSTDATE.className	= "NOINPUT"   : .txtROLLSTDATE.readOnly		= True : .imgCalEndar4.disabled  = True 
				.txtROLLEDDATE.className	= "NOINPUT"   : .txtROLLEDDATE.readOnly		= True : .imgCalEndar5.disabled  = True 
				'��۽ð�
				.txtBRDSTTIME.className		= "NOINPUT"   : .txtBRDSTTIME.readOnly		= True
				.txtBRDEDTIME.className		= "NOINPUT"   : .txtBRDEDTIME.readOnly		= True
				'�ñ�
				.cmbTYPHOUR.disabled		= True : .cmbTYPHOUR.disabled = True
				'�ʼ�
				.txtCMLAN.className			= "NOINPUT_R" : .txtCMLAN.readOnly			= True
				'��ۿ���
				.chkBRDMON.disabled			= True : .chkBRDTUE.disabled		= True
				.chkBRDWED.disabled			= True : .chkBRDTHU.disabled		= True
				.chkBRDFRI.disabled			= True : .chkBRDSAT.disabled		= True
				.chkBRDSUN.disabled			= True 
				'��������
				.txtADLOCALFLAG.className	= "NOINPUT_L" : .txtADLOCALFLAG.readOnly	= True
				'���౸��
				.txtBRDDIV.className		= "NOINPUT_L" : .txtBRDDIV.readOnly			= True
				'û�౸��
				.txtADSTOCFLAG.className	= "NOINPUT_L" : .txtADSTOCFLAG.readOnly		= True
				'������
				.txtINPUT_AREAFLAGNAME.className = "NOINPUT_L" : .txtINPUT_AREAFLAGNAME.readOnly		= True
				
				'�ܰ�
				.txtPRICE.className			= "NOINPUT_R" : .txtPRICE.readOnly			= True 
				'Ƚ��
				.txtCNT.className			= "NOINPUT_R" : .txtCNT.readOnly			= True 
				'�ݾ�
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly			= True
				.chkVOCH_TYPE.disabled = True
			else 

				'������
				.txtCLIENTNAME.className	 = "INPUT_L" : .txtCLIENTNAME.readOnly	= FALSE : .ImgCLIENTCODE.disabled = FALSE
				.txtCLIENTCODE.className	 = "INPUT_L" : .txtCLIENTCODE.readOnly	= FALSE
				'�귣��
				.txtSUBSEQNAME.className	 = "INPUT_L" : .txtSUBSEQNAME.readOnly	= FALSE : .ImgSUBSEQCODE.disabled = FALSE
				.txtSUBSEQ.className		 = "INPUT_L" : .txtSUBSEQ.readOnly		= FALSE 
				'�����
				.txtCLIENTSUBNAME.className	= "INPUT_L" : .txtCLIENTSUBNAME.readOnly	= FALSE : .imgCLIENTSUBCODE.disabled	 = FALSE
				.txtCLIENTSUBCODE.className	= "INPUT_L" : .txtCLIENTSUBCODE.readOnly	= FALSE 
				'��
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= FALSE : .ImgTIMCODE.disabled	 = FALSE
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= FALSE
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly		= FALSE : .imgDEPT_CD.disabled	 = FALSE
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= FALSE
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= FALSE : .ImgMEDCODE.disabled	 = FALSE
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= FALSE
				'��ü��
				.txtREAL_MED_NAME.className	= "INPUT_L" : .txtREAL_MED_NAME.readOnly	= FALSE : .ImgREAL_MED_CODE.disabled	 = FALSE
				.txtREAL_MED_CODE.className	= "INPUT_L" : .txtREAL_MED_CODE.readOnly	= FALSE
				'����
				.txtPROGRAM.className		= "INPUT_L" : .txtPROGRAM.readOnly		= FALSE
				'ǰ��
				.txtSTD.className			= "INPUT_L" : .txtSTD.readOnly		= FALSE
				'����
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly		= FALSE : .ImgMATTERCODE.disabled = FALSE
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly		= FALSE
				'���۴���
				.txtEXCLIENTNAME.className	= "INPUT_L" : .txtEXCLIENTNAME.readOnly	= FALSE : .ImgEXCLIENTCODE.disabled = FALSE
				.txtEXCLIENTCODE.className	= "INPUT_L" : .txtEXCLIENTCODE.readOnly	= FALSE
				'��ü����
				.cmbINPUT_MEDFLAG.disabled	= FALSE 
				'����
				.chkSPONSOR.disabled		= FALSE 
				'����Ⱓ
				.txtTBRDSTDATE.className	= "INPUT"   : .txtTBRDSTDATE.readOnly		= FALSE : .imgCalEndar.disabled  = FALSE 
				.txtTBRDEDDATE.className	= "INPUT"   : .txtTBRDEDDATE.readOnly		= FALSE : .imgCalEndar1.disabled  = FALSE 
				'����Ⱓ
				.txtROLLSTDATE.className	= "INPUT"   : .txtROLLSTDATE.readOnly		= FALSE : .imgCalEndar4.disabled  = FALSE 
				.txtROLLEDDATE.className	= "INPUT"   : .txtROLLEDDATE.readOnly		= FALSE : .imgCalEndar5.disabled  = FALSE 
				'��۽ð�
				.txtBRDSTTIME.className		= "INPUT"   : .txtBRDSTTIME.readOnly		= FALSE
				.txtBRDEDTIME.className		= "INPUT"   : .txtBRDEDTIME.readOnly		= FALSE
				'�ñ�
				.cmbTYPHOUR.disabled		= FALSE : .cmbTYPHOUR.disabled = FALSE
				'�ʼ�
				.txtCMLAN.className			= "INPUT_R" : .txtCMLAN.readOnly			= FALSE
				'��ۿ���
				.chkBRDMON.disabled			= FALSE : .chkBRDTUE.disabled		= FALSE
				.chkBRDWED.disabled			= FALSE : .chkBRDTHU.disabled		= FALSE
				.chkBRDFRI.disabled			= FALSE : .chkBRDSAT.disabled		= FALSE
				.chkBRDSUN.disabled			= FALSE 
				'��������
				.txtADLOCALFLAG.className	= "INPUT_L" : .txtADLOCALFLAG.readOnly	= FALSE
				'���౸��
				.txtBRDDIV.className		= "INPUT_L" : .txtBRDDIV.readOnly			= FALSE
				'û�౸��
				.txtADSTOCFLAG.className	= "INPUT_L" : .txtADSTOCFLAG.readOnly		= FALSE
				'������
				.txtINPUT_AREAFLAGNAME.className = "INPUT_L" : .txtINPUT_AREAFLAGNAME.readOnly		= FALSE
				
				'�ܰ�
				.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly			= FALSE 
				'Ƚ��
				.txtCNT.className			= "INPUT_R" : .txtCNT.readOnly			= FALSE 
				'�ݾ�
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= FALSE
				.chkVOCH_TYPE.disabled = FALSE
			End If
		else
			
				'������
				.txtCLIENTNAME.className	 = "INPUT_L" : .txtCLIENTNAME.readOnly	= FALSE : .ImgCLIENTCODE.disabled = FALSE
				.txtCLIENTCODE.className	 = "INPUT_L" : .txtCLIENTCODE.readOnly	= FALSE
				'�귣��
				.txtSUBSEQNAME.className	 = "INPUT_L" : .txtSUBSEQNAME.readOnly	= FALSE : .ImgSUBSEQCODE.disabled = FALSE
				.txtSUBSEQ.className		 = "INPUT_L" : .txtSUBSEQ.readOnly		= FALSE 
				'�����
				.txtCLIENTSUBNAME.className	= "INPUT_L" : .txtCLIENTSUBNAME.readOnly	= FALSE : .imgCLIENTSUBCODE.disabled	 = FALSE
				.txtCLIENTSUBCODE.className	= "INPUT_L" : .txtCLIENTSUBCODE.readOnly	= FALSE 
				'��
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= FALSE : .ImgTIMCODE.disabled	 = FALSE
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= FALSE
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly		= FALSE : .imgDEPT_CD.disabled	 = FALSE
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= FALSE
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= FALSE : .ImgMEDCODE.disabled	 = FALSE
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= FALSE
				'��ü��
				.txtREAL_MED_NAME.className	= "INPUT_L" : .txtREAL_MED_NAME.readOnly	= FALSE : .ImgREAL_MED_CODE.disabled	 = FALSE
				.txtREAL_MED_CODE.className	= "INPUT_L" : .txtREAL_MED_CODE.readOnly	= FALSE
				'����
				.txtPROGRAM.className		= "INPUT_L" : .txtPROGRAM.readOnly		= FALSE
				'ǰ��
				.txtSTD.className			= "INPUT_L" : .txtSTD.readOnly		= FALSE
				'����
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly		= FALSE : .ImgMATTERCODE.disabled = FALSE
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly		= FALSE
				'���۴���
				.txtEXCLIENTNAME.className	= "INPUT_L" : .txtEXCLIENTNAME.readOnly	= FALSE : .ImgEXCLIENTCODE.disabled = FALSE
				.txtEXCLIENTCODE.className	= "INPUT_L" : .txtEXCLIENTCODE.readOnly	= FALSE
				'��ü����
				.cmbINPUT_MEDFLAG.disabled	= FALSE 
				'����
				.chkSPONSOR.disabled		= FALSE 
				'����Ⱓ
				.txtTBRDSTDATE.className	= "INPUT"   : .txtTBRDSTDATE.readOnly		= FALSE : .imgCalEndar.disabled  = FALSE 
				.txtTBRDEDDATE.className	= "INPUT"   : .txtTBRDEDDATE.readOnly		= FALSE : .imgCalEndar1.disabled  = FALSE 
				'����Ⱓ
				.txtROLLSTDATE.className	= "INPUT"   : .txtROLLSTDATE.readOnly		= FALSE : .imgCalEndar4.disabled  = FALSE 
				.txtROLLEDDATE.className	= "INPUT"   : .txtROLLEDDATE.readOnly		= FALSE : .imgCalEndar5.disabled  = FALSE 
				'��۽ð�
				.txtBRDSTTIME.className		= "INPUT"   : .txtBRDSTTIME.readOnly		= FALSE
				.txtBRDEDTIME.className		= "INPUT"   : .txtBRDEDTIME.readOnly		= FALSE
				'�ñ�
				.cmbTYPHOUR.disabled		= FALSE : .cmbTYPHOUR.disabled = FALSE
				'�ʼ�
				.txtCMLAN.className			= "INPUT_R" : .txtCMLAN.readOnly			= FALSE
				'��ۿ���
				.chkBRDMON.disabled			= FALSE : .chkBRDTUE.disabled		= FALSE
				.chkBRDWED.disabled			= FALSE : .chkBRDTHU.disabled		= FALSE
				.chkBRDFRI.disabled			= FALSE : .chkBRDSAT.disabled		= FALSE
				.chkBRDSUN.disabled			= FALSE 
				'��������
				.txtADLOCALFLAG.className	= "INPUT_L" : .txtADLOCALFLAG.readOnly	= FALSE
				'���౸��
				.txtBRDDIV.className		= "INPUT_L" : .txtBRDDIV.readOnly			= FALSE
				'û�౸��
				.txtADSTOCFLAG.className	= "INPUT_L" : .txtADSTOCFLAG.readOnly		= FALSE
				'������
				.txtINPUT_AREAFLAGNAME.className = "INPUT_L" : .txtINPUT_AREAFLAGNAME.readOnly		= FALSE
				
				'�ܰ�
				.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly		= FALSE 
				'Ƚ��
				.txtCNT.className			= "INPUT_R" : .txtCNT.readOnly			= FALSE 
				'�ݾ�
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= FALSE
				.chkVOCH_TYPE.disabled = FALSE
				
		End If
		
		'�������� ������
		IF .chkSPONSOR.checked THEN
			.txtCOMMI_RATE.readOnly = "FALSE"
			.txtCOMMI_RATE.className = "INPUT_R"
			.txtCOMMISSION.readOnly = "FALSE"
			.txtCOMMISSION.className = "INPUT_R"
			
		ELSE		
			.txtCOMMI_RATE.readOnly = "true"
			.txtCOMMI_RATE.className = "NOINPUT_R"
			.txtCOMMISSION.readOnly = "TRUE"
			.txtCOMMISSION.className = "NOINPUT_R"
		END IF
	End With
End Sub


'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i
   	Dim strYEARMON, strCLIENTCODE, strCLIENTNAME
   	Dim strTIMCODE, strTIMNAME
	Dim strREAL_MED_NAME , strREAL_MED_CODE
	Dim strMATTERCODE, strMATTERNAME
	Dim strVOCH_TYPE
	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME      = .txtTIMNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value
		strMATTERCODE    = .txtMATTERCODE1.value
		strMATTERNAME    = .txtMATTERNAME1.value
		strVOCH_TYPE	 = .cmbVOCH_TYPE1.value
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
	
		vntData = mobjMDETELEC_TRAN.SelectRtn(gstrConfigXml, mlngRowCnt, mlngColCnt,strYEARMON, _
											  strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, _
											  strREAL_MED_CODE, strREAL_MED_NAME, strMATTERCODE, _
											  strMATTERNAME, strVOCH_TYPE)
												
   			
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt >0 then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG	
				
   				For i = 1 To .sprSht.MaxRows
   					If mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR", i) = "1" Then
   						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HCCFFFF, &H000000,False
   					End If
   				Next
				
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				InitPageData
   				PreSearchFiledValue strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strREAL_MED_NAME, strREAL_MED_CODE, strMATTERCODE, strMATTERNAME
   			end if
   			
	   		gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		mstrPROCESS = True
   	end with
End Sub


'��ȸ�� �̹� ��ȸ�� �����ͼ� �ٽ� ������
Sub PreSearchFiledValue (strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strREAL_MED_CODE, strREAL_MED_NAME, strMATTERCODE, strMATTERNAME)
	frmThis.txtYEARMON1.value = strYEARMON
	frmThis.txtCLIENTCODE1.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
	frmThis.txtTIMCODE1.value = strTIMCODE
	frmThis.txtTIMNAME1.value = strTIMNAME
	frmThis.txtREAL_MED_CODE1.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME1.value = strREAL_MED_NAME
	frmThis.txtMATTERCODE1.value = strMATTERCODE
	frmThis.txtMATTERNAME1.value = strMATTERNAME
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
	
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strSEQ 
	Dim strYEARMON
	Dim strSPONSOR
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	
	
	with frmThis
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | MEDNAME | MATTERNAME | MEDCODE | MATTERCODE | TBRDSTDATE | TBRDEDDATE | CNT | PRICE | AMT | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | CLIENTSUBCODE | CLIENTSUBNAME | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | REAL_MED_CODE | REAL_MED_NAME | PROGRAM | STD | EXCLIENTCODE | EXCLIENTNAME | INPUT_MEDFLAG | SPONSOR | ROLLSTDATE | ROLLEDDATE | BRDSTTIME | BRDEDTIME | TYPHOUR | CMLAN | BRDMON | BRDTUE | BRDWED | BRDTHU | BRDFRI | BRDSAT | BRDSUN | ADLOCALFLAG | BRDDIV | ADSTOCFLAG | INPUT_AREAFLAGNAME | COMMI_RATE | COMMISSION | VOCH_TYPE | TRU_TRANS_NO | COMMI_TRANS_NO ")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		'������ Validation
		if DataValidation =false then exit sub
			
		intRtn = mobjMDETELEC_TRAN.ProcessRtn(gstrConfigXml,vntData)

   		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   	end with
End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	dim i,j
	
	DataValidation = false
	with frmThis
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
	End with
	DataValidation = true
End Function

'****************************************************************************************
' �������ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function Clientcode_FieldCheck ()
	Clientcode_FieldCheck = false
	Dim vntData
   	Dim i, strCols
   	
	with frmThis
  	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDETELEC_TRAN.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value, "CUST")
		
		if mlngRowCnt =0 then
			gErrorMsgBox "�����ָ� Ȯ�� �ϼ���", ""
			.txtCLIENTCODE.focus
			Exit Function
   		end if
   	End with
   	Clientcode_FieldCheck = true
End Function

'****************************************************************************************
' û���� �ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function REAL_MED_CODE_FieldCheck ()
	REAL_MED_CODE_FieldCheck = false
	Dim vntData
   	Dim i, strCols
   	
	with frmThis
  	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDETELEC_TRAN.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtREAL_MED_CODE.value, "REAL")
		
		if mlngRowCnt =0 then
			gErrorMsgBox "��ü���ڵ带 Ȯ�� �ϼ���", ""
			.txtREAL_MED_CODE.focus
			'.txtREAL_MED_CODE.style.backgroundColor = "#ccccff"
			
			exit Function
   		end if
   		'.txtREAL_MED_CODE.style.backgroundColor = "WHITE"
   	End with
   	REAL_MED_CODE_FieldCheck = true
End Function

'****************************************************************************************
' ��ü�� �ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function MEDCODE_FieldCheck ()
	MEDCODE_FieldCheck = false
	Dim vntData
   	Dim i, strCols
   	
	with frmThis
  	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDETELEC_TRAN.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtMEDCODE.value, "MED")
		
		if mlngRowCnt =0 then
			gErrorMsgBox "��ü���ڵ带 Ȯ�� �ϼ���", ""
			.txtMEDCODE.focus
			'.txtMEDCODE.style.backgroundColor = "#ccccff"
			
			exit Function
   		end if
   		'.txtMEDCODE.style.backgroundColor = "WHITE"
   	End with
   	MEDCODE_FieldCheck = true
End Function



'------------------------------------------
' ��ü ������ �� ��Ʈ�� ����
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim intSelCnt, intRtn, i
	dim strYEARMON, strSEQ

	with frmThis
		strSEQFLAG = False
		
		If .txtTRU_TRANS_NO.value <> "" Or .txtCOMMI_TRANS_NO.value <> "" Then
			gErrorMsgBox "�ŷ������� �����ϴ� �����Դϴ�." & vbcrlf & "�ŷ������� ���� ���� �Ͻʽÿ�.","�����ȳ�!"
			Exit Sub
		End IF 

		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",vntData(i))
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",vntData(i)) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",vntData(i)) <> "" Then
					gErrorMsgBox "�ŷ������� ���� �մϴ�." & vbcrlf & "�켱 �ŷ����� �� ���� �Ͻʿ�","�����ȳ�!"
					Exit For
				End If
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)))
				intRtn = mobjMDETELEC_TRAN.DeleteRtn(gstrConfigXml,strYEARMON, strSEQ)
				strSEQFLAG = True
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			End IF
		next
		
		gOkMsgBox intSelCnt & "���� �ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn
		End If
	End with
	err.clear
End Sub

'��ȣ�� Ŭ�����Ѵ�.
Sub CleanField (objField1, objField2)
	If frmThis.sprSht.MaxRows > 0 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TRU_TRANS_NO",frmThis.sprSht.ActiveRow) = "" and _
		   mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"COMMI_TRANS_NO",frmThis.sprSht.ActiveRow) = "" Then
			
			if isobject(objField1) then 
				objField1.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField1.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			end if
			if isobject(objField2) then 
				objField2.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField2.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			End If
		End If
	ELSE
		if isobject(objField1) then 
			objField1.value = ""
		end if
		if isobject(objField2) then 
			objField2.value = ""
		End If
	End IF
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD style="WIDTH: 100%" vAlign="top" height="100%">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
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
											<td class="TITLE">����û�� ����</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 302px"
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
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 30px" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="txtYEARMON1" style="WIDTH: 62px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1,'')"
												width="62">�� ��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 98px" width="98"><INPUT class="INPUT" id="txtYEARMON1" style="WIDTH: 80px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" size="10" name="txtYEARMON1">
											</TD>
											<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="50">������</TD>
											<TD class="SEARCHDATA" style="WIDTH: 222px; HEIGHT: 19pt" width="222"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�����ָ�" style="WIDTH: 145px; HEIGHT: 22px"
													maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
												width="50">��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 222px; HEIGHT: 19pt" width="229"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 145px; HEIGHT: 22px" maxLength="100"
													name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" maxLength="6"
													size="6" name="txtTIMCODE1"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 57px; HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME1, txtMATTERCODE1)"
												width="57">�����</TD>
											<td class="SEARCHDATA" style="HEIGHT: 19pt"><INPUT class="INPUT_L" id="txtMATTERNAME1" title="�����" style="WIDTH: 145px; HEIGHT: 22px"
													maxLength="100" size="30" name="txtMATTERNAME1"> <IMG id="ImgMATTERCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgMATTERCODE1"> <INPUT class="INPUT_L" id="txtMATTERCODE1" title="�����ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													maxLength="10" size="4" name="txtMATTERCODE1">
											</td>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" title="cmbVOCH_TYPE1" style="WIDTH: 62px; CURSOR: hand" onclick="vbscript:Call gCleanField(cmbVOCH_TYPE1,'')"
												width="62">�� ��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 98px" width="98">
												<SELECT style="WIDTH: 90px" id="cmbVOCH_TYPE1" title="����" name="cmbVOCH_TYPE1">
													<OPTION selected value="">��ü</OPTION>
													<OPTION value="0">����Ź</OPTION>
													<OPTION value="1">����</OPTION>
													<OPTION value="2">�Ϲ�</OPTION>
												</SELECT>
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
												width="50">��ü��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 222px" width="222"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="��ü���" style="WIDTH: 145px; HEIGHT: 22px"
													maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													maxLength="6" name="txtREAL_MED_CODE1"></TD>
											<td class="SEARCHDATA" colSpan="4"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
													align="right" border="0" name="imgQuery">&nbsp;
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 15px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="500" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif'
																	align='absMiddle' border='0' name='imgTableUp'></span> &nbsp;&nbsp;&nbsp;&nbsp;�հ� 
															: <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																readOnly maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																src="../../../images/imgNew.gIF" border="0" name="imgNew"></td>
														<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
																height="20" alt="�ڷḦ �������� �մϴ�.." src="../../../images/imglistcopy.gIF" border="0"
																name="Imgcopy"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
														<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
																name="imgDelete"></td>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<!--���̺��� �������°��� �����ش�-->
									<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
										<TR>
											<TD align="left" width="100%" height="1"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"></TD>
							</TR>
						</TABLE>
						<table height="80%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD id="tblBody1" style="HEIGHT: 100%" vAlign="top" align="left" colSpan="2">
									<TABLE id="tblData" style="WIDTH: 353px; HEIGHT: 469px" cellSpacing="1" cellPadding="0"
										border="0">
										<TR>
											<TD class="LABEL" title="����� �����մϴ�." align="right" width="76">�� ��</TD>
											<TD class="DATA" style="WIDTH: 100px"><INPUT dataFld="YEARMON" class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 95px; HEIGHT: 22px"
													accessKey="NUM,M" dataSrc="#xmlBind" maxLength="6" size="9" name="txtYEARMON">&nbsp;</TD>
											<TD class="LABEL" title="�Ϸù�ȣ�� �����մϴ�." style="WIDTH: 76px" align="right">�Ϸù�ȣ</TD>
											<TD class="DATA" style="WIDTH: 101px"><INPUT dataFld="SEQ" class="NOINPUT_R" id="txtSEQ" title="�ϳ��ȣ" style="WIDTH: 94px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly size="11" name="txtSEQ"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="�����ָ� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTCODE,txtCLIENTNAME)"
												align="right">������</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="27" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" maxLength="6" size="4" name="txtCLIENTCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="�귣�� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtSUBSEQ,txtSUBSEQNAME)"
												align="right">�귣��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="SUBSEQNAME" class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="22" name="txtSUBSEQNAME"> <IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgSUBSEQCODE">
												<INPUT dataFld="SUBSEQ" class="INPUT_L" id="txtSUBSEQ" title="�귣���ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="6" name="txtSUBSEQ"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="����θ� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTSUBCODE,txtCLIENTSUBNAME)"
												align="right">�����</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="CLIENTSUBNAME" class="INPUT_L" id="txtCLIENTSUBNAME" title="����θ�" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="11" name="txtCLIENTSUBNAME"> <IMG id="imgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgCLIENTSUBCODE">
												<INPUT dataFld="CLIENTSUBCODE" class="INPUT_L" id="txtCLIENTSUBCODE" title="�����" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="6" size="9" name="txtCLIENTSUBCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="���� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTIMCODE,txtTIMNAME)"
												align="right">��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="TIMNAME" class="INPUT_L" id="txtTIMNAME" title="����" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgTIMCODE"> <INPUT dataFld="TIMCODE" class="INPUT_L" id="txtTIMCODE" title="���ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="6" size="6" name="txtTIMCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="���μ��� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPT_CD,txtDEPT_NAME)"
												align="right">���μ�</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPT_NAME" title="���μ���" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="11" name="txtDEPT_NAME"> <IMG id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgDEPT_CD">
												<INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPT_CD" title="���μ�" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="9" name="txtDEPT_CD"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="��ü�� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEDCODE,txtMEDNAME)"
												align="right">��ü��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="MEDNAME" class="INPUT_L" id="txtMEDNAME" title="��ü��" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="18" name="txtMEDNAME"> <IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgMEDCODE"> <INPUT dataFld="MEDCODE" class="INPUT_L" id="txtMEDCODE" title="��ü���ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" maxLength="6" size="9" name="txtMEDCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="��ü�縦 �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREAL_MED_CODE,txtREAL_MED_NAME)"
												align="right">��ü��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="REAL_MED_NAME" class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="13" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE">
												<INPUT dataFld="REAL_MED_CODE" class="INPUT_L" id="txtREAL_MED_CODE" title="��ü���ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" maxLength="6" size="10" name="txtREAL_MED_CODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="������ �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtPROGRAM,'')"
												align="right">����</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="PROGRAM" class="INPUT_L" id="txtPROGRAM" title="����" style="WIDTH: 277px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="38" name="txtPROGRAM"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="ǰ���� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtSTD,'')"
												align="right">ǰ��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="STD" class="INPUT_L" id="txtSTD" title="ǰ��" style="WIDTH: 277px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" maxLength="100" size="38" name="txtSTD"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="������� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMATTERCODE,txtMATTERNAME)"
												align="right">����</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="�����" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="22" name="txtMATTERNAME"> <IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMATTERCODE">
												<INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="������ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" size="10" name="txtMATTERCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="���۴���縦 �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEXCLIENTCODE,txtEXCLIENTNAME)"
												align="right">���۴���</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="EXCLIENTNAME" class="INPUT_L" id="txtEXCLIENTNAME" title="�������" style="WIDTH: 198px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="22" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">
												<INPUT dataFld="EXCLIENTCODE" class="INPUT_L" id="txtEXCLIENTCODE" title="�������ڵ�" style="WIDTH: 61px; HEIGHT: 22px"
													dataSrc="#xmlBind" size="10" name="txtEXCLIENTCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" align="right">��ü����</TD>
											<TD class="DATA"><SELECT dataFld="INPUT_MEDFLAG" class="INPUT" id="cmbINPUT_MEDFLAG" title="��ü����" style="WIDTH: 99px"
													dataSrc="#xmlBind" name="cmbINPUT_MEDFLAG"></SELECT></TD>
											<TD class="LABEL" align="right">��������</TD>
											<TD class="DATA" style="HEIGHT: 19pt">&nbsp;&nbsp;&nbsp; <INPUT dataFld="SPONSOR" id="chkSPONSOR" dataSrc="#xmlBind" type="checkbox" name="chkSPONSOR"
													CHECKED></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="����Ⱓ�� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTBRDSTDATE,txtTBRDEDDATE)"
												align="right">����Ⱓ</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="TBRDSTDATE" class="INPUT" id="txtTBRDSTDATE" title="����Ⱓ" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="11" name="txtTBRDSTDATE">&nbsp;<IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar">&nbsp;~&nbsp;<INPUT dataFld="TBRDEDDATE" class="INPUT" id="txtTBRDEDDATE" title="����Ⱓ" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="9" name="txtTBRDEDDATE">&nbsp;<IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar1">&nbsp;</TD>
										</TR>
										<TR>
											<TD class="LABEL" title="����Ⱓ�� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtROLLSTDATE,txtROLLEDDATE)"
												align="right">����Ⱓ</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="ROLLSTDATE" class="INPUT" id="txtROLLSTDATE" title="����Ⱓ" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="9" name="txtROLLSTDATE">&nbsp;<IMG id="imgCalEndar4" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar4">&nbsp;~&nbsp;<INPUT dataFld="ROLLEDDATE" class="INPUT" id="txtROLLEDDATE" title="����Ⱓ" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="9" name="txtROLLEDDATE">&nbsp;<IMG id="imgCalEndar5" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar5">&nbsp;</TD>
										</TR>
										<TR>
											<TD class="LABEL" title="��۱Ⱓ�� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtBRDSTTIME,txtBRDEDTIME)">��۽ð�</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="BRDSTTIME" class="INPUT" id="txtBRDSTTIME" title="��۽ð�" style="WIDTH: 100px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="5" size="11" name="txtBRDSTTIME">&nbsp;~<INPUT dataFld="BRDEDTIME" class="INPUT" id="txtBRDEDTIME" title="��۽ð�" style="WIDTH: 100px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="5" size="13" name="txtBRDEDTIME"></TD>
										</TR>
										<TR>
											<TD class="LABEL" align="right">�ñ�</TD>
											<TD class="DATA"><SELECT dataFld="TYPHOUR" class="INPUT" id="cmbTYPHOUR" title="�ñ�" style="WIDTH: 99px" dataSrc="#xmlBind"
													name="cmbTYPHOUR"></SELECT></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCMLAN,'')">�ʼ�</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="CMLAN" class="INPUT_R" id="txtCMLAN" title="�ʼ�" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" maxLength="8" size="3" name="txtCMLAN"></TD>
										</TR>
										<TR>
											<TD class="LABEL" align="right">��ۿ���</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="BRDMON" id="chkBRDMON" dataSrc="#xmlBind" type="checkbox" name="chkBRDMON">&nbsp;��
												<INPUT dataFld="BRDTUE" id="chkBRDTUE" dataSrc="#xmlBind" type="checkbox" name="chkBRDTUE">&nbsp;ȭ
												<INPUT dataFld="BRDWED" id="chkBRDWED" dataSrc="#xmlBind" type="checkbox" name="chkBRDWED">&nbsp;��
												<INPUT dataFld="BRDTHU" id="chkBRDTHU" dataSrc="#xmlBind" type="checkbox" name="chkBRDTHU">&nbsp;��
												<INPUT dataFld="BRDFRI" id="chkBRDFRI" dataSrc="#xmlBind" type="checkbox" name="chkBRDFRI">&nbsp;��
												<INPUT dataFld="BRDSAT" id="chkBRDSAT" dataSrc="#xmlBind" type="checkbox" name="chkBRDSAT">&nbsp;��
												<INPUT dataFld="BRDSUN" id="chkBRDSUN" dataSrc="#xmlBind" type="checkbox" name="chkBRDSUN">&nbsp;��</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtADLOCALFLAG,'')"
												align="right">��������</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="ADLOCALFLAG" class="INPUT_L" id="txtADLOCALFLAG" title="��������" style="WIDTH: 282px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="38" name="txtADLOCALFLAG"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="���౸���� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtBRDDIV,'')"
												align="right">���౸��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="BRDDIV" class="INPUT_L" id="txtBRDDIV" title="���౸��" style="WIDTH: 282px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="38" name="txtBRDDIV"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="û�౸���� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call CleanField(txtADSTOCFLAG,'')"
												align="right">û�౸��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="ADSTOCFLAG" class="INPUT_L" id="txtADSTOCFLAG" title="��������" style="WIDTH: 282px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="38" name="txtADSTOCFLAG"></TD>
										</TR>
										<TR>
											<TD class="LABEL" title="�����縦 �����մϴ�." style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtINPUT_AREAFLAGNAME,'')"
												align="right">������</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="INPUT_AREAFLAGNAME" class="INPUT_L" id="txtINPUT_AREAFLAGNAME" title="������"
													style="WIDTH: 282px; HEIGHT: 22px" dataSrc="#xmlBind" maxLength="100" size="38" name="txtINPUT_AREAFLAGNAME"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtPRICE,'')"
												align="right">�� ��</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="PRICE" class="INPUT_R" id="txtPRICE" title="�ܰ�" style="WIDTH: 99px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="20" size="8" name="txtPRICE"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCNT,'')">Ƚ&nbsp;��</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="CNT" class="INPUT_R" id="txtCNT" title="ȸ��" style="WIDTH: 99px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="15" size="8" name="txtCNT"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtINPUT_AREAFLAGNAME,'')">�� 
												��</TD>
											<TD class="DATA"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="�ݾ�" style="WIDTH: 99px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="20" size="5" name="txtAMT"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField('','')">AOR����</TD>
											<TD class="DATA">&nbsp;&nbsp;&nbsp;<INPUT id="chkVOCH_TYPE" dataSrc="#xmlBind" dataFld="VOCH_TYPE" type="checkbox" name="chkVOCH_TYPE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMI_RATE,'')">��������</TD>
											<TD class="DATA"><INPUT dataFld="COMMI_RATE" class="INPUT_R" id="txtCOMMI_RATE" title="�ܰ�" style="WIDTH: 73px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="20" size="6" name="txtCOMMI_RATE">(%)</TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMISSION,'')">������</TD>
											<TD class="DATA"><INPUT dataFld="COMMISSION" class="INPUT_R" id="txtCOMMISSION" title="ȸ��" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" maxLength="15" size="8" name="txtCOMMISSION"></TD>
										</TR>
										<TR>
											<TD class="LABEL">����Ź��ȣ</TD>
											<TD class="DATA"><INPUT dataFld="TRU_TRANS_NO" class="NOINPUT_R" id="txtTRU_TRANS_NO" title="����Ź��ȣ" style="WIDTH: 99px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly maxLength="20" size="11" name="txtTRU_TRANS_NO"></TD>
											<TD class="LABEL">�������ȣ</TD>
											<TD class="DATA"><INPUT dataFld="COMMI_TRANS_NO" class="NOINPUT_R" id="txtCOMMI_TRANS_NO" title="�������ȣ"
													style="WIDTH: 99px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly maxLength="20" size="5" name="txtCOMMI_TRANS_NO"></TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 80%; HEIGHT: 100%" vAlign="top" align="left">
									<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="22489">
										<PARAM NAME="_ExtentY" VALUE="20531">
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
								</TD>
							</TR>
							<tr>
								<TD class="BOTTOMSPLIT" id="lblStatus" width="100%" colSpan="5" height="10"></TD>
							</tr>
						</table>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
