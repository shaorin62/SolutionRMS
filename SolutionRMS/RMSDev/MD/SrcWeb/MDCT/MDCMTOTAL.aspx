<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMTOTAL.aspx.vb" Inherits="MD.MDCMTOTAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����û�� ���/��ȸ</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : MD/TOTAL û��ȭ��(MDCMTOTAL)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMTOTAL.aspx
'��      �� : TOTAL��ü TOTAL Process ó��
'�Ķ�  ���� : 
'Ư��  ���� : ����ó��(���߼��� Row Coyp)
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/11 By HWNAG DUCK SU
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDMTTOTAL
Dim mobjMDCOGET 
Dim mstrCheck
Dim mcomecalender , mcomecalender2, mcomecalender3
Dim mstrPROCESS	'�ű��̸� True ��ȸ�� False
Dim mstrHIDDEN

CONST meTAB = 9
mstrHIDDEN = 0

mstrPROCESS = False
mstrCheck = True
mcomecalender = FALSE
mcomecalender2 = FALSE
mcomecalender3 = FALSE
	
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�Է��ʵ带 �����' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody").style.display = "inline"
			document.getElementById("tblSheet").style.height = "62%"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='�Է��ʵ� ��ġ��' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
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
Sub imgNEW_onclick ()
	initpageData
	Call sprSht_Keydown(meINS_ROW, 0)	
	mstrPROCESS = False
end Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' ���������Ѵ�.
'-----------------------------------------------------------------------------------------
Sub Imgcopy_onclick ()
	Dim intRtn
   	Dim vntData
	Dim intSelCnt,  i
	
	Dim strCHK, strYEARMON, strGFLAG, strSEQ, strDEMANDDAY, strCLIENTCODE, strCLIENTNAME, strGREATCODE, strGREATNAME
	Dim strTIMCODE, strTIMNAME, strSUBSEQ, strSUBSEQNAME, strMEDCODE, strMEDNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strAMT, strCOMMISSION, strCOMMI_RATE, strTBRDSTDATE, strTBRDEDDATE, strMPP_CODE, strMPP_NAME
	Dim strPROGRAM, strMATTERCODE, strMATTERNAME, strGRID_CNT, strVOCH_TYPE, strTRU_TAX_FLAG, strDEPT_CD, strDEPT_NAME 
	Dim strEXCLIENTCODE, strEXCLIENTNAME, strCLIENTSUBCODE, strCLIENTSUBNAME, strMEMO
	
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
		
		' ���Ŀ� �������� �ٲ� ���
		'strSEQ				=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",.sprSht.ActiveRow)
		
		strYEARMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON ",.sprSht.ActiveRow)
		strDEMANDDAY  		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY ",.sprSht.ActiveRow)
		strCLIENTCODE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		strCLIENTNAME 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",.sprSht.ActiveRow)
		strGREATCODE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"GREATCODE",.sprSht.ActiveRow)
		strGREATNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"GREATNAME",.sprSht.ActiveRow)
		strTIMCODE 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",.sprSht.ActiveRow)
		strTIMNAME 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",.sprSht.ActiveRow)
		strSUBSEQ 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
		strSUBSEQNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",.sprSht.ActiveRow)
		strMEDCODE 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",.sprSht.ActiveRow)
		strMEDNAME 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",.sprSht.ActiveRow)
		strMATTERCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",.sprSht.ActiveRow)
		strMATTERNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",.sprSht.ActiveRow)
		strREAL_MED_CODE 	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",.sprSht.ActiveRow)
		strREAL_MED_NAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",.sprSht.ActiveRow)
		strAMT 				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",.sprSht.ActiveRow)
		strCOMMISSION 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",.sprSht.ActiveRow)
		strCOMMI_RATE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",.sprSht.ActiveRow)
		strTBRDSTDATE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",.sprSht.ActiveRow)
		strTBRDEDDATE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",.sprSht.ActiveRow)
		strMPP_CODE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MPP_CODE",.sprSht.ActiveRow)
		strMPP_NAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MPP_NAME",.sprSht.ActiveRow)
		strPROGRAM 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",.sprSht.ActiveRow)
		strGRID_CNT 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CNT",.sprSht.ActiveRow)
		strVOCH_TYPE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",.sprSht.ActiveRow)
		strTRU_TAX_FLAG 	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow)
		strDEPT_CD			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",.sprSht.ActiveRow)
		strDEPT_NAME 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",.sprSht.ActiveRow)
		strEXCLIENTCODE 	=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",.sprSht.ActiveRow)
		strEXCLIENTNAME 	=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",.sprSht.ActiveRow)
		strCLIENTSUBCODE  	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow)
		strCLIENTSUBNAME  	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",.sprSht.ActiveRow)
		strMEMO 			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",.sprSht.ActiveRow)
	
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, meINS_ROW, 0, -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strYEARMON
		mobjSCGLSpr.SetTextBinding .sprSht,"GFLAG",.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, strCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, strCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",.sprSht.ActiveRow, strGREATCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",.sprSht.ActiveRow, strGREATNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, strTIMCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",.sprSht.ActiveRow, strTIMNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, strSUBSEQ
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, strSUBSEQNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",.sprSht.ActiveRow, strMEDCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",.sprSht.ActiveRow, strMEDNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",.sprSht.ActiveRow, strMATTERCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",.sprSht.ActiveRow, strMATTERNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, strREAL_MED_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, strREAL_MED_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",.sprSht.ActiveRow, strCOMMISSION
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, strCOMMI_RATE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",.sprSht.ActiveRow, strTBRDEDDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",.sprSht.ActiveRow, strMPP_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",.sprSht.ActiveRow, strMPP_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM",.sprSht.ActiveRow, strPROGRAM
		mobjSCGLSpr.SetTextBinding .sprSht,"CNT",.sprSht.ActiveRow, strGRID_CNT
		mobjSCGLSpr.SetTextBinding .sprSht,"VOCH_TYPE",.sprSht.ActiveRow, strVOCH_TYPE
		mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, strTRU_TAX_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",.sprSht.ActiveRow, strDEPT_CD
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",.sprSht.ActiveRow, strDEPT_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, strEXCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, strEXCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow, strCLIENTSUBCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",.sprSht.ActiveRow, strCLIENTSUBNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",.sprSht.ActiveRow, strMEMO

		gXMLSetFlag xmlBind, meUPD_TRANS
		mstrPROCESS = False
   	end With
end Sub


'����Ʈ
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i
	Dim chkcnt
	Dim strYEARMON
	Dim strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME
	Dim strSUBSEQ, strSUBSEQNAME, strMEDFLAG, strGFLAG
	Dim strSEQ
	
	Dim Con1, Con2, Con3
	Dim Con4, Con5, Con6
	Dim Con7, Con8, Con9	
	Dim Con10, Con11, Con12
	Dim Con13
	
	with frmThis
		Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = "" : Con7 = ""
		Con8 = "" : Con9 = "" : Con10 = "" : Con11 = "" : Con12 = "" : Con13 = ""
		
		if frmThis.sprSht.MaxRows = 0 then
			gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
			Exit Sub
		end if
		
		ModuleDir = "MD"

		ReportName = "MDCMTOTAL_MEDIUM.rpt"
		
		strYEARMON		 = .txtYEARMON1.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCLIENTNAME	 = .txtCLIENTNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value
		strTIMCODE		 = .txtTIMCODE1.value
		strTIMNAME		 = .txtTIMNAME1.value
		strMEDCODE		 = .txtMEDCODE1.value
		strMEDNAME		 = .txtMEDNAME1.value
		strSUBSEQ		 = .txtSUBSEQ1.value
		strSUBSEQNAME	 = .txtSUBSEQNAME1.value
		
		If strYEARMON <> ""			Then Con1  = " AND (YEARMON = '" & strYEARMON & "') "
		If strCLIENTCODE <> ""		Then Con2  = " AND (CLIENTCODE = '" & strCLIENTCODE & "')"
		If strCLIENTNAME <> ""		Then Con3  = " AND (DBO.SC_GET_HIGHCUSTNAME_FUN(CLIENTCODE) LIKE '%" & strCLIENTNAME & "%') "
		If strREAL_MED_CODE <> ""	Then Con4  = " AND (REAL_MED_CODE = '" & strREAL_MED_CODE & "') "
		If strREAL_MED_NAME <> ""	Then Con5  = " AND (DBO.SC_GET_HIGHCUSTNAME_FUN(REAL_MED_CODE) LIKE '%" & strREAL_MED_NAME & "%') "
		If strTIMCODE <> ""			Then Con6  = " AND (TIMCODE = '" & strTIMCODE & "') "
		If strTIMNAME <> ""			Then Con7  = " AND (DBO.SC_GET_CUSTNAME_FUN(TIMCODE) LIKE '%" & strTIMNAME & "%') "
		If strMEDCODE <> ""			Then Con8  = " AND (MEDCODE = '" & strMEDCODE & "')"
		If strMEDNAME <> ""			Then Con9  = " AND (DBO.SC_GET_CUSTNAME_FUN(MEDCODE) LIKE '%" & strMEDNAME & "%') "
		If strSUBSEQ <> ""			Then Con10 = " AND (SUBSEQ = '" & strSUBSEQ & "')"
		If strSUBSEQNAME <> ""		Then Con11 = " AND (DBO.SC_GET_SUBSEQNAME_FUN(SUBSEQ) LIKE '%" & strSUBSEQNAME & "%') "
		
		chkcnt=0
		For i=1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				if chkcnt = 0 then
					strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				else
					strSEQ = strSEQ & "," & mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)  
				end if 
				chkcnt = chkcnt +1
			End If
			
		Next
		
		if chkcnt <> 0 then
			Con12 = " AND ( SEQ IN (" & strSEQ &"))"
		end if 
        
		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & Con4 & ":" & Con5 & ":" & Con6 & ":" & Con7 & ":" & Con8 & ":" & Con9 & ":" & Con10 & ":" & Con11 & ":" & Con12
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt

	end with  
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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
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
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
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

'��ü�� �˾� ��ư
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE1_POP	
	Dim vntRet
	Dim vntInParams
	With frmThis

		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value),"MED_GEN")
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
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
			
			vntData = mobjMDCOGET.GetREAL_MED_CODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "MED_GEN")
			
			If not gDoErrorRtn ("GetREAL_MED_CODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
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
			
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
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

'��ü- ä�� �˾� ��ư
Sub ImgMEDCODE1_onclick
	Call MEDCODE1_POP()
End Sub

'���� ������List ��������
Sub MEDCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value),trim(.txtMEDCODE1.value), trim(.txtMEDNAME1.value), "MED_GEN")
	    
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), _
												trim(.txtMEDCODE1.value),trim(.txtMEDNAME1.value), "MED_GEN")
			
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
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,455)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ1.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME1.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE1.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtTIMNAME1.value = trim(vntRet(5,0))	' �����ָ� ǥ��
			SELECTRTN
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
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
					SELECTRTN
				Else
					Call SUBSEQCODE1_POP()
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
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtGREATCODE.value = trim(vntRet(4,0))
			.txtGREATNAME.value = trim(vntRet(5,0))
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
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
			vntData = mobjMDCOGET.GetHIGHCUSTCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE_ALL") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					.txtGREATCODE.value = trim(vntData(4,1))
					.txtGREATNAME.value = trim(vntData(5,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
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

		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value),"MED_GEN")
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetREAL_MED_CODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "MED_GEN")
			
			If not gDoErrorRtn ("GetREAL_MED_CODE") Then
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
			.txtGREATCODE.value = trim(vntRet(6,0))
			.txtGREATNAME.value = trim(vntRet(7,0))
					
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(7,0))
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
			vntData = mobjMDCOGET.GetTIMCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
									 		trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE_ALL") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTSUBCODE.value = trim(vntData(2,1))
					.txtCLIENTSUBNAME.value = trim(vntData(3,1))
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					.txtGREATCODE.value = trim(vntData(6,1))
					.txtGREATNAME.value = trim(vntData(7,1))
					
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(2,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(7,1))
						
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


'��ü��-ä�� �˾� ��ư-------
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'���� ������List ��������
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis   
	
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value), trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "MED_GEN")
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
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
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
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
												trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "MED_GEN")
			
			
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
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_CODE",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MPP_NAME",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
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


'�귣��
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_ALL.aspx",vntInParams , 640,435)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))	
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	
			.txtCLIENTCODE.value = trim(vntRet(2,0))	
			.txtCLIENTNAME.value = trim(vntRet(3,0))	
			.txtGREATCODE.value = trim(vntRet(4,0))	
			.txtGREATNAME.value = trim(vntRet(5,0))	
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
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
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
			vntData = mobjMDCOGET.Get_BrandInfo_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
													trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo_ALL") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	
					.txtCLIENTNAME.value = trim(vntData(3,1))	
					.txtGREATCODE.value = trim(vntData(4,1))
					.txtGREATNAME.value = trim(vntData(5,1))
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
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
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

'����� ��ư �˾�
Sub ImgMATTERCODE_onclick
	Call MATTERCODE_POP()
End Sub

Sub MATTERCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTNAME.value), trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
							trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "A2") '<< �޾ƿ��°��
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
			.txtGREATCODE.value = trim(vntRet(14,0))	' ����ó�ڵ� ǥ��
			.txtGREATNAME.value = trim(vntRet(15,0))	' ����ó�� ǥ��
			
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
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(14,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(15,0))
				
				
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
            
			vntData = mobjMDCOGET.GetMATTER_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtCLIENTNAME.value),trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
												trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "A2")
			If not gDoErrorRtn ("GetMATTER") Then
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
					.txtGREATCODE.value	 = trim(vntData(14,1))	' ����ó�ڵ� ǥ��
					.txtGREATNAME.value	 = trim(vntData(15,1))	' ����ó�� ǥ��
					
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
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(14,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(15,1))
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
			.txtGREATCODE.value = trim(vntRet(5,0))	'�ڵ�� ǥ��
			.txtGREATNAME.value = trim(vntRet(6,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(6,0))
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
			vntData = mobjMDCOGET.GetCLIENTSUBCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
		
			If not gDoErrorRtn ("GetCLIENTSUBCODE_ALL") Then
			
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,1))	'Code�� ����
					.txtCLIENTSUBNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(3,1))	'Code�� ����
					.txtCLIENTNAME.value = trim(vntData(4,1))	'�ڵ�� ǥ��
					.txtGREATCODE.value = trim(vntData(5,1))	'�ڵ�� ǥ��
					.txtGREATNAME.value = trim(vntData(6,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(6,1))
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

'����ó �˾�
Sub ImgGREATCODE_onclick
	Call GREATCODE_POP()
End Sub

Sub GREATCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtGREATCODE.value),trim(.txtGREATNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMGREATCUSTPOP.aspx",vntInParams , 413,440)
		
		If isArray(vntRet) Then
		    .txtGREATCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtGREATNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtGREATCODE
		End If
	end With
End Sub

Sub txtGREATNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetGREATCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtGREATCODE.value,.txtGREATNAME.value)
		
			If not gDoErrorRtn ("GetGREATCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtGREATCODE.value = trim(vntData(0,1))	'Code�� ����
					.txtGREATNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call GREATCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

' MPP�ڵ��˾� ��ư
Sub ImgMPP_onclick
	Call MPP_POP()
End Sub

'���� ������List ��������
Sub MPP_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtMPP.value, .txtMPP_NAME.value) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMMPPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtMPP.value = vntRet(0,0) and .txtMPP_NAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtMPP.value = vntRet(0,0)
			.txtMPP_NAME.value = vntRet(1,0)
			
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP",.sprSht.ActiveRow, .txtMPP.value
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",.sprSht.ActiveRow, .txtMPP_NAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtPROGRAM.focus()
			gSetChangeFlag .txtMPP                      ' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMPP_NAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
		
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetMPP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtMPP.value,.txtMPP_NAME.value)
			if not gDoErrorRtn ("GetREALMEDNO") then
				If mlngRowCnt = 1 Then
					.txtMPP.value = vntData(0,0)
					.txtMPP_NAME.value = vntData(1,0)
					
					if .sprSht.ActiveRow >0 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP",.sprSht.ActiveRow, .txtMPP.value
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",.sprSht.ActiveRow, .txtMPP_NAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtPROGRAM.focus()
					gSetChangeFlag .txtMPP
				Else
					Call MPP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' ������ �޷�
'****************************************************************************************
'û����
Sub imgCalEndar_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender2 = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar,"txtDEMANDDAY_onchange()"
	mcomecalender2 = false
	gXMLDataChanged xmlBind         
End Sub

'������
Sub imgCalEndar1_onclick
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtTBRDSTDATE,frmThis.imgCalEndar1,"txtTBRDSTDATE_onchange()"
	mcomecalender = false
	gXMLDataChanged xmlBind
End Sub

'������
Sub imgCalEndar2_onclick
	mcomecalender3 = true
	gShowPopupCalEndar frmThis.txtTBRDEDDATE,frmThis.imgCalEndar2,"txtTBRDEDDATE_onchange()"
	mcomecalender3 = false
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

Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
	'û���ϼ��� ������� ��������
		DateClean frmThis.txtYEARMON.value
		
		frmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEMANDDAY_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		txtDEMANDDAY_onchange
		frmThis.txtMATTERNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtMATTERCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMEDNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtMEDCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCNT_onkeydown
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

Sub txtTBRDSTDATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTBRDEDDATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTBRDEDDATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTIMNAME.focus()	
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTIMCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtSUBSEQNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSUBSEQ_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCLIENTSUBNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTSUBCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtGREATNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtGREATCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtREAL_MED_NAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtREAL_MED_CODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEXCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEXCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEPT_NAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEPT_CD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbVOCH_TYPE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
Sub cmbVOCH_TYPE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.chkTRU_TAX_FLAG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub chkTRU_TAX_FLAG_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPROGRAM.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
Sub txtPROGRAM_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMEMO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
Sub txtMEMO_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		FrmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************
Sub txtMATTERNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERNAME",frmThis.sprSht.ActiveRow, frmThis.txtMATTERNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtMATTERCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MATTERCODE",frmThis.sprSht.ActiveRow, frmThis.txtMATTERCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSUBSEQNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSUBSEQ_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, frmThis.txtSUBSEQ.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtTIMNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, frmThis.txtTIMNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtTIMCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, frmThis.txtTIMCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtDEMANDDAY_onchange
	Dim strdate 
	Dim strDEMANDDAY
	strdate = ""
	strDEMANDDAY =""
	With frmThis
		strdate=.txtDEMANDDAY.value
	
		If mcomecalender2 Then
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

Sub txtMEDNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDNAME",frmThis.sprSht.ActiveRow, frmThis.txtMEDNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtMEDCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEDCODE",frmThis.sprSht.ActiveRow, frmThis.txtMEDCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_CODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_CODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtDEPT_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtDEPT_CD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_CD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCOMMI_RATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, frmThis.txtCOMMI_RATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCOMMISSION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub cmbVOCH_TYPE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, frmThis.cmbVOCH_TYPE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub chkTRU_TAX_FLAG_onclick
	TRU_TAX_Flag_Disable
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtYEARMON_onchange	
	DateClean frmThis.txtYEARMON.value
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCNT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CNT",frmThis.sprSht.ActiveRow, frmThis.txtCNT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtTBRDSTDATE_onchange
	Dim strdate 
	Dim strTBRDSTDATE, strTBRDSTDATE2
	Dim strOLDYEARMON
	strdate = ""
	strTBRDSTDATE =""
	strTBRDSTDATE2 = ""

	With frmThis
		strdate=.txtTBRDSTDATE.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender Then
			strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strTBRDSTDATE2 = strdate
		else
			If len(strdate) = 4 Then
				strTBRDSTDATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strTBRDSTDATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strTBRDSTDATE2 = strdate
			elseif len(strdate) = 3 Then
				strTBRDSTDATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strTBRDSTDATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strTBRDSTDATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strTBRDSTDATE2 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE2
			DateClean_TBRDSTDATE strTBRDSTDATE
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With

	gSetChange
End Sub

Sub txtTBRDEDDATE_onchange
	Dim strdate 
	Dim TBRDEDDATE, TBRDEDDATE2
	Dim strOLDYEARMON
	strdate = ""
	TBRDEDDATE =""
	TBRDEDDATE2 = ""

	With frmThis
		strdate=.txtTBRDEDDATE.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender3 Then
			TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			TBRDEDDATE2 = strdate
		else
			If len(strdate) = 4 Then
				TBRDEDDATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				TBRDEDDATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				TBRDEDDATE2 = strdate
			elseif len(strdate) = 3 Then
				TBRDEDDATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				TBRDEDDATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				TBRDEDDATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				TBRDEDDATE2 = strdate
			End If
		End If
		
		If frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, TBRDEDDATE2
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		End If
	END With
End Sub
Sub txtCLIENTSUBNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtCLIENTSUBCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtGREATNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, frmThis.txtGREATNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtGREATCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, frmThis.txtGREATCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtREAL_MED_CODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, frmThis.txtREAL_MED_CODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtEXCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtEXCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtPROGRAM_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PROGRAM",frmThis.sprSht.ActiveRow, frmThis.txtPROGRAM.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtMEMO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

'-----------------------------------------------------------------------------------------
' CHK_TRU_TAX_FALG
'-----------------------------------------------------------------------------------------
 Sub TRU_TAX_Flag_Disable
	With frmThis
		If .chkTRU_TAX_FLAG.checked = False Then
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, 0
			End If
		else
			If .sprSht.ActiveRow > 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, 1
			End If
		End If	
	End With	
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ݾ�
Sub txtAMT_onblur
	With frmThis
		COMMI_RATE_Cal
		Call gFormatNumber(.txtAMT,0,True)
	end With
End Sub

'��������
Sub txtCOMMI_RATE_onblur
	With frmThis
		COMMI_RATE_Cal
	end With
End Sub

'������
Sub txtCOMMISSION_onblur
	With frmThis
		If frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMISSION",frmThis.sprSht.ActiveRow, frmThis.txtCOMMISSION.value
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		End If
		Call gFormatNumber(.txtCOMMISSION,0,True)
	end With
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ���ֱ� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ݾ�
Sub txtAMT_onfocus
	With frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end With
End Sub

'������
Sub txtCOMMISSION_onfocus
	With frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end With
End Sub


'****************************************************************************************
' ������ ���
'****************************************************************************************
'�������ڵ����
Sub COMMI_RATE_Cal
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,dblCOMMI_RATE
	
	With frmThis
		intAMT = .txtAMT.value
		
		If intAMT= "" Then  Exit Sub

		If .txtCOMMI_RATE.value ="" Then
			.txtCOMMI_RATE.value = 15
			dblCOMMI_RATE	= .txtCOMMI_RATE.value
		else
			dblCOMMI_RATE	= .txtCOMMI_RATE.value
		End If
			
		.txtCOMMISSION.value = intAMT * dblCOMMI_RATE /100
		
		txtCOMMI_RATE_onchange
		txtCOMMISSION_onchange
		
		gSetChangeFlag .txtAMT
		gSetChangeFlag .txtCOMMI_RATE
		gSetChangeFlag .txtCOMMISSION
	End With
	txtCOMMISSION_onblur
End Sub

'�������� �ڵ����
Sub COMMISSION_Cal
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,intCOMMISSION,dblCOMMI_RATE
	
	With frmThis
		If .txtAMT.value = "" then Exit Sub
		If .txtCOMMISSION.value = "" then Exit Sub
		
		intAMT = int(.txtAMT.value)
		intCOMMISSION = int(.txtCOMMISSION.value)
		
		If intAMT = 0 OR intAMT < intCOMMISSION Then
			.txtCOMMI_RATE.value = 0
		ELSE
			If intCOMMISSION <> "" AND intAMT <> "" Then
				dblCOMMI_RATE = gRound((intCOMMISSION /  intAMT * 100),2)
				.txtCOMMI_RATE.value = dblCOMMI_RATE
   			ELSE
   				.txtCOMMI_RATE.value = 0
			End If
		End If
		
		txtCOMMI_RATE_onchange
		txtCOMMISSION_onchange
		
		gSetChangeFlag .txtAMT
		gSetChangeFlag .txtCOMMI_RATE
		gSetChangeFlag .txtCOMMISSION
	End With
	txtCOMMISSION_onblur
End Sub


'�ݾ׿��� ���ͽ� ������ �ڵ����
Sub txtAMT_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtCNT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'������������ ���ͽ� ������ �ڵ����
Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtTBRDSTDATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'�����ῡ�� ���ͽ� �������� �ڵ����
Sub txtCOMMISSION_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMISSION_Cal
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
		
			sprShtToFieldBinding Col,Row
		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
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

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If mstrPROCESS = True Then
			frmThis.sprSht.MaxRows = 0
		End If
		
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, frmThis.txtYEARMON.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEMANDDAY",frmThis.sprSht.ActiveRow, frmThis.txtDEMANDDAY.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDSTDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDSTDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TBRDEDDATE",frmThis.sprSht.ActiveRow, frmThis.txtTBRDEDDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, "15"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GFLAG",frmThis.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TRU_TAX_FLAG",frmThis.sprSht.ActiveRow, "1"
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		strRow = frmThis.sprSht.ActiveRow
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,false,"YEARMON",1,strRow,false
		
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
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strCOLUMN = "COMMISSION"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
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
   	Dim strSTD_STEP, strSTD_CM, strSTD_FACE, strSTD_PAGE, strPRICE
   	Dim strAMT
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"YEARMON")  Then 
			.txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
			call DateClean_SHEET(.txtYEARMON.value ,Row)
			
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDDAY") Then .txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE")  Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
															  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntData(4,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntData(5,1)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						.txtCLIENTCODE.value = vntData(0,1)
						.txtCLIENTNAME.value = vntData(1,1)
						.txtGREATCODE.value = vntData(4,1)
						.txtGREATNAME.value = vntData(5,1)
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBCODE")  Then .txtMPP.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
			
				If not gDoErrorRtn ("GetCLIENTSUBCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,0))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,0))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(3,0))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(4,0))
						.txtCLIENTSUBCODE.value = trim(vntData(0,0))	'Code�� ����
						.txtCLIENTSUBNAME.value = trim(vntData(1,0))	'�ڵ�� ǥ��
						.txtCLIENTCODE.value = trim(vntData(3,0))	'Code�� ����
						.txtCLIENTNAME.value = trim(vntData(4,0))	'�ڵ�� ǥ��
						
						.txtCLIENTSUBNAME.focus
						.sprSht.focus 
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME"), Row
						.txtCLIENTSUBNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
				
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTCODE")  Then .txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
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
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMCODE")  Then .txtTIMCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetTIMCODE_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, trim(vntData(6,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, trim(vntData(7,1))
						
						.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
						.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
						.txtCLIENTSUBCODE.value = trim(vntData(2,1)) 
						.txtCLIENTSUBNAME.value = trim(vntData(3,1)) 
						.txtCLIENTCODE.value = trim(vntData(4,1))
						.txtCLIENTNAME.value = trim(vntData(5,1))
						.txtGREATCODE.value = trim(vntData(6,1))
						.txtGREATNAME.value = trim(vntData(7,1))
			
						.txtTIMNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtTIMNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQ")  Then .txtSUBSEQ.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_BrandInfo_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", strCodeName, _
														  "", "")
				If not gDoErrorRtn ("Get_BrandInfo_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntData(4,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntData(5,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntData(6,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntData(7,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(8,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntData(9,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntData(10,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntData(11,1)
						
						
						.txtSUBSEQ.value = vntData(0,1)
						.txtSUBSEQNAME.value = vntData(1,1)
						.txtCLIENTCODE.value = vntData(2,1)
						.txtCLIENTNAME.value = vntData(3,1)
						.txtGREATCODE.value = vntData(4,1)
						.txtGREATNAME.value =vntData(5,1)
						.txtTIMCODE.value = vntData(6,1)
						.txtTIMNAME.value = vntData(7,1)
						.txtCLIENTSUBCODE.value = vntData(8,1)
						.txtCLIENTSUBNAME.value = vntData(9,1)
						.txtDEPT_CD.value = vntData(10,1)
						.txtDEPT_NAME.value = vntData(11,1)
						
						
						.txtSUBSEQNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME"), Row
						.txtSUBSEQNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDCODE")  Then .txtMEDCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, ""
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, "","", strCode, strCodeName, "MED_GEN")
	
				If not gDoErrorRtn ("GetMEDGUBNCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntData(4,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",Row, vntData(5,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",Row, vntData(6,1)
						
						.txtMEDCODE.value = vntData(0,1)
						.txtMEDNAME.value = vntData(1,1)
						.txtREAL_MED_CODE.value = vntData(3,1)
						.txtREAL_MED_NAME.value = vntData(4,1)
						
						.txtMEDNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME"), Row
						.txtMEDNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_CODE")  Then .txtREAL_MED_CODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, ""
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetREAL_MED_CODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(strCode),trim(strCodeName), "MED_GEN")
			
				If not gDoErrorRtn ("mobjMDCOGET.GetREAL_MED_CODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						.txtREAL_MED_CODE.value = trim(vntData(0,1))
						.txtREAL_MED_NAME.value = trim(vntData(1,1))	
						
						.txtREAL_MED_NAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME"), Row
						.txtREAL_MED_NAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERCODE")  Then .txtMATTERCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MATTERNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "", "", strCodeName, "", "A2")

				If not gDoErrorRtn ("GetMATTER") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, trim(vntData(6,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, trim(vntData(7,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, trim(vntData(8,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, trim(vntData(9,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(10,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(11,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(12,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, trim(vntData(13,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, trim(vntData(14,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, trim(vntData(15,1))
						
						
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
						.txtGREATCODE.value	 = trim(vntData(14,1))	' ����ó�ڵ� ǥ��
						.txtGREATNAME.value	 = trim(vntData(15,1))	' ����ó�� ǥ��
						
						.txtMATTERNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME"), Row
						.txtMATTERNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_CD") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.txtDEPT_CD.value = trim(vntData(0,1))
						.txtDEPT_NAME.value = trim(vntData(1,1))
						
						.txtDEPT_NAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.txtDEPT_NAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATCODE")  Then .txtGREATCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"GREATCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"GREATNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetGREATCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(strCode),trim(strCodeName))
					
				If not gDoErrorRtn ("GetGREATCUSTCODE") Then
					If mlngRowCnt = 1 Then
						
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						
						.txtGREATCODE.value = trim(vntData(0,1))	'Code�� ����
						.txtGREATNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
						
						.txtGREATNAME.focus
						.sprSht.focus
						
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME"), Row
						.txtGREATNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End IF
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MPP_CODE") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MPP_CODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MPP_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MPP_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",Row, ""
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetMPP(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode, strCodeName)
			
				if not gDoErrorRtn ("GetMPP") then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",.sprSht.ActiveRow,  trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",.sprSht.ActiveRow,  trim(vntData(1,1))
						
						.txtPROGRAM.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MPP_NAME"), Row
						.txtPROGRAM.focus()
						.sprSht.focus 
					End If
   				end if
   			End If
		End If
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"), Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE"), Row)
			.txtCOMMI_RATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION"), Row)
			.txtCOMMISSION.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TBRDSTDATE")  Then 
			Dim strdate
			Dim strPUB_DATE
			Dim strYEARMON
			
			strdate = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
			strYEARMON = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			
			
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",Row, strdate
			mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",Row, strYEARMON
			
			DateClean_SHEET_TBRDDATE strYEARMON, Row
			
			.txtTBRDSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
			.txtTBRDEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		End If 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TBRDEDDATE")  Then .txtTBRDEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PROGRAM")  Then .txtPROGRAM.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"VOCH_TYPE") Then .cmbVOCH_TYPE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO") Then .txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TRU_TAX_FLAG") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",Row) = "1" Then
				.chkTRU_TAX_FLAG.checked = True
			Else
				.chkTRU_TAX_FLAG.checked = False
			End If
		End If
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'------------------------------------------------------
'SHEET CHANG �˾�������� �������϶� �˾�
'------------------------------------------------------
Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP_ALL.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntRet(5,0)
				
				.txtCLIENTCODE.value = vntRet(0,0)		
				.txtCLIENTNAME.value = vntRet(1,0)
				.txtGREATCODE.value = vntRet(4,0)
				.txtGREATNAME.value = vntRet(5,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then			
			vntInParams = array("", "", "" , TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCLIENTSUBPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				
				.txtCLIENTSUBCODE.value = trim(vntRet(0,0))	'Code�� ����
				.txtCLIENTSUBNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				.txtCLIENTCODE.value = trim(vntRet(3,0))	'Code�� ����
				.txtCLIENTNAME.value = trim(vntRet(4,0))	'�ڵ�� ǥ��
				
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
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then		
			vntInParams = array("","","", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)), "MED_GEN")
			
			vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
		
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",Row, vntRet(6,0)
				
				.txtMEDCODE.value = vntRet(0,0)
				.txtMEDNAME.value = vntRet(1,0)
				.txtREAL_MED_CODE.value = vntRet(3,0)
				.txtREAL_MED_NAME.value = vntRet(4,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)),"MED_GEN")
			
		    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				
				.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code�� ����
				.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"GREATNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMGREATCUSTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				
				.txtGREATCODE.value = trim(vntRet(0,0))	'Code�� ����
				.txtGREATNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
		
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)) , "", "")
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP_ALL.aspx",vntInParams , 640,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(6,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(8,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(9,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(10,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(11,0)
				
				.txtSUBSEQ.value = trim(vntRet(0,0))	
				.txtSUBSEQNAME.value = trim(vntRet(1,0))	
				.txtCLIENTCODE.value = trim(vntRet(2,0))	
				.txtCLIENTNAME.value = trim(vntRet(3,0))	
				.txtGREATCODE.value = trim(vntRet(4,0))	
				.txtGREATNAME.value = trim(vntRet(5,0))	
				.txtTIMCODE.value = trim(vntRet(6,0))	
				.txtTIMNAME.value = trim(vntRet(7,0))	
				.txtCLIENTSUBCODE.value = trim(vntRet(8,0))	
				.txtCLIENTSUBNAME.value = trim(vntRet(9,0))
				.txtDEPT_CD.value = trim(vntRet(10,0))	
				.txtDEPT_NAME.value = trim(vntRet(11,0))	
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then			
			vntInParams = array("", "" , "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP_ALL.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntRet(6,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntRet(7,0)
				
		
			
				.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
				.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
				.txtCLIENTSUBCODE.value = trim(vntRet(2,0))       ' �ڵ�� ǥ��
				.txtCLIENTSUBNAME.value = trim(vntRet(3,0))       ' �ڵ�� ǥ��
				.txtCLIENTCODE.value = trim(vntRet(4,0))    ' �ڵ�� ǥ��
				.txtCLIENTNAME.value = trim(vntRet(5,0))    ' �ڵ�� ǥ��
				.txtGREATCODE.value = trim(vntRet(6,0))
				.txtGREATNAME.value = trim(vntRet(7,0))
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then			
			vntInParams = array("","" , "", "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MATTERNAME",Row)), "", "A2")
			
			vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(6,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(8,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(9,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(10,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(11,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(12,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(13,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, vntRet(14,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"GREATNAME",Row, vntRet(15,0)
				
				.txtMATTERCODE.value = trim(vntRet(0,0))	' �����ڵ� ǥ��
				.txtMATTERNAME.value = trim(vntRet(1,0))	' ����� ǥ��
				.txtCLIENTCODE.value = trim(vntRet(2,0))	' �������ڵ� ǥ��
				.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
				.txtTIMCODE.value = trim(vntRet(4,0))		' ���ڵ� ǥ��
				.txtTIMNAME.value = trim(vntRet(5,0))		' ���� ǥ��
				.txtSUBSEQ.value = trim(vntRet(6,0))		' �귣�� ǥ��
				.txtSUBSEQNAME.value = trim(vntRet(7,0))	' �귣��� ǥ��
				.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' ���ۻ��ڵ� ǥ��
				.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' ���ۻ�� ǥ��
				.txtDEPT_CD.value = trim(vntRet(10,0))		' �μ��ڵ� ǥ��
				.txtDEPT_NAME.value = trim(vntRet(11,0))	' �μ��� ǥ��
				.txtCLIENTSUBCODE.value = trim(vntRet(12,0))	' ������ڵ� ǥ��
				.txtCLIENTSUBNAME.value = trim(vntRet(13,0))	' ����θ� ǥ��
				.txtGREATCODE.value = trim(vntRet(14,0))	' ����ó�ڵ� ǥ��
				.txtGREATNAME.value = trim(vntRet(15,0))	' ����ó�� ǥ��
				
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
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtCLIENTNAME.focus
		.sprSht.Focus
	
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MPP_NAME") Then			
			vntInParams = array( "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MPP_NAME",Row)))
		
			vntRet = gShowModalWindow("../MDCO/MDCMMPPPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"MPP_NAME",Row, vntRet(1,0)
			
			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
	End With
End Sub

'------------------------------------------------------
'��Ʈ ��������
'------------------------------------------------------
Sub SHEET_COMMI_RATE_Cal (Col, Row)
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,dblCOMMI_RATE, intCOMMISSION
	
	With frmThis
		If Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
			dblCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
			
			If intAMT = 0 OR intAMT < intCOMMISSION Then
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 0
				.txtCOMMI_RATE.value = 0
			else
				If intAMT <> 0 AND intCOMMISSION <> 0 AND dblCOMMI_RATE = 0.00 Then
					dblCOMMI_RATE = gRound((intCOMMISSION /  intAMT * 100),2)
   					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
   					.txtCOMMI_RATE.value = dblCOMMI_RATE
				ELSE
					dblCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
					intCOMMISSION = intAMT * dblCOMMI_RATE /100
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
					.txtCOMMISSION.value = intCOMMISSION
				End If
			End If
			
		ELSEIF Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_RATE") Then
		
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			
			If intAMT = 0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, "0"
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, "0"
				.txtCOMMI_RATE.value = 0
				.txtCOMMISSION.value = 0
			ELSE
				dblCOMMI_RATE = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
				intCOMMISSION = intAMT * dblCOMMI_RATE /100
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
				.txtCOMMISSION.value = intCOMMISSION
			End If
			
		ELSEIF Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
		
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
			
			If intAMT = 0 OR intAMT < intCOMMISSION Then
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, "0"
				.txtCOMMI_RATE.value = 0
			ELSE
				If intCOMMISSION <> "" AND intAMT <> "" Then
					dblCOMMI_RATE = gRound((intCOMMISSION /  intAMT * 100),2)
   					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
   					.txtCOMMI_RATE.value = dblCOMMI_RATE
   				ELSE
   					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, "0"
   					.txtCOMMI_RATE.value = 0
				End If
			End If
		End If
		
	End With
End Sub

'-------------------------------------------------
''��Ʈ�� �������ѷο��� ������ ��� �ʴ��� ���ε�
'-------------------------------------------------
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '�׸��� �����Ͱ� ������ ������.
		
		.txtYEARMON.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtSEQ.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		.txtMATTERNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		.txtMATTERCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		.txtSUBSEQ.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		.txtSUBSEQNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtTIMNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row)
		.txtTIMCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtDEMANDDAY.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		.txtMEDCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		.txtMEDNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		.txtREAL_MED_CODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtREAL_MED_NAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtDEPT_CD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		.txtDEPT_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtPROGRAM.value       =   mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",Row)
		.txtCNT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"CNT",Row)
		.txtMEMO.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtCOMMI_RATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtCOMMISSION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		.cmbVOCH_TYPE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row)
		.txtTBRDSTDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
		.txtTBRDEDDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",Row) then
			.chkTRU_TAX_FLAG.checked =true
		else
			.chkTRU_TAX_FLAG.checked =false
		end if
		.txtCLIENTSUBCODE.value =   mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",Row)
		.txtCLIENTSUBNAME.value =   mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",Row)
		.txtGREATCODE.value		=   mobjSCGLSpr.GetTextBinding(.sprSht,"GREATCODE",Row)
		.txtGREATNAME.value		=   mobjSCGLSpr.GetTextBinding(.sprSht,"GREATNAME",Row)
		.txtEXCLIENTCODE.value	=   mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		.txtEXCLIENTNAME.value	=   mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		
   	end With
   
	Call gFormatNumber(frmThis.txtAMT,0,True)
	Call gFormatNumber(frmThis.txtCOMMISSION,0,True)
	Call Field_Lock ()
End Function


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjMDMTTOTAL	= gCreateRemoteObject("cMDCT.ccMDCTTOTAL_MEDIUM")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 52, 0, 0, 0,0
		
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | AMT | COMMISSION | COMMI_RATE | MATTERCODE | MATTERNAME | TBRDSTDATE | TBRDEDDATE | PROGRAM | CNT | MPP_CODE | MPP_NAME | MEMO | VOCH_TYPE | TRU_TAX_FLAG | TRU_TRANS_NO | COMMI_TRANS_NO | GREATCODE | GREATNAME | TIMCODE | TIMNAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | CLIENTSUBCODE | CLIENTSUBNAME | OLDYEARMON | OLDSEQ | REAL_MED_BISNO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | MATTERNAME1 | DEPT_NAME1 | EXCLIENTNAME1 | AMT1 | COMMISSION1"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|û����|G|����|�������ڵ�|�����ָ�|�귣���ڵ�|�귣��|ä���ڵ�|ä��|��ü���ڵ�|��ü���|��ü�����ڹ�ȣ|����ݾ�|������|��������|�����ڵ�|�����|������|������|���α׷�|�ʼ�|MPP�ڵ�|MPP��|���|��ǥ����|VAT|����Ź�ŷ���ȣ|������ŷ���ȣ|����ó�ڵ�|����ó��|���ڵ�|����|���ۻ��ڵ�|���ۻ��|�μ��ڵ�|�μ���|������ڵ�|����θ�|OLDYEARMON|OLDSEQ|����ڹ�ȣ|�ŷ�ó��|��ü��|Client�μ���|�귣���|�����|�����μ�|Cre����|��ü��|���������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   6|     8|0|   0|         0|      13|         0|     12|      0|  12|         0|      12|               0|      10|    10|       8|       0|    12|     9|     9|      12|   6|      0|   12|  15|       8|  4|            12|             12|        0|      0|      0|   0|         0|       0|       0|     0|         0|       0|         0|     0|         0|       0|     0|           0|       0|     0|       0|      0|     0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK|TRU_TAX_FLAG "
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE ", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | COMMISSION | AMT1 | COMMISSION1", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "GFLAG | SEQ | TRU_TRANS_NO | COMMI_TRANS_NO | REAL_MED_BISNO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | MATTERNAME1 | DEPT_NAME1 | EXCLIENTNAME1 | AMT1 | COMMISSION1"   
		mobjSCGLSpr.ColHidden .sprSht, " SEQ | CLIENTCODE | GREATCODE | GREATNAME | MEDCODE | REAL_MED_CODE | SUBSEQ | TIMCODE | DEPT_CD | GFLAG | DEPT_NAME | CLIENTSUBNAME | MPP_CODE | OLDYEARMON | OLDSEQ", True
		mobjSCGLSpr.ColHidden .sprSht, " SEQ | CLIENTCODE | SUBSEQ | MEDCODE | REAL_MED_CODE | MATTERCODE | MPP_CODE | GREATCODE | GREATNAME | TIMCODE | TIMNAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | CLIENTSUBCODE | CLIENTSUBNAME | OLDYEARMON | OLDSEQ | REAL_MED_BISNO ", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | CNT | TBRDSTDATE | TBRDEDDATE| GFLAG | TRU_TRANS_NO | COMMI_TRANS_NO | REAL_MED_BISNO1",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME | SUBSEQNAME | MEDNAME | REAL_MED_NAME | MATTERNAME | DEPT_NAME | MPP_NAME | EXCLIENTNAME | MEMO | PROGRAM",-1,-1,0,2,false '����
		.sprSht.style.visibility = "visible"

    End With
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDMTTOTAL = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	With frmThis
		.sprSht.MaxRows = 0
		
		'�ʱⰪ ����
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		.txtYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	
		.txtTBRDSTDATE.value = gNowDate2
		
		
		'û������� �������� | ���۳���� ��������
		DateClean .txtYEARMON.value
		DateClean_TBRDSTDATE Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	
		
		'�⺻�� ����
		
		.txtCOMMI_RATE.value = "15"
		.chkTRU_TAX_FLAG.checked = True
		.txtYEARMON.focus
		.cmbVOCH_TYPE.value = "0"
		
		'Sheet�ʱ�ȭ
		Get_COMBO_VALUE
		Field_Lock	
	End With
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

	With frmThis
		.txtDEMANDDAY.value = date2
		If .sprSht.maxRows >= 1 Then
		mobjSCGLSpr.SetTextBinding .sprSht, "DEMANDDAY" , .sprSht.ActiveRow , .txtDEMANDDAY.value
		End If
		
	End With
End Sub

'������ ������ ��ȸ���� ����
Sub DateClean_TBRDSTDATE (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
		
	With frmThis
		if strYEARMON = "" THEN EXIT SUB
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		.txtTBRDEDDATE.value = date2
		
		If .sprSht.maxRows > 0 Then
			mobjSCGLSpr.SetTextBinding .sprSht, "TBRDEDDATE" , .sprSht.ActiveRow , .txtTBRDEDDATE.value
		End If
		
	End With
End Sub

Sub DateClean_SHEET (strYEARMON, Row)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		.txtDEMANDDAY.value = date2
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",Row, date2
	End With
End Sub

Sub DateClean_SHEET_TBRDDATE (strYEARMON, Row)
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	With frmThis
		.txtTBRDEDDATE.value = date2
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",Row, date2
	End With
End Sub


'-----------------------------------------------------------------------------------------
' �׸��� �޺��ڽ� ����
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData, vntData_VOCH, vntData_DUTY
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_VOCH = mobjMDMTTOTAL.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE",,,vntData_VOCH,,60 
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
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",.sprSht.ActiveRow) <> "" Then
				.txtYEARMON.className       = "NOINPUT_L" : .txtYEARMON.readOnly		= True 
			End If
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",.sprSht.ActiveRow) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> ""  Then
				'�⵵
				.txtYEARMON.className       = "NOINPUT_L" : .txtYEARMON.readOnly		= True 
				'�ʼ�
				.txtCNT.className			= "NOINPUT_L" : .txtCNT.readOnly			= True 
				'�濵�Ⱓ
				.txtTBRDSTDATE.className	= "NOINPUT_L" : .txtTBRDSTDATE.readOnly		= True : .imgCalEndar.disabled	 = True
				.txtTBRDEDDATE.className	= "NOINPUT_L" : .txtTBRDEDDATE.readOnly		= True : .imgCalEndar1.disabled  = True
				'����
				.txtMATTERNAME.className	= "NOINPUT_L" : .txtMATTERNAME.readOnly		= True : .ImgMATTERCODE.disabled = True
				.txtMATTERCODE.className	= "NOINPUT_L" : .txtMATTERCODE.readOnly		= True
				'�귣��
				.txtSUBSEQNAME.className	= "NOINPUT_L" : .txtSUBSEQNAME.readOnly		= True : .ImgSUBSEQCODE.disabled = True
				.txtSUBSEQ.className		= "NOINPUT_L" : .txtSUBSEQ.readOnly			= True
				'��
				.txtTIMNAME.className		= "NOINPUT_L" : .txtTIMNAME.readOnly		= True : .ImgTIMCODE.disabled	 = True
				.txtTIMCODE.className		= "NOINPUT_L" : .txtTIMCODE.readOnly		= True
				'���ۻ�
				.txtEXCLIENTNAME.className	= "NOINPUT_L" : .txtEXCLIENTNAME.readOnly		= True : .ImgEXCLIENTCODE.disabled = True
				.txtEXCLIENTCODE.className	= "NOINPUT_L" : .txtEXCLIENTCODE.readOnly		= True 
				'������
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				'���μ�
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .imgDEPT_CD.disabled = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				'û����
				.txtDEMANDDAY.className		= "NOINPUT"   : .txtDEMANDDAY.readOnly		= True : .imgCalEndar2.disabled  = True 
				'��ü
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled	 = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True
				'��ü��
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				'����ó
				.txtGREATNAME.className		= "NOINPUT_L" : .txtGREATNAME.readOnly		= True : .ImgGREATCODE.disabled = True
				.txtGREATCODE.className		= "NOINPUT_L" : .txtGREATCODE.readOnly		= True 
				'CIC/�����
				.txtCLIENTSUBNAME.className= "NOINPUT_L" : .txtCLIENTSUBNAME.readOnly	= True : .ImgCLIENTSUBCODE.disabled = True
				.txtCLIENTSUBCODE.className= "NOINPUT_L" : .txtCLIENTSUBCODE.readOnly	= True
				'���α׷�
				.txtPROGRAM.className		= "NOINPUT_L" : .txtPROGRAM.readOnly		= True
				'���/�ݾ�/��������/������
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True
				.txtAMT.className		= "NOINPUT_R" : .txtAMT.readOnly		= True
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True 
				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				'/VAT����
				.cmbVOCH_TYPE.disabled		= True : .chkTRU_TAX_FLAG.disabled = True

			else 
				
				'�ʼ�
				.txtCNT.className			= "INPUT_L" : .txtCNT.readOnly			= False 
				'�濵�Ⱓ
				.txtTBRDSTDATE.className	= "INPUT_L" : .txtTBRDSTDATE.readOnly	= False : .imgCalEndar.disabled	  = False
				.txtTBRDEDDATE.className	= "INPUT_L" : .txtTBRDEDDATE.readOnly	= False : .imgCalEndar1.disabled  = False
				'����
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : .ImgMATTERCODE.disabled = False
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
				'�귣��
				.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
				.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
				'��
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
				'���ۻ�
				.txtEXCLIENTNAME.className	= "INPUT_L" : .txtEXCLIENTNAME.readOnly	= False : .ImgEXCLIENTCODE.disabled = False
				.txtEXCLIENTCODE.className	= "INPUT_L" : .txtEXCLIENTCODE.readOnly	= False
				'������
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
				'û����
				.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
				'��ü��
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
				'����ó
				.txtGREATNAME.className		= "INPUT_L" : .txtGREATNAME.readOnly	= False : .ImgGREATCODE.disabled = False
				.txtGREATCODE.className		= "INPUT_L" : .txtGREATCODE.readOnly	= False 
				'CIC/�����
				.txtCLIENTSUBNAME.className= "INPUT_L" : .txtCLIENTSUBNAME.readOnly= False : .ImgCLIENTSUBCODE.disabled = False
				.txtCLIENTSUBCODE.className= "INPUT_L" : .txtCLIENTSUBCODE.readOnly= False
				'���α׷�
				.txtPROGRAM.className		= "INPUT_L" : .txtPROGRAM.readOnly		= False
				'���/�ܰ�/�ݾ�/��������/������
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
				.txtAMT.className		= "INPUT_R" : .txtAMT.readOnly		= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
				'����/����/ ��ǥ����/����/VAT����/�鼼����
				.cmbVOCH_TYPE.disabled		= False : .chkTRU_TAX_FLAG.disabled = False
			End If
		else
				'�⵵
				.txtYEARMON.className       = "INPUT_L" : .txtYEARMON.readOnly		= False 
				'�ʼ�
				.txtCNT.className			= "INPUT_L" : .txtCNT.readOnly			= False 
				'�濵�Ⱓ
				.txtTBRDSTDATE.className	= "INPUT_L" : .txtTBRDSTDATE.readOnly	= False : .imgCalEndar.disabled	  = False
				.txtTBRDEDDATE.className	= "INPUT_L" : .txtTBRDEDDATE.readOnly	= False : .imgCalEndar1.disabled  = False
				'����
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : .ImgMATTERCODE.disabled = False
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
				'�귣��
				.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
				.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
				'��
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
				'���ۻ�
				.txtEXCLIENTNAME.className	= "INPUT_L" : .txtEXCLIENTNAME.readOnly	= False : .ImgEXCLIENTCODE.disabled = False
				.txtEXCLIENTCODE.className	= "INPUT_L" : .txtEXCLIENTCODE.readOnly	= False
				'������
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
				'û����
				.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
				'��ü��
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
				'����ó
				.txtGREATNAME.className		= "INPUT_L" : .txtGREATNAME.readOnly	= False : .ImgGREATCODE.disabled = False
				.txtGREATCODE.className		= "INPUT_L" : .txtGREATCODE.readOnly	= False 
				'CIC/�����
				.txtCLIENTSUBNAME.className= "INPUT_L" : .txtCLIENTSUBNAME.readOnly= False : .ImgCLIENTSUBCODE.disabled = False
				.txtCLIENTSUBCODE.className= "INPUT_L" : .txtCLIENTSUBCODE.readOnly= False
				'���α׷�
				.txtPROGRAM.className		= "INPUT_L" : .txtPROGRAM.readOnly		= False
				'���/�ܰ�/�ݾ�/��������/������
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
				.txtAMT.className		= "INPUT_R" : .txtAMT.readOnly		= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
				'����/����/ ��ǥ����/����/VAT����/�鼼����
				.cmbVOCH_TYPE.disabled		= False : .chkTRU_TAX_FLAG.disabled = False
		
		End If
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim vntData2
	Dim strYEARMON, strCLIENTCODE,strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strTIMCODE, strTIMNAME,strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME, strMATTERNAME, strMEMO
   	Dim strMEDFLAG, strGFLAG, strVOCH_TYPE
   	Dim i, strCols
   	Dim strRows
	Dim intCnt, intCnt2
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		 = .txtYEARMON1.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strCLIENTNAME	 = .txtCLIENTNAME1.value
		strREAL_MED_CODE = .txtREAL_MED_CODE1.value
		strREAL_MED_NAME = .txtREAL_MED_NAME1.value
		strTIMCODE		 = .txtTIMCODE1.value
		strTIMNAME		 = .txtTIMNAME1.value
		strMEDCODE		 = .txtMEDCODE1.value
		strMEDNAME		 = .txtMEDNAME1.value
		strSUBSEQ		 = .txtSUBSEQ1.value
		strSUBSEQNAME	 = .txtSUBSEQNAME1.value
		strMATTERNAME	 = .txtMATTERNAME1.value
		strMEMO			 = .txtMEMO1.value
		strVOCH_TYPE	 = .cmbVOCH_TYPE1.value
	
		vntData = mobjMDMTTOTAL.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
										  strYEARMON, _
										  strCLIENTCODE, strCLIENTNAME, _
										  strREAL_MED_CODE, strREAL_MED_NAME, _
										  strTIMCODE, strTIMNAME, _
										  strMEDCODE,strMEDNAME, _
										  strSUBSEQ,strSUBSEQNAME,"", strMATTERNAME, strMEMO, strVOCH_TYPE)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt > 0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				
   				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",intCnt) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> ""  Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,41,True
   				
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   				'���� �˻��� ��Ƴ���
   				PreSearchFiledValue strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME
   			End If
   		End If

   		mstrPROCESS = True
   	end With
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME)
	With frmThis
		.txtYEARMON1.value		= strYEARMON
		.txtCLIENTCODE1.value	= strCLIENTCODE
		.txtCLIENTNAME1.value	= strCLIENTNAME
		.txtREAL_MED_CODE1.value= strREAL_MED_CODE
		.txtREAL_MED_NAME1.value= strREAL_MED_NAME
		.txtTIMCODE1.value		= strTIMCODE
		.txtTIMNAME1.value		= strTIMNAME
		.txtMEDCODE1.value		= strMEDCODE
		.txtMEDNAME1.value		= strMEDNAME
		.txtSUBSEQ1.value		= strSUBSEQ
		.txtSUBSEQNAME1.value	= strSUBSEQNAME

	End With
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

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strSEQ 
	Dim strYEARMON, strGFLAG, strVATFLAG
	Dim strPROJECTION
	Dim strSPONSOR
	Dim strMANAGENO
	Dim strDUTYFLAG
	Dim strDataCHK
	Dim lngCol, lngRow
	With frmThis
	
   		if  .sprSht.MaxRows = 0 then 
   			gErrorMsgBox "�ű��� ���� �����մϴ�.","����ȳ�"
   			exit sub
   		End if
   		
   		'������ Validation
		If DataValidation =False Then exit Sub
		'On error resume Next
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "YEARMON | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | EXCLIENTCODE | EXCLIENTNAME ",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ������/��/�귣��/ä��/��ü��/����/���ۻ� �� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | DEMANDDAY | GFLAG | SEQ | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | COMMISSION | COMMI_RATE | MATTERCODE | MATTERNAME | TBRDSTDATE | TBRDEDDATE | PROGRAM | CNT | MPP_CODE | MPP_NAME | MEMO | VOCH_TYPE | TRU_TAX_FLAG | TRU_TRANS_NO | COMMI_TRANS_NO | GREATCODE | GREATNAME | TIMCODE | TIMNAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | CLIENTSUBCODE | CLIENTSUBNAME | OLDYEARMON | OLDSEQ ")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		intRtn = mobjMDMTTOTAL.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
			.sprSht.focus()
   		End If
   	end With
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
	Dim vntData
   	Dim i, strCols
   	
	'On error resume Next
	With frmThis
		'������ Validation
		If not gDataValidation(frmThis) Then exit Function
   	End With
		DataValidation = True
End Function

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",i) <> "" Then
					gErrorMsgBox "�����Ͻ� " & i & "���� �ڷ�� �ŷ���ǥ�� ���� �մϴ�." & vbcrlf & "���� �ŷ���ǥ�� ���� �Ͻʽÿ�!","�����ȳ�!"
					exit Sub
				else 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",i) = "1" Then
						gErrorMsgBox "�����Ͻ� " & i & "���� �ڷ�� ���ε� �ڷ��Դϴ�." & vbcrlf & "���� �������ó�� �Ͻʽÿ�!","�����ȳ�!"
						exit Sub
					End If
					lngchkCnt = lngchkCnt +1
				End If
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
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDMTTOTAL.DeleteRtn(gstrConfigXml,dblSEQ, strYEARMON)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
	err.clear	
End Sub


'��ȣ�� Ŭ�����Ѵ�.
Sub CleanField (objField1, objField2)
	If frmThis.sprSht.MaxRows > 0 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"GFLAG",frmThis.sprSht.ActiveRow) = "0" Then
			
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
	End IF
End Sub

'�÷��׸� �ʱ�ȭ �Ѵ�.
Sub CleanFieldflag (objField1)
	If frmThis.sprSht.MaxRows > 0 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"GFLAG",frmThis.sprSht.ActiveRow) = "0" Then
			
			if isobject(objField1) then 
				objField1.value = "0"
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField1.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			end if
		End If
	End IF
End Sub


-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="118" background="../../../images/back_p.gIF"
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
											<td class="TITLE">û����� - ����û��</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 246px"
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
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE class="SEARCHDATA" id="tblKey" height="48" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, txtSEQ)"
									width="50">�� ��</TD>
								<TD class="SEARCHDATA" style="HEIGHT: 19pt" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 78px; HEIGHT: 22px" accessKey="NUM"
										maxLength="6" size="7" name="txtYEARMON1"><INPUT dataFld="SEQ" class="NOINPUT_L" id="txtSEQ" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										dataSrc="#xmlBind" maxLength="6" size="3" name="txtSEQ"></TD>
								<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">������</TD>
								<TD class="SEARCHDATA" style="HEIGHT: 19pt" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�����ָ�" style="WIDTH: 123px; HEIGHT: 22px"
										maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
									width="50">��</TD>
								<TD class="SEARCHDATA" style="HEIGHT: 19pt" width="200"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 123px; HEIGHT: 22px" maxLength="100"
										name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
										align="absMiddle" border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" maxLength="6"
										size="6" name="txtTIMCODE1"></TD>
								<TD class="SEARCHLABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1, txtSUBSEQ1)"
									width="50">�귣��</TD>
								<td class="SEARCHDATA" style="HEIGHT: 19pt" colspan="2"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="�귣���" style="WIDTH: 140px; HEIGHT: 22px"
										maxLength="100" size="18" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgSUBSEQ1"> <INPUT class="INPUT_L" id="txtSUBSEQ1" title="�귣���ڵ�" style="WIDTH: 55px; HEIGHT: 22px"
										maxLength="8" name="txtSUBSEQ1">
								</td>
							</TR>
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME1, '')">�����</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMATTERNAME1" title="�����" style="WIDTH: 120px; HEIGHT: 22px"
										maxLength="200" name="txtMATTERNAME1" size="26">
									<SELECT style="Z-INDEX: 0; WIDTH: 65px" id="cmbVOCH_TYPE1" title="����" name="cmbVOCH_TYPE1">
										<OPTION selected value="">��ü</OPTION>
										<OPTION value="0">����Ź</OPTION>
										<OPTION value="1">����</OPTION>
										<OPTION value="2">�Ϲ�</OPTION>
									</SELECT></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
									width="50">��ü��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="��ü���" style="WIDTH: 123px; HEIGHT: 22px"
										maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME1, txtMEDCODE1)"
									width="50">ä��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtMEDNAME1" title="ä�θ�" style="WIDTH: 123px; HEIGHT: 22px"
										maxLength="100" size="15" name="txtMEDNAME1"> <IMG id="ImgMEDCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgMEDCODE1"> <INPUT class="INPUT_L" id="txtMEDCODE1" title="ä���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										maxLength="6" size="2" name="txtMEDCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEMO1, '')">���</TD>
								<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMEMO1" title="���" style="WIDTH: 140px; HEIGHT: 22px" maxLength="200"
										name="txtMEMO1" size="26"></TD>
								<td>
									<IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
										alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery">&nbsp;
								</td>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 25px"></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="500" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" style="vAlign:'absmiddle'"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id="imgTableUp" style="CURSOR: hand" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgTableUp.gif"
														align="absMiddle" border="0" name="imgTableUp"></span> &nbsp;&nbsp;&nbsp;&nbsp;�հ� 
												: <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="top" align="right" height="28">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" height="20" alt="�ڷḦ �μ��մϴ�."
													src="../../../images/imgCho.gIF" border="0" name="imgCho"></TD>
											<TD><IMG id="imgNEW" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" height="20" alt="�ڷḦ �μ��մϴ�."
													src="../../../images/imgNew.gIF" border="0" name="imgNEW"></TD>
											<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imglistcopy.gIF" border="0" name="Imgcopy"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �μ��մϴ�."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 72px" vAlign="top" align="center">
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtYEARMON, '')"
												width="70">���</TD>
											<TD class="DATA" width="120"><INPUT dataFld="YEARMON" class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 118px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" size="13"
													name="txtYEARMON">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEMANDDAY, '')"
												width="50">û����</TD>
											<TD class="DATA" width="104"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="û����" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" maxLength="10" size="10" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" src="../../../images/btnCalEndar.gIF" height="16" align="absMiddle" border="0" name="imgCalEndar"></TD>
											<TD class="LABEL" style="HEIGHT: 22px; CURSOR: hand" onclick="vbscript:Call CleanField(txtTIMNAME,txtTIMCODE)"
												width="70">��</TD>
											<TD class="DATA" width="270"><INPUT dataFld="TIMNAME" class="INPUT_L" id="txtTIMNAME" title="����" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgTIMCODE"> <INPUT dataFld="TIMCODE" class="INPUT_L" id="txtTIMCODE" title="���ڵ�" style="WIDTH: 56px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtTIMCODE"></TD>
											<TD class="LABEL" style="HEIGHT: 22px; CURSOR: hand" onclick="vbscript:Call CleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="70">���ۻ�</TD>
											<TD class="DATA"><INPUT dataFld="EXCLIENTNAME" class="INPUT_L" id="txtEXCLIENTNAME" title="���ۻ��" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">
												<INPUT dataFld="EXCLIENTCODE" class="INPUT_L" id="txtEXCLIENTCODE" title="���ۻ��ڵ�" style="WIDTH: 55px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtEXCLIENTCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMATTERNAME, txtMATTERCODE)"
												width="70">�����</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="�����" style="WIDTH: 200px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="500" size="30" name="txtMATTERNAME"> <IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMATTERCODE">
												<INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="�����ڵ�" style="WIDTH: 58px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtMATTERCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtSUBSEQNAME, txtSUBSEQ)"
												width="70">�귣��</TD>
											<TD class="DATA"><INPUT dataFld="SUBSEQNAME" class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="32" name="txtSUBSEQNAME"> <IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgSUBSEQCODE">
												<INPUT dataFld="SUBSEQ" class="INPUT_L" id="txtSUBSEQ" title="�귣���ڵ�" style="WIDTH: 56px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtSUBSEQ"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPT_NAME, txtDEPT_CODE)"
												width="70">���μ�</TD>
											<TD class="DATA"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPT_NAME" title="���μ���" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtDEPT_NAME"> <IMG id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgDEPT_CD">
												<INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPT_CD" title="���μ��ڵ�" style="WIDTH: 55px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtDEPT_CODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEDNAME, txtMEDCODE)"
												width="70">ä��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="MEDNAME" class="INPUT_L" id="txtMEDNAME" title="ä�θ�" style="WIDTH: 200px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtMEDNAME"> <IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgMEDCODE"> <INPUT dataFld="MEDCODE" class="INPUT_L" id="txtMEDCODE" title="ä���ڵ�" style="WIDTH: 58px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtMEDCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)"
												width="70">CIC/�����</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTSUBNAME" class="INPUT_L" id="txtCLIENTSUBNAME" title="CIC/����θ�" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtCLIENTSUBNAME"> <IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTSUBCODE">
												<INPUT dataFld="CLIENTSUBCODE" class="INPUT_L" id="txtCLIENTSUBCODE" title="CIC/������ڵ�"
													style="WIDTH: 56px; HEIGHT: 22px" dataSrc="#xmlBind" maxLength="10" size="4" name="txtCLIENTSUBCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanFieldflag(cmbVOCH_TYPE)"
												width="70">û��</TD>
											<TD class="DATA"><SELECT dataFld="VOCH_TYPE" id="cmbVOCH_TYPE" title="û��" style="WIDTH: 85px" dataSrc="#xmlBind"
													name="cmbVOCH_TYPE">
													<OPTION value="0" selected>����Ź</OPTION>
													<OPTION value="1">����</OPTION>
													<OPTION value="2">�Ϲ�</OPTION>
													<OPTION value="3">AOR</OPTION>
												</SELECT></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtAMT, '')"
												width="70">����ݾ�</TD>
											<TD class="DATA" style="WIDTH: 120px"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="����ݾ�" style="WIDTH: 118px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="50" size="13" name="txtAMT"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCNT, '')"
												width="50">�ʼ�</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="CNT" class="INPUT_R" id="txtCNT" title="�ʼ�" style="WIDTH: 99px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="10" name="txtCNT">
											</TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="70">������</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="33" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 56px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtCLIENTCODE"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call gCleanField(chkTRU_TAX_FLAG, '')"
												width="70">VAT</TD>
											<TD class="DATA"><INPUT id="chkTRU_TAX_FLAG" title="VAT����" type="checkbox" CHECKED name="chkTRU_TAX_FLAG"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMISSION, '')"
												width="70">������</TD>
											<TD class="DATA" style="WIDTH: 120px"><INPUT dataFld="COMMISSION" class="INPUT_R" id="txtCOMMISSION" title="������" style="WIDTH: 118px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="50" size="30" name="txtCOMMISSION"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMI_RATE, '')"
												width="50">(%)</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="COMMI_RATE" class="INPUT_R" id="txtCOMMI_RATE" title="��������" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" maxLength="10" size="9" name="txtCOMMI_RATE">%</TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtGREATNAME, txtGREATCODE)"
												width="70">����ó</TD>
											<TD class="DATA" vAlign="middle"><INPUT dataFld="GREATNAME" class="INPUT_L" id="txtGREATNAME" title="����ó��" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtGREATNAME"> <IMG id="ImgGREATCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgGREATCODE">
												<INPUT dataFld="GREATCODE" class="INPUT_L" id="txtGREATCODE" title="����ó�ڵ�" style="WIDTH: 56px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtGREATCODE"></TD>
											<TD class="LABEL" style="HEIGHT: 25px; CURSOR: hand" onclick="vbscript:Call CleanField(txtPROGRAM, '')"
												width="70">���α׷�</TD>
											<TD class="DATA"><INPUT dataFld="PROGRAM" class="INPUT_L" id="txtPROGRAM" title="���α׷�" style="WIDTH: 253px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="500" size="45" name="txtPROGRAM"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTBRDSTDATE, txtTBRDEDDATE)"
												width="70">�濵�Ⱓ</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="TBRDSTDATE" class="INPUT" id="txtTBRDSTDATE" title="��۽�����" style="WIDTH: 110px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="2" name="txtTBRDSTDATE">&nbsp;<IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar1">
												~ <INPUT dataFld="TBRDEDDATE" class="INPUT" id="txtTBRDEDDATE" title="���������" style="WIDTH: 107px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" maxLength="10" size="2" name="txtTBRDEDDATE">&nbsp;<IMG id="imgCalEndar12" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar2"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)"
												width="70">��ü��</TD>
											<TD class="DATA" vAlign="middle"><INPUT dataFld="REAL_MED_NAME" class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 190px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="100" size="30" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE">
												<INPUT dataFld="REAL_MED_CODE" class="INPUT_L" id="txtREAL_MED_CODE" title="��ü���ڵ�" style="WIDTH: 56px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="10" size="4" name="txtREAL_MED_CODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEMO, '')"
												width="70">���</TD>
											<TD class="DATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="���" style="WIDTH: 253px; HEIGHT: 22px"
													dataSrc="#xmlBind" maxLength="120" size="12" name="txtMEMO"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblSheet" height="62%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="12752">
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
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
