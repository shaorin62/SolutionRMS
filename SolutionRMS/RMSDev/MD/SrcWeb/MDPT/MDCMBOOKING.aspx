<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMBOOKING.aspx.vb" Inherits="MD.MDCMBOOKING" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����û�� ���/��ȸ</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : MD/��ŷ ȭ��(MDCMBOOKING)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMBOOKING.aspx
'��      �� : �μ��ü Booking Process ó��
'�Ķ�  ���� : 
'Ư��  ���� : ����ó��(���߼��� Row Coyp)
'----------------------------------------------------------------------------------------
'HISTORY    :1) Old Ver. Kim Tae Yup
'			 2) 2008/08/14 By Kim Tae Ho
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
Dim mobjBOOK, mobjMDCOGET 
Dim mstrCheck
Dim mstrPub
Dim mcomecalender, mcomecalender2
Dim mstrPROCESS	'�ű��̸� True ��ȸ�� False
Dim mstrPROCESS2 '��ȸ�����̸� True �űԻ�12���̸� False
Dim mstrHIDDEN
Dim mstrSUM
mstrSUM = 0
CONST meTAB = 9
mstrPROCESS = False
mstrPROCESS2 = True
mstrCheck = True
mcomecalender = FALSE
mcomecalender2 = FALSE
mstrHIDDEN = 0
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			'document.getElementById("SizeOrSdt").innerHTML="������"
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

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i
	Dim chkcnt
	Dim strYEARMON
	Dim strSEQ
	Dim strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME
	Dim strSUBSEQ, strSUBSEQNAME, strMEDFLAG, strGFLAG, strVOCH_TYPE
	
	Dim Con1, Con2, Con3
	Dim Con4, Con5, Con6
	Dim Con7, Con8, Con9	
	Dim Con10, Con11, Con12
	Dim Con13, Con14, Con15
	
	with frmThis
		Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = "" : Con6 = "" : Con7 = ""
		Con8 = "" : Con9 = "" : Con10 = "" : Con11 = "" : Con12 = "" : Con13 = "" : Con14 = "" : Con15 = ""
		
		if frmThis.sprSht.MaxRows = 0 then
			gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
			Exit Sub
		end if
		
		ModuleDir = "MD"
		IF .cmbMED_FLAG1.value = "MP01" THEN
			ReportName = "MDCMBOOKING.rpt"
		ELSE
			ReportName = "MDCMBOOKING_MP02.rpt"
		END IF
		
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
		strMEDFLAG		 = .cmbMED_FLAG1.value
		strGFLAG		 = .cmbGFLAG1.value
		strVOCH_TYPE	 = .cmbVOCH_TYPE1.value
		
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
		If strMEDFLAG <> ""			Then Con12 = " AND (MED_FLAG = '" & strMEDFLAG & "')"
		If strGFLAG <> ""			Then Con13 = " AND (GFLAG = '" & strGFLAG & "')"
		If strVOCH_TYPE <> ""		Then 
			If strVOCH_TYPE = "PROJECTION" Then
				Con14 = " AND (PROJECTION = 'Y')"
			Else
				Con14 = " AND (VOCH_TYPE = '" & strVOCH_TYPE & "')"
			End If
		End If
		
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
			Con15 = " AND ( SEQ IN (" & strSEQ &"))"
		End if 

		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & Con4 & ":" & Con5 & ":" & Con6 & ":" & Con7 & ":" & Con8 & ":" & Con9 & ":" & Con10 & ":" & Con11 & ":" & Con12 & ":" & Con13 & ":" & Con14 & ":" & Con15
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
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
' ���������Ѵ�.
'-----------------------------------------------------------------------------------------
Sub Imgcopy_onclick ()
	Dim intRtn
   	Dim vntData
	Dim intSelCnt,  i
	Dim strCHK, strGFLAGNAME, strYEARMON, strSEQ, strMED_FLAG, strDIVMEDIA, strPUB_DATE, strDEMANDDAY, strCLIENTCODE, strCLIENTNAME
	Dim strMEDCODE, strMEDNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strSUBSEQ, strSUBSEQNAME, strTIMCODE, strTIMNAME, strMATTERCODE, strMATTERNAME
	Dim strDEPT_CD, strDEPT_NAME, strPUB_FACE, strEXECUTE_FACE, strSTD_STEP, strSTD_CM, strSTD_FACE, strSTD, strSTD_PAGE, strCOL_DEG
	Dim strPROJECTION, strPRICE, strAMT, strCOMMI_RATE, strCOMMISSION, strVOCH_TYPE, strRECEIPT_GUBUN, strTRU_TAX_FLAG, strDUTYFLAG
	Dim strMEMO, strTRU_TRANS_NO, strCOMMI_TRANS_NO, strGFLAG, strEXCLIENTCODE, strEXCLIENTNAME
	
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
		
		strYEARMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",.sprSht.ActiveRow)
		strMED_FLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",.sprSht.ActiveRow)
		strDIVMEDIA			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DIVMEDIA",.sprSht.ActiveRow)
		strPUB_DATE 		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",.sprSht.ActiveRow)
		strDEMANDDAY		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",.sprSht.ActiveRow)
		strCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
		strCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",.sprSht.ActiveRow)
		strMEDCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",.sprSht.ActiveRow)
		strMEDNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",.sprSht.ActiveRow)
		strREAL_MED_CODE	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",.sprSht.ActiveRow)
		strREAL_MED_NAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",.sprSht.ActiveRow)
		strSUBSEQ			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow)
		strSUBSEQNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",.sprSht.ActiveRow)
		strTIMCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",.sprSht.ActiveRow)
		strTIMNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",.sprSht.ActiveRow)
		strMATTERCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",.sprSht.ActiveRow)
		strMATTERNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",.sprSht.ActiveRow)
		strDEPT_CD			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",.sprSht.ActiveRow)
		strDEPT_NAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",.sprSht.ActiveRow)
		strPUB_FACE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",.sprSht.ActiveRow)
		strEXECUTE_FACE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXECUTE_FACE",.sprSht.ActiveRow)
		strSTD_STEP			=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",.sprSht.ActiveRow)
		strSTD_CM			=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",.sprSht.ActiveRow)
		strSTD_FACE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_FACE",.sprSht.ActiveRow)
		strSTD				=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD",.sprSht.ActiveRow)
		strSTD_PAGE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_PAGE",.sprSht.ActiveRow)
		strCOL_DEG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"COL_DEG",.sprSht.ActiveRow)
		strPROJECTION		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTION",.sprSht.ActiveRow)
		strPRICE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",.sprSht.ActiveRow)
		strAMT				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",.sprSht.ActiveRow)
		strCOMMI_RATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",.sprSht.ActiveRow)
		strCOMMISSION		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",.sprSht.ActiveRow)
		strVOCH_TYPE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",.sprSht.ActiveRow)
		strRECEIPT_GUBUN	=	mobjSCGLSpr.GetTextBinding(.sprSht,"RECEIPT_GUBUN",.sprSht.ActiveRow)
		strTRU_TAX_FLAG		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow)
		strDUTYFLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"DUTYFLAG",.sprSht.ActiveRow)
		strMEMO				=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",.sprSht.ActiveRow)
		strEXCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",.sprSht.ActiveRow)
		strEXCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",.sprSht.ActiveRow)
	
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, meINS_ROW, 0, -1, 1)
		
		Call Get_SUBCOMBO_VALUE2(strMED_FLAG,frmThis.sprSht.ActiveRow)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"GFLAGNAME",.sprSht.ActiveRow, "����"
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strYEARMON
		mobjSCGLSpr.SetTextBinding .sprSht,"MED_FLAG",.sprSht.ActiveRow, strMED_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"DIVMEDIA",.sprSht.ActiveRow, strDIVMEDIA
		mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",.sprSht.ActiveRow, strPUB_DATE
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, strCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, strCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",.sprSht.ActiveRow, strMEDCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",.sprSht.ActiveRow, strMEDNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, strREAL_MED_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, strREAL_MED_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",.sprSht.ActiveRow, strSUBSEQ
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, strSUBSEQNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, strTIMCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",.sprSht.ActiveRow, strTIMNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",.sprSht.ActiveRow, strMATTERCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",.sprSht.ActiveRow, strMATTERNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",.sprSht.ActiveRow, strDEPT_CD
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",.sprSht.ActiveRow, strDEPT_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"PUB_FACE",.sprSht.ActiveRow, strPUB_FACE
		mobjSCGLSpr.SetTextBinding .sprSht,"EXECUTE_FACE",.sprSht.ActiveRow, strEXECUTE_FACE
		mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",.sprSht.ActiveRow, strSTD_STEP
		mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",.sprSht.ActiveRow, strSTD_CM
		mobjSCGLSpr.SetTextBinding .sprSht,"STD_FACE",.sprSht.ActiveRow, strSTD_FACE
		mobjSCGLSpr.SetTextBinding .sprSht,"STD",.sprSht.ActiveRow, strSTD
		mobjSCGLSpr.SetTextBinding .sprSht,"STD_PAGE",.sprSht.ActiveRow, strSTD_PAGE
		mobjSCGLSpr.SetTextBinding .sprSht,"COL_DEG",.sprSht.ActiveRow, strCOL_DEG
		mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTION",.sprSht.ActiveRow, strPROJECTION
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",.sprSht.ActiveRow, strPRICE
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, strCOMMI_RATE
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",.sprSht.ActiveRow, strCOMMISSION
		mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTION",.sprSht.ActiveRow, strPROJECTION
		mobjSCGLSpr.SetTextBinding .sprSht,"VOCH_TYPE",.sprSht.ActiveRow, strVOCH_TYPE
		mobjSCGLSpr.SetTextBinding .sprSht,"RECEIPT_GUBUN",.sprSht.ActiveRow, strRECEIPT_GUBUN
		mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, strTRU_TAX_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"DUTYFLAG",.sprSht.ActiveRow, strDUTYFLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",.sprSht.ActiveRow, strMEMO
		mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TRANS_NO",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"GFLAG",.sprSht.ActiveRow, "M"
		
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, strEXCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, strEXCLIENTNAME

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
   		Dim i, strCols
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
   		Dim i, strCols
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
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
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
   		Dim i, strCols
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
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' �ڵ�� ǥ��
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
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
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
											trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					.txtCLIENTCODE.value = trim(vntData(4,1))
					.txtCLIENTNAME.value = trim(vntData(5,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))	
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))	
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
							trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "MED_PRINT")
	    
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
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
											trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "MED_PRINT")
			
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

'�귣��
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtTIMNAME.value = trim(vntRet(5,0))	' �����ָ� ǥ��
			.txtDEPT_CD.value = trim(vntRet(8,0))	' �����ָ� ǥ��
			.txtDEPT_NAME.value = trim(vntRet(9,0))	' �����ָ� ǥ��
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(3,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntRet(5,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(8,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(9,0))
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
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
												trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	' �������ڵ�
					.txtCLIENTNAME.value = trim(vntData(3,1))	' ������
					.txtTIMCODE.value = trim(vntData(4,1))		' ���ڵ�
					.txtTIMNAME.value = trim(vntData(5,1))		' ����
					.txtDEPT_CD.value = trim(vntData(8,1))		' �μ��ڵ�
					.txtDEPT_NAME.value = trim(vntData(9,1))	' �μ���
					
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQ",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUBSEQNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMCODE",frmThis.sprSht.ActiveRow, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TIMNAME",frmThis.sprSht.ActiveRow, trim(vntData(5,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(8,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(9,1))
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
							trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "B", TRIM(.txtMATTERCODE.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
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
			.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' ���ۻ�� ǥ��
			.txtDEPT_CD.value = trim(vntRet(10,0))		' �μ��ڵ� ǥ��
			.txtDEPT_NAME.value = trim(vntRet(11,0))	' �μ��� ǥ��
			
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
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME.value),trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
											trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "B")
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

'****************************************************************************************
' ������ �޷�
'****************************************************************************************
Sub imgCalEndar1_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtPUB_DATE,frmThis.imgCalEndar1,"txtPUB_DATE_onchange()"
	Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PUB_DATE"), frmThis.sprSht.ActiveRow)
	mcomecalender = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalEndar2_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender2 = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar2,"txtDEMANDDAY_onchange()"
	mcomecalender2 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
Sub txtMATTERCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtSUBSEQNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSUBSEQ_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtTIMNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtTIMCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCLIENTNAME1.focus()()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPUB_DATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPUB_DATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEMANDDAY.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
	mcomecalender = false
End Sub

Sub txtDEMANDDAY_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMEDNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
	mcomecalender2 = false
End Sub

Sub txtMEDCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtREAL_MED_NAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtREAL_MED_CODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtDEPT_NAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEPT_CD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPUB_FACE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPUB_FACE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEXECUTE_FACE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEXECUTE_FACE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		If frmThis.cmbMED_FLAG.value = "MP01" Then
			frmThis.txtSTD_STEP.focus()
		ELSE
			frmThis.txtSTD.focus()
		End If
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSTD_STEP_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtSTD_CM.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSTD_CM_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtSTD_FACE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSTD_FACE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbCOL_DEG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSTD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtSTD_PAGE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtSTD_PAGE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbCOL_DEG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtMEMO_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPRICE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPRICE_onkeydown
	If window.event.keyCode = meEnter Or window.event.keyCode = meTab Then
		priceCal
	End If
End Sub

Sub txtAMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtCOMMISSION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbVOCH_TYPE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub chkPROJECTION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMEMO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub chkRECEIPT_GUBUN_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.chkTRU_TAX_FLAG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub chkTRU_TAX_FLAG_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		'frmThis.cmbDUTYFLAG.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub cmbCOL_DEG_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.chkPROJECTION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub cmbMED_FLAG_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbDIVMEDIA.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
Sub cmbDIVMEDIA_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtMATTERNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub
Sub cmbVOCH_TYPE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.chkRECEIPT_GUBUN.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub cmbDUTYFLAG_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.chkGFLAG1.focus()
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

Sub txtPUB_DATE_onchange
	Dim strdate 
	Dim strPUB_DATE, strPUB_DATE2
	Dim strOLDYEARMON
	strdate = ""
	strPUB_DATE =""
	strPUB_DATE2 = ""
	With frmThis
		strdate=.txtPUB_DATE.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender Then
			strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strPUB_DATE2 = strdate
		else
			If len(strdate) = 4 Then
				strPUB_DATE = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strPUB_DATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strPUB_DATE2 = strdate
			elseif len(strdate) = 3 Then
				strPUB_DATE = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strPUB_DATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strPUB_DATE = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strPUB_DATE2 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			strOLDYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",.sprSht.ActiveRow)
			IF mstrPROCESS THEN
				If strOLDYEARMON  <> strPUB_DATE Then
					gErrorMsgBox "�������� ����� ������ �� �����ϴ�.",""
					.txtPUB_DATE.value = strdate
					EXIT Sub
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",.sprSht.ActiveRow, strPUB_DATE2
					mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strPUB_DATE
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				End If
			ELSE
				mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",.sprSht.ActiveRow, strPUB_DATE2
				mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strPUB_DATE
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			END IF
			Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PUB_DATE"), frmThis.sprSht.ActiveRow)
		else 
			.txtYEARMON.value = strPUB_DATE
			DateClean strPUB_DATE
			Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PUB_DATE"), frmThis.sprSht.ActiveRow)
		End If
	End With
	gSetChange
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
Sub txtPUB_FACE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PUB_FACE",frmThis.sprSht.ActiveRow, frmThis.txtPUB_FACE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtEXECUTE_FACE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXECUTE_FACE",frmThis.sprSht.ActiveRow, frmThis.txtEXECUTE_FACE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSTD_STEP_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_STEP",frmThis.sprSht.ActiveRow, frmThis.txtSTD_STEP.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSTD_CM_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_CM",frmThis.sprSht.ActiveRow, frmThis.txtSTD_CM.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSTD_FACE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_FACE",frmThis.sprSht.ActiveRow, frmThis.txtSTD_FACE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSTD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD",frmThis.sprSht.ActiveRow, frmThis.txtSTD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtSTD_PAGE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_PAGE",frmThis.sprSht.ActiveRow, frmThis.txtSTD_PAGE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtMEMO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MEMO",frmThis.sprSht.ActiveRow, frmThis.txtMEMO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub txtPRICE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PRICE",frmThis.sprSht.ActiveRow, frmThis.txtPRICE.value
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

Sub chkPROJECTION_onClick
	If frmThis.sprSht.ActiveRow >0 Then
		if frmThis.chkPROJECTION.checked = true then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PROJECTION",frmThis.sprSht.ActiveRow, "1"
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PROJECTION",frmThis.sprSht.ActiveRow, "0"
		end if
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub chkRECEIPT_GUBUN_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"RECEIPT_GUBUN",frmThis.sprSht.ActiveRow, frmThis.chkRECEIPT_GUBUN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub chkTRU_TAX_FLAG_onchange
	DutyFlag_Disable
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub cmbCOL_DEG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COL_DEG",frmThis.sprSht.ActiveRow, frmThis.cmbCOL_DEG.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub cmbMED_FLAG_onchange
	Dim strMED_FLAGNAME
	Call SUBCOMBO_TYPE()
	
	With frmThis
		If .cmbMED_FLAG.value = "MP01" Then
			document.getElementById("SizeOrSdt").innerHTML="������"
			pnlSIZE.style.display = "inline"
			pnlSTD.style.display = "none"

			.txtSTD_STEP.value = "15"
			.txtSTD_CM.value = "37.0"
			.txtSTD_FACE.value = "1"
			.txtSTD.value = ""
			.txtSTD_PAGE.value = ""
			If .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",.sprSht.ActiveRow, .txtSTD_STEP.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",.sprSht.ActiveRow, .txtSTD_CM.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_FACE",.sprSht.ActiveRow, .txtSTD_FACE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",.sprSht.ActiveRow, .txtSTD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_PAGE",.sprSht.ActiveRow, .txtSTD_PAGE.value
			End If
			
			gXMLNewBinding frmThis,xmlBind,"#xmlBind"
			
		elseif .cmbMED_FLAG.value = "MP02" Then
			document.getElementById("SizeOrSdt").innerHTML="�԰�"
			pnlSIZE.style.display = "none"
			pnlSTD.style.display = "inline"
			
			.txtSTD_STEP.value = ""
			.txtSTD_CM.value = ""
			.txtSTD_FACE.value = ""
			.txtSTD.value = ""
			.txtSTD_PAGE.value = "1"
			If .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",.sprSht.ActiveRow, .txtSTD_STEP.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",.sprSht.ActiveRow, .txtSTD_CM.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_FACE",.sprSht.ActiveRow, .txtSTD_FACE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",.sprSht.ActiveRow, .txtSTD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_PAGE",.sprSht.ActiveRow, .txtSTD_PAGE.value
			End If
			
			gXMLNewBinding frmThis,xmlBind,"#xmlBind"
		End If
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"MED_FLAG",.sprSht.ActiveRow, .cmbMED_FLAG.value
			Call Get_SUBCOMBO_VALUE(.cmbMED_FLAG.value)
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	end With
	gSetChange
End Sub

Sub cmbDIVMEDIA_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVMEDIA",frmThis.sprSht.ActiveRow, frmThis.cmbDIVMEDIA.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub cmbVOCH_TYPE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, frmThis.cmbVOCH_TYPE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub
Sub cmbDUTYFLAG_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DUTYFLAG",frmThis.sprSht.ActiveRow, frmThis.cmbDUTYFLAG.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub


'����/�鼼 ���� ����(�ΰ����� ���϶� ������ �� �ִ�.)
Sub DutyFlag_Disable
	With frmThis
		If .chkTRU_TAX_FLAG.checked = False Then
			.cmbDUTYFLAG.value = "Y"
			.cmbDUTYFLAG.disabled = False
			If .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, 0
				mobjSCGLSpr.SetTextBinding .sprSht,"DUTYFLAG",.sprSht.ActiveRow, .cmbDUTYFLAG.value
				mobjSCGLSpr.SetCellsLock2 .sprSht,False,"DUTYFLAG",.sprSht.ActiveRow,.sprSht.ActiveRow,False
			End If
		else
			.cmbDUTYFLAG.value = ""
			.cmbDUTYFLAG.disabled = True
			If .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TRU_TAX_FLAG",.sprSht.ActiveRow, 1
				mobjSCGLSpr.SetTextBinding .sprSht,"DUTYFLAG",.sprSht.ActiveRow, .cmbDUTYFLAG.value
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,"DUTYFLAG",.sprSht.ActiveRow,.sprSht.ActiveRow,False
			End If
		End If	
	End With	
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ܰ�
Sub txtPRICE_onblur
	With frmThis
		Call gFormatNumber(.txtPRICE,0,True)
		priceCal
	end With
End Sub

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
Sub txtPRICE_onfocus
	With frmThis
		.txtPRICE.value = Replace(.txtPRICE.value,",","")
	end With
End Sub

Sub txtAMT_onfocus
	With frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end With
End Sub

Sub txtCOMMISSION_onfocus
	With frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end With
End Sub


'****************************************************************************************
' ������ ���
'****************************************************************************************
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

Sub priceCal
	Dim strSTD_STEP
	Dim strSTD_CM
	Dim strSTD_FACE
	Dim strSTD_PAGE
	Dim strPRICE
	Dim strAMT
	'On error resume Next
	With frmThis
		strSTD_STEP = .txtSTD_STEP.value
		strSTD_CM	= .txtSTD_CM.value
		strSTD_FACE = .txtSTD_FACE.value
		strSTD_PAGE = .txtSTD_PAGE.value
		strPRICE	= .txtPRICE.value
		
		If .cmbMED_FLAG.value = "MP01" Then
			If strSTD_STEP <> "" AND  strSTD_CM <> "" AND  strSTD_FACE <> "" AND  strPRICE <> "" Then
				strAMT	= CDBL(strSTD_STEP) *  CDBL(strSTD_CM) *  CDBL(strSTD_FACE) *  CDBL(strPRICE)
			End If
		ELSE
			If strSTD_PAGE <> "" AND  strPRICE <> "" Then
				strAMT	= CDBL(strSTD_PAGE) * CDBL(strPRICE)
			End If
		End If
		
		.txtAMT.value = strAMT
		txtAMT_onchange
		COMMI_RATE_Cal
		
		.txtAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
   	end With
End Sub

'������������ ���ͽ� ������ �ڵ����
Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'�ݾ׿��� ���ͽ� ������ �ڵ����
Sub txtAMT_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
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
		If mstrPROCESS = True Then
			frmThis.sprSht.MaxRows = 0
		End If
		frmThis.txtSUMAMT.value = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		Call Get_SUBCOMBO_VALUE2("MP01",frmThis.sprSht.ActiveRow)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,False,frmThis.sprSht.ActiveRow,5,5,True
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GFLAGNAME",frmThis.sprSht.ActiveRow, "����"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GFLAG",frmThis.sprSht.ActiveRow, "M"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MED_FLAG",frmThis.sprSht.ActiveRow, "MP01"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVMEDIA",frmThis.sprSht.ActiveRow, "MPDIV01"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_STEP",frmThis.sprSht.ActiveRow, "15"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_CM",frmThis.sprSht.ActiveRow, "37.0"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_FACE",frmThis.sprSht.ActiveRow, "1"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STD_PAGE",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_RATE",frmThis.sprSht.ActiveRow, "15"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VOCH_TYPE",frmThis.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TRU_TAX_FLAG",frmThis.sprSht.ActiveRow, "1"
		DutyFlag_Disable
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COL_DEG",frmThis.sprSht.ActiveRow, "C/L"
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PUB_DATE",frmThis.sprSht.ActiveRow, gNowDate2
		Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PUB_DATE"), frmThis.sprSht.ActiveRow)
		
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
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MED_FLAG") Then
			.cmbMED_FLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row)
			Call Get_SUBCOMBO_VALUE2(mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row), Row)
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "MP01" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",Row, "15"
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",Row, "37.0"
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_FACE",Row, "1"
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_PAGE",Row, ""
			ELSE
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_FACE",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",Row, ""
				mobjSCGLSpr.SetTextBinding .sprSht,"STD_PAGE",Row, "1"
			End If
			'.cmbDIVMEDIA.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVMEDIA",Row)
			mobjSCGLSpr.SetTextBinding .sprSht,"DIVMEDIA",Row, .cmbDIVMEDIA.value
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVMEDIA")  Then .cmbDIVMEDIA.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVMEDIA",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PUB_DATE") Then	
			Dim strdate
			Dim strPUB_DATE
			Dim strYEARMON
			strdate = mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",Row)
			strYEARMON = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row) <> "" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row) <> strYEARMON Then
					gErrorMsgBox "�������� ����� ������ �� �����ϴ�.",""
					mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",Row, strdate
					EXIT Sub
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",Row, strdate
					mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",Row, strYEARMON
					DateClean_SHEET strYEARMON, Row
					.txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
					.txtPUB_DATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",Row)
					.txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
				End If
			Else
				mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",Row, strdate
				mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",Row, strYEARMON
				DateClean_SHEET strYEARMON, Row
				.txtYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
				.txtPUB_DATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",Row)
				.txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDDAY") Then .txtDEMANDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
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
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQ") Then .txtSUBSEQ.value = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", strCodeName, _
													  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row), mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row))

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
						
						.txtSUBSEQ.value = vntData(0,1)
						.txtSUBSEQNAME.value = vntData(1,1)
						.txtCLIENTCODE.value = vntData(2,1)
						.txtCLIENTNAME.value = vntData(3,1)
						.txtTIMCODE.value = vntData(4,1)
						.txtTIMNAME.value = vntData(5,1)
						.txtDEPT_CD.value = vntData(8,1)
						.txtDEPT_NAME.value = vntData(9,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUBSEQNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMCODE") Then .txtTIMCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row), mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row), "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, trim(vntData(5,1))
						
						.txtTIMCODE.value = trim(vntData(0,1))	    ' Code�� ����
						.txtTIMNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
						.txtCLIENTCODE.value = trim(vntData(4,1))
						.txtCLIENTNAME.value = trim(vntData(5,1))
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERCODE") Then .txtMATTERCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MATTERNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MATTERCODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row) = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row), _
												mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row), mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row), _
												mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row), strCodeName, mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row), "B")

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
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, trim(vntData(8,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(10,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(11,1))
						
						
						.txtMATTERCODE.value = trim(vntData(0,1))	' �����ڵ� ǥ��
						.txtMATTERNAME.value = trim(vntData(1,1))	' ����� ǥ��
						.txtCLIENTCODE.value = trim(vntData(2,1))	' �������ڵ� ǥ��
						.txtCLIENTNAME.value = trim(vntData(3,1))	' �����ָ� ǥ��
						.txtTIMCODE.value	 = trim(vntData(4,1))	' ���ڵ� ǥ��
						.txtTIMNAME.value	 = trim(vntData(5,1))	' ���� ǥ��
						.txtSUBSEQ.value	 = trim(vntData(6,1))	' �귣�� ǥ��
						.txtSUBSEQNAME.value = trim(vntData(7,1))	' �귣��� ǥ��
						.txtEXCLIENTCODE.value = trim(vntData(8,1))	' ���ۻ��ڵ� ǥ��
						.txtEXCLIENTNAME.value = trim(vntData(9,1))	' ���ۻ��ڵ� ǥ��
						.txtDEPT_CD.value	 = trim(vntData(10,1))	' �μ��ڵ� ǥ��
						.txtDEPT_NAME.value	 = trim(vntData(11,1))	' �μ��� ǥ��
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
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
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PUB_FACE") Then .txtPUB_FACE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXECUTE_FACE") Then .txtEXECUTE_FACE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EXECUTE_FACE",Row)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"STD_STEP") Then .txtSTD_STEP.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"STD_CM") Then .txtSTD_CM.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"STD_FACE") Then .txtSTD_FACE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_FACE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"STD") Then .txtSTD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"STD_PAGE") Then .txtSTD_PAGE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_PAGE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COL_DEG") Then .cmbCOL_DEG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COL_DEG",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PROJECTION") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTION",Row) = "1" Then
				.chkPROJECTION.checked = True
			Else
				.chkPROJECTION.checked = False
			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then 
			strSTD_STEP = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",Row)
			strSTD_CM	= mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",Row)
			strSTD_FACE = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_FACE",Row)
			strSTD_PAGE = mobjSCGLSpr.GetTextBinding(.sprSht,"STD_PAGE",Row)
			strPRICE	= mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row) = "MP01" Then
				If strSTD_STEP <> "" AND  strSTD_CM <> "" AND  strSTD_FACE <> "" AND  strPRICE <> "" Then
					strAMT	= CDBL(strSTD_STEP) *  CDBL(strSTD_CM) *  CDBL(strSTD_FACE) *  CDBL(strPRICE)
				End If
			ELSE 
				If strSTD_PAGE <> "" AND  strPRICE <> "" Then
					strAMT	= CDBL(strSTD_PAGE) * CDBL(strPRICE)
				End If
			End If
			mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, strAMT
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"), Row)
			.txtPRICE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
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
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"VOCH_TYPE") Then .cmbVOCH_TYPE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"RECEIPT_GUBUN") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"RECEIPT_GUBUN",Row) = "1" Then
				.chkRECEIPT_GUBUN.checked = True
			Else
				.chkRECEIPT_GUBUN.checked = False
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TRU_TAX_FLAG") Then 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",Row) = "1" Then
				.chkTRU_TAX_FLAG.checked = True
				.cmbDUTYFLAG.value = ""
				.cmbDUTYFLAG.disabled = True
				mobjSCGLSpr.SetTextBinding .sprSht,"DUTYFLAG",Row, .cmbDUTYFLAG.value
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,"DUTYFLAG",Row,Row,False
			Else
				.chkTRU_TAX_FLAG.checked = False
				.cmbDUTYFLAG.value = "Y"
				.cmbDUTYFLAG.disabled = False
				mobjSCGLSpr.SetTextBinding .sprSht,"DUTYFLAG",Row, .cmbDUTYFLAG.value
				mobjSCGLSpr.SetCellsLock2 .sprSht,False,"DUTYFLAG",Row,Row,False
			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DUTYFLAG") Then .cmbDUTYFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DUTYFLAG",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO") Then .txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

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
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(8,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(9,0)
				
				.txtSUBSEQ.value = trim(vntRet(0,0))		' �귣�� ǥ��
				.txtSUBSEQNAME.value = trim(vntRet(1,0))	' �귣��� ǥ��
				.txtCLIENTCODE.value = trim(vntRet(2,0))	' ������ ǥ��
				.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
				.txtTIMCODE.value = trim(vntRet(4,0))	' �����ָ� ǥ��
				.txtTIMNAME.value = trim(vntRet(5,0))	' �����ָ� ǥ��
				.txtDEPT_CD.value = trim(vntRet(8,0))	' �����ָ� ǥ��
				.txtDEPT_NAME.value = trim(vntRet(9,0))	' �����ָ� ǥ��
				
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
				
				.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code�� ����
				.txtTIMNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
				.txtCLIENTCODE.value = trim(vntRet(4,0))    ' �ڵ�� ǥ��
				.txtCLIENTNAME.value = trim(vntRet(5,0))    ' �ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MATTERNAME") Then			
			vntInParams = array("","" , "", "", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MATTERNAME",Row)), "", "B","")
			
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
		.txtCLIENTNAME1.focus()
		.sprSht.Focus
	End With
End Sub

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
			If Col = 4 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		Elseif Row = 0 and Col = 1 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			Elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			
			For intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	End With
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
		.txtSEQ.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		.txtMATTERNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",Row)
		.txtMATTERCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		.txtSUBSEQ.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row)
		.txtSUBSEQNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtTIMNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",Row)
		.txtTIMCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtPUB_DATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",Row)
		.txtDEMANDDAY.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",Row)
		.txtMEDCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
		.txtMEDNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",Row)
		.txtREAL_MED_CODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
		.txtREAL_MED_NAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtDEPT_CD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		.txtDEPT_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtPUB_FACE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",Row)
		.txtEXECUTE_FACE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXECUTE_FACE",Row)
		.txtMEMO.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row)
		.txtPRICE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtCOMMI_RATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtCOMMISSION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		
		.cmbCOL_DEG.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COL_DEG",Row)
		.cmbMED_FLAG.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",Row)
		
		.cmbVOCH_TYPE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",Row)
		
		Call SUBCOMBO_TYPE()
		.cmbDIVMEDIA.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DIVMEDIA",Row)
		
		If .cmbMED_FLAG.value = "MP01" Then
			document.getElementById("SizeOrSdt").innerHTML="������"
			pnlSIZE.style.display = "inline"
			pnlSTD.style.display = "none"
			
			mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_CM", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", False
			mobjSCGLSpr.ColHidden .sprSht, "STD", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", True
		
			.txtSTD_STEP.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",Row)
			.txtSTD_CM.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",Row)
			.txtSTD_FACE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_FACE",Row)
			.txtSTD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row)
			.txtSTD_PAGE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_PAGE",Row)
		ELSE
			document.getElementById("SizeOrSdt").innerHTML="�԰�"
			pnlSIZE.style.display = "none"
			pnlSTD.style.display = "inline"
			
			mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_CM", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", True
			mobjSCGLSpr.ColHidden .sprSht, "STD", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", False
			
			.txtSTD_STEP.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",Row)
			.txtSTD_CM.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",Row)
			.txtSTD_FACE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_FACE",Row)
			.txtSTD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row)
			.txtSTD_PAGE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"STD_PAGE",Row)
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TAX_FLAG",Row) = "1" Then
			.chkTRU_TAX_FLAG.checked = True
			.cmbDUTYFLAG.value = ""
			.cmbDUTYFLAG.disabled = True
		ELSE
			.chkTRU_TAX_FLAG.checked = False
			.cmbDUTYFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DUTYFLAG",Row)
			.cmbDUTYFLAG.disabled = False
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",Row) = "M" Then
			.chkGFLAG1.checked = True
			.chkGFLAG2.checked = False
			.chkGFLAG3.checked = False
			.chkGFLAG4.checked = False
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",Row) = "B" Then
			.chkGFLAG1.checked = False
			.chkGFLAG2.checked = True
			.chkGFLAG3.checked = False
			.chkGFLAG4.checked = False
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",Row) = "J" Then
			.chkGFLAG1.checked = False
			.chkGFLAG2.checked = False
			.chkGFLAG3.checked = True
			.chkGFLAG4.checked = False
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",Row) = "S" Then
			.chkGFLAG1.checked = False
			.chkGFLAG2.checked = False
			.chkGFLAG3.checked = False
			.chkGFLAG4.checked = True
		ELSE 
			.chkGFLAG1.checked = False
			.chkGFLAG2.checked = False
			.chkGFLAG3.checked = False
			.chkGFLAG4.checked = False
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTION",Row) = "1" Then
			.chkPROJECTION.checked = True
		ELSE
			.chkPROJECTION.checked = False
		End If
		
		If mobjSCGLSpr.GetTextBinding(.sprSht,"RECEIPT_GUBUN",Row) = "1" Then
			.chkRECEIPT_GUBUN.checked = True
		ELSE
			.chkRECEIPT_GUBUN.checked = False
		End If
   	end With
   
	Call gFormatNumber(frmThis.txtPRICE,0,True)
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
	set mobjBOOK		= gCreateRemoteObject("cMDPT.ccMDPTBOOKING")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 57, 0, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | GFLAGNAME | YEARMON | SEQ | MED_FLAG | DIVMEDIA | PUB_DATE | DEMANDDAY | CLIENTCODE |  CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BUSINO | SUBSEQ | SUBSEQNAME | MATTERCODE | MATTERNAME | AMT | COMMI_RATE | EXECUTE_FACE | STD_STEP | STD_CM | MEMO | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | PUB_FACE | STD_FACE | STD | STD_PAGE | COL_DEG | PROJECTION | PRICE | COMMISSION | VOCH_TYPE | RECEIPT_GUBUN | TRU_TAX_FLAG | DUTYFLAG | TRU_TRANS_NO | COMMI_TRANS_NO | GFLAG | EXCLIENTCODE | EXCLIENTNAME | REAL_MED_BUSINO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | MATTERNAME1 |DEPT_NAME1 | EXCLIENTNAME1 | AMT1 | COMMISSION1 | MATTERUSER"
											  '  1|          2|        3|    4|	        5|	       6|         7|          8|           9|           10|       11|       12|             13|             14|               15|	   16|          17|          18|          19|   20|          21|            22|       23|       24|       25|         26|        27|        28|      29|        30|   31|        32|       33|          34|     35|          36|         37|             38|            39|        40|    41|            42|              43|     44|            45|            46
		mobjSCGLSpr.SetHeader .sprSht,		 "����|G|���|����|��ü����|����|������|û����|�������ڵ�|�����ָ�|��ü�ڵ�|��ü��|��ü���ڵ�|��ü���|��ü�����ڹ�ȣ|�귣���ڵ�|�귣���|�����ڵ�|�����|�ݾ�|��������|�����|��|Cm|���|���ڵ�|����|�μ��ڵ�|�μ���|û���|��|�԰�|Page|����|����|�ܰ�|������|��ǥ����|����|VAT|�鼼����|����Ź�ŷ���ȣ|������ŷ���ȣ|GFLAG|���۴�����ڵ�|���۴�����|����ڹ�ȣ|�ŷ�ó��|��ü��|Client�μ���|�귣���|�����|�����μ�|Cre����|��ü��|���������|��������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|3|   0|   4|       6|   6|     8|     8|         0|      11|       0|    10|         0|      11|               0|         0|      11|       0|    11|   9|       4|     7| 3| 4|  12|     0|  10|       0|    10|    10| 4|   7|   5|   5|   4|   8|    10|       7|   4|  4|       7|            12|            12|    0|             0|           0|         0|       0|     0|           0|       0|     0|       0|      0|     0|         0|        12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | PROJECTION | RECEIPT_GUBUN | TRU_TAX_FLAG "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "COL_DEG", -1, -1, "C/L" & vbTab & "B/W" , 10, 40, False, False
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE | DEMANDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GFLAGNAME | YEARMON | MED_FLAG | DIVMEDIA | CLIENTCODE |  CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | DEPT_CD | DEPT_NAME | PUB_FACE | EXECUTE_FACE | STD | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | EXCLIENTNAME | REAL_MED_BUSINO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | MATTERNAME1 | DEPT_NAME1 | EXCLIENTNAME1 | MATTERUSER", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "STD_CM", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | STD_STEP | STD_FACE | STD_PAGE | PRICE | AMT | COMMISSION | AMT1 | COMMISSION1", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "GFLAGNAME | YEARMON | SEQ | TRU_TRANS_NO | COMMI_TRANS_NO | GFLAG | MATTERUSER"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE | MEDCODE | REAL_MED_CODE | SUBSEQ | TIMCODE | MATTERCODE | DEPT_CD | GFLAG | EXCLIENTCODE", True
		'mobjSCGLSpr.ColHidden .sprSht, "REAL_MED_BUSINO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | DEPT_NAME1 | EXCLIENTNAME1 | AMT1 | COMMISSION1", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | GFLAGNAME | STD | TRU_TRANS_NO | COMMI_TRANS_NO | REAL_MED_BUSINO1 | MATTERUSER",-1,-1,2,2,False
		
		.sprSht.style.visibility = "visible"

    End With
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjBOOK = Nothing
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
		.txtYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	'û���
		.txtPUB_DATE.value = gNowDate2
		
		'û���ϼ��� ������� ��������
		DateClean .txtYEARMON.value
		
		'�μ����� ����
		Call SUBCOMBO_TYPE()
		'�⺻�� ����
		.txtSTD_STEP.value = "15"
		.txtSTD_CM.value = "37.0"
		.txtSTD_FACE.value = "1"
		.txtCOMMI_RATE.value = "15"
		.chkPROJECTION.checked = False
		.chkRECEIPT_GUBUN.checked = False
		.chkTRU_TAX_FLAG.checked = True
		
		'������/�԰��Է��ʵ� ����
		document.getElementById("SizeOrSdt").innerHTML="������"
		pnlSIZE.style.display = "inline"
		pnlSTD.style.display = "none"
		
		mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", False
		mobjSCGLSpr.ColHidden .sprSht, "STD_CM", False
		mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", False
		mobjSCGLSpr.ColHidden .sprSht, "STD", True
		mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", True
				
		'Sheet�ʱ�ȭ
		.txtYEARMON1.focus
		
		Field_Lock
		DutyFlag_Disable
		Get_COMBO_VALUE
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
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",Row, date2
	End With
End Sub

'-----------------------------------------------------------------------------------------
' SUBCOMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()
	Dim vntPUB_FACE
	Dim strMED_FLAG
	Dim vntMED_FLAG_DIVMEDIA
	With frmThis   
		strMED_FLAG = "MP_" & .cmbMED_FLAG.value
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
       	
       	vntMED_FLAG_DIVMEDIA = mobjBOOK.GetDataType_DIVMEDIA(gstrConfigXml, mlngRowCnt, mlngColCnt, .cmbMED_FLAG.value)
		If not gDoErrorRtn ("GetDataTypeChange") Then 
			 gLoadComboBox .cmbDIVMEDIA, vntMED_FLAG_DIVMEDIA, False
   		End If  
   		gSetChange
   	end With   
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
		
		vntData = mobjBOOK.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		vntData_VOCH = mobjBOOK.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		vntData_DUTY = mobjBOOK.Get_COMBODUGY_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "MED_FLAG",,,vntData,,50 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE",,,vntData_VOCH,,60 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DUTYFLAG",,,vntData_DUTY,,60 
			mobjSCGLSpr.TypeComboBox = True 
			'Call Get_SUBCOMBO_VALUE("MP01")
   		End If
   	End With
End Sub

'-----------------------------------------------------------------------------------------
' �׸��� ���� �޺� ����
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strMED_FLAG)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
       	vntData = mobjBOOK.GetDataType_DIVMEDIA(gstrConfigXml, mlngRowCnt, mlngColCnt, strMED_FLAG)
		If not gDoErrorRtn ("GetDataType_DIVMEDIA") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DIVMEDIA",,,vntData,,80 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

Sub Get_SUBCOMBO_VALUE2(strMED_FLAG, Row)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
       	vntData = mobjBOOK.GetDataType_DIVMEDIA(gstrConfigXml, mlngRowCnt, mlngColCnt, strMED_FLAG)
		If not gDoErrorRtn ("GetDataType_DIVMEDIA") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DIVMEDIA",Row,Row,vntData,,80 
			gLoadComboBox .cmbDIVMEDIA, vntData, False
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub


Sub Set_RowCOMBO(strMED_FLAG, Row)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
       	vntData = mobjBOOK.GetDataType_DIVMEDIA(gstrConfigXml, mlngRowCnt, mlngColCnt, strMED_FLAG)
		mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DIVMEDIA",Row,Row,vntData,,80 
		mobjSCGLSpr.TypeComboBox = True 
   		gSetChange
   	end With   
End Sub


'-----------------------------------------------------------------------------------------
' Field_Lock  �ŷ�������ȣ�� ���ݰ�꼭 ��ȣ�� ������ �����Ҽ� ������ �ʵ带 ReadOnlyó��
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",.sprSht.ActiveRow) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> "" Then
				'����
				.cmbMED_FLAG.disabled = True : .cmbDIVMEDIA.disabled = True
				'����
				.txtMATTERNAME.className	= "NOINPUT_L" : .txtMATTERNAME.readOnly		= True : .ImgMATTERCODE.disabled = True
				.txtMATTERCODE.className	= "NOINPUT_L" : .txtMATTERCODE.readOnly		= True
				'�귣��
				.txtSUBSEQNAME.className	= "NOINPUT_L" : .txtSUBSEQNAME.readOnly		= True : .ImgSUBSEQCODE.disabled = True
				.txtSUBSEQ.className		= "NOINPUT_L" : .txtSUBSEQ.readOnly			= True
				'��
				.txtTIMNAME.className		= "NOINPUT_L" : .txtTIMNAME.readOnly		= True : .ImgTIMCODE.disabled	 = True
				.txtTIMCODE.className		= "NOINPUT_L" : .txtTIMCODE.readOnly		= True
				'û����
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				'������/û����
				.txtPUB_DATE.className		= "NOINPUT"   : .txtPUB_DATE.readOnly		= True : .imgCalEndar1.disabled  = True 
				.txtDEMANDDAY.className		= "NOINPUT"   : .txtDEMANDDAY.readOnly		= True : .imgCalEndar2.disabled  = True 
				'��ü
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled	 = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True
				'��ü��
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				'���μ�
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .imgDEPT_CD.disabled	 = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				'û���
				.txtPUB_FACE.className		= "NOINPUT_L" : .txtPUB_FACE.readOnly		= True
				'�����
				.txtEXECUTE_FACE.className	= "NOINPUT_L" : .txtEXECUTE_FACE.readOnly	= True
				'������/�԰�
				.txtSTD_STEP.className		= "NOINPUT_R" : .txtSTD_STEP.readOnly		= True
				.txtSTD_CM.className		= "NOINPUT_R" : .txtSTD_CM.readOnly			= True
				.txtSTD_FACE.className		= "NOINPUT_R" : .txtSTD_FACE.readOnly		= True
				.txtSTD.className			= "NOINPUT_R" : .txtSTD.readOnly			= True
				.txtSTD_PAGE.className		= "NOINPUT_R" : .txtSTD_PAGE.readOnly		= True
				'���/�ܰ�/�ݾ�/��������/������
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True
				.txtPRICE.className			= "NOINPUT_R" : .txtPRICE.readOnly			= True 
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly			= True
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True 
				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				'����/����/ ��ǥ����/����/VAT����/�鼼����
				.cmbCOL_DEG.disabled		= True : .chkPROJECTION.disabled	= True
				.cmbVOCH_TYPE.disabled		= True : .chkRECEIPT_GUBUN.disabled = True
				.chkTRU_TAX_FLAG.disabled	= True : .cmbDUTYFLAG.disabled		= True
			else 
				'����
				.cmbMED_FLAG.disabled = False : .cmbDIVMEDIA.disabled = False
				'����
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : .ImgMATTERCODE.disabled = False
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
				'�귣��
				.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
				.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
				'��
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
				'û����
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
				'������/û����
				.txtPUB_DATE.className		= "INPUT"   : .txtPUB_DATE.readOnly		= False : .imgCalEndar1.disabled  = False 
				.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
				'��ü
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
				'��ü��
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
				'���μ�
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
				'û���
				.txtPUB_FACE.className		= "INPUT_L" : .txtPUB_FACE.readOnly		= False
				'�����
				.txtEXECUTE_FACE.className	= "INPUT_L" : .txtEXECUTE_FACE.readOnly	= False
				'������/�԰�
				.txtSTD_STEP.className		= "INPUT_R" : .txtSTD_STEP.readOnly		= False
				.txtSTD_CM.className		= "INPUT_R" : .txtSTD_CM.readOnly		= False
				.txtSTD_FACE.className		= "INPUT_R" : .txtSTD_FACE.readOnly		= False
				.txtSTD.className			= "INPUT_R" : .txtSTD.readOnly			= False
				.txtSTD_PAGE.className		= "INPUT_R" : .txtSTD_PAGE.readOnly		= False
				'���/�ܰ�/�ݾ�/��������/������
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
				.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly		= False 
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
				'����/����/ ��ǥ����/����/VAT����/�鼼����
				.cmbCOL_DEG.disabled		= False : .chkPROJECTION.disabled	 = False
				.cmbVOCH_TYPE.disabled		= False : .chkRECEIPT_GUBUN.disabled = False
				.chkTRU_TAX_FLAG.disabled	= False
				If .chkTRU_TAX_FLAG.checked = True Then
					.cmbDUTYFLAG.disabled	= True
				ELSE
					.cmbDUTYFLAG.disabled	= False
				End If
			End If
		else
			'����
			.cmbMED_FLAG.disabled = False : .cmbDIVMEDIA.disabled = False
			'����
			.txtMATTERNAME.className		= "INPUT_L" : .txtMATTERNAME.readOnly		= False : .ImgMATTERCODE.disabled = False
			.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
			'�귣��
			.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
			.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
			'��
			.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
			.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
			'û����
			.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
			.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
			'������/û����
			.txtPUB_DATE.className		= "INPUT"   : .txtPUB_DATE.readOnly		= False : .imgCalEndar1.disabled  = False 
			.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
			'��ü
			.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
			.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
			'��ü��
			.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
			.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
			'���μ�
			.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
			.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
			'û���
			.txtPUB_FACE.className		= "INPUT_L" : .txtPUB_FACE.readOnly		= False
			'�����
			.txtEXECUTE_FACE.className	= "INPUT_L" : .txtEXECUTE_FACE.readOnly	= False
			'������/�԰�
			.txtSTD_STEP.className		= "INPUT_R" : .txtSTD_STEP.readOnly		= False
			.txtSTD_CM.className		= "INPUT_R" : .txtSTD_CM.readOnly		= False
			.txtSTD_FACE.className		= "INPUT_R" : .txtSTD_FACE.readOnly		= False
			.txtSTD.className			= "INPUT_R" : .txtSTD.readOnly			= False
			.txtSTD_PAGE.className		= "INPUT_R" : .txtSTD_PAGE.readOnly		= False
			'���/�ܰ�/�ݾ�/��������/������
			.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
			.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly		= False 
			.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= False
			.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
			.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
			'����/����/ ��ǥ����/����/VAT����/�鼼����
			.cmbCOL_DEG.disabled		= False : .chkPROJECTION.disabled	 = False
			.cmbVOCH_TYPE.disabled		= False : .chkRECEIPT_GUBUN.disabled = False
			.chkTRU_TAX_FLAG.disabled	= False
			If .chkTRU_TAX_FLAG.checked = True Then
				.cmbDUTYFLAG.disabled	= True
			ELSE
				.cmbDUTYFLAG.disabled	= False
			End If
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
	Dim strTIMCODE, strTIMNAME,strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME
   	Dim strMEDFLAG, strGFLAG, strVOCH_TYPE
   	Dim i, strCols
   	Dim strRows
	Dim intCnt, intCnt2
	Dim strtemp
	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
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
		strMEDFLAG		 = .cmbMED_FLAG1.value
		strGFLAG		 = .cmbGFLAG1.value
		strVOCH_TYPE	 = .cmbVOCH_TYPE1.value
		
		If strMEDFLAG = "MP01" Then
			mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_CM", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", False
			mobjSCGLSpr.ColHidden .sprSht, "STD", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", True
		ELSE 
			mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_CM", True
			mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", True
			mobjSCGLSpr.ColHidden .sprSht, "STD", False
			mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", False
		End If

		'Call Get_SUBCOMBO_VALUE(strMEDFLAG)

		vntData = mobjBOOK.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
									strYEARMON, _
									strCLIENTCODE, strCLIENTNAME, _
									strREAL_MED_CODE, strREAL_MED_NAME, _
									strTIMCODE, strTIMNAME, _
									strMEDCODE, strMEDNAME, _
									strSUBSEQ, strSUBSEQNAME, _
									strMEDFLAG, strGFLAG, strVOCH_TYPE)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
	   			For intCnt = 1 To .sprSht.MaxRows
	   			
	   				'for�� �ѹ����� �ּ�ȭ �ϱ����� ���⿡ ��ġ
	   				strtemp = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVMEDIA",intCnt)
	   				Call Set_RowCOMBO (mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",intCnt), intCnt)
	   				mobjSCGLSpr.SetTextBinding .sprSht,"DIVMEDIA",intCnt,strtemp
	   				
					If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",intCnt) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> ""  Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,44,True
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   				PreSearchFiledValue strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME, strMEDFLAG, strGFLAG, strVOCH_TYPE
   			End If
   			
   			
   			
   		End If
   		Layout_change
   		mstrPROCESS = True
   	end With
End Sub

Sub Layout_change ()
	Dim intCnt
	With frmThis
	For intCnt = 1 To .sprSht.MaxRows 
'		If mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",intCnt) = "Y" Then
'		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
'		End If
	Next 
	End With
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strTIMCODE, strTIMNAME, strMEDCODE, strMEDNAME, strSUBSEQ, strSUBSEQNAME, strMEDFLAG, strGFLAG, strVOCH_TYPE)
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
		.cmbMED_FLAG1.value		= strMEDFLAG
		.cmbGFLAG1.value		= strGFLAG
		.cmbVOCH_TYPE1.value	= strVOCH_TYPE
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
	Dim lngCol, lngRow , i
	With frmThis
   		'������ Validation
		'If DataValidation =False Then exit Sub
		'On error resume Next
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "PUB_DATE | DEMANDDAY | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | DEPT_CD | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME",lngCol, lngRow, False) 

		If strDataCHK = False Then
			for i = 1 to .sprSht.MaxRows
				gErrorMsgBox lngRow & " ���� ������/û����/������/��ü/��ü��/�귣��/��/����/���ۻ�/�μ��� �ʼ� �Է»����Դϴ�.","����ȳ�"
				Exit Sub	
			next
		End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | GFLAGNAME | YEARMON | SEQ | MED_FLAG | DIVMEDIA | PUB_DATE | DEMANDDAY | CLIENTCODE |  CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | DEPT_CD | DEPT_NAME | PUB_FACE | EXECUTE_FACE | STD_STEP | STD_CM | STD_FACE | STD | STD_PAGE | COL_DEG | PROJECTION | PRICE | AMT | COMMI_RATE | COMMISSION | VOCH_TYPE | RECEIPT_GUBUN | TRU_TAX_FLAG | DUTYFLAG | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | GFLAG | EXCLIENTCODE | MATTERUSER")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		intRtn = mobjBOOK.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
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
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		If not gDataValidation(frmThis) Then exit Function
   		
   		'If Clientcode_FieldCheck =False Then exit Function
   		'If REAL_MED_CODE_FieldCheck =False Then exit Function
   		'If MEDCODE_FieldCheck =False Then exit Function
   	End With
	DataValidation = True
End Function

'****************************************************************************************
' �������ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function Clientcode_FieldCheck ()
	Clientcode_FieldCheck = False
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjBOOK.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value, "CUST")
		
		If mlngRowCnt =0 Then
			gErrorMsgBox "�������ڵ带 Ȯ�� �Ͻÿ�",""
			.txtCLIENTCODE.focus
			exit Function
   		End If
   	End With
   	Clientcode_FieldCheck = True
End Function
'****************************************************************************************
' ��ü���ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function REAL_MED_CODE_FieldCheck ()
	REAL_MED_CODE_FieldCheck = False
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjBOOK.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtREAL_MED_CODE.value, "REAL")
		
		If mlngRowCnt =0 Then
			gErrorMsgBox "��ü���ڵ带 Ȯ���Ͻÿ�",""
			.txtREAL_MED_CODE.focus
			exit Function
   		End If
   	End With
   	REAL_MED_CODE_FieldCheck = True
End Function
'****************************************************************************************
' ��ü���ڵ��� ���翩�� Ȯ��
'****************************************************************************************
Function MEDCODE_FieldCheck ()
	MEDCODE_FieldCheck = False
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
  	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjBOOK.DataValidation(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtMEDCODE.value, "MED")
		
		If mlngRowCnt =0 Then
			gErrorMsgBox "��ü�ڵ带 Ȯ���Ͻÿ�",""
			.txtMEDCODE.focus
			exit Function
   		End If
   	End With
   	MEDCODE_FieldCheck = True
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
					If mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",i) = "B" Then
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
					intRtn = mobjBOOK.DeleteRtn(gstrConfigXml,dblSEQ, strYEARMON)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
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
								<TD align="left" width="400" height="20">
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
											<td class="TITLE">�μ� û�����</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE class="SEARCHDATA" id="tblKey" height="48" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, txtSEQ)"
									width="50">�� ��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 78px; HEIGHT: 22px" accessKey="NUM"
										type="text" maxLength="6" size="7" name="txtYEARMON1"><INPUT dataFld="SEQ" class="NOINPUT_L" id="txtSEQ" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtSEQ" readOnly></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">������</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
									<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
									width="50">��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 123px; HEIGHT: 22px" type="text"
										maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
										maxLength="6" size="6" name="txtTIMCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1, txtSUBSEQ1)"
									width="50">�귣��</TD>
								<td class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="�귣���" style="WIDTH: 140px; HEIGHT: 22px"
										type="text" maxLength="100" size="18" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgSUBSEQ1"> <INPUT class="INPUT_L" id="txtSUBSEQ1" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="8" name="txtSUBSEQ1" size="3">
								</td>
							</TR>
							<TR>
								<TD class="SEARCHDATA" colSpan="2"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
									width="50">��ü��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="��ü���" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME1, txtMEDCODE1)"
									width="50">��ü��</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtMEDNAME1" title="��ü��" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" name="txtMEDNAME1"> <IMG id="ImgMEDCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgMEDCODE1"> <INPUT class="INPUT_L" id="txtMEDCODE1" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" size="2" name="txtMEDCODE1"></TD>
								<td class="SEARCHDATA" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
										align="right" border="0" name="imgQuery"><SELECT id="cmbMED_FLAG1" title="��������" style="WIDTH: 65px" name="cmbMED_FLAG1">
										<OPTION value="" selected>��ü</OPTION>
										<OPTION value="MP01">�Ź�</OPTION>
										<OPTION value="MP02">����</OPTION>
									</SELECT><SELECT id="cmbGFLAG1" title="��������" style="WIDTH: 65px" name="cmbGFLAG1">
										<OPTION value="" selected>��ü</OPTION>
										<OPTION value="M">����</OPTION>
										<OPTION value="B">����</OPTION>
										<OPTION value="J">����</OPTION>
										<OPTION value="S">����</OPTION>
									</SELECT><SELECT id="cmbVOCH_TYPE1" title="����" style="WIDTH: 65px" name="cmbVOCH_TYPE1">
										<OPTION value="" selected>��ü</OPTION>
										<OPTION value="0">����Ź</OPTION>
										<OPTION value="1">����</OPTION>
										<OPTION value="2">�Ϲ�</OPTION>
										<OPTION value="PROJECTION">����</OPTION>
									</SELECT>
								</td>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 20px"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="500" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" vAlign="absmiddle"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif'
														align='absMiddle' border='0' name='imgTableUp'></span> &nbsp;&nbsp;&nbsp;&nbsp;�հ� 
												: <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
												<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="top" align="right" height="28">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgCho.gIF"
													border="0" name="imgCho"></TD>
											<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgNew.gIF"
													border="0" name="imgREG"></TD>
											<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													alt="�ڷḦ �μ��մϴ�." src="../../../images/imglistcopy.gIF" border="0" name="Imgcopy"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgSave.gIF"
													border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													alt="�ڷḦ �μ��մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
								<TD style="WIDTH: 100%; HEIGHT: 120px" vAlign="top" align="center">
									<TABLE class="DATA" id="tblHidden" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" width="50">����</TD>
											<TD class="DATA" width="200"><SELECT dataFld="MED_FLAG" id="cmbMED_FLAG" title="��ü����" style="WIDTH: 85px" dataSrc="#xmlBind"
													name="cmbMED_FLAG">
													<OPTION value="MP01" selected>�Ź�</OPTION>
													<OPTION value="MP02">����</OPTION>
												</SELECT><SELECT dataFld="DIVMEDIA" id="cmbDIVMEDIA" title="�����" style="WIDTH: 111px" dataSrc="#xmlBind"
													name="cmbDIVMEDIA"></SELECT><INPUT dataFld="YEARMON" id="txtYEARMON" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtYEARMON"></TD>
											<TD class="LABEL" width="50">������</TD>
											<TD class="DATA" width="200"><INPUT dataFld="PUB_DATE" class="INPUT" id="txtPUB_DATE" title="������" style="WIDTH: 123px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="16" name="txtPUB_DATE">&nbsp;<IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar1"><INPUT dataFld="EXCLIENTCODE" id="txtEXCLIENTCODE" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtEXCLIENTCODE"><INPUT dataFld="EXCLIENTNAME" id="txtEXCLIENTNAME" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtEXCLIENTNAME">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtPUB_FACE, '')"
												width="50">û���</TD>
											<TD class="DATA" width="200"><INPUT dataFld="PUB_FACE" class="INPUT_R" id="txtPUB_FACE" title="û���" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="50" name="txtPUB_FACE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtPRICE, '')"
												width="50">�ܰ�</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="PRICE" class="INPUT_R" id="txtPRICE" title="�ܰ�" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="9" size="9" name="txtPRICE">
														</TD>
														<td align="right"><SELECT dataFld="VOCH_TYPE" id="cmbVOCH_TYPE" style="WIDTH: 85px" dataSrc="#xmlBind" name="cmbVOCH_TYPE">
																<OPTION value="0" selected>����Ź</OPTION>
																<OPTION value="1">����</OPTION>
																<OPTION value="2">�Ϲ�</OPTION>
																<OPTION value="3">AOR</OPTION>
															</SELECT>
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMATTERNAME, txtMATTERCODE)">�����</TD>
											<TD class="DATA"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="�귣���" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtMATTERNAME"> <IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMATTERCODE">
												<INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" name="txtMATTERCODE"></TD>
											<TD class="LABEL">û����</TD>
											<TD class="DATA"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="û����" style="WIDTH: 123px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="16" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalEndar2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar2"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEXECUTE_FACE, '')">�����</TD>
											<TD class="DATA"><INPUT dataFld="EXECUTE_FACE" class="INPUT_R" id="txtEXECUTE_FACE" title="�����" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="18" name="txtEXECUTE_FACE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtAMT, '')">�ݾ�</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="�ݾ�" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="9" name="txtAMT">
														</TD>
														<td class="DATA_RIGHT" align="right">����<INPUT id="chkRECEIPT_GUBUN" title="����" type="checkbox" name="chkRECEIPT_GUBUN">
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtSUBSEQNAME, txtSUBSEQ)">�귣��</TD>
											<TD class="DATA"><INPUT dataFld="SUBSEQNAME" class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtSUBSEQNAME"> <IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgSUBSEQCODE">
												<INPUT dataFld="SUBSEQ" class="INPUT_L" id="txtSUBSEQ" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" name="txtSUBSEQ"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEDNAME, txtMEDCODE)">��ü��</TD>
											<TD class="DATA"><INPUT dataFld="MEDNAME" class="INPUT_L" id="txtMEDNAME" title="��ü��" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="13" name="txtMEDNAME"> <IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMEDCODE">
												<INPUT dataFld="MEDCODE" class="INPUT_L" id="txtMEDCODE" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="2" name="txtMEDCODE"></TD>
											<TD class="LABEL" id="SizeOrSdt"></TD>
											<TD class="DATA">
												<DIV id="pnlSIZE" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><INPUT dataFld="STD_STEP" class="INPUT_R" id="txtSTD_STEP" title="��" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" size="1" name="txtSTD_STEP">��<INPUT dataFld="STD_CM" class="INPUT_R" id="txtSTD_CM" title="CM" style="WIDTH: 42px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="5" size="1" name="txtSTD_CM">cm&nbsp;
													<INPUT dataFld="STD_FACE" class="INPUT_R" id="txtSTD_FACE" title="��" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" size="1" name="txtSTD_FACE"></DIV>
												<DIV id="pnlSTD" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><INPUT dataFld="STD" class="INPUT_R" id="txtSTD" title="�԰�" style="WIDTH: 83px; HEIGHT: 22px"
														accessKey="" dataSrc="#xmlBind" type="text" maxLength="10" name="txtSTD">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT dataFld="STD_PAGE" class="INPUT_R" id="txtSTD_PAGE" title="������" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" name="txtSTD_PAGE">
													P</DIV>
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMI_RATE, '')">��������</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD class="DATA" width="92"><INPUT dataFld="COMMI_RATE" class="INPUT_R" id="txtCOMMI_RATE" title="��������" style="WIDTH: 64px; HEIGHT: 22px"
																dataSrc="#xmlBind" type="text" maxLength="6" size="5" name="txtCOMMI_RATE">%
														</TD>
														<td class="DATA_RIGHT" align="right">VAT<INPUT id="chkTRU_TAX_FLAG" title="VAT����" type="checkbox" CHECKED name="chkTRU_TAX_FLAG">
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTIMNAME, txtTIMCODE)">��</TD>
											<TD class="DATA"><INPUT dataFld="TIMNAME" class="INPUT_L" id="txtTIMNAME" title="����" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="20" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgTIMCODE">
												<INPUT dataFld="TIMCODE" class="INPUT_L" id="txtTIMCODE" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" size="6" name="txtTIMCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)">��ü��</TD>
											<TD class="DATA"><INPUT dataFld="REAL_MED_NAME" class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="7" name="txtREAL_MED_NAME">
												<IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE">
												<INPUT dataFld="REAL_MED_CODE" class="INPUT_L" id="txtREAL_MED_CODE" title="��ü���ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" name="txtREAL_MED_CODE">
											</TD>
											<TD class="LABEL">����</TD>
											<TD class="DATA"><SELECT dataFld="COL_DEG" id="cmbCOL_DEG" title="����" style="WIDTH: 84px" dataSrc="#xmlBind"
													name="cmbCOL_DEG">
													<OPTION value="B/W">B/W</OPTION>
													<OPTION value="C/L" selected>C/L</OPTION>
												</SELECT>&nbsp;<INPUT id="chkPROJECTION" title="����" type="checkbox" name="chkPROJECTION">����</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMISSION, '')">������</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="COMMISSION" class="INPUT_R" id="txtCOMMISSION" title="������" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="12" name="txtCOMMISSION">
														</TD>
														<td align="right"><SELECT dataFld="DUTYFLAG" id="cmbDUTYFLAG" style="WIDTH: 85px" dataSrc="#xmlBind" name="cmbDUTYFLAG">
																<OPTION value="Y" selected>����</OPTION>
																<OPTION value="N">�鼼</OPTION>
															</SELECT>
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME, txtCLIENTCODE)">û����</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPT_NAME, txtDEPT_CD)">���μ�</TD>
											<TD class="DATA"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPT_NAME" title="���μ���" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="6" name="txtDEPT_NAME">
												<IMG id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgDEPT_CD">
												<INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPT_CD" title="���μ��ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtDEPT_CD"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEMO, '')">���</TD>
											<TD class="DATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="���" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="120" size="12" name="txtMEMO"></TD>
											<TD class="LABEL">����</TD>
											<TD class="DATA"><INPUT id="chkGFLAG1" disabled type="radio" value="chkGFLAG1" name="chkGFLAG">����
												<INPUT id="chkGFLAG2" disabled type="radio" value="chkGFLAG2" name="chkGFLAG">����
												<INPUT id="chkGFLAG3" disabled type="radio" value="chkGFLAG3" name="chkGFLAG">����
												<INPUT id="chkGFLAG4" disabled type="radio" value="chkGFLAG4" name="chkGFLAG">����</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--BodySplit End-->
						</TABLE>
						<TABLE id="tblSheet" height="65%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td class="DATA" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
