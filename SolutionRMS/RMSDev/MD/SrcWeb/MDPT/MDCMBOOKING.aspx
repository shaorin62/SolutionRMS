<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMBOOKING.aspx.vb" Inherits="MD.MDCMBOOKING" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>개별청약 등록/조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : MD/부킹 화면(MDCMBOOKING)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMBOOKING.aspx
'기      능 : 인쇄매체 Booking Process 처리
'파라  메터 : 
'특이  사항 : 복사처리(다중선택 Row Coyp)
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
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjBOOK, mobjMDCOGET 
Dim mstrCheck
Dim mstrPub
Dim mcomecalender, mcomecalender2
Dim mstrPROCESS	'신규이면 True 조회면 False
Dim mstrPROCESS2 '조회상태이면 True 신규상12태이면 False
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
' 이벤트 프로시져 
'=========================================================================================
'입력 필드 숨기기
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			'document.getElementById("SizeOrSdt").innerHTML="사이즈"
			document.getElementById("tblBody").style.display = "inline"
			document.getElementById("tblSheet").style.height = "65%"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
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
' 명령 버튼 클릭 이벤트
'-----------------------------------
'조회버튼
Sub imgQuery_onclick
	If frmThis.txtYEARMON1.value = "" Then
		gErrorMsgBox "조회년월을 입력하시오","조회안내"
		exit Sub
	End If
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

'신규버튼
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
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
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
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
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
' 내역복사한다.
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
			gErrorMsgBox "내역복사는 한건만 가능합니다.",""
			Exit Sub
		elseif strCNT2 =0 Then
			gErrorMsgBox "내역복사할 로우를 선택하시오.",""
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
		mobjSCGLSpr.SetTextBinding .sprSht,"GFLAGNAME",.sprSht.ActiveRow, "미정"
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
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	On error resume next
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
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

'매체사 팝업 버튼
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE1.value), trim(.txtREAL_MED_NAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
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

'팁 팝업 버튼
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE1.value = vntRet(0,0) and .txtTIMNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTIMCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtTIMNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtCLIENTCODE1.value = trim(vntRet(4,0))       ' 코드명 표시
			.txtCLIENTNAME1.value = trim(vntRet(5,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
											trim(.txtTIMCODE1.value),trim(.txtTIMNAME1.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME1.value = trim(vntData(1,1))       ' 코드명 표시
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

'매체 팝업 버튼
Sub ImgMEDCODE1_onclick
	Call MEDCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub MEDCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtMEDCODE1.value), trim(.txtMEDNAME1.value), "MED_PRINT")
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE1.value = vntRet(0,0) and .txtMEDNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMEDCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtMEDNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtREAL_MED_CODE1.value = trim(vntRet(3,0))       ' 코드명 표시
			.txtREAL_MED_NAME1.value = trim(vntRet(4,0))       ' 코드명 표시
			SELECTRTN
			
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), _
												trim(.txtMEDCODE1.value),trim(.txtMEDNAME1.value), "MED_PRINT")
			
			If not gDoErrorRtn ("GetMEDGUBNCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE1.value = trim(vntData(0,1))	    ' Code값 저장
					.txtMEDNAME1.value = trim(vntData(1,1))       ' 코드명 표시
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

'브랜드
Sub ImgSUBSEQ1_onclick
	Call SUBSEQCODE1_POP()
End Sub

Sub SUBSEQCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ1.value), trim(.txtSUBSEQNAME1.value), trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtSUBSEQ1.value = trim(vntRet(0,0))		' 브랜드 표시
			.txtSUBSEQNAME1.value = trim(vntRet(1,0))	' 브랜드명 표시
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' 광고주 표시
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' 광고주명 표시
			.txtTIMCODE1.value = trim(vntRet(4,0))	' 광고주명 표시
			.txtTIMNAME1.value = trim(vntRet(5,0))	' 광고주명 표시
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
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ1.value),trim(.txtSUBSEQNAME1.value),  _
												trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ1.value = trim(vntData(0,1))
					.txtSUBSEQNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(2,1))		' 광고주 표시
					.txtCLIENTNAME1.value = trim(vntData(3,1))	' 광고주
					.txtTIMCODE1.value = trim(vntData(4,1))	' 광고주
					.txtTIMNAME1.value = trim(vntData(5,1))	' 광고주
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
' 팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
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

'매체사 팝업 버튼
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' 코드명 표시
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
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

'팁 팝업 버튼
Sub ImgTIMCODE_onclick
	Call TIMCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value), _
							trim(.txtTIMCODE.value), trim(.txtTIMNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMTIMPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtTIMCODE.value = vntRet(0,0) and .txtTIMNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtTIMNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtCLIENTCODE.value = trim(vntRet(4,0))       ' 코드명 표시
			.txtCLIENTNAME.value = trim(vntRet(5,0))       ' 코드명 표시
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

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), _
											trim(.txtTIMCODE.value),trim(.txtTIMNAME.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME.value = trim(vntData(1,1))       ' 코드명 표시
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

'매체 팝업 버튼
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value), _
							trim(.txtMEDCODE.value), trim(.txtMEDNAME.value), "MED_PRINT")
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDGBNPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMEDCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtMEDNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			.txtREAL_MED_CODE.value = trim(vntRet(3,0))       ' 코드명 표시
			.txtREAL_MED_NAME.value = trim(vntRet(4,0))       ' 코드명 표시
			
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

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMEDGUBNCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), _
											trim(.txtMEDCODE.value),trim(.txtMEDNAME.value), "MED_PRINT")
			
			If not gDoErrorRtn ("GetMEDGUBNCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtMEDNAME.value = trim(vntData(1,1))       ' 코드명 표시
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

'브랜드
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	' 브랜드명 표시
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주 표시
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
			.txtTIMCODE.value = trim(vntRet(4,0))	' 광고주명 표시
			.txtTIMNAME.value = trim(vntRet(5,0))	' 광고주명 표시
			.txtDEPT_CD.value = trim(vntRet(8,0))	' 광고주명 표시
			.txtDEPT_NAME.value = trim(vntRet(9,0))	' 광고주명 표시
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
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
												trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	' 광고주코드
					.txtCLIENTNAME.value = trim(vntData(3,1))	' 광고주
					.txtTIMCODE.value = trim(vntData(4,1))		' 팀코드
					.txtTIMNAME.value = trim(vntData(5,1))		' 팀명
					.txtDEPT_CD.value = trim(vntData(8,1))		' 부서코드
					.txtDEPT_NAME.value = trim(vntData(9,1))	' 부서명
					
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

'소재명 버튼 팝업
Sub ImgMATTERCODE_onclick
	Call MATTERCODE_POP()
End Sub

Sub MATTERCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTNAME.value), trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
							trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "B", TRIM(.txtMATTERCODE.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
		If isArray(vntRet) Then
			If .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtMATTERCODE.value = trim(vntRet(0,0))	' 소재코드 표시
			.txtMATTERNAME.value = trim(vntRet(1,0))	' 소재명 표시
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주코드 표시
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
			.txtTIMCODE.value = trim(vntRet(4,0))		' 팀코드 표시
			.txtTIMNAME.value = trim(vntRet(5,0))		' 팀명 표시
			.txtSUBSEQ.value = trim(vntRet(6,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(7,0))	' 브랜드명 표시
			.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' 제작사코드 표시
			.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' 제작사명 표시
			.txtDEPT_CD.value = trim(vntRet(10,0))		' 부서코드 표시
			.txtDEPT_NAME.value = trim(vntRet(11,0))	' 부서명 표시
			
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
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME.value),trim(.txtTIMNAME.value), trim(.txtSUBSEQNAME.value),"", _
											trim(.txtMATTERNAME.value), trim(.txtDEPT_NAME.value), "B")
			If not gDoErrorRtn ("GetMATTER") Then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntData(0,1))	' 소재코드 표시
					.txtMATTERNAME.value = trim(vntData(1,1))	' 소재명 표시
					.txtCLIENTCODE.value = trim(vntData(2,1))	' 광고주코드 표시
					.txtCLIENTNAME.value = trim(vntData(3,1))	' 광고주명 표시
					.txtTIMCODE.value	 = trim(vntData(4,1))	' 팀코드 표시
					.txtTIMNAME.value	 = trim(vntData(5,1))	' 팀명 표시
					.txtSUBSEQ.value	 = trim(vntData(6,1))	' 브랜드 표시
					.txtSUBSEQNAME.value = trim(vntData(7,1))	' 브랜드명 표시
					.txtEXCLIENTCODE.value = trim(vntData(8,1))	' 제작사코드 표시
					.txtEXCLIENTNAME.value = trim(vntData(9,1))	' 제작사명 표시
					.txtDEPT_CD.value	 = trim(vntData(10,1))	' 부서코드 표시
					.txtDEPT_NAME.value	 = trim(vntData(11,1))	' 부서명 표시
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

'담당부서 팝업 
Sub imgDEPT_CD_onclick
	Call DEPT_CD_POP()
End Sub

Sub DEPT_CD_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtDEPT_NAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtDEPT_CD.value = trim(vntRet(0,0))	'Code값 저장
			.txtDEPT_NAME.value = trim(vntRet(1,0))	'코드명 표시
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
			'Long Type의 ByRef 변수의 초기화
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
' 게재일 달력
'****************************************************************************************
Sub imgCalEndar1_onclick
	'CalEndar를 화면에 표시
	mcomecalender = true
	gShowPopupCalEndar frmThis.txtPUB_DATE,frmThis.imgCalEndar1,"txtPUB_DATE_onchange()"
	Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PUB_DATE"), frmThis.sprSht.ActiveRow)
	mcomecalender = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalEndar2_onclick
	'CalEndar를 화면에 표시
	mcomecalender2 = true
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalEndar2,"txtDEMANDDAY_onchange()"
	mcomecalender2 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'****************************************************************************************
' 입력필드 키다운 이벤트
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
' 입력필드 체인지 이벤트
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
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
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
					gErrorMsgBox "게재일의 년월은 수정할 수 없습니다.",""
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
			document.getElementById("SizeOrSdt").innerHTML="사이즈"
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
			document.getElementById("SizeOrSdt").innerHTML="규격"
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


'영세/면세 구분 세팅(부가세가 무일때 선택할 수 있다.)
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
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'단가
Sub txtPRICE_onblur
	With frmThis
		Call gFormatNumber(.txtPRICE,0,True)
		priceCal
	end With
End Sub

'금액
Sub txtAMT_onblur
	With frmThis
		COMMI_RATE_Cal
		Call gFormatNumber(.txtAMT,0,True)
	end With
End Sub

'수수료율
Sub txtCOMMI_RATE_onblur
	With frmThis
		COMMI_RATE_Cal
	end With
End Sub

'수수료
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
' 천단위 나눔점 없애기 ( 단가, 금액, 수수료)
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
' 수수료 계산
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

'수수료율에서 엔터시 수수료 자동계산
Sub txtCOMMI_RATE_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtCOMMISSION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'금액에서 엔터시 수수료 자동계산
Sub txtAMT_onkeydown
	If window.event.keyCode = meEnter OR window.event.keyCode = meTAB Then
		COMMI_RATE_Cal
		frmThis.txtCOMMI_RATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
'--------------------------------------------------
'쉬트 키다운
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
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GFLAGNAME",frmThis.sprSht.ActiveRow, "미정"
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
					gErrorMsgBox "게재일의 년월은 수정할 수 없습니다.",""
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
			'명이 수정되면 코드를 지운다.
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
						
						.txtREAL_MED_CODE.value = trim(vntData(0,1))	    ' Code값 저장
						.txtREAL_MED_NAME.value = trim(vntData(1,1))       ' 코드명 표시
						
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
						
						.txtTIMCODE.value = trim(vntData(0,1))	    ' Code값 저장
						.txtTIMNAME.value = trim(vntData(1,1))       ' 코드명 표시
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
						
						
						.txtMATTERCODE.value = trim(vntData(0,1))	' 소재코드 표시
						.txtMATTERNAME.value = trim(vntData(1,1))	' 소재명 표시
						.txtCLIENTCODE.value = trim(vntData(2,1))	' 광고주코드 표시
						.txtCLIENTNAME.value = trim(vntData(3,1))	' 광고주명 표시
						.txtTIMCODE.value	 = trim(vntData(4,1))	' 팀코드 표시
						.txtTIMNAME.value	 = trim(vntData(5,1))	' 팀명 표시
						.txtSUBSEQ.value	 = trim(vntData(6,1))	' 브랜드 표시
						.txtSUBSEQNAME.value = trim(vntData(7,1))	' 브랜드명 표시
						.txtEXCLIENTCODE.value = trim(vntData(8,1))	' 제작사코드 표시
						.txtEXCLIENTNAME.value = trim(vntData(9,1))	' 제작사코드 표시
						.txtDEPT_CD.value	 = trim(vntData(10,1))	' 부서코드 표시
						.txtDEPT_NAME.value	 = trim(vntData(11,1))	' 부서명 표시
						
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
	'변경 플래그 설정
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
				
				.txtSUBSEQ.value = trim(vntRet(0,0))		' 브랜드 표시
				.txtSUBSEQNAME.value = trim(vntRet(1,0))	' 브랜드명 표시
				.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주 표시
				.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
				.txtTIMCODE.value = trim(vntRet(4,0))	' 광고주명 표시
				.txtTIMNAME.value = trim(vntRet(5,0))	' 광고주명 표시
				.txtDEPT_CD.value = trim(vntRet(8,0))	' 광고주명 표시
				.txtDEPT_NAME.value = trim(vntRet(9,0))	' 광고주명 표시
				
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
				
				.txtTIMCODE.value = trim(vntRet(0,0))	    ' Code값 저장
				.txtTIMNAME.value = trim(vntRet(1,0))       ' 코드명 표시
				.txtCLIENTCODE.value = trim(vntRet(4,0))    ' 코드명 표시
				.txtCLIENTNAME.value = trim(vntRet(5,0))    ' 코드명 표시
				
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
				
				.txtMATTERCODE.value = trim(vntRet(0,0))	' 소재코드 표시
				.txtMATTERNAME.value = trim(vntRet(1,0))	' 소재명 표시
				.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주코드 표시
				.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
				.txtTIMCODE.value = trim(vntRet(4,0))		' 팀코드 표시
				.txtTIMNAME.value = trim(vntRet(5,0))		' 팀명 표시
				.txtSUBSEQ.value = trim(vntRet(6,0))		' 브랜드 표시
				.txtSUBSEQNAME.value = trim(vntRet(7,0))	' 브랜드명 표시
				.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' 제작사코드 표시
				.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' 제작사코드 표시
				.txtDEPT_CD.value = trim(vntRet(10,0))		' 부서코드 표시
				.txtDEPT_NAME.value = trim(vntRet(11,0))	' 부서명 표시
				
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
				
				.txtDEPT_CD.value = trim(vntRet(0,0))	'Code값 저장
				.txtDEPT_NAME.value = trim(vntRet(1,0))	'코드명 표시
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		sprShtToFieldBinding Col, Row
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
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


'시트에 데이터한로우의 정보를 헤더 필더에 바인딩
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '그리드 데이터가 없으면 나간다.
		
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
			document.getElementById("SizeOrSdt").innerHTML="사이즈"
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
			document.getElementById("SizeOrSdt").innerHTML="규격"
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
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성	
	set mobjBOOK		= gCreateRemoteObject("cMDPT.ccMDPTBOOKING")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 57, 0, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | GFLAGNAME | YEARMON | SEQ | MED_FLAG | DIVMEDIA | PUB_DATE | DEMANDDAY | CLIENTCODE |  CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BUSINO | SUBSEQ | SUBSEQNAME | MATTERCODE | MATTERNAME | AMT | COMMI_RATE | EXECUTE_FACE | STD_STEP | STD_CM | MEMO | TIMCODE | TIMNAME | DEPT_CD | DEPT_NAME | PUB_FACE | STD_FACE | STD | STD_PAGE | COL_DEG | PROJECTION | PRICE | COMMISSION | VOCH_TYPE | RECEIPT_GUBUN | TRU_TAX_FLAG | DUTYFLAG | TRU_TRANS_NO | COMMI_TRANS_NO | GFLAG | EXCLIENTCODE | EXCLIENTNAME | REAL_MED_BUSINO1 | REAL_MED_NAME1 | MEDNAME1 | TIMNAME1 | SUBSEQNAME1 | MATTERNAME1 |DEPT_NAME1 | EXCLIENTNAME1 | AMT1 | COMMISSION1 | MATTERUSER"
											  '  1|          2|        3|    4|	        5|	       6|         7|          8|           9|           10|       11|       12|             13|             14|               15|	   16|          17|          18|          19|   20|          21|            22|       23|       24|       25|         26|        27|        28|      29|        30|   31|        32|       33|          34|     35|          36|         37|             38|            39|        40|    41|            42|              43|     44|            45|            46
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|G|년월|순번|매체구분|구분|게재일|청구일|광고주코드|광고주명|매체코드|매체명|매체사코드|매체사명|매체사사업자번호|브랜드코드|브랜드명|소재코드|소재명|금액|수수료율|집행면|단|Cm|비고|팀코드|팀명|부서코드|부서명|청약면|면|규격|Page|색도|돌출|단가|수수료|전표구분|접수|VAT|면세구분|위수탁거래번호|수수료거래번호|GFLAG|제작대행사코드|제작대행사명|사업자번호|거래처명|매체명|Client부서명|브랜드명|소재명|실적부서|Cre조직|매체비|대행수수료|소재등록자"
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
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjBOOK = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	With frmThis
		.sprSht.MaxRows = 0
		
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		.txtYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	'청년월
		.txtPUB_DATE.value = gNowDate2
		
		'청구일세팅 게재월의 마지막일
		DateClean .txtYEARMON.value
		
		'인쇄종류 세팅
		Call SUBCOMBO_TYPE()
		'기본값 세팅
		.txtSTD_STEP.value = "15"
		.txtSTD_CM.value = "37.0"
		.txtSTD_FACE.value = "1"
		.txtCOMMI_RATE.value = "15"
		.chkPROJECTION.checked = False
		.chkRECEIPT_GUBUN.checked = False
		.chkTRU_TAX_FLAG.checked = True
		
		'사이즈/규격입력필드 세팅
		document.getElementById("SizeOrSdt").innerHTML="사이즈"
		pnlSIZE.style.display = "inline"
		pnlSTD.style.display = "none"
		
		mobjSCGLSpr.ColHidden .sprSht, "STD_STEP", False
		mobjSCGLSpr.ColHidden .sprSht, "STD_CM", False
		mobjSCGLSpr.ColHidden .sprSht, "STD_FACE", False
		mobjSCGLSpr.ColHidden .sprSht, "STD", True
		mobjSCGLSpr.ColHidden .sprSht, "STD_PAGE", True
				
		'Sheet초기화
		.txtYEARMON1.focus
		
		Field_Lock
		DutyFlag_Disable
		Get_COMBO_VALUE
	End With
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'청구일 조회조건 생성
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
' SUBCOMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()
	Dim vntPUB_FACE
	Dim strMED_FLAG
	Dim vntMED_FLAG_DIVMEDIA
	With frmThis   
		strMED_FLAG = "MP_" & .cmbMED_FLAG.value
		On error resume Next
		'Long Type의 ByRef 변수의 초기화
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
' 그리드 콤보박스 설정
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData, vntData_VOCH, vntData_DUTY
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
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
' 그리드 서브 콤보 설정
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strMED_FLAG)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type의 ByRef 변수의 초기화
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
		'Long Type의 ByRef 변수의 초기화
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
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
       	vntData = mobjBOOK.GetDataType_DIVMEDIA(gstrConfigXml, mlngRowCnt, mlngColCnt, strMED_FLAG)
		mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DIVMEDIA",Row,Row,vntData,,80 
		mobjSCGLSpr.TypeComboBox = True 
   		gSetChange
   	end With   
End Sub


'-----------------------------------------------------------------------------------------
' Field_Lock  거래명세서번호나 세금계산서 번호가 있으면 수정할수 없도록 필드를 ReadOnly처리
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",.sprSht.ActiveRow) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",.sprSht.ActiveRow) <> "" Then
				'구분
				.cmbMED_FLAG.disabled = True : .cmbDIVMEDIA.disabled = True
				'소재
				.txtMATTERNAME.className	= "NOINPUT_L" : .txtMATTERNAME.readOnly		= True : .ImgMATTERCODE.disabled = True
				.txtMATTERCODE.className	= "NOINPUT_L" : .txtMATTERCODE.readOnly		= True
				'브랜드
				.txtSUBSEQNAME.className	= "NOINPUT_L" : .txtSUBSEQNAME.readOnly		= True : .ImgSUBSEQCODE.disabled = True
				.txtSUBSEQ.className		= "NOINPUT_L" : .txtSUBSEQ.readOnly			= True
				'팀
				.txtTIMNAME.className		= "NOINPUT_L" : .txtTIMNAME.readOnly		= True : .ImgTIMCODE.disabled	 = True
				.txtTIMCODE.className		= "NOINPUT_L" : .txtTIMCODE.readOnly		= True
				'청구지
				.txtCLIENTNAME.className	= "NOINPUT_L" : .txtCLIENTNAME.readOnly		= True : .ImgCLIENTCODE.disabled = True
				.txtCLIENTCODE.className	= "NOINPUT_L" : .txtCLIENTCODE.readOnly		= True
				'게재일/청구일
				.txtPUB_DATE.className		= "NOINPUT"   : .txtPUB_DATE.readOnly		= True : .imgCalEndar1.disabled  = True 
				.txtDEMANDDAY.className		= "NOINPUT"   : .txtDEMANDDAY.readOnly		= True : .imgCalEndar2.disabled  = True 
				'매체
				.txtMEDNAME.className		= "NOINPUT_L" : .txtMEDNAME.readOnly		= True : .ImgMEDCODE.disabled	 = True
				.txtMEDCODE.className		= "NOINPUT_L" : .txtMEDCODE.readOnly		= True
				'매체사
				.txtREAL_MED_NAME.className = "NOINPUT_L" : .txtREAL_MED_NAME.readOnly	= True : .ImgREAL_MED_CODE.disabled = True
				.txtREAL_MED_CODE.className = "NOINPUT_L" : .txtREAL_MED_CODE.readOnly	= True
				'담당부서
				.txtDEPT_NAME.className		= "NOINPUT_L" : .txtDEPT_NAME.readOnly		= True : .imgDEPT_CD.disabled	 = True
				.txtDEPT_CD.className		= "NOINPUT_L" : .txtDEPT_CD.readOnly		= True
				'청약면
				.txtPUB_FACE.className		= "NOINPUT_L" : .txtPUB_FACE.readOnly		= True
				'집행면
				.txtEXECUTE_FACE.className	= "NOINPUT_L" : .txtEXECUTE_FACE.readOnly	= True
				'사이즈/규격
				.txtSTD_STEP.className		= "NOINPUT_R" : .txtSTD_STEP.readOnly		= True
				.txtSTD_CM.className		= "NOINPUT_R" : .txtSTD_CM.readOnly			= True
				.txtSTD_FACE.className		= "NOINPUT_R" : .txtSTD_FACE.readOnly		= True
				.txtSTD.className			= "NOINPUT_R" : .txtSTD.readOnly			= True
				.txtSTD_PAGE.className		= "NOINPUT_R" : .txtSTD_PAGE.readOnly		= True
				'비고/단가/금액/수수료율/수수료
				.txtMEMO.className			= "NOINPUT_L" : .txtMEMO.readOnly			= True
				.txtPRICE.className			= "NOINPUT_R" : .txtPRICE.readOnly			= True 
				.txtAMT.className			= "NOINPUT_R" : .txtAMT.readOnly			= True
				.txtCOMMI_RATE.className	= "NOINPUT_R" : .txtCOMMI_RATE.readOnly		= True 
				.txtCOMMISSION.className	= "NOINPUT_R" : .txtCOMMISSION.readOnly		= True
				'색도/돌출/ 전표구분/접수/VAT유무/면세구분
				.cmbCOL_DEG.disabled		= True : .chkPROJECTION.disabled	= True
				.cmbVOCH_TYPE.disabled		= True : .chkRECEIPT_GUBUN.disabled = True
				.chkTRU_TAX_FLAG.disabled	= True : .cmbDUTYFLAG.disabled		= True
			else 
				'구분
				.cmbMED_FLAG.disabled = False : .cmbDIVMEDIA.disabled = False
				'소재
				.txtMATTERNAME.className	= "INPUT_L" : .txtMATTERNAME.readOnly	= False : .ImgMATTERCODE.disabled = False
				.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
				'브랜드
				.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
				.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
				'팀
				.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
				.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
				'청구지
				.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
				.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
				'게재일/청구일
				.txtPUB_DATE.className		= "INPUT"   : .txtPUB_DATE.readOnly		= False : .imgCalEndar1.disabled  = False 
				.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
				'매체
				.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
				.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
				'매체사
				.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
				.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
				'담당부서
				.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
				.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
				'청약면
				.txtPUB_FACE.className		= "INPUT_L" : .txtPUB_FACE.readOnly		= False
				'집행면
				.txtEXECUTE_FACE.className	= "INPUT_L" : .txtEXECUTE_FACE.readOnly	= False
				'사이즈/규격
				.txtSTD_STEP.className		= "INPUT_R" : .txtSTD_STEP.readOnly		= False
				.txtSTD_CM.className		= "INPUT_R" : .txtSTD_CM.readOnly		= False
				.txtSTD_FACE.className		= "INPUT_R" : .txtSTD_FACE.readOnly		= False
				.txtSTD.className			= "INPUT_R" : .txtSTD.readOnly			= False
				.txtSTD_PAGE.className		= "INPUT_R" : .txtSTD_PAGE.readOnly		= False
				'비고/단가/금액/수수료율/수수료
				.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
				.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly		= False 
				.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= False
				.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
				.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
				'색도/돌출/ 전표구분/접수/VAT유무/면세구분
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
			'구분
			.cmbMED_FLAG.disabled = False : .cmbDIVMEDIA.disabled = False
			'소재
			.txtMATTERNAME.className		= "INPUT_L" : .txtMATTERNAME.readOnly		= False : .ImgMATTERCODE.disabled = False
			.txtMATTERCODE.className	= "INPUT_L" : .txtMATTERCODE.readOnly	= False
			'브랜드
			.txtSUBSEQNAME.className	= "INPUT_L" : .txtSUBSEQNAME.readOnly	= False : .ImgSUBSEQCODE.disabled = False
			.txtSUBSEQ.className		= "INPUT_L" : .txtSUBSEQ.readOnly		= False
			'팀
			.txtTIMNAME.className		= "INPUT_L" : .txtTIMNAME.readOnly		= False : .ImgTIMCODE.disabled	  = False
			.txtTIMCODE.className		= "INPUT_L" : .txtTIMCODE.readOnly		= False
			'청구지
			.txtCLIENTNAME.className	= "INPUT_L" : .txtCLIENTNAME.readOnly	= False : .ImgCLIENTCODE.disabled = False
			.txtCLIENTCODE.className	= "INPUT_L" : .txtCLIENTCODE.readOnly	= False
			'게재일/청구일
			.txtPUB_DATE.className		= "INPUT"   : .txtPUB_DATE.readOnly		= False : .imgCalEndar1.disabled  = False 
			.txtDEMANDDAY.className		= "INPUT"   : .txtDEMANDDAY.readOnly	= False : .imgCalEndar2.disabled  = False 
			'매체
			.txtMEDNAME.className		= "INPUT_L" : .txtMEDNAME.readOnly		= False : .ImgMEDCODE.disabled	  = False
			.txtMEDCODE.className		= "INPUT_L" : .txtMEDCODE.readOnly		= False
			'매체사
			.txtREAL_MED_NAME.className = "INPUT_L" : .txtREAL_MED_NAME.readOnly= False : .ImgREAL_MED_CODE.disabled = False
			.txtREAL_MED_CODE.className = "INPUT_L" : .txtREAL_MED_CODE.readOnly= False
			'담당부서
			.txtDEPT_NAME.className		= "INPUT_L" : .txtDEPT_NAME.readOnly	= False : .imgDEPT_CD.disabled = False
			.txtDEPT_CD.className		= "INPUT_L" : .txtDEPT_CD.readOnly		= False
			'청약면
			.txtPUB_FACE.className		= "INPUT_L" : .txtPUB_FACE.readOnly		= False
			'집행면
			.txtEXECUTE_FACE.className	= "INPUT_L" : .txtEXECUTE_FACE.readOnly	= False
			'사이즈/규격
			.txtSTD_STEP.className		= "INPUT_R" : .txtSTD_STEP.readOnly		= False
			.txtSTD_CM.className		= "INPUT_R" : .txtSTD_CM.readOnly		= False
			.txtSTD_FACE.className		= "INPUT_R" : .txtSTD_FACE.readOnly		= False
			.txtSTD.className			= "INPUT_R" : .txtSTD.readOnly			= False
			.txtSTD_PAGE.className		= "INPUT_R" : .txtSTD_PAGE.readOnly		= False
			'비고/단가/금액/수수료율/수수료
			.txtMEMO.className			= "INPUT_L" : .txtMEMO.readOnly			= False
			.txtPRICE.className			= "INPUT_R" : .txtPRICE.readOnly		= False 
			.txtAMT.className			= "INPUT_R" : .txtAMT.readOnly			= False
			.txtCOMMI_RATE.className	= "INPUT_R" : .txtCOMMI_RATE.readOnly	= False 
			.txtCOMMISSION.className	= "INPUT_R" : .txtCOMMISSION.readOnly	= False
			'색도/돌출/ 전표구분/접수/VAT유무/면세구분
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
' 데이터 조회
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
		'Sheet초기화
		.sprSht.MaxRows = 0
		intCnt2 = 1
		
		'Long Type의 ByRef 변수의 초기화
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
				
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	   			For intCnt = 1 To .sprSht.MaxRows
	   			
	   				'for문 한번으로 최소화 하기위해 여기에 배치
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
   				'검색시에 첫행을 MASTER와 바인딩 시키기 위함
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
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
'이전 검색어를 담아 놓는다.
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
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
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
' 데이터 처리
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
   		'데이터 Validation
		'If DataValidation =False Then exit Sub
		'On error resume Next
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "PUB_DATE | DEMANDDAY | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | DEPT_CD | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME",lngCol, lngRow, False) 

		If strDataCHK = False Then
			for i = 1 to .sprSht.MaxRows
				gErrorMsgBox lngRow & " 줄의 게재일/청구일/광고주/매체/매체사/브랜드/팀/소재/제작사/부서는 필수 입력사항입니다.","저장안내"
				Exit Sub	
			next
		End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | GFLAGNAME | YEARMON | SEQ | MED_FLAG | DIVMEDIA | PUB_DATE | DEMANDDAY | CLIENTCODE |  CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | DEPT_CD | DEPT_NAME | PUB_FACE | EXECUTE_FACE | STD_STEP | STD_CM | STD_FACE | STD | STD_PAGE | COL_DEG | PROJECTION | PRICE | AMT | COMMI_RATE | COMMISSION | VOCH_TYPE | RECEIPT_GUBUN | TRU_TAX_FLAG | DUTYFLAG | MEMO | TRU_TRANS_NO | COMMI_TRANS_NO | GFLAG | EXCLIENTCODE | MATTERUSER")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtn = mobjBOOK.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
   		End If
   	end With
End Sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
	Dim vntData
   	Dim i, strCols
   	
	'On error resume Next
	With frmThis
		'Master 입력 데이터 Validation : 필수 입력항목 검사
   		If not gDataValidation(frmThis) Then exit Function
   		
   		'If Clientcode_FieldCheck =False Then exit Function
   		'If REAL_MED_CODE_FieldCheck =False Then exit Function
   		'If MEDCODE_FieldCheck =False Then exit Function
   	End With
	DataValidation = True
End Function

'****************************************************************************************
' 광고주코드의 존재여부 확인
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
			gErrorMsgBox "광고주코드를 확인 하시오",""
			.txtCLIENTCODE.focus
			exit Function
   		End If
   	End With
   	Clientcode_FieldCheck = True
End Function
'****************************************************************************************
' 매체사코드의 존재여부 확인
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
			gErrorMsgBox "매체사코드를 확인하시오",""
			.txtREAL_MED_CODE.focus
			exit Function
   		End If
   	End With
   	REAL_MED_CODE_FieldCheck = True
End Function
'****************************************************************************************
' 매체명코드의 존재여부 확인
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
			gErrorMsgBox "매체코드를 확인하시오",""
			.txtMEDCODE.focus
			exit Function
   		End If
   	End With
   	MEDCODE_FieldCheck = True
End Function

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '실제데이터여부 플레
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",i) <> "" Then
					gErrorMsgBox "선택하신 " & i & "행의 자료는 거래명세표가 존재 합니다." & vbcrlf & "먼저 거래명세표를 삭제 하십시오!","삭제안내!"
					exit Sub
				else 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"GFLAG",i) = "B" Then
						gErrorMsgBox "선택하신 " & i & "행의 자료는 승인된 자료입니다." & vbcrlf & "먼저 승인취소처리 하십시오!","삭제안내!"
						exit Sub
					End If
					lngchkCnt = lngchkCnt +1
				End If
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
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
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
	err.clear	
End Sub

'번호를 클리어한다.
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
											<td class="TITLE">인쇄 청약관리</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
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
									width="50">년 월</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 78px; HEIGHT: 22px" accessKey="NUM"
										type="text" maxLength="6" size="7" name="txtYEARMON1"><INPUT dataFld="SEQ" class="NOINPUT_L" id="txtSEQ" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
										dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtSEQ" readOnly></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">광고주</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
									<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
									width="50">팀</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtTIMNAME1" title="팀명" style="WIDTH: 123px; HEIGHT: 22px" type="text"
										maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" type="text"
										maxLength="6" size="6" name="txtTIMCODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1, txtSUBSEQ1)"
									width="50">브랜드</TD>
								<td class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="브랜드명" style="WIDTH: 140px; HEIGHT: 22px"
										type="text" maxLength="100" size="18" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgSUBSEQ1"> <INPUT class="INPUT_L" id="txtSUBSEQ1" title="시퀀스코드" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="8" name="txtSUBSEQ1" size="3">
								</td>
							</TR>
							<TR>
								<TD class="SEARCHDATA" colSpan="2"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
									width="50">매체사</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="매체사명" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" name="txtREAL_MED_CODE1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME1, txtMEDCODE1)"
									width="50">매체명</TD>
								<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtMEDNAME1" title="매체명" style="WIDTH: 123px; HEIGHT: 22px"
										type="text" maxLength="100" name="txtMEDNAME1"> <IMG id="ImgMEDCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="ImgMEDCODE1"> <INPUT class="INPUT_L" id="txtMEDCODE1" title="매체명코드" style="WIDTH: 53px; HEIGHT: 22px"
										type="text" maxLength="6" size="2" name="txtMEDCODE1"></TD>
								<td class="SEARCHDATA" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
										align="right" border="0" name="imgQuery"><SELECT id="cmbMED_FLAG1" title="제작종류" style="WIDTH: 65px" name="cmbMED_FLAG1">
										<OPTION value="" selected>전체</OPTION>
										<OPTION value="MP01">신문</OPTION>
										<OPTION value="MP02">잡지</OPTION>
									</SELECT><SELECT id="cmbGFLAG1" title="제작종류" style="WIDTH: 65px" name="cmbGFLAG1">
										<OPTION value="" selected>전체</OPTION>
										<OPTION value="M">미정</OPTION>
										<OPTION value="B">배정</OPTION>
										<OPTION value="J">집행</OPTION>
										<OPTION value="S">실적</OPTION>
									</SELECT><SELECT id="cmbVOCH_TYPE1" title="구분" style="WIDTH: 65px" name="cmbVOCH_TYPE1">
										<OPTION value="" selected>전체</OPTION>
										<OPTION value="0">위수탁</OPTION>
										<OPTION value="1">협찬</OPTION>
										<OPTION value="2">일반</OPTION>
										<OPTION value="PROJECTION">돌출</OPTION>
									</SELECT>
								</td>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 20px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="500" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" vAlign="absmiddle"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id='imgTableUp' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/imgTableUp.gif'
														align='absMiddle' border='0' name='imgTableUp'></span> &nbsp;&nbsp;&nbsp;&nbsp;합계 
												: <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
												<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
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
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="자료를 인쇄합니다." src="../../../images/imgCho.gIF"
													border="0" name="imgCho"></TD>
											<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="자료를 인쇄합니다." src="../../../images/imgNew.gIF"
													border="0" name="imgREG"></TD>
											<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													alt="자료를 인쇄합니다." src="../../../images/imglistcopy.gIF" border="0" name="Imgcopy"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" alt="자료를 인쇄합니다." src="../../../images/imgSave.gIF"
													border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													alt="자료를 인쇄합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
											<TD class="LABEL" width="50">구분</TD>
											<TD class="DATA" width="200"><SELECT dataFld="MED_FLAG" id="cmbMED_FLAG" title="매체구분" style="WIDTH: 85px" dataSrc="#xmlBind"
													name="cmbMED_FLAG">
													<OPTION value="MP01" selected>신문</OPTION>
													<OPTION value="MP02">잡지</OPTION>
												</SELECT><SELECT dataFld="DIVMEDIA" id="cmbDIVMEDIA" title="게재면" style="WIDTH: 111px" dataSrc="#xmlBind"
													name="cmbDIVMEDIA"></SELECT><INPUT dataFld="YEARMON" id="txtYEARMON" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtYEARMON"></TD>
											<TD class="LABEL" width="50">게재일</TD>
											<TD class="DATA" width="200"><INPUT dataFld="PUB_DATE" class="INPUT" id="txtPUB_DATE" title="게재일" style="WIDTH: 123px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="16" name="txtPUB_DATE">&nbsp;<IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar1"><INPUT dataFld="EXCLIENTCODE" id="txtEXCLIENTCODE" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtEXCLIENTCODE"><INPUT dataFld="EXCLIENTNAME" id="txtEXCLIENTNAME" style="WIDTH: 8px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" name="txtEXCLIENTNAME">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtPUB_FACE, '')"
												width="50">청약면</TD>
											<TD class="DATA" width="200"><INPUT dataFld="PUB_FACE" class="INPUT_R" id="txtPUB_FACE" title="청약면" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="50" name="txtPUB_FACE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtPRICE, '')"
												width="50">단가</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="PRICE" class="INPUT_R" id="txtPRICE" title="단가" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="9" size="9" name="txtPRICE">
														</TD>
														<td align="right"><SELECT dataFld="VOCH_TYPE" id="cmbVOCH_TYPE" style="WIDTH: 85px" dataSrc="#xmlBind" name="cmbVOCH_TYPE">
																<OPTION value="0" selected>위수탁</OPTION>
																<OPTION value="1">협찬</OPTION>
																<OPTION value="2">일반</OPTION>
																<OPTION value="3">AOR</OPTION>
															</SELECT>
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMATTERNAME, txtMATTERCODE)">소재명</TD>
											<TD class="DATA"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="브랜드명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtMATTERNAME"> <IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMATTERCODE">
												<INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="시퀀스코드" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" name="txtMATTERCODE"></TD>
											<TD class="LABEL">청구일</TD>
											<TD class="DATA"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="청구일" style="WIDTH: 123px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="16" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalEndar2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalEndar2"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEXECUTE_FACE, '')">집행면</TD>
											<TD class="DATA"><INPUT dataFld="EXECUTE_FACE" class="INPUT_R" id="txtEXECUTE_FACE" title="집행면" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="18" name="txtEXECUTE_FACE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtAMT, '')">금액</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="금액" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="9" name="txtAMT">
														</TD>
														<td class="DATA_RIGHT" align="right">접수<INPUT id="chkRECEIPT_GUBUN" title="돌출" type="checkbox" name="chkRECEIPT_GUBUN">
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtSUBSEQNAME, txtSUBSEQ)">브랜드</TD>
											<TD class="DATA"><INPUT dataFld="SUBSEQNAME" class="INPUT_L" id="txtSUBSEQNAME" title="브랜드명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtSUBSEQNAME"> <IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgSUBSEQCODE">
												<INPUT dataFld="SUBSEQ" class="INPUT_L" id="txtSUBSEQ" title="시퀀스코드" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" name="txtSUBSEQ"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEDNAME, txtMEDCODE)">매체명</TD>
											<TD class="DATA"><INPUT dataFld="MEDNAME" class="INPUT_L" id="txtMEDNAME" title="매체명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="13" name="txtMEDNAME"> <IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgMEDCODE">
												<INPUT dataFld="MEDCODE" class="INPUT_L" id="txtMEDCODE" title="매체명코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="2" name="txtMEDCODE"></TD>
											<TD class="LABEL" id="SizeOrSdt"></TD>
											<TD class="DATA">
												<DIV id="pnlSIZE" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><INPUT dataFld="STD_STEP" class="INPUT_R" id="txtSTD_STEP" title="단" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" size="1" name="txtSTD_STEP">단<INPUT dataFld="STD_CM" class="INPUT_R" id="txtSTD_CM" title="CM" style="WIDTH: 42px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="5" size="1" name="txtSTD_CM">cm&nbsp;
													<INPUT dataFld="STD_FACE" class="INPUT_R" id="txtSTD_FACE" title="단" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" size="1" name="txtSTD_FACE"></DIV>
												<DIV id="pnlSTD" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><INPUT dataFld="STD" class="INPUT_R" id="txtSTD" title="규격" style="WIDTH: 83px; HEIGHT: 22px"
														accessKey="" dataSrc="#xmlBind" type="text" maxLength="10" name="txtSTD">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT dataFld="STD_PAGE" class="INPUT_R" id="txtSTD_PAGE" title="페이지" style="WIDTH: 40px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="3" name="txtSTD_PAGE">
													P</DIV>
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMI_RATE, '')">수수료율</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD class="DATA" width="92"><INPUT dataFld="COMMI_RATE" class="INPUT_R" id="txtCOMMI_RATE" title="수수료율" style="WIDTH: 64px; HEIGHT: 22px"
																dataSrc="#xmlBind" type="text" maxLength="6" size="5" name="txtCOMMI_RATE">%
														</TD>
														<td class="DATA_RIGHT" align="right">VAT<INPUT id="chkTRU_TAX_FLAG" title="VAT유무" type="checkbox" CHECKED name="chkTRU_TAX_FLAG">
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtTIMNAME, txtTIMCODE)">팀</TD>
											<TD class="DATA"><INPUT dataFld="TIMNAME" class="INPUT_L" id="txtTIMNAME" title="팀명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="20" name="txtTIMNAME"> <IMG id="ImgTIMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgTIMCODE">
												<INPUT dataFld="TIMCODE" class="INPUT_L" id="txtTIMCODE" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="6" size="6" name="txtTIMCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREAL_MED_NAME, txtREAL_MED_CODE)">매체사</TD>
											<TD class="DATA"><INPUT dataFld="REAL_MED_NAME" class="INPUT_L" id="txtREAL_MED_NAME" title="매체사명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="7" name="txtREAL_MED_NAME">
												<IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE">
												<INPUT dataFld="REAL_MED_CODE" class="INPUT_L" id="txtREAL_MED_CODE" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" name="txtREAL_MED_CODE">
											</TD>
											<TD class="LABEL">색도</TD>
											<TD class="DATA"><SELECT dataFld="COL_DEG" id="cmbCOL_DEG" title="색도" style="WIDTH: 84px" dataSrc="#xmlBind"
													name="cmbCOL_DEG">
													<OPTION value="B/W">B/W</OPTION>
													<OPTION value="C/L" selected>C/L</OPTION>
												</SELECT>&nbsp;<INPUT id="chkPROJECTION" title="돌출" type="checkbox" name="chkPROJECTION">돌출</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCOMMISSION, '')">수수료</TD>
											<TD class="DATA">
												<TABLE cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
													<TR>
														<TD width="92"><INPUT dataFld="COMMISSION" class="INPUT_R" id="txtCOMMISSION" title="수수료" style="WIDTH: 92px; HEIGHT: 22px"
																accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="13" size="12" name="txtCOMMISSION">
														</TD>
														<td align="right"><SELECT dataFld="DUTYFLAG" id="cmbDUTYFLAG" style="WIDTH: 85px" dataSrc="#xmlBind" name="cmbDUTYFLAG">
																<OPTION value="Y" selected>영세</OPTION>
																<OPTION value="N">면세</OPTION>
															</SELECT>
														</td>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME, txtCLIENTCODE)">청구지</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPT_NAME, txtDEPT_CD)">담당부서</TD>
											<TD class="DATA"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPT_NAME" title="담당부서명" style="WIDTH: 123px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="6" name="txtDEPT_NAME">
												<IMG id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgDEPT_CD">
												<INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPT_CD" title="담당부서코드" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtDEPT_CD"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtMEMO, '')">비고</TD>
											<TD class="DATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="비고" style="WIDTH: 199px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="120" size="12" name="txtMEMO"></TD>
											<TD class="LABEL">발행</TD>
											<TD class="DATA"><INPUT id="chkGFLAG1" disabled type="radio" value="chkGFLAG1" name="chkGFLAG">미정
												<INPUT id="chkGFLAG2" disabled type="radio" value="chkGFLAG2" name="chkGFLAG">배정
												<INPUT id="chkGFLAG3" disabled type="radio" value="chkGFLAG3" name="chkGFLAG">집행
												<INPUT id="chkGFLAG4" disabled type="radio" value="chkGFLAG4" name="chkGFLAG">실적</TD>
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
