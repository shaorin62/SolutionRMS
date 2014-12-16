<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORLIST.aspx.vb" Inherits="MD.MDCMOUTDOORLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>개별청약 승인/조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : MD/OUTDOORLIST 청약승인화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMOUTDOORLIST.aspx
'기      능 : 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/23 By Hwang Duck su
			:2) 2009/09/28 By Kim Tae Yub
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDOTOUTDOOR
Dim mstrPROCESS	'신규이면 True 조회면 False
Dim mobjMDCMGET
Dim mstrCheck
Dim mstrCheck2

CONST meTAB = 9
mstrPROCESS = False

mstrCheck = True
mstrCheck2 = True

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtYEARMON1.value = "" then
		gErrorMsgBox "청구년월을 입력하시오","조회안내"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
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
	'Window_OnUnload
End Sub

Sub imgSetting_onclick
	Call ProcessRtn_ConfirmOK()
End Sub

Sub ImgConfirmCancel_onclick
	Call ProcessRtn_ConfirmCancel()
End Sub

Sub ImgAORSave_onclick
	Call AOR_Confirm("CONFIRM")
End Sub

Sub ImgAORSaveCancel_onclick
	Call AOR_Confirm("CANCEL")
End Sub

'-----------------------------------------------------------------------------------------
' 내역복사한다.
'-----------------------------------------------------------------------------------------
Sub Imgcopy_onclick ()
	Dim intRtn
   	Dim vntData
	Dim intSelCnt,  i
	Dim strYEARMON, strGUBUN, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strREAL_MED_CODE, strREAL_MED_NAME, strREAL_MED_BISNO
	Dim strMEDCODE, strMEDNAME, strMED_FLAG, strHIGHSUBSEQ, strSUBSEQNAME, strDEMANDDAY, strTBRDSTDATE
	Dim strTBRDEDDATE, strGBN_FLAG, strTITLE, strMATTERNAME, strTOTALAMT, strAMT, strOUT_AMT
	Dim strCOMMI_RATE, strCOMMISSION, strMED_GBN, strLOCATION, strMEMO, strCONTIDX, strMDIDX, strCYEAR, strCMONTH, strSIDE, strPORTAL_SEQ, strCOMMI_TAX_FLAG
	 
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
		
		strYEARMON			=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",strCNT)
		strGUBUN			=	"미승인"
		strCLIENTCODE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",strCNT)
		strCLIENTNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",strCNT)
		strTIMCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMCODE",strCNT)
		strTIMNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TIMNAME",strCNT)
		strREAL_MED_CODE	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",strCNT)
		strREAL_MED_NAME	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",strCNT)
		strREAL_MED_BISNO	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",strCNT)	
		strMEDCODE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",strCNT)
		strMEDNAME			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",strCNT)
		strMED_FLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_FLAG",strCNT)
		strHIGHSUBSEQ		=	mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSUBSEQ",strCNT)
		strSUBSEQNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",strCNT)
		strDEMANDDAY		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",strCNT)
		strTBRDSTDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",strCNT)
		strTBRDEDDATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",strCNT)
		strGBN_FLAG			=	mobjSCGLSpr.GetTextBinding(.sprSht,"GBN_FLAG",strCNT)
		strTITLE			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TITLE",strCNT)
		strMATTERNAME		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERNAME",strCNT)
		strTOTALAMT			=	mobjSCGLSpr.GetTextBinding(.sprSht,"TOTALAMT",strCNT)
		strAMT				=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",strCNT)
		strOUT_AMT			=	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",strCNT)
		strCOMMI_RATE		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",strCNT)
		strCOMMISSION		=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",strCNT)
		strMED_GBN			=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_GBN",strCNT)
		strLOCATION			=	mobjSCGLSpr.GetTextBinding(.sprSht,"LOCATION",strCNT)
		strMEMO				=	mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",strCNT)
		strCOMMI_TAX_FLAG	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TAX_FLAG",strCNT)
		
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, meINS_ROW, 0, -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, strYEARMON
		mobjSCGLSpr.SetTextBinding .sprSht,"GUBUN",.sprSht.ActiveRow, strGUBUN
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, strCLIENTCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, strCLIENTNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",.sprSht.ActiveRow, strTIMCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",.sprSht.ActiveRow, strTIMNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow, strREAL_MED_CODE
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",.sprSht.ActiveRow, strREAL_MED_NAME
		mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_BISNO",.sprSht.ActiveRow, strREAL_MED_BISNO
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",.sprSht.ActiveRow, strMEDCODE
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",.sprSht.ActiveRow, strMEDNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"MED_FLAG",.sprSht.ActiveRow, strMED_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"HIGHSUBSEQ",.sprSht.ActiveRow, strHIGHSUBSEQ
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, strSUBSEQNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, strDEMANDDAY
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",.sprSht.ActiveRow, strTBRDSTDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",.sprSht.ActiveRow, strTBRDEDDATE
		mobjSCGLSpr.SetTextBinding .sprSht,"GBN_FLAG",.sprSht.ActiveRow, strGBN_FLAG
		mobjSCGLSpr.SetTextBinding .sprSht,"TITLE",.sprSht.ActiveRow, strTITLE		
		mobjSCGLSpr.SetTextBinding .sprSht,"MATTERNAME",.sprSht.ActiveRow, strMATTERNAME
		mobjSCGLSpr.SetTextBinding .sprSht,"TOTALAMT",.sprSht.ActiveRow, strTOTALAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT
		mobjSCGLSpr.SetTextBinding .sprSht,"OUT_AMT",.sprSht.ActiveRow, strOUT_AMT
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, strCOMMI_RATE
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",.sprSht.ActiveRow, strCOMMISSION
		mobjSCGLSpr.SetTextBinding .sprSht,"MED_GBN",.sprSht.ActiveRow, strMED_GBN
		mobjSCGLSpr.SetTextBinding .sprSht,"LOCATION",.sprSht.ActiveRow, strLOCATION
		mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",.sprSht.ActiveRow, strMEMO
		
		mobjSCGLSpr.SetTextBinding .sprSht,"CONTIDX",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"MDIDX",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"CYEAR",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"CMONTH",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"SIDE",.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding .sprSht,"PORTAL_SEQ",.sprSht.ActiveRow, ""
		
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_TAX_FLAG",.sprSht.ActiveRow, strCOMMI_TAX_FLAG
		
		mobjSCGLSpr.SetCellsLock2 .sprSht,False,"TOTALAMT | AMT | OUT_AMT",.sprSht.ActiveRow,.sprSht.ActiveRow,False

		mstrPROCESS = False
   	end With
end Sub


'프린트
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
	Dim intCnt
	Dim lngCHK, lngCHKSUM
	Dim intRtn
	Dim strYEARMON, strSEQ, strNUM, strUSERID
	Dim vntData
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
			Exit Sub
		end if
		
		intRtn = mobjMDOTOUTDOOR.DeleteRtn_temp(gstrConfigXml)
		
		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",intCnt) = "미승인" THEN
   					gErrorMsgBox "승인상태인 데이터만 출력이 가능합니다.","인쇄안내!"
					Exit Sub
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "인쇄할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		ModuleDir = "MD"

		ReportName = "MDCMOUTDOOR_MEDIUM_Y_NEW.rpt"
		
		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				strYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intCnt)
				strSEQ		= mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",intCnt)
				strNUM		= intCnt
				strUSERID = ""
				vntData = mobjMDOTOUTDOOR.ProcessRtn_TEMP(gstrConfigXml,strYEARMON, strSEQ, strNUM, strUSERID)
			END IF
		Next
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "call printSetTimeout('" & strYEARMON & "', '" & strSEQ & "')", 10000
	end with  
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout(strYEARMON, strSEQ)
	Dim intRtn, intRtn2
	With frmThis
		intRtn = mobjMDOTOUTDOOR.DeleteRtn_temp(gstrConfigXml)
	End With
End Sub

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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
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
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			SelectRtn
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
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE1.value),trim(.txtREAL_MED_NAME1.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = trim(vntData(0,1))
					.txtREAL_MED_NAME1.value = trim(vntData(1,1))
					SelectRtn
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
			SelectRtn
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
			vntData = mobjMDCMGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), _
											trim(.txtTIMCODE1.value),trim(.txtTIMNAME1.value))
			
			If not gDoErrorRtn ("GetTIMCODE") Then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))	    ' Code값 저장
					.txtTIMNAME1.value = trim(vntData(1,1))       ' 코드명 표시
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(5,1))
					SelectRtn
				Else
					Call TIMCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK")  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			NEXT
		ELSEIF  Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_TAX_FLAG")  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_TAX_FLAG"), mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_TAX_FLAG"),,, , , , , , mstrCheck2
			if mstrCheck2 = True then 
				mstrCheck2 = False
			elseif mstrCheck2 = False then 
				mstrCheck2 = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"COMMI_TAX_FLAG"), intcnt
			next
		end if
	end with
End Sub  

'시트 더블클릭 
sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strATTR01
	Dim vntInParams
	Dim vntRet

	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		ElseIf Row > 0 and Col > 0 then
			strATTR01 =  mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",Row)

			If mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",Row) = "승인" Then
				vntInParams = array(strATTR01) '<< 받아오는경우
				vntRet = gShowModalWindow("MDCMOUTDOORDIVPOP.aspx",vntInParams , 590,430)
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			else
				gErrorMsgBox "승인되지 않은 데이터는 상세 내역의 데이터를 수정하시거나 변경 하실수 없습니다.","승인안내!"
				exit sub
			end if 
		end if
		
	end with
end sub




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
	
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") OR _
		   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
		   
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strCOLUMN = "COMMISSION"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") Then
				strCOLUMN = "TOTALAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
				strCOLUMN = "OUT_AMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) OR _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") or  _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTALAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
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
   	Dim strTOTALAMT, strAMT, strOUT_AMT
   	Dim strCOMMI_RATE, strCOMMISSION
   	
   	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"), Row)
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then 
			Call SHEET_COMMI_RATE_Cal (mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT"), Row)
		End If
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub SHEET_COMMI_RATE_Cal (Col, Row)
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim intAMT,intOUT_AMT
	Dim dblCOMMI_RATE
	Dim intCOMMISSION
	With frmThis
	
		If Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
			If intAMT <> 0 AND intOUT_AMT <> 0 Then
				intCOMMISSION = intAMT - intOUT_AMT
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
				dblCOMMI_RATE = gRound((intCOMMISSION / intAMT),2)
   				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
			ELSE
				IF intAMT = 0 THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intAMT
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 1
				END IF
				
			End If
		ELSEIF Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
			If intAMT <> 0 AND intOUT_AMT <> 0 Then
				intCOMMISSION = intAMT - intOUT_AMT
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intCOMMISSION
				dblCOMMI_RATE = gRound((intCOMMISSION / intAMT),2)
   				mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, dblCOMMI_RATE
			ELSE
				IF intAMT = 0 THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row, intAMT
					mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row, 1
				END IF
				
			End If
		End If
	End With
end Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성									
	set mobjMDOTOUTDOOR	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR")
	set mobjMDCMGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 41, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_BISNO | REAL_MED_NAME | MEDCODE | MEDNAME | MED_FLAG | DEPT_CD | DEPT_NAME | HIGHSUBSEQ | SUBSEQNAME | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | COMMI_TAX_FLAG | COMMI_TRANS_NO | ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|년월|번호|구분|광고주코드|광고주|팀코드|팀|외주처코드|사업자번호|외주처|외주처코드|외주처|매체구분코드|담당부서코드|담당부서명|브랜드코드|브랜드명|청구일자|계약시작일|계약종료일|매출구분|계약명|소재명|총계약금액|월청구금액|월외주비|내수율|내수액|제작종류|장소|비고|포탈계약번호|포탈매체번호|포탈년도|포탈월|면구분|AOR지정|부가세유무|거래명세서번호|상세번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   0|   5|         0|    13|     0|10|         0|        13|    13|         0|     0|           0|			 0|        13|         0|      10|       8|         8|         8|       9|    15|    15|        10|        10|      10|     6|     9|      10|  10|  10|           0|           0|       0|     0|     0|      6|       6  |            10|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | COMMI_TAX_FLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOTALAMT | AMT | OUT_AMT | COMMISSION ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | MEDCODE | MEDNAME | MED_FLAG | DEPT_CD | DEPT_NAME | HIGHSUBSEQ | SUBSEQNAME | GBN_FLAG | TITLE | MATTERNAME | MED_GBN | LOCATION | MEMO | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | ATTR01", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | MEDCODE | MEDNAME | MED_FLAG | HIGHSUBSEQ | SUBSEQNAME | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | COMMI_TRANS_NO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GUBUN | PORTAL_SEQ",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | CLIENTCODE | TIMCODE | REAL_MED_CODE | MEDCODE | MEDNAME | MED_FLAG | HIGHSUBSEQ | GBN_FLAG | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | ATTR01", true
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDOTOUTDOOR = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtYEARMON1.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		.txtCLIENTNAME1.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim strYEARMON, strCLIENTCODE,strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
	Dim strTIMCODE, strTIMNAME, strTITLE
   	Dim strGUBUN
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

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
		strTITLE		 = .txtTITLENAME1.value
		strGUBUN		 = .cmbGUBUN.value
		
		vntData = mobjMDOTOUTDOOR.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											strYEARMON, _
											strCLIENTCODE, strCLIENTNAME, _
											strREAL_MED_CODE, strREAL_MED_NAME, _
											strTIMCODE, strTIMNAME, strTITLE, strGUBUN)

		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			
			AMT_SUM
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE	
   		end if
   		mstrPROCESS = True
   	end with
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


'------------------------------------------
' 승인 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG 
	
	strFLAG = "CONFIRM"
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 승인이 불가능 합니다.","승인안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",intCnt) = "승인" THEN
   					gErrorMsgBox "미승인상태인 데이터만 승인이 가능합니다.","저장안내!"
					Exit Sub
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "승인할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		'여기서 부터 문제
		'if DataValidation =false then exit sub
	    '데이터 Validation End
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_BISNO | REAL_MED_NAME | MEDCODE | MEDNAME | MED_FLAG | DEPT_CD | DEPT_NAME | HIGHSUBSEQ | SUBSEQNAME | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | COMMI_TAX_FLAG | COMMI_TRANS_NO | ATTR01 ")
		
		intRtn = mobjMDOTOUTDOOR.ProcessRtn(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOUTDOOR_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 승인취소 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmCancel
    Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG 
	
	strFLAG = "CANCEL"
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 승인취소이 불가능 합니다.","승인취소안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",intCnt) = "미승인" THEN
   					gErrorMsgBox "승인상태인 데이터만 승인취소가 가능합니다.","저장안내!"
					Exit Sub					
   				END IF

   				if mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",intCnt) = "승인" AND mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",intCnt) <> "" THEN
   					gErrorMsgBox "청구진행 하지않은 데이터만 가능합니다.","저장안내!"
					Exit Sub
   				END IF
   				
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "승인취소할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_BISNO | REAL_MED_NAME | MEDCODE | MEDNAME | MED_FLAG | HIGHSUBSEQ | SUBSEQNAME | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | COMMI_TAX_FLAG | COMMI_TRANS_NO | ATTR01 ")
		
		intRtn = mobjMDOTOUTDOOR.ProcessRtn(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOUTDOOR_OK") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인취소" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub


'------------------------------------------
' AOR 저장로직
'------------------------------------------
Sub AOR_Confirm (strFLAG)
	Dim intRtn
   	Dim vntData
	Dim strYEARMON, strSEQ
	Dim i
	Dim lngchkCnt
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 AOR지정이 불가능 합니다.","승인안내!"
			Exit Sub
		end if
		
   		lngchkCnt = 0
   		
   		For i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",i) = "미승인" THEN
   					gErrorMsgBox "승인상태인 데이터만 AOR지정 및 취소가 가능합니다.","AOR지정및취소안내!"
					Exit Sub					
   				END IF
   				
				strYEARMON  = mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",i)
				strSEQ		= mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",i)
				
				vntData = mobjMDOTOUTDOOR.VOCHNO_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strSEQ) 
				
				If mlngRowCnt > 0 Then
					'gErrorMsgBox i & "행의 데이터는 전표가 발생되어 AOR 지정 및 취소를 할수 없습니다.","AOR지정및취소안내!"
					'Exit Sub
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next

		If lngchkCnt = 0 Then
			gErrorMsgBox "AOR 지정할 데이터를 선택 하십시오.","AOR지정안내!"
			Exit Sub
		End If
		'여기서 부터 문제
		'if DataValidation =false then exit sub
	    '데이터 Validation End
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | GUBUN | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | REAL_MED_CODE | REAL_MED_NAME | MEDCODE | MEDNAME | MED_FLAG | HIGHSUBSEQ | SUBSEQNAME | DEMANDDAY | TBRDSTDATE | TBRDEDDATE | GBN_FLAG | TITLE | MATTERNAME | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MED_GBN | LOCATION | MEMO | CONTIDX | MDIDX | CYEAR | CMONTH | SIDE | PORTAL_SEQ | COMMI_TAX_FLAG")
		
		intRtn = mobjMDOTOUTDOOR.AOR_Confirm(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("AOR_Confirm") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngchkCnt & " 건의 자료가 처리" & mePROC_DONE
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
													<TABLE cellSpacing="0" cellPadding="0" width="197" background="../../../images/back_p.gIF"
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
												<td class="TITLE">청약관리 - 개별청약 조회 및 승인</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Wait Button Start-->
										<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton2" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
														height="20" alt="화면을 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
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
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
													width="50">청구년월</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="년월조회" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
														maxLength="6" size="10" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="50">광고주</TD>
												<TD class="SEARCHDATA" width="250"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 173px; HEIGHT: 22px"
														maxLength="100" align="left" size="22" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
														maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
													width="50">팀</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="팀명" style="WIDTH: 173px; HEIGHT: 22px" maxLength="100"
														size="22" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
														align="absMiddle" border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" maxLength="6"
														size="6" name="txtTIMCODE1"></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">구분</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbGUBUN" title="구분" style="WIDTH: 96px" name="cmbGUBUN">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="Y">승인</OPTION>
														<OPTION value="N">미승인</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)">매체사</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="매체사명" style="WIDTH: 173px; HEIGHT: 22px"
														maxLength="100" size="7" name="txtREAL_MED_NAME1"> <IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="매체사코드" style="WIDTH: 53px; HEIGHT: 22px"
														maxLength="6" name="txtREAL_MED_CODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTITLENAME1, '')">계약명</TD>
												<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtTITLENAME1" title="계약명" style="WIDTH: 244px; HEIGHT: 22px"
														maxLength="100" size="36" name="txtTITLENAME1"></TD>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD align="left" width="400" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="Imgcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
																	alt="자료를 인쇄합니다." src="../../../images/imglistcopy.gIF" border="0" name="Imgcopy"></TD>
															<TD><IMG id="ImgAORSave" onmouseover="JavaScript:this.src='../../../images/ImgAORSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAORSave.gIF'"
																	height="20" alt="자료를승인처리합니다." src="../../../images/ImgAORSave.gIF" border="0" name="ImgAORSave"></TD>
															<td><IMG id="ImgAORSaveCancel" onmouseover="JavaScript:this.src='../../../images/ImgAORSaveCancelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAORSaveCancel.gIF'"
																	height="20" alt="승인처리를 취소합니다." src="../../../images/ImgAORSaveCancel.gif" border="0"
																	name="ImgAORSaveCancel"></td>
															<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
																	height="20" alt="자료를승인처리합니다." src="../../../images/imgAgree.gIF" border="0" name="imgSetting"></TD>
															<td><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCancelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCancel.gIF'"
																	height="20" alt="승인처리를 취소합니다." src="../../../images/imgAgreeCancel.gif" border="0"
																	name="ImgConfirmCancel"></td>
															<td><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="개별 거래명세서를 출력합니다.." src="../../../images/imgPrint.gIF" border="0"
																	name="imgPrint"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
											ms_positioning="GridLayout">
											<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="15372">
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
								<!--List End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
