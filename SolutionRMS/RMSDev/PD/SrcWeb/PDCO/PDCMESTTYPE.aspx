<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMESTTYPE.aspx.vb" Inherits="PD.PDCMESTTYPE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>견적 TYPE유형 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
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
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
'전역변수 설정
Dim mobjPDCOESTTYPE
Dim mobjPDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag
Dim mobjSCCOGET
Dim mstrCheck
CONST meTAB = 9
'=========================================================================================
' 이벤트 프로시져 
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
Sub imgExcel1_onclick()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht1
		end with
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
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgTypeNew_onclick()
	with frmThis
		'Field_TypeChange("T")
		gClearAllObject frmThis
		initpageData
		.txtTYPENAME.focus()
	End with
End Sub

Sub imgRowDel_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtnAll
	gFlowWait meWAIT_OFF
End Sub

Sub Field_TypeChange(byval strCHK)
	with frmThis
		If strCHK = "T" Then
			.txtTYPENAME.className = "INPUT_L"
			.txtTYPENAME.readOnly = false
			.txtCLIENTCODE.className = "INPUT"
			.txtCLIENTCODE.readOnly = false
			.ImgCLIENTCODE.disabled = false
			.txtCLIENTNAME.className = "INPUT_L"
			.txtCLIENTNAME.readOnly = false
		ElseIf strCHK = "F" Then
			.txtTYPENAME.className = "NOINPUT_L"
			.txtTYPENAME.readOnly = true
			.txtCLIENTCODE.className = "NOINPUT"
			.txtCLIENTCODE.readOnly = true
			.ImgCLIENTCODE.disabled = true
			.txtCLIENTNAME.className = "NOINPUT_L"
			.txtCLIENTNAME.readOnly = true
		End If
	End with
End Sub




Sub imgTableUP_onclick
	Dim strRow
	
	with frmThis
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "선택된 데이터가 없습니다.","이동안내!"
			Exit Sub
		end if 
		if strRow = 1 then exit sub
		
		'자기자신을 넘긴다.
		sprSht_UpCopy strRow
	
	end with
End Sub

Sub imgTableDown_onclick

	with frmThis
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "선택된 데이터가 없습니다.","이동안내!"
			Exit Sub
		end if 
		
		if strRow = (.sprSht.MaxRows) then exit sub	
		
		sprSht_DownCopy strRow
	end with
End Sub


Sub sprSht_UpCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row셋팅	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'msgbox strRow
		'돌면서 자신과 printseq만큼 위에꺼 복사
		for i=1 to .sprSht_copy.MaxRows
'CHK|PREESTNO|PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN|DETAIL_BTN|SUBDETAIL|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|PRODUCTIONCOMMISSION|HDRSEQ			
			'msgbox strRow & "-" & strPRINT_SEQ
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow- strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow-strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"HDRSEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"HDRSEQ",strRow -strPRINT_SEQ)
			
			
			'기본으로 Y를 박는다 ' <-필요없음!! 한번이라도 이동했다면 전체 저장해야함.
			'mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
'			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
'				strCopySeq = mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
'				exit for
'			End if 	
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				'msgbox "1일때"
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ 

'CHK|PREESTNO|PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN|DETAIL_BTN|SUBDETAIL|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|PRODUCTIONCOMMISSION|HDRSEQ			
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				'mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"HDRSEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"HDRSEQ",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				'msgbox "아닐때"
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i) + strPRINT_SEQ
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				'mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"HDRSEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"HDRSEQ",i)
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


Sub sprSht_DownCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row셋팅	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'돌면서 자신과 printseq만큼 위에꺼 복사
		for i=1 to .sprSht_copy.MaxRows
			'msgbox strRow & "+" & strPRINT_SEQ & "=" & strRow+strPRINT_SEQ
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow+ strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow+strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"HDRSEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"HDRSEQ",strRow +strPRINT_SEQ)
			
			
			'기본으로 MOVEFLAG 에 Y를 박는다.
			'mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
'			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
'				strCopySeq = mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
'				exit for
'			End if 	
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				'msgbox "1일때"
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ 
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				'mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"HDRSEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"HDRSEQ",i)
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				'msgbox "아닐때"
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i) - strPRINT_SEQ
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				'mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"HDRSEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"HDRSEQ",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
				'msgbox mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 		
			
			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
		
		If .txtSUSURATE.value = 0 Then
			.txtSUSUAMT.readOnly = false
		Else
			.txtSUSUAMT.readOnly = true
		End If
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub

Sub txtSUSURATE_onblur
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		
	End with
End Sub

Sub CleanFieldRate
	with frmThis
		.txtSUSURATE.value = 0
	End With
End Sub

Sub txtCOMMISSION_onfocus
	with frmThis
		.txtCOMMISSION.value = Replace(.txtCOMMISSION.value,",","")
	end with
End Sub
Sub txtCOMMISSION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMISSION,0,true)
	end with
End Sub
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
		
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
Sub txtNONECOMMISSION_onfocus
	with frmThis
		.txtNONECOMMISSION.value = Replace(.txtNONECOMMISSION.value,",","")
	end with
End Sub
Sub txtNONECOMMISSION_onblur
	with frmThis
		call gFormatNumber(.txtNONECOMMISSION,0,true)
	end with
End Sub
Sub txtAMT_onfocus
	with frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'서버업무객체 생성	
	Set mobjPDCOESTTYPE = gCreateRemoteObject("cPDCO.ccPDCOESTTYPE")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'화면 초기값 설정
	'SEARCHCOMBO_TYPE
	InitPageData	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht1.maxrows = 0
	.sprSht.maxrows = 0
	.txtSUSURATE.value = 0
	.txtSUSUAMT.value = 0
	.txtSEQ.style.visibility = "hidden"
	'Field_TypeChange("F")
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
	End with
End Sub

Sub Grid_Layout()
	Dim intGBN
	
	
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'*** 헤더조회 Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 10, 0, 5
		mobjSCGLSpr.SpreadDataField .sprSht1,    "SEQ|TYPENAME|CLIENTCODE|CLIENTNAME|COMMISSION|NONECOMMISSION|SUSURATE|SUSUAMT|AMT|SUMAMT"
		mobjSCGLSpr.SetHeader .sprSht1,		    "No.|견적유형명|광고주코드|광고주|COMMISSION|NONECOMMISSION|수수료율|수수료|금액|견적금액"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "4  |20        |0         |20    |12        |12            |8       |12    |12  |13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "COMMISSION|NONECOMMISSION|SUSUAMT|AMT|SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "SUSURATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "SEQ",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "TYPENAME|CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"SEQ|TYPENAME|CLIENTCODE|CLIENTNAME|COMMISSION|NONECOMMISSION|SUSURATE|SUSUAMT|AMT|SUMAMT"
		mobjSCGLSpr.ColHidden .sprSht1, "CLIENTCODE", true
		mobjSCGLSpr.SetScrollBar .sprSht1,2,True,0,-1
		
		
		'**************************************************
		'***상세내역 Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 25, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|PREESTNO|PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN|DETAIL_BTN|SUBDETAIL|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|PRODUCTIONCOMMISSION|HDRSEQ"
		mobjSCGLSpr.SetHeader .sprSht,		  "선택|가견적번호|이동|순번|대분류|중분류|견적항목코드|견적항목명|견적명|내역|커미션|수량|단가|금액|수수료금액|저장구분|상세견적|상세견적여부|가짜순번|상세저장여부|상세부분여부|다이렉트플레그|간접비|TYPENO"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","    4|       10|   4|   4|     8|    12|        8 |2|        15|12    |  20|     6|  12|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     |10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		'mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"상세견적", "DETAIL_BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK|COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRINT_SEQ|QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMCODENAME|STD|FAKENAME", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|PREESTNO|DETAIL_BTN|IMESEQ|SAVEFLAG|DETAILYNFLAG|HDRSEQ"
		mobjSCGLSpr.ColHidden .sprSht, "GBN|SUBDETAIL|INDIRECFLAG|PRODUCTIONCOMMISSION|PREESTNO|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|HDRSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME|CLASSNAME|FAKENAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PRINT_SEQ|ITEMCODE|ITEMCODESEQ",-1,-1,2,2,false
		
		
		
		'**************************************************
		'***상세내역 Sheet copy
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_copy
		mobjSCGLSpr.SpreadLayout .sprSht_copy, 25, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_copy, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_copy, "CHK|PREESTNO|PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN|DETAIL_BTN|SUBDETAIL|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|PRODUCTIONCOMMISSION|HDRSEQ"
		mobjSCGLSpr.SetHeader .sprSht_copy,		  "선택|가견적번호|이동|순번|대분류|중분류|견적항목코드|견적항목명|견적명|내역|커미션|수량|단가|금액|수수료금액|저장구분|상세견적|상세견적여부|가짜순번|상세저장여부|상세부분여부|다이렉트플레그|간접비|TYPENO"
		mobjSCGLSpr.SetColWidth .sprSht_copy, "-1","    4|       10|   4|   4|     8|    12|        8 |2|        15|12    |  20|     6|  12|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     |10"
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_copy,"..", "BTN"
		'mobjSCGLSpr.SetCellTYpeButton2 .sprSht_copy,"상세견적", "DETAIL_BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_copy, "CHK|COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "PRINT_SEQ|QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_copy, "ITEMCODENAME|STD|FAKENAME", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht_copy, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_copy, true, "PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|PREESTNO|DETAIL_BTN|IMESEQ|SAVEFLAG|DETAILYNFLAG|HDRSEQ"
		mobjSCGLSpr.ColHidden .sprSht_copy, "GBN|SUBDETAIL|INDIRECFLAG|PRODUCTIONCOMMISSION|PREESTNO|IMESEQ|SAVEFLAG|DETAILYNFLAG|INDIRECFLAG|HDRSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "DIVNAME|CLASSNAME|FAKENAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "PRINT_SEQ|ITEMCODE|ITEMCODESEQ",-1,-1,2,2,false
		
		
		
	End with
	'DateClean
	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	pnlTab3.style.visibility = "visible" 
End Sub
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODESEARCH_POP()
End Sub

Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
'-----------------------------------------------------------------------------------------
' 광고주팝업(조회)
'-----------------------------------------------------------------------------------------
Sub CLIENTCODESEARCH_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)	
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value),"A")
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				
				Else
					Call CLIENTCODESEARCH_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 광고주팝업(입력)
'-----------------------------------------------------------------------------------------
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)	
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),"A")
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
				
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'검색조건 시작일
Sub imgFrom_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtFrom,.imgFrom,"txtFrom_onchange()"
		gSetChange
	end with
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub

'검색조건 종료일
Sub imgTo_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtTo,.imgTo,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtTo_onchange
	gSetChange
End Sub


Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCOESTTYPE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value,.txtSEARCHTYPENAME.value)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht1, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				sprSHT1_Click 1,1
   			Else
   				.sprSht1.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_ProcessRtn (Byval strHDRSEQ)
   	Dim vntData
   	Dim i, strCols

    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjPDCOESTTYPE.SelectRtn_ProcessRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strHDRSEQ)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht1, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				sprSHT1_Click 1,1
   			Else
   				.sprSht1.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub


Sub SelectRtn_REAL(byval strHDRSEQ)
	With frmThis
		IF not SelectRtn_Head (strHDRSEQ) Then Exit Sub
		
		CALL SelectRtn_Detail (strHDRSEQ)
		txtSUSUAMT_onblur
	    txtCOMMISSION_onblur
		txtSUMAMT_onblur
		txtNONECOMMISSION_onblur
		txtAMT_onblur
	End With 
End SUb

Function SelectRtn_Head (ByVal strHDRSEQ)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCOESTTYPE.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strHDRSEQ)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			frmThis.sprSht.MaxRows = 0	
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function


Function SelectRtn_Detail(ByVal strHDRSEQ)
	Dim vntData
   	Dim i, strCols
   	Dim intColorCnt
   	

    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjPDCOESTTYPE.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strHDRSEQ)
		
		if not gDoErrorRtn ("SelectRtn_DTL") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				Detail_Yn
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
   			Else
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Function

Sub Detail_Yn()
	Dim intCnt
	with frmThis
		For intCnt =1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",intCnt) = "Y" Then
				'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",intCnt,intCnt,false
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",intCnt,intCnt,false
				'버튼형태로 변경
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",intCnt,intCnt,,false
				Else
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",intCnt,intCnt,,false
				End If
				
			Else
				'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
				'mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DETAIL_BTN|QTY|PRICE|AMT",Row,Row,false	
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",intCnt,intCnt,false
				'일반형태로 변경
				'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "DETAIL_BTN",Row,Row,255,,,,,False
				mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",intCnt,intCnt,0,,,,,,,,False
			End If
		Next
	End With
End Sub


Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFrom.value = date1
		.txtTo.value = date2
	End With
End Sub
Sub EndPage()
	set mobjPDCOESTTYPE = Nothing
	set mobjPDCMGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
'Sheet 
'-----------------------------------------------------------------------------------------
'행추가 버튼클릭
Sub imgRowAdd_onclick ()
	with frmThis
		If .txtTYPENAME.className = "NOINPUT_L" AND .txtSEQ.value = "" Then
			gErrorMsgBox "신규등록 및 조회대상이 아닌내역은 행추가를 할수 없습니다.","처리안내"
			Exit Sub
		End If
	End with
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

Sub sprSht_Keydown(KeyCode, Shift)

	Dim intRtn
	
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select
End Sub



Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, "9999999999" 
		mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",.sprSht.ActiveRow, "N"
		If .sprSht.MaxRows = 1 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow,1
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",.sprSht.ActiveRow-1) + 1
		End If
		
		If .txtSEQ.value <> "" Then
			mobjSCGLSpr.SetTextBinding .sprSht,"HDRSEQ",.sprSht.ActiveRow,.txtSEQ.value 
		End If
	End With
End Sub
'스프레드의 모든 항목의 값이 변경 될때 발생 하는 이벤트 입니다.
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, strCols
	Dim strCode, strCodeName
	Dim strQTY, strPRICE, strAMT
	Dim lngPrice
	Dim lngVALUE
	Dim lngVALUE1
	Dim lngVALUE2

	with frmThis
				'Long Type의 ByRef 변수의 초기화
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				strCode = ""
				strCodeName = ""
				IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
					strCode = ""
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",.sprSht.ActiveRow)
					vntData = mobjPDCMGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",strCodeName)
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntData(4,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntData(7,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntData(8,1)	
						
						'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
						If vntData(7,1) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
						
							'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
							mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",Row,Row,false
							'수량,단가,금액은 0처리
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
							'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
							If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
							Else
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
							End If
						Else
							'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",Row,Row,false
							'일반형태로 변경
							mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
						End If	
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						SUSUAMT_CHANGEVALUE2
						BUDGET_AMT_SUM
					Else
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End If
					.txtSUSURATE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
					.sprSht.Focus	
					mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
				'수량로직	
				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strPRICE <> "" And strAMT = "" Then
   						lngVALUE = strQTY * strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE
   					ElseIf strPRICE = "" And strAMT <> "" Then
   						lngVALUE1 = gRound(strAMT/strQTY,0)
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngVALUE1
   					ElseIf strPRICE <> "" And strAMT <> "" Then
   						lngVALUE2 = strQTY * strPRICE
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE2
   					End IF
   					Call SUSUAMT_CHANGEVALUE(Row)
   					BUDGET_AMT_SUM
   				'단가 로직
   				ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
   					strQTY		= mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",.sprSht.ActiveRow)
					strPRICE   = mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",.sprSht.ActiveRow)
					strAMT = strQTY * strPRICE
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT	
					Call SUSUAMT_CHANGEVALUE(Row)
					BUDGET_AMT_SUM
				'금액로직	
   				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strAMT = 0 Then
   						mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, strAMT
   					Else 
   						If strQTY <> 0  Then
   							lngPrice = gRound(strAMT/strQTY,0)
   							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngPrice
   						End IF
   					End IF
   					Call SUSUAMT_CHANGEVALUE(Row)
   					BUDGET_AMT_SUM
   				Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMIFLAG") Then
   					Call SUSUAMT_CHANGEVALUE2
   					BUDGET_AMT_SUM
				END IF
	end with
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
Sub SUSUAMT_CHANGEVALUE(ByVal Row)
	Dim strAMT,strCOMMIFLAG
	with frmThis
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		if strCOMMIFLAG = "1" Then
			'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
	End with
End SUb
Sub SUSUAMT_CHANGEVALUE2
Dim intCnt
Dim strAMT,strCOMMIFLAG
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		if strCOMMIFLAG = "1" Then
		'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
	Next
	
	End with
End Sub
Sub txtSUSURATE_onchange
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSURATE  
	End with
End Sub
Sub txtSUSUAMT_onchange
	with frmThis
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSUAMT  
	End with
End Sub

Sub BUDGET_AMT_SUM
	'총합계 변수
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM, intAMTSUB
	Dim lngSUSU
	'수수료 계산 변수
	Dim intCnt,intSUSU,intSUSUSUM 
	'COMMISSION 계산 변수
	Dim intCnt1,intCOM,intCOMSUM 
	'NONECOMMISSION 계산변수
	Dim intCnt2,intNON,intNONSUM 
	
	with frmThis
	
		IntAMTSUM = 0
		IntPRICESUM = 0
		intSUSU = 0
		intSUSUSUM = 0
		intCOM = 0
		intCOMSUM = 0
		intNON = 0
		intNONSUM = 0
		intAMTSUB = 0
		For intCnt = 1 To .sprSht.MaxRows
			intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
			intSUSUSUM = intSUSUSUM + intSUSU
		Next
		If .txtSUSURATE.value <> 0 Then
			.txtSUSUAMT.value = intSUSUSUM
		End If
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		intAMTSUB = IntAMTSUM
		If .txtSUSURATE.value <> 0 Then
			IntAMTSUM = IntAMTSUM + intSUSUSUM
		Else
			IntAMTSUM = IntAMTSUM + .txtSUSUAMT.value 
		End If
		.txtSUMAMT.value = IntAMTSUM
		
		.txtAMT.value = intAMTSUB
		
		For intCnt1 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG", intCnt1) = "1" Then
				
				intCOM = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intCOMSUM = intCOMSUM + intCOM
			Else
				
				intNON = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intNONSUM = intNONSUM + intNON
			end if
		Next
		.txtCOMMISSION.value = intCOMSUM
		.txtNONECOMMISSION.value = intNONSUM
		txtAMT_onblur
		txtSUSUAMT_onblur
		txtCOMMISSION_onblur
		txtSUMAMT_onblur
		txtNONECOMMISSION_onblur
		
	End With
End Sub
'스프레드의 행을 더블 클릭 시 발생
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'스프레드의 항목이 변할시 어떠한 함수를 태우고자 할때 사용
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding( sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)	
				
				'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
				If vntRet(7,0) = "Y" Then
				
					'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",Row,Row,false
					'수량,단가,금액은 0처리
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",Row,Row,false
					'일반형태로 변경
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If	
				'mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
	
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
			End IF
			
			.txtSUSURATE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub
'스프레드 내 버튼을 클릭 하였을때 발생 하는 이벤트
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strPREESTNO
	Dim strGBN
	Dim strITEMCODE
	Dim dblAMT
	Dim vntData
	Dim strRTN
	Dim strColCnt
	Dim strRowCnt
	Dim intCnt
	Dim strChk
	Dim intTempChkCnt
	Dim dblSUSUAMT
	Dim intColorCnt
	
	Dim strOldITEMCODE
	with frmThis
	    '외주항목코드 선택 버튼 부분
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			strOldITEMCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row)
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)
				
				'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
				If vntRet(7,0) = "Y" Then
				
					'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",Row,Row,false
					'수량,단가,금액은 0처리
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",Row,Row,false
					'일반형태로 변경
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
			End IF
			.txtSUSURATE.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			
		'CF견적의 부문을 선택시 버튼처리....	
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DETAIL_BTN") Then
			'상세견적 저장 및 조회 팝업 호출
			
			
			'*******************************************1.   간접비 부문 처리와 그외 부문 처리로 분기처리!
			
				
				strPREESTNO = "9999999999"		
				strGBN="F"
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) = "" Then
					strITEMCODESEQ = 0
				Else
					strITEMCODESEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row)
				End If
				
				dblAMT = mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",Row)
				vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",Row),strPREESTNO,strGBN,strITEMCODESEQ,"F",Trim(.txtSEQ.value))
				vntRet = gShowModalWindow("PDCMESTTYPE_SUBITEM.aspx",vntInParams , 1149,650)
				
				mlngTempRowCnt=clng(0): mlngTempColCnt=clng(0)
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) <> "" Then
					vntData = mobjPDCOESTTYPE.SelectRtn_DtlSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row),.txtSEQ.value)
				Else
					vntData = mobjPDCOESTTYPE.SelectRtn_TempSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row))
				End If
				
				
				If mlngTempRowCnt > 0 Then
					strRTN = Cstr(vntData(0,1))
				Else
					strRTN = "0"
				End If
				
					
				If strRTN <> Cstr(dblAMT) Or strRTN = 0 Then
					'값을 받아오는 것이 아니라, PD_SUBITEM_DTL 또는 PD_SUBITEM_INPUT 에서 가져온다.
					
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, strRTN
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, strRTN
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
					If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
						dblSUSUAMT = strRTN * .txtSUSURATE.value * 0.01
						mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
					End If
					BUDGET_AMT_SUM
					'팝업을 눌렀다면 반드시 Y 로바뀌도록 변경 (추후 자동저장을 생각해보자!!!!!)
					mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
				End IF
				
				
				
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
				
				mobjSCGLSpr.CellChanged .sprSht, SAVEFLAG,Row
			End If
	End with
End Sub

Sub sprSht1_Keyup(KeyCode, Shift)
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
	
	'키가 움질일때 바인딩
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprSht1_Click frmThis.sprSht1.ActiveCol,frmThis.sprSht1.ActiveRow
	End If
End Sub

Sub sprSht1_Click(ByVal Col, ByVal Row)
	Dim intcnt,intCnt2
	Dim strHDRSEQ
	with frmThis
	
		'쉬트바인딩 프로젝트-JOB 
		strHDRSEQ = mobjSCGLSpr.GetTextBinding( .sprSht1,"SEQ",.sprSht1.ActiveRow)
		IF strHDRSEQ <> "" Then
			SelectRtn_REAL(strHDRSEQ)
		End If
	end with

End Sub


'------------------------------------------
' 데이터 저장
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim strPREESTNO
	Dim intCnt
	Dim intRtnDtl
	Dim vntData
	Dim strAGREEYEARMON
	Dim strJOBNO
	Dim intSearchRtn
	Dim strCHKCONFIRM
	'헤더와 디테일을 연결
	Dim strSEQ
	strSEQ = ""
	Dim strPRODUCTIONCHK
	
	if DataValidation =false then exit sub
	
	strMasterData = gXMLGetBindingData (xmlBind)
	
	with frmThis
		'저장의 공통사항 처리
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|PREESTNO|PRINT_SEQ|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN|IMESEQ|SAVEFLAG|HDRSEQ")		
		
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
			Else
				gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				exit sub
			End If
		End If
			
			strSEQ = .txtSEQ.value 
			intRtn = mobjPDCOESTTYPE.ProcessRtn_ESTTYPE(gstrConfigXml,strMasterData,vntData,strSEQ)
				
			if not gDoErrorRtn ("ProcessRtn_ESTTYPE") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
				
				SelectRtn_ProcessRtn(strSEQ)
				
			End If
		
	End with
End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	
	'On error resume next
	with frmThis
  		'Field 필수 입력 항목 검사
  		If .txtTYPENAME.value = "" Then
			gErrorMsgBox "견적타입명을 입력하십시오.","저장안내"
			Exit Function
		End If
	
		
		'Sheet 필수 입력 항목 검사 
		If .sprSht.MaxRows = 0 Then
				gErrorMsgBox "저장할 상세 내역이 존재 하지 않습니다.","저장안내"
				Exit Function
		End IF
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		for intCnt = 1 to .sprSht.MaxRows
   		'DIVNAME|CLASSNAME|ITEMCODE,ITEMCODENAME
			if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNAME",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNAME",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODENAME",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 외주항목 내용 을 확인하십시오","입력오류"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function
'헤더내역 삭제
Sub DeleteRtnAll
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2,lngCnt
	Dim strHDRSEQ
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)
		IF gDoErrorRtn ("DeleteRtnAll") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		lngCnt =0
		
		for i = intSelCnt-1 to 0 step -1
		
				strHDRSEQ = mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))
				
				intRtn = mobjPDCOESTTYPE.DeleteRtn_ALL(gstrConfigXml,strHDRSEQ)
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				
   			End IF
		next
		'1건이라도 삭제건이 있다면 메세지 출력
		If lngCnt <> 0 Then
			gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
		End If
		SelectRtn
	End with
End Sub
'부문 내역 삭제
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2,lngCnt
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
	Dim strITEMCODE
	Dim strDETAILYNFLAG
	Dim intCnt_Prod
	Dim intCnt_ProdChk
	Dim strPRODUCTSUSUCHK
	Dim strHDRSEQ
	
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		
		
		lngCnt =0
		intRtn2 = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)) <> ""  Then
		
				strHDRSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"HDRSEQ",vntData(i))
				strITEMCODESEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)))
				strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",vntData(i))
				strDETAILYNFLAG = mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",vntData(i))
				'마지막 체크 변경
				intRtn2 = mobjPDCOESTTYPE.DeleteRtn_EST(gstrConfigXml,strHDRSEQ, strITEMCODESEQ, strITEMCODE,strDETAILYNFLAG)
			Else
				
				strHDRSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"HDRSEQ",vntData(i))
				
				
				strITEMCODESEQ = 0
				strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",vntData(i))
				strDETAILYNFLAG = mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",vntData(i))
				
				intRtn2 = mobjPDCOESTTYPE.DeleteRtn_ESTTempDel(gstrConfigXml,strHDRSEQ, strITEMCODESEQ, strITEMCODE,strDETAILYNFLAG)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				gWriteText "", "[" & strITEMCODESEQ & "] 자료가 삭제되었습니다."
   			End IF
		next
		'헤더재계산
		Call SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		'저장되어있는 값이 있으면 DB 에 헤더재계산 값을 저장 
		If intRtn2 = 0 Then
   		Else
   		
			DelProc
		End If
		'1건이라도 삭제건이 있다면 메세지 출력
		If lngCnt <> 0 Then
			gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
		End If
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	End with
	err.clear
End Sub

Sub DelProc
	Dim intHDR
	Dim strMasterData
	Dim strSEQ
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis	
		strSEQ = .txtSEQ.value
		intHDR = mobjPDCOESTTYPE.ProcessRtn_DelProc(gstrConfigXml,strMasterData,strSEQ,"UPDATE")
				if not gDoErrorRtn ("ProcessRtn_DelProc") then
					SelectRtn_ProcessRtn(strSEQ)
				End If
	End with
End Sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" HEIGHT="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD valign="top">
						<!--1단-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="115" background="../../../images/back_p.gIF"
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
											<td class="TITLE">견적 TYPE 유형관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
					</TD>
				</TR>
				<tr>
					<td>
						<TABLE cellSpacing="0" cellPadding="0" width="1050" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
					</td>
				</tr>
				<tr>
					<td height="100%" valign="top">
						<TABLE id="tblBody" style="WIDTH: 100%;HEIGHT: 90%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="100">견적유형명
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 209px"><INPUT class="INPUT_L" id="txtSEARCHTYPENAME" style="WIDTH: 208px; HEIGHT: 22px" type="text"
														size="29"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 107px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1,txtCLIENTCODE1)"
													width="107">광고주
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 328px"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 240px; HEIGHT: 22px"
														type="text" maxLength="255" align="left" size="34"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5"></TD>
												<TD align="right"><INPUT dataFld="SEQ" class="INPUT_R" id="txtSEQ" style="WIDTH: 8px; HEIGHT: 22px" accessKey="NUM"
														dataSrc="#xmlBind" type="text" size="1" name="txtSEQ"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
														width="54" border="0" name="imgQuery"><IMG id="imgTypeNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다." src="../../../images/imgNew.gIF"
														border="0" name="imgTypeNew"><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" height="20" alt="자료를 삭제합니다."
														src="../../../images/imgDelete.gIF" border="0" name="imgDelete"><IMG id="imgExcel1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"
														width="54" border="0" name="imgExcel1"></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR vAlign="top" align="left">
								<TD vAlign="top" height="25%">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; TOP: 0px; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht1" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: relative; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht1" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="3254">
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
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="100">견적유형명
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 209px"><INPUT dataFld="TYPENAME" class="INPUT_L" id="txtTYPENAME" style="WIDTH: 208px; HEIGHT: 22px"
														type="text" size="29" dataSrc="#xmlBind" name="txtTYPENAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 107px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
													width="107">광고주
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 211px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 144px; HEIGHT: 22px"
														type="text" maxLength="255" align="left" size="18" name="txtCLIENTNAME" dataFld="CLIENTNAME" dataSrc="#xmlBind">
													<IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
														src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE"><INPUT dataFld="CLIENTCODE" class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 48px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="6" align="left" size="2" name="txtCLIENTCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField('','')">
												</TD>
												<TD align="right" colspan="2"><input id="txtPRINT_SEQ" style="VISIBILITY: hidden; WIDTH: 5px" type="text" maxLength="2"
														value="1" name="txtPRINT_SEQ" accessKey="NUM,"><IMG id="imgTableUp" style="CURSOR: hand" alt="자료를 올립니다." src="../../../images/imgTableUp.gif"
														align="absMiddle" border="0" name="imgTableUp"> <IMG id="imgTableDown" style="CURSOR: hand" alt="자료를 내립니다." src="../../../images/imgTableDown.gif"
														align="absMiddle" border="0" name="imgTableDown">  <IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" height="20" alt="자료입력을 위해 행을추가합니다." src="../../../images/imgRowAdd.gIF"
														align="absMiddle" border="0" name="imgRowAdd"><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF"
														border="0" name="imgRowDel" align="absMiddle"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" border="0"
														name="imgSave" align="absMiddle"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"
														width="54" border="0" name="imgExcel" align="absMiddle"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" width="100">COMMISSION
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 209px"><INPUT dataFld="COMMISSION" class="NOINPUT_R" id="txtCOMMISSION" style="WIDTH: 208px; HEIGHT: 22px; TEXT-ALIGN: right"
														type="text" size="29" name="txtCOMMISSION" accessKey="NUM" dataSrc="#xmlBind" readOnly></TD>
												<TD class="SEARCHLABEL" width="107" style="WIDTH: 107px">NONECOMMISSION
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 211px"><INPUT dataFld="NONECOMMISSION" class="NOINPUT_R" id="txtNONECOMMISSION" style="WIDTH: 208px; HEIGHT: 22px; TEXT-ALIGN: right"
														type="text" size="29" name="txtNONECOMMISSION" dataSrc="#xmlBind" readOnly accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" width="70" style="WIDTH: 70px">금액</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" style="WIDTH: 208px; HEIGHT: 22px; TEXT-ALIGN: right"
														type="text" size="29" name="txtAMT" dataSrc="#xmlBind" readOnly accessKey="NUM"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" width="100">수수료비율
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 209px"><INPUT dataFld="SUSURATE" class="INPUT_R" id="txtSUSURATE" style="WIDTH: 208px; HEIGHT: 22px"
														type="text" size="29" name="txtSUSURATE" dataSrc="#xmlBind"></TD>
												<TD class="SEARCHLABEL" width="107" style="WIDTH: 107px">수수료
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 211px"><INPUT dataFld="SUSUAMT" class="INPUT_R" id="txtSUSUAMT" style="WIDTH: 208px; HEIGHT: 22px"
														type="text" size="29" name="txtSUSUAMT" dataSrc="#xmlBind"></TD>
												<TD class="SEARCHLABEL" width="70" style="WIDTH: 70px">합계금액</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="SUMAMT" class="NOINPUT_R" id="txtSUMAMT" style="WIDTH: 208px; HEIGHT: 22px; TEXT-ALIGN: right"
														type="text" size="29" name="txtSUMAMT" dataSrc="#xmlBind" readOnly accessKey="NUM"></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR vAlign="top" align="left">
								<TD vAlign="top" height="60%">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: relative; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="7646">
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
							<TR vAlign="top" align="left">
								<TD vAlign="top" height="0%">
									<DIV id="pnlTab3" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_copy" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: relative; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht_copy" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="7646">
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
								<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
						<!--2단끝--></td>
				</tr>
			</TABLE>
		</FORM>
	</body>
</HTML>
