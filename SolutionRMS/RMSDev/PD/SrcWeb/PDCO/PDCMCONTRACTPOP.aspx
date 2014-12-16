<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACTPOP.aspx.vb" Inherits="PD.PDCMCONTRACTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>계약서 등록 및 확정</title>
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMCONTRACT, mobjPDCMGET
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
mALLCHECK = TRUE
mstrCheck=True
'=========================================================================================
' 이벤트 프로시져 
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
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
	
End Sub

Sub imgNew_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
Dim intCnt
Dim lngCnt
Dim lngSumCnt
	with frmThis
	
	lngCnt = 0
	lngSumCnt = 0
	For intCnt = 1 To .sprSht.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1"  Then
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CONFIRMFLAG", intCnt) = "Y" Then
				gErrorMsgBox intCnt & " 행은 검수 확인 내용입니다.확인내역을 변경하여주십시오.","삭제안내"
				Exit Sub
			End If
			lngCnt = 1
			lngSumCnt = lngSumCnt + lngCnt
		End if
	Next
	If lngSumCnt = 0 Then
		gErrorMsgBox "선택된 데이터가 없습니다.","삭제안내"
		Exit Sub
	End If
	End with
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
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	Dim intCnt2
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht1.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
	For intCnt2 = 1 To frmThis.sprSht1.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXNO",intCnt2) <> "" THEN
			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSNO",intCnt2) & " 에 대하여" &vbcrlf & "세금계산서번호가 존재하는 내역은 재출력할 수 없습니다.","인쇄안내!"
			Exit Sub
		End If
	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다.
		'md_trans_temp삭제 시작
		intRtn = mobjPDCMCONTRACT.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		
		vntData = mobjPDCMCONTRACT.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
	
		strcntsum = 0
		IF not gDoErrorRtn ("Get_CATVTRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt
			
			for i=1 to 3
				strUSERID = ""
				vntDataTemp = mobjPDCMCONTRACT.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
				
		window.setTimeout "printSetTimeout", 10000
	
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMCATVTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'청구일 조회조건 생성
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

Sub STEDClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtSTDATE.value = date1
		.txtEDDATE.value = date2
	End With
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

Sub imgSTDATE_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtSTDATE,.imgSTDATE,"txtSTDATE_onchange()"
		gSetChange
	end with
End Sub

Sub txtSTDATE_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value  = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtCONTRACTDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value  = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgEDDATE_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtEDDATE,.imgEDDATE,"txtEDDATE_onchange()"
		gSetChange
	end with
End Sub

Sub txtEDDATE_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgDELIVERYDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtDELIVERYDAY,.imgDELIVERYDAY,"txtDELIVERYDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtDELIVERYDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVERYDAY",frmThis.sprSht.ActiveRow, frmThis.txtDELIVERYDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgTESTDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtTESTDAY,.imgTESTDAY,"txtTESTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtTESTDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTDAY",frmThis.sprSht.ActiveRow, frmThis.txtTESTDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtLOCALAREA,txtAMT,txtTESTMENT,txtCOMENT

Sub txtLOCALAREA_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"LOCALAREA",frmThis.sprSht.ActiveRow, frmThis.txtLOCALAREA.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtAMT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
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

Sub txtTESTMENT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTMENT",frmThis.sprSht.ActiveRow, frmThis.txtTESTMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtCOMENT_Onchange

	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMENT",frmThis.sprSht.ActiveRow, frmThis.txtCOMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtPAYMENTGBN_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PAYMENTGBN",frmThis.sprSht.ActiveRow, frmThis.txtPAYMENTGBN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbENDGBN_onchange
'txtCONTRACTNO,cmbTEST
	with frmThis
		If .cmbENDGBN.value = "T" Then
			.txtCONTRACTNO.style.visibility = "visible"
			.cmbTEST.style.visibility = "visible"
			.txtJOBNO.style.visibility = "hidden"
			.txtJOBNAME.style.visibility = "hidden"
			.ImgJOBNO.style.visibility = "hidden"
		Elseif  .cmbENDGBN.value = "F" Then
			.txtCONTRACTNO.style.visibility = "hidden"
			.cmbTEST.style.visibility = "hidden"
			.txtJOBNO.style.visibility = "visible"
			.txtJOBNAME.style.visibility = "visible"
			.ImgJOBNO.style.visibility = "visible"
		Elseif  .cmbENDGBN.value = "" Then
			.txtCONTRACTNO.style.visibility = "visible"
			.cmbTEST.style.visibility = "hidden"
			.txtJOBNO.style.visibility = "visible"
			.txtJOBNAME.style.visibility = "visible"
			.ImgJOBNO.style.visibility = "visible"
		End If
	End with
	SelectRtn
End Sub
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i


	'서버업무객체 생성	
	set mobjPDCMCONTRACT	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "380px"
	pnlTab1.style.left= "7px"
	
	

	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    Input_Layout
	pnlTab1.style.visibility = "visible"
	frmThis.txtCONTRACTNO.style.visibility = "hidden"
	'화면 초기값 설정
	InitPageData	
	
	'이곳에 파라미터 받기
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	with frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtCONTRACTNO.value = vntInParam(i)	'CC Code or Name
			end select
		next
	.cmbENDGBN.selectedIndex = 0 
	cmbENDGBN_onchange
	
	
	Call gCleanField("txtFrom","txtTo")
	End with
	 
	SelectRtn
	
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

Sub Input_Layout
	gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'거래명세서 생성 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONTRACTNO|OUTSCODE|OUTSNAME|JOBNO|JOBNAME|ADJAMT|JOBGUBN|CREPART|RANKTRANS|SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|계약서번호|외주처코드|외주처|JOBNO|JOB명|금액|제작부문|제작분류|랭크|순번"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6  |10        |0         |30    |12   |30   |14  |15      |15      |0   |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " OUTSCODE|OUTSNAME|JOBNO|JOBNAME|RANKTRANS|JOBGUBN|CREPART|CONTRACTNO", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"OUTSCODE|OUTSNAME|JOBNO|JOBNAME|JOBGUBN|CREPART|ADJAMT"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE|RANKTRANS|SEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|OUTSCODE|JOBGUBN|CREPART|CHK",-1,-1,2,2,false
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|ITEMNAME",-1,-1,0,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSNAME|CONTRACTNO"
		.txtOUTSCODE.style.visibility = "hidden"
	    		
    End With    
End Sub

Sub Select_Layout
	Dim strComboList
	gSetSheetDefaultColor() 
	With frmThis
		strComboList =  "계약서 미확인" & vbTab & "계약서 확인"
		'******************************************************************
		'거래명세서 생성 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONTRACTNO|CONTRACTNAME|CONTRACTDAY|LOCALAREA|STDATE|EDDATE|AMT|DELIVERYDAY|TESTDAY|PAYMENTGBN|TESTMENT|COMENT|OUTSCODE|CONFIRMFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		"선택|계약서번호|계약명|계약일|납품장소|용역시작일|용역종료일|계약금액|납품일|검수일|대금지급방법|검수결과|특약사항|외주처코드|계약서확인"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "6|10        |18    |8     |13      |10        |10        |12      |9     |9     |9           |9       |10     |0         |13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "LOCALAREA|PAYMENTGBN|TESTMENT|COMENT|CONFIRMFLAG|CONTRACTNO", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "STDATE|EDDATE|DELIVERYDAY|TESTDAY|CONTRACTDAY"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DELIVERYDAY|TESTDAY|CONTRACTDAY"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE", true
		mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO|TESTDAY|PAYMENTGBN", false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK",-1,-1,2,2,false
	    mobjSCGLSpr.SetCellAlign2 .sprSht, "CONTRACTNAME",-1,-1,0,2,false
		'mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSNAME"
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,15,15,-1,-1,strComboList
		mobjSCGLSpr.CellGroupingEach .sprSht,"CONTRACTNAME|LOCALAREA",,,,0
		
	    		
    End With    
End Sub
'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		DateClean
		STEDClean
		.txtDELIVERYDAY.value = gNowDate
		.txtTESTDAY.value = gNowDate
		.txtCONTRACTDAY.value = gNowDate
		.txtLOCALAREA.value = "에스케이 플래닛(주) 본사"
		'.txtCONTRACTNO.style.visibility = "hidden"
		.cmbTEST.style.visibility = "hidden"
		.txtTESTMENT.value  = ""
		.txtCOMENT.value  = ""
		.txtPAYMENTGBN.value = ""
		.txtAMT.value  = 0
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub
'****************************************************************************************
' 이벤트 처리
'****************************************************************************************
Sub sprSht_Change(ByVal Col, ByVal Row)
	
	Dim intCnt
	Dim lngAMT
	Dim lngSUMAMT
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	if Col = 1 Then
		lngAMT = 0
		lngSUMAMT = 0
		
		For intCnt = 1 To frmThis.sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1" And frmThis.cmbENDGBN.value = "F" Then
				lngAMT = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJAMT", intCnt)		
				lngSUMAMT = lngSUMAMT + lngAMT
			End if
		Next
		frmThis.txtAMT.value = lngSUMAMT
		txtAMT_onblur
	End if
End Sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	dim intcnt
	with frmThis
		if .cmbENDGBN.value = "" then
			exit Sub
		End if
		If Row = 0 and Col = 1  then 
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
				
			next
			For intCnt = 1 To .sprSht.MaxRows
				If  .cmbENDGBN.value = "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If			
			Next
		Elseif Row > 0 and Col > 0 then
			If .cmbENDGBN.value  = "T" Then
			sprShtToFieldBinding Col,Row
			End IF
		end if
		If .cmbENDGBN.value = "F" then
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		End If
	end with
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '그리드 데이터가 없으면 나간다.
			.txtLOCALAREA.value = mobjSCGLSpr.GetTextBinding(.sprSht,"LOCALAREA",Row)
			.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
			.txtEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			.txtDELIVERYDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYDAY",Row)
			.txtTESTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTDAY",Row)
			.txtPAYMENTGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PAYMENTGBN",Row)
			.txtTESTMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTMENT",Row)
			.txtCOMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMENT",Row)
			.txtOUTSCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",Row)
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
			.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
		If .txtAMT.value <> "" Then
			txtAMT_onblur
		End If
	End with
End Function
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
'-----------------------------------------------------------------------------------------
' 외주처 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub imgOUTSCODE_onclick
	Call SEARCHOUT_POP()
End Sub

'실제 데이터List 가져오기
Sub SEARCHOUT_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtOUTSCODE1.value), trim(.txtOUTSNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtOUTSCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtOUTSNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE1.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE1.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' JOB 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'실제 데이터List 가져오기
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtJOBNO.value = trim(vntRet(0,0))  ' Code값 저장
			.txtJOBNAME.value = trim(vntRet(1,0))  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
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
'-----------------------------------------------------------------------------------------
' 데이터조회
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strGBN
	Dim strOUTSCODE
	Dim strOUTSNAME
	Dim strFROM
	Dim strTO
	Dim strJOBNO
	Dim strJOBNAME
	Dim vntData
	Dim intCnt
	Dim strCONFIRM
	Dim strCONTRACTNO
	'On error resume next
	with frmThis
		.sprSht.MaxRows = 0
		strGBN = .cmbENDGBN.value 
		strOUTSCODE = TRIM(.txtOUTSCODE1.value)
		strOUTSNAME =  TRIM(.txtOUTSNAME.value)
		strJOBNO = TRIM(.txtJOBNO.value)
		strJOBNAME =  TRIM(.txtJOBNAME.value)
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strCONTRACTNO = .txtCONTRACTNO.value 
		
		If Len(strCONTRACTNO) = 10 Then
			strCONTRACTNO = MID(strCONTRACTNO,1,7) & "-" & MID(strCONTRACTNO,8,3)
		End if
		
		strCONFIRM = .cmbTEST.value
		
		
		
		IF strGBN = "F" THEN  '미완료조회
		 Call Input_Layout()
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strJOBNO,strJOBNAME)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE		
   					If mlngRowCnt > 0 Then
   					
   						For intCnt = 1 To .sprSht.MaxRows
								If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
								End If
						Next	
						initpageData
   					Else
   						.sprSht.MaxRows = 0
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO", true
   					.imgDelete.disabled = true
   					.imgSave.disabled = false
   			end if
		ELSEIF strGBN = "T" THEN  '완료조회
			Call Input_Layout()
			Call Select_Layout()
		
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn_EXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strCONFIRM,strCONTRACTNO)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE			
   					If mlngRowCnt > 0 Then	
   						sprShtToFieldBinding 1,1
   					Else
   						.sprSht.MaxRows = 0
   						
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNAME", false	
   					.imgDelete.disabled = false
   					.imgSave.disabled = false
   			end if
   		ELSEIF strGBN = "" THEN  '전체조회
   			Call Input_Layout()
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strJOBNO,strJOBNAME,strCONTRACTNO)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE		
   					If mlngRowCnt > 0 Then
   					
   						For intCnt = 1 To .sprSht.MaxRows
								If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
								End If
						Next	
						initpageData
   					Else
   						.sprSht.MaxRows = 0
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO", false	
   					.imgDelete.disabled = true
   					.imgSave.disabled = true
   			end if
		END IF
   	end with
End Sub
'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strGUBN
	Dim strOUTSCODE
	Dim strCONTRACTNAME
	Dim strSAVEFLAG
	Dim strCOMENT 
	Dim intCnt2
	Dim lngCNTSUM
	Dim lngCNT
	
	strMasterData = gXMLGetBindingData (xmlBind)
		with frmThis
		
		If .cmbENDGBN.value  = "F" Then
			strSAVEFLAG = "F"
		Elseif .cmbENDGBN.value = "T" Then
			strSAVEFLAG = "T"
		End If
		
		'txtCONTRACTNAME,txtCONTRACTDAY
		
		If .sprSht.MaxRows = 0 Then
				gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
				Exit Sub
		End IF
		lngCNTSUM = 0
		lngCNT = 0
	
		For intCnt2 = 1 To .sprSht.MaxRows
			lngCNT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2)
			lngCNTSUM = lngCNTSUM + lngCNT
		Next
		If lngCNTSUM = 0 Then
			gErrorMsgBox "선택되어진 자료가 없습니다.","저장안내"
			Exit Sub
		End if
		
		If .cmbENDGBN.value ="F" Then
			if DataValidation =false then exit sub
		End If
		If strSAVEFLAG = "F" Then
			If .txtCONTRACTNAME.value = "" Then
				gErrorMsgBox "계약명을 넣어주십시오.","저장안내"
				Exit Sub
			End If
			If .txtCONTRACTDAY.value = "" Then
				gErrorMsgBox "계약일을 넣어주십시오.","저장안내"
				Exit Sub
			End If
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|JOBNO|SEQ")
		Elseif strSAVEFLAG = "T" then
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|CONTRACTNO|CONTRACTNAME|LOCALAREA|STDATE|EDDATE|AMT|DELIVERYDAY|TESTDAY|PAYMENTGBN|TESTMENT|COMENT|OUTSCODE|CONFIRMFLAG|CONTRACTDAY")
		End If
		
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
				gErrorMsgBox "선택된 " & meNO_DATA,"저장안내"
				exit Sub
			Else
				gErrorMsgBox "변경된 입력필드 " & meNO_DATA,"저장안내"
				exit sub
			End If
		End If
		strGUBN = ""
		If strSAVEFLAG = "F" then
			For intCnt = 1 to .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
					strGUBN = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt)
					strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
					'strCONTRACTNAME =  mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt)
				End If
				If strGUBN <> "" AND strOUTSCODE <> "" AND strCONTRACTNAME <> "" Then
					Exit For
				End If
			Next
			strCONTRACTNAME = .txtCONTRACTNAME.value 
			
			strGUBN = MID(strGUBN,1,1)
			
			If strGUBN = "" Then
				gErrorMsgBox "선택되어진JOB번호가 없습니다.","저장안내"
				Exit Sub
			End If
		End If
		strCOMENT = .txtCOMENT.value 
	
		intRtn = mobjPDCMCONTRACT.ProcessRtn(gstrConfigXml,strMasterData,vntData,strGUBN,strOUTSCODE,strCONTRACTNAME,strSAVEFLAG,strCOMENT )
			if not gDoErrorRtn ("ProcessRtn") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
				SelectRtn
			End If
		End with
End Sub

Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim strOUTSCODE
   	Dim lngCnt
   	Dim strSTDINT
	'On error resume next
	with frmThis
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		strSTDINT = ""
   		for intCnt = 1 To .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
   				strSTDINT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
   				
   				If strSTDINT <> ""  Then
   					Exit For
   				End If
   			End if
   		Next
  
   		for intCnt = 1 to .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
				if strSTDINT <> mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) Then
					gErrorMsgBox intCnt & " 번째 행의 외주처를 확인하십시오." & vbcrlf & "단일외주처 일경우에만 저장이 가능합니다.","입력오류"
					Exit Function
				End If
			End If
		next
   	
   	End with
	DataValidation = true
End Function
'자료삭제
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
	Dim strRow
	Dim strCONTRACTNO
	with frmThis
	
		
		
		'선택된 자료를 끝에서 부터 삭제
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		for i = .sprSht.MaxRows to 1 step -1
		
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "Y" Then
				gErrorMsgBox "확정견적은 삭제하실수 없으며, 상세내역에서 확정을 취소후 삭제하십시오.","삭제안내"
				Exit Sub
			End if
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
				
					strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					intRtn = mobjPDCMCONTRACT.DeleteRtn(gstrConfigXml,strCONTRACTNO)
				End IF
				
   			End If
   			IF not gDoErrorRtn ("DeleteRtn") then
					mobjSCGLSpr.DeleteRow .sprSht,i
					
   			End IF
		next
		gWriteText lblstatus, "자료가 " & intRtn & " 건 삭제되었습니다."
		'선택 블럭을 해제
		'mobjSCGLSpr.DeselectBlock .sprSht
		'strRow = .sprSht.ActiveRow
		SelectRtn
		'mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		'Call sprSht_Click(1,strRow)
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
			<TABLE id="tblForm" style="WIDTH: 100%" HEIGHT="100%" cellSpacing="0" cellPadding="0" border="0">
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
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;계약관리</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<TABLE id="tblButton1"  cellSpacing="0" cellPadding="2"
											 border="0"ALIGN="right">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
														height="20" alt="화면을 닫습니다." src="../../../images/imgClose.gIF" border="0" name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="83">기간</TD>
									<TD class="SEARCHDATA" style="WIDTH: 249px"><INPUT class="INPUT" id="txtFrom" title="계약검색 시작일자" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtFrom"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="계약검색 종료일자" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtTo"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgTo">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand">완료구분</TD>
									<TD class="SEARCHDATA" style="WIDTH: 135px; CURSOR: hand"><SELECT id="cmbENDGBN" style="WIDTH: 128px" name="cmbENDGBN">
											<OPTION value="">전체</OPTION>
											<OPTION value="F" selected>미완료</OPTION>
											<OPTION value="T">완료</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE1)">외주처</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtOUTSNAME" title="외주처명 조회" style="WIDTH: 224px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtOUTSNAME"><IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtOUTSCODE1" title="외주처코드조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtOUTSCODE1"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')"
										width="83">계약서번호</TD>
									<TD class="SEARCHDATA" style="WIDTH: 249px"><INPUT class="INPUT_L" id="txtCONTRACTNO" title="JOB명 조회" style="WIDTH: 240px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand">계약서확인</TD>
									<TD class="SEARCHDATA" style="WIDTH: 135px; CURSOR: hand"><SELECT id="cmbTEST" style="WIDTH: 128px" name="cmbTEST">
											<OPTION value="" selected>전체</OPTION>
											<OPTION value="계약서 미확인">계약서 미확인</OPTION>
											<OPTION value="계약서 확인">계약서 확인</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)">JOB명</TD>
									<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtJOBNAME" title="JOB명 조회" style="WIDTH: 224px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtJOBNO" title="JOBNO 조회" style="WIDTH: 65px; HEIGHT: 22px" type="text"
											maxLength="7" align="left" size="3" name="txtJOBNO"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left"  height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;계약서 등록 및 확정</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD>
												<!--		
												<td><IMG id="imgTestOK" onmouseover="JavaScript:this.src='../../../images/imgTestOKOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTestOK.gIF'"
														height="20" alt="검수를 확인처리 합니다." src="../../../images/imgTestOK.gIF" border="0" name="imgTestOK"></td>
												<td><IMG id="imgTestCancel" onmouseover="JavaScript:this.src='../../../images/imgTestCancelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTestCancel.gIF'"
														height="20" alt="검수를 취소합니다." src="../../../images/imgTestCancel.gIF" border="0" name="imgTestCancel"></td>-->
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							
							
							<!---->
							<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
							
							
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="LEFT" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')"
													width="85"><FONT face="굴림">계약명</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px" width="251"></FONT><INPUT dataFld="CONTRACTNAME" id="txtCONTRACTNAME" style="WIDTH: 240px; HEIGHT: 21px" accessKey="M"
														dataSrc="#xmlBind" type="text" size="33" name="txtCONTRACTNAME" title="계약명"></TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTDAY,'')"
													width="89"><FONT face="굴림">계약일</FONT></TD>
												<TD class="DATA" width="257"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="계약일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="M,DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" alt="ImgCONTRACTDAY" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="ImgCONTRACTDAY"></TD>
												<TD class="LABEL" width="90" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtLOCALAREA,'')"><FONT face="굴림">납품장소</FONT></TD>
												<TD class="DATA" width="257"><FONT face="굴림"><INPUT dataFld="LOCALAREA" class="INPUT_L" id="txtLOCALAREA" title="납품장소" style="WIDTH: 251px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="36" name="txtLOCALAREA"></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDELIVERYDAY, '')"><FONT face="굴림">납품일</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px"><FONT face="굴림"></FONT><INPUT dataFld="DELIVERYDAY" class="INPUT" id="txtDELIVERYDAY" title="납품일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDELIVERYDAY"><IMG id="imgDELIVERYDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgDELIVERYDAY">&nbsp;
													<INPUT dataFld="OUTSCODE" id="txtOUTSCODE" title="외주처코드_숨김" style="WIDTH: 121px; HEIGHT: 21px"
														dataSrc="#xmlBind" type="text" size="14" name="txtOUTSCODE">
												</TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSTDATE, txtEDDATE)"><FONT face="굴림"><FONT face="굴림">용역기간</FONT></FONT></TD>
												<TD class="DATA"></FONT><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="용역기간 시작일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtSTDATE"><IMG id="imgSTDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgSTDATE">&nbsp;~ <INPUT dataFld="EDDATE" class="INPUT" id="txtEDDATE" title="용역기간 종료일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtEDDATE"><IMG id="imgEDDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgEDDATE"></FONT></TD>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAMT, '')"><FONT face="굴림"><FONT face="굴림">계약금액</FONT></FONT></TD>
												<TD class="DATA"></FONT><FONT face="굴림"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="계약금액" style="WIDTH: 251px; HEIGHT: 22px"
															accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtAMT"></FONT></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTESTMENT, '')"><FONT face="굴림">검수결과</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px"><FONT face="굴림"><INPUT dataFld="TESTMENT" class="INPUT_L" id="txtTESTMENT" title="검수결과" style="WIDTH: 240px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" maxLength="255" size="35" name="txtTESTMENT"></FONT>
												</TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTESTDAY, '')"><FONT face="굴림"><FONT face="굴림">검수일</FONT></FONT></TD>
												<TD class="DATA"><INPUT dataFld="TESTDAY" class="INPUT" id="txtTESTDAY" title="검수일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtTESTDAY"><IMG id="imgTESTDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgTESTDAY"></TD>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPAYMENTGBN,'')"><FONT face="굴림">대금지급방법</FONT></TD>
												<TD class="DATA"><INPUT dataFld="PAYMENTGBN" class="INPUT_L" id="txtPAYMENTGBN" title="대금지급방법" style="WIDTH: 251px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="255" size="37" name="txtPAYMENTGBN"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand; HEIGHT: 130px" onclick="vbscript:Call gCleanField(txtCOMENT,'')"><FONT face="굴림">특약사항</FONT></TD>
												<TD class="DATA" colSpan="5"><TEXTAREA dataFld="COMENT" id="txtCOMENT" style="WIDTH: 952px" dataSrc="#xmlBind" name="txtCOMENT"
														rows="8" wrap="hard" cols="116"></TEXTAREA></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="11642">
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
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
