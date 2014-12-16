<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVCOMMI.aspx.vb" Inherits="MD.MDCMCATVCOMMI" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CATV 수수료 거래명세표 발행</title>
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
Dim mobjMDCMCATVCOMMI, mobjMDCMGET
Dim mstrCheck
Dim mALLCHECK
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

Sub imgClose_onclick ()
	Window_OnUnload
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
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDCMCATVCOMMI.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMCATVCOMMI_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		vntData	= mobjMDCMCATVCOMMI.Get_CATVCOMMI_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_CATVCOMMI_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum
			for i=1 to 2
				strUSERID = ""
				vntDataTemp = mobjMDCMCATVCOMMI.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "printSetTimeout", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMCATVCOMMI.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		If .chkSPONSOR.checked= TRUE Then
			strSPONSOR = "Y"
		else
			strSPONSOR = ""
		end if
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value, "commi","CATV", strSPONSOR)
		vntRet = gShowModalWindow("../MDCO/MDCMCOMMIREALMEDPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = vntRet(1,0)		        ' Code값 저장
			.txtREAL_MED_NAME1.value = vntRet(2,0)             ' 코드명 표시
			IF vntRet(3,0) = "완료" THEN
				window.event.keyCode = meEnter
				txtTRANSNO_onkeydown
			ELSE
				.txtTRANSNO.value = ""
			END IF
			gSetChangeFlag .txtREAL_MED_CODE                ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strSPONSOR
   		
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			If .chkSPONSOR.checked= TRUE Then
				strSPONSOR = "Y"
			else
				strSPONSOR = ""
			end if
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtTRANSNO.value,.txtREAL_MED_NAME1.value,"ALL","commi", "CATV", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = vntData(0,0)
					.txtREAL_MED_NAME1.value = vntData(1,0)
				Else
					Call REAL_MED_CODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 거래처번호팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgTRU_onclick
	Call TRU_POP()
End Sub

Sub txtTRANSNO_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strTRANSYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strTRANSYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value, .txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value, "commi", "CATV", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)  ' Code값 저장
					.txtTRANSNO.value = vntData(1,0)  ' 코드명 표시
					.txtREAL_MED_CODE.value = vntData(2,0)  ' 코드명 표시
					.txtREAL_MED_NAME1.value = vntData(3,0)  ' 코드명 표시
					'Call SelectRtn ()
				Else
					Call TRU_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRU_POP
	dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
			strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
	
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value, .txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value, "commi", "CATV") '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code값 저장
			.txtTRANSNO.value = vntRet(1,0)  ' 코드명 표시
			.txtREAL_MED_CODE.value = vntRet(2,0)  ' 코드명 표시
			.txtREAL_MED_NAME1.value = vntRet(3,0)  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'단가
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'금액
Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

'수수료
Sub txtSUMAMTVAT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMTVAT,0,true)
	end with
End Sub

Sub txtTRANSYEARMON_onblur
	With frmThis
		if .txtTRANSNO.value ="" then
			If .txtTRANSYEARMON.value <> "" AND Len(.txtTRANSYEARMON.value) = 6 Then DateClean
		end if
	End With
End Sub

'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------
Sub imgCalDemandday_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub


'청구년월
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'발행일
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
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

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub

'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_SUM, NewTop, NewLeft
End Sub

'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM	
	End with
end sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMDCMCATVCOMMI	= gCreateRemoteObject("cMDCT.ccMDCTCATVCOMMI")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "152px"
	pnlTab1.style.left= "7px"
	
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "152px"
	pnlTab2.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		'*********************************
		'수수료시트
		'*********************************
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 1, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,   "CHK | YEARMON | SEQ | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | CLIENTCODE | CLIENTNAME | CLIENTBISNO | SUBSEQ | MEDCODE | MEDNAME | AMT | COMMI_RATE | COMMISSION  | MEMO | DEPT_CD | COMMI_TAX_FLAG | TRANSRANK| ATTR01 |SPONSOR|COMMI_TRANS_NO"
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|YEARMON|SEQ|REAL_MED_CODE|매체사|매체사사업자번호|CLIENTCODE|광고주|광고주사업번호|시퀀스|MEDCODE|채널명|취급액|수수료율|수수료|비고|DEPT_CD|COMMI_TAX_FLAG|TRANSRANK|ATTR01|SPONSOR|COMMI_TRANS_NO"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|	   0|  0|	         0|    20|	            13|         0|    20|		     13|     6|      0|    12|    12|      10|    12|  13|      0|             0|        0|     0|      0|            0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13" 
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " CLIENTNAME | MEDNAME | REAL_MED_NAME | MEMO", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | COMMISSION  ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME|MEDNAME|MEMO",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "REAL_MED_BISNO|CLIENTBISNO|SUBSEQ",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ |CLIENTCODE | REAL_MED_CODE | MEDCODE|DEPT_CD|COMMI_TAX_FLAG|SPONSOR|TRANSRANK|COMMI_TRANS_NO|ATTR01", true
		
		'합계 표시 그리드 기본화면 구성
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 22, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "CHK | YEARMON | SEQ | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | CLIENTCODE | CLIENTNAME | CLIENTBISNO | SUBSEQ | MEDCODE | MEDNAME | AMT | COMMI_RATE | COMMISSION  | MEMO | DEPT_CD | COMMI_TAX_FLAG | TRANSRANK| ATTR01 |SPONSOR|COMMI_TRANS_NO"
		mobjSCGLSpr.SetText .sprSht_SUM, 5, 1, "합      계"
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"1|3|5",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|COMMISSION", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_SUM, "YEARMON | SEQ |CLIENTCODE | REAL_MED_CODE | MEDCODE|DEPT_CD|COMMI_TAX_FLAG|SPONSOR|TRANSRANK|COMMI_TRANS_NO|ATTR01", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
	    
	    '수수료거래명세 조회 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 22, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht1, "TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE |CLIENTNAME| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME| DEPT_CD| DEMANDDAY| PRINTDAY| AMT| SUSURATE| SUSU| VAT| MEMO| MED_FLAG| SPONSOR| TAXYEARMON| TAXNO| TRUST_SEQ"
		mobjSCGLSpr.SetHeader .sprSht1,		"TRANSYEARMON|TRANSNO|SEQ|CLIENTCODE|광고주명|MEDCODE|매체명|REAL_MED_CODE|매체사|DEPT_CD|청구일자|발행일자|취급액|수수료율|금액|부가세|비고| MED_FLAG| SPONSOR|TAXYEARMON|TAXNO|TRUST_SEQ" 
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "       0|	    0|	0|	   0|	     18|      0|    18|	           0|     0|      0|       0|       8|  10|       7|    10|    10|   10"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "DEMANDDAY| PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT| SUSU| VAT|SUSURATE", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "CLIENTNAME|MEDNAME|CLIENTNAME", -1, -1, 20
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "CLIENTNAME|MEDNAME|MEMO",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht1, "TRANSYEARMON|TRANSNO|SEQ| CLIENTCODE|MEDCODE|REAL_MED_CODE|REAL_MED_NAME|DEPT_CD|DEMANDDAY|MED_FLAG| SPONSOR |TAXYEARMON|TAXNO|TRUST_SEQ", true
		
    End With
    
	pnlTab1.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)

	'기본값 설정
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
				case 1 : .txtREAL_MED_CODE.value = vntInParam(i)
				case 2 : .txtREAL_MED_NAME1.value = vntInParam(i)			'조회추가필드
				case 3 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				case 4 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 5 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjMDCMCATVCOMMI = Nothing
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
		.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean
		'.txtDEMANDDAY.value = gNowDate
		.txtPRINTDAY.value  = gNowDate
		.sprSht.MaxRows = 0	
		.sprSht1.MaxRows = 0
		
		.txtDEMANDDAY.readOnly = "FALSE"
		.txtDEMANDDAY.className = "INPUT"
		.imgCalDemandday.disabled = FALSE

	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON, strCOMMIYEARMON
	Dim intTRANSNO, intCOMMINO
	Dim intRANKTRANS
	Dim intCnt,bsdiv, bsdiv1
	Dim intColFlag, intColFlag1
	Dim chkcnt
	chkcnt = 0
	
	with frmThis
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "발행일은 필수 입력 사항 입니다.",""
			Exit Sub
		End If
		
		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "거래명세서를 생성할 데이터를 체크 하십시오",""
			exit sub
		end if

		 '저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
		
		For intCnt = 1 To .sprSht.MaxRows
		'최대값
			bsdiv1 = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
			IF intColFlag1 < bsdiv1 THEN
				intColFlag1 = bsdiv1
			END IF
		Next
		
   		'데이터 Validation
   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "상세항목 이 없습니다.",""
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | REAL_MED_CODE | REAL_MED_NAME | REAL_MED_BISNO | CLIENTCODE | CLIENTNAME | CLIENTBISNO | SUBSEQ | MEDCODE | MEDNAME | AMT | COMMI_RATE | COMMISSION  | MEMO | DEPT_CD | COMMI_TAX_FLAG | TRANSRANK| ATTR01 |SPONSOR|COMMI_TRANS_NO")
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		intCOMMINO = 0
		strCOMMIYEARMON =.txtTRANSYEARMON.value
		
		intRtn = mobjMDCMCATVCOMMI.ProcessRtn(gstrConfigXml,strMasterData,vntData, intCOMMINO,strCOMMIYEARMON, intColFlag1)

		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			'InitPageData
			If intRtn <> 0  Then
				.txtTRANSNO.value = intCOMMINO
				.txtTRANSYEARMON.value = strCOMMIYEARMON
				selectRtn
			Else
				initpagedata
			End If
   		end if
   	end with
End Sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	with frmThis

  	End with
	DataValidation = true
End Function

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' 거래명세서 발행 조회[최초입력조회]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData1
	Dim strTRANSYEARMON, strREAL_MED_CODE, strREAL_MED_NAME, strTRANSNO
	Dim strPRINTDAY
	Dim strSPONSOR
   	Dim i, strCols
   	Dim IngsusuColCnt, IngsusuRowCnt
   	
	'On error resume next
	with frmThis
		If .txtTRANSYEARMON.value = "" Then
			gErrorMsgBox "조회시 년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If
		
		strTRANSNO = ""
		strTRANSNO = .txtTRANSNO.value
'		IF strTRANSNO = "" THEN
'			IF  .txtREAL_MED_CODE.value = "" THEN
'				gErrorMsgBox "조회시 청구지는 반드시 넣어야 합니다.",""
'				Exit SUb
'			END IF
'		END IF
		
		If .chkSPONSOR.checked = True Then
			strSPONSOR = "Y"
		Else
			strSPONSOR = ""
		End If
		'Sheet초기화
		.sprSht.MaxRows = 0
		.sprSht1.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON		= .txtTRANSYEARMON.value
		strTRANSNO			= .txtTRANSNO.value
		strREAL_MED_CODE	= .txtREAL_MED_CODE.value
		strREAL_MED_NAME	= .txtREAL_MED_NAME1.value
		
		IF strTRANSNO <> "" THEN
			IF not SelectRtn_HDR (strTRANSYEARMON, strTRANSNO, strREAL_MED_CODE) Then Exit Sub
			
			pnlTab1.style.visibility = "HIDDEN"
			pnlTab2.style.visibility = "visible"
			
			.txtDEMANDDAY.readOnly = "TRUE"
			.txtDEMANDDAY.className = "NOINPUT"
			.imgCalDemandday.disabled = True

			'쉬트 조회
			if SelectRtn_DTL (strTRANSYEARMON, strTRANSNO, strREAL_MED_CODE) then
				.txtTRANSYEARMON.value = strTRANSYEARMON
				.txtTRANSNO.value = strTRANSNO
				.txtREAL_MED_CODE.value = strREAL_MED_CODE
				.txtREAL_MED_NAME1.value = strREAL_MED_NAME
			end if
		ELSE
		
			InitPageData
			vntData1 = mobjMDCMCATVCOMMI.SelectRtn_SUSU(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON,strREAL_MED_CODE,strSPONSOR)
			
			pnlTab1.style.visibility = "visible"
			pnlTab2.style.visibility = "HIDDEN"
			
			.txtDEMANDDAY.readOnly = "FALSE"
			.txtDEMANDDAY.className = "INPUT"
			.imgCalDemandday.disabled = FALSE

			if not gDoErrorRtn ("SelectRtn") then
				if mlngRowCnt > 0 then
					mobjSCGLSpr.SetClipbinding .sprSht, vntData1, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	   				PreSearchFiledValue strTRANSYEARMON, strREAL_MED_CODE, strREAL_MED_NAME
   				else
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   					'InitPageData
   					PreSearchFiledValue strTRANSYEARMON, strREAL_MED_CODE, strREAL_MED_NAME
   				end if
   				DateClean
   				AMT_SUM '합계그리드 표시
   			end if
		END IF
		
   	end with
End Sub

Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strREAL_MED_CODE)
	dim vntData
	on error resume next

	'초기화
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMGET.Get_CATVCOMMI_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strREAL_MED_CODE)
	
	IF not gDoErrorRtn ("Get_PRINTCOMMI_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 거래명세번호에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			gWriteText "", "선택한 거래명세번호에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			SelectRtn_HDR = True
			
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByVal strYEARMON,ByVal strTRANSNO, ByVal strREAL_MED_CODE)
	dim vntData
	Dim strCOUNT
	Dim intCnt
	'on error resume next

	'초기화
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMGET.Get_CATVCOMMI_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strREAL_MED_CODE)
	
	IF not gDoErrorRtn ("Get_PRINTCOMMI_LIST") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		strCOUNT = "0"
		For intCnt = 1 To frmThis.sprSht1.MaxRows
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"SPONSOR", intCnt) = "Y" Then
				frmThis.chkSPONSOR.checked = True
				strCOUNT = "1"
				Exit For
			End If	
		Next
		
		If strCOUNT = "0" Then
		frmThis.chkSPONSOR.checked = False
		End If
		
		
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG

		SelectRtn_DTL = True
		gWriteText "", "선택한 거래명세번호건의 상세내역에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strREAL_MED_CODE, strREAL_MED_NAME)
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME1.value = strREAL_MED_NAME
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMOUNT, IntAMTSUM, IntCOMMISSION, IntCOMMISSIONSUM
	'AMOUNT|COMMISSION
	With frmThis
		IntAMTSUM = 0
		IntCOMMISSIONSUM = 0

		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMOUNT = 0
				IntCOMMISSION = 0
				
				IntAMOUNT		= mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt) '금액
				IntCOMMISSION	= mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION", lngCnt)	'부가세
				
				IntAMTSUM		 = IntAMTSUM + IntAMOUNT
				IntCOMMISSIONSUM = IntCOMMISSIONSUM + IntCOMMISSION
			Next
		end if
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"COMMISSION",1, IntCOMMISSIONSUM		
		ELSE
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, 0
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"COMMISSION",1, 0		
		end if
	End With
End Sub

'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	with frmThis
		strDESCRIPTION = ""
		
		IF .sprSht1.MaxRows = 0 THEN
			gErrorMsgBox "삭제할 건의 상세내역이 없습니다.","삭제안내!"
			Exit Sub
		END IF
		
		For intCnt2 = 1 To .sprSht1.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",intCnt2) & " 에 대하여" &vbcrlf & "세금계산서번호가 존재하는 내역은 삭제가 되지 않습니다.","삭제안내!"
				Exit Sub
			End If
		Next
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0
		
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		'mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE |CLIENTNAME| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME| DEPT_CD| DEMANDDAY| PRINTDAY| AMT| SUSURATE| SUSU| VAT| MEMO| MED_FLAG| SPONSOR| TAXYEARMON| TAXNO| TRUST_SEQ")
		
		'선택된 자료를 끝에서 부터 삭제
		strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
	
		intRtn = mobjMDCMCATVCOMMI.DeleteRtn(gstrConfigXml,vntData, strTRANSYEARMON, strTRANSNO)

		IF not gDoErrorRtn ("DeleteRtn") then
			If strDESCRIPTION <> "" Then
				gErrorMsgBox strDESCRIPTION,"삭제안내!"
				Exit Sub
			End If
			for i = .sprSht1.MaxRows to 1 step -1
				mobjSCGLSpr.DeleteRow .sprSht1,i
			next
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", strTRANSYEARMON & "-" & strTRANSNO & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제 
		mobjSCGLSpr.DeselectBlock .sprSht1
		'SelectRtn
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
			<TABLE id="tblForm" style="WIDTH: 793px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD style="WIDTH: 300px" align="left" width="427" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">CATV&nbsp;수수료거래명세 생성 및 발행</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 495px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="203" border="0">
										<TR>
											<TD></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
													src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다."
													src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgDelete.gIF" width="54"
													border="0" name="imgDelete"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
							<!--Top Define Table End-->
							<!--Input Define Table End--></TABLE>
						<TABLE id="tblBody" style="WIDTH: 792px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="굴림">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)"
													width="91">년 월</TD>
												<TD class="DATA" width="312"><INPUT class="INPUT" id="txtTRANSYEARMON" title="거래명세년월" style="WIDTH: 72px; HEIGHT: 22px"
														accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON"><IMG id="ImgTRU" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTRU"><INPUT class="INPUT" id="txtTRANSNO" title="거래명세번호" style="WIDTH: 72px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTRANSNO"></TD>
												<TD class="LABEL" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"
													width="90"><FONT face="굴림">발행일자</FONT></TD>
												<TD class="DATA" width="312"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="발행일" style="WIDTH: 96px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtPRINTDAY"><IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgCalPrintday">
												</TD>
											</TR>
											<tr>
												<TD class="LABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_CODE, txtREAL_MED_NAME1)"
													title="청구지코드,명 삭제">청구지
												</TD>
												<TD class="DATA"><INPUT class="INPUT" id="txtREAL_MED_CODE" title="코드조회" style="WIDTH: 72px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtREAL_MED_CODE"><IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgREAL_MED_CODE"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="코드명" style="WIDTH: 217px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="30" name="txtREAL_MED_NAME1"></TD>
												<TD class="LABEL" style="WIDTH: 90px; CURSOR: hand">협찬구분
												</TD>
												<TD class="DATA">&nbsp;&nbsp;<INPUT id="chkSPONSOR" type="checkbox" name="chkSPONSOR">&nbsp;협찬&nbsp; 
													생성 및 조회</TD>
											</tr>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 794px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="right" border="0">
										<TR>
											<TD class="LABEL" width="90"><FONT face="굴림">청구지</FONT></TD>
											<TD class="DATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtREAL_MED_NAME" title="광고주명" style="WIDTH: 172px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtREAL_MED_NAME">
											</TD>
											<TD class="LABEL" width="90"><FONT face="굴림">담당부서</FONT></TD>
											<TD class="DATA" width="173"><FONT face="굴림"><INPUT dataFld="DEPT_NAME" class="NOINPUT_L" id="txtDEPT_NAME" title="브랜드코드" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtDEPT_NAME"></FONT>
											</TD>
											<TD class="LABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')"><FONT face="굴림">청구일자</FONT></TD>
											<TD class="DATA" width="173"><FONT face="굴림"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="브랜드명" style="WIDTH: 104px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="12" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></FONT></TD>
										</TR>
										<TR>
											<TD class="LABEL"><FONT face="굴림">대행금액</FONT></TD>
											<TD class="DATA"><FONT face="굴림"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="광고금액" style="WIDTH: 135px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="16" name="txtAMT"></FONT>
											</TD>
											<TD class="LABEL"><FONT face="굴림">부가세</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="부가세" style="WIDTH: 135px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="17" name="txtVAT"></TD>
											<TD class="LABEL"><FONT face="굴림">계</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT dataFld="SUMAMTVAT" class="NOINPUT_R" id="txtSUMAMTVAT" title="계" style="WIDTH: 154px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="20" name="txtSUMAMTVAT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End--></TABLE>
					</TD>
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 791px"><FONT face="굴림"></FONT></TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start-->
				<TR>
					<TD class="LISTFRAME" style="WIDTH: 794px; HEIGHT: 360px" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 786px; HEIGHT: 336px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="20796">
								<PARAM NAME="_ExtentY" VALUE="9155">
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
							<OBJECT id="sprSht_SUM" style="WIDTH: 786px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
					<PARAM NAME="_Version" VALUE="393216">
					<PARAM NAME="_ExtentX" VALUE="20796">
					<PARAM NAME="_ExtentY" VALUE="635">
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
						<DIV id="pnlTab2" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht1" style="WIDTH: 786px; HEIGHT: 360px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="20796">
								<PARAM NAME="_ExtentY" VALUE="9419">
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
				<!--
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 794px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
					</TR>
					-->
				<!--BodySplit End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
