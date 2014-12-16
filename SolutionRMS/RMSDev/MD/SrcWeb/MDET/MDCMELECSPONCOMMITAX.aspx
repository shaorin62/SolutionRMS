<%@ Page CodeBehind="MDCMELECSPONCOMMITAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECSPONCOMMITAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 협찬광고 수수료 세금계산서 발행</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/표준샘플/스프레드쉬트
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : SpreadSheet를 이용한 조회/입력/수정/삭제/인쇄 처리 표준 샘플
'파라  메터 : 
'특이  사항 : 표준샘플을 위해 만든 것임
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
'****************************************************************************************
-->
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet 정보 --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECSPONCOMMITAX , mobjMDCMGET
Dim mstrCheck
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
Sub imgFind_onclick
	GETTAXPOP
End Sub

Sub imgQuery_onclick
	if frmThis.txtTRANSYEARMON1.value = "" then
	    gErrorMsgBox "년월 입력하시오",""
		exit Sub
	end if
	If LEN(frmThis.txtTRANSYEARMON1.value) <> 6 Then
		 gErrorMsgBox "년월은 6자리 입니다",""
		exit Sub
	End If
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim VATFLAG
	Dim FLAG
	Dim strUSERID
	
	IF frmThis.sprSht.MaxRows = 0 then
		gFlowWait meWAIT_ON
		with frmThis		
			ModuleDir = "MD"
			ReportName = "TAXNO_BLACK.rpt"
						
			IF .cmbFLAG.value = "receipt" THEN
				FLAG = "Y"
			ELSE
				FLAG = "N"
			END IF
						
			Params = FLAG
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
		end with
		gFlowWait meWAIT_OFF
	else
	
		'체크가 된 데이터가 있는지 없는지 체크한다.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
				intCount = 1
			end if
		next
		
		'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
		if intCount = 0 then
			gErrorMsgBox "선택된 데이터가 없습니다. 인쇄할 데이터를 체크하시오",""
			Exit Sub
		end if
		
		gFlowWait meWAIT_ON
		with frmThis
			'인쇄버튼을 클릭하기 전에 md_tax_temp테이블에 내용을 삭제한다
			'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
			'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
			'md_trans_temp삭제 시작
			intRtn = mobjMDCMELECSPONCOMMITAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp삭제 끝
			
			ModuleDir = "MD"
			ReportName = "TAX_NEW.rpt"
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
					strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"VAT",i) = 0 THEN
						VATFLAG = "N"
					ELSE
						VATFLAG = "Y"
					END IF
					IF .cmbFLAG.value = "receipt" THEN
						FLAG = "Y"
					ELSE
						FLAG = "N"
					END IF
					strUSERID = ""
					
					vntDataTemp = mobjMDCMELECSPONCOMMITAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			Params = "MD_TAXELEC_TEMP" & ":" & strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
			'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	END IF
End Sub

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMELECSPONCOMMITAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub	

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub ImgTaxCre_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub btnCOMMISSION_onclick ()
	Dim intCnt
	Dim intRtn
	Dim strDEMANDDAY
	Dim strTAXYEARMON
	With frmThis
		If .rdT.checked = True Then
			gErrorMsgBox "청구일 적용은 미완료상태 에서 적용됩니다.","처리안내!"
			Exit Sub
		End If
		
		strDEMANDDAY = .txtDEMANDDAY.value
		strTAXYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		intRtn = gYesNoMsgbox("선택된 항목의 청구일을 변경 하시겠습니까?","변경 확인")
		IF intRtn <> vbYes then exit Sub
			
		For intCnt = 1 To .sprSht.MaxRows
			If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
			mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
			mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
			End If
		Next
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 매체사코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
'청구일 조회조건 생성
Sub DateClean
Dim date1
Dim date2
Dim strDATE
	strDATE = MID(frmThis.txtTRANSYEARMON1.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
		.txtDEMANDDAY.value = date2
		
	End With
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtTRANSYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "commi", "ELECSPON") 
		vntRet = gShowModalWindow("../MDCO/MDCMTAXCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(1,0) and .txtCLIENTNAME1.value = vntRet(2,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code값 저장
			.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			IF vntRet(3,0) = "완료" THEN
				'window.event.keyCode = meEnter
				'txtTRANSNO1_onkeydown
			ELSE
				.txtTRANSNO1.value = ""
			END IF
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTAXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"commi","ELECSPON")
											  'gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtCUSTNO.value,.txtCUSTNAME.value, mtranscommiflag, mtransTblflag
			if not gDoErrorRtn ("GetTAXCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = vntData(1,0)
					.txtCLIENTNAME1.value = vntData(2,0)
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
' 거래명세서팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' 거래명세서팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'1건을 찾을경우 즉시 조회
Sub txtTRANSNO1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON1.value <> "" Or Len(.txtTRANSYEARMON1.value) = 6 Then
				strYEARMON = .txtTRANSYEARMON1.value
			End If
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, .txtTRANSNO1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "commi", "ELECSPON","0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON1.value = vntData(0,0)  ' Code값 저장
					.txtTRANSNO1.value = vntData(1,0)  ' 코드명 표시
					.txtCLIENTCODE1.value = vntData(2,0)  ' 코드명 표시
					.txtCLIENTNAME1.value = vntData(3,0)  ' 코드명 표시
					'Call SelectRtn ()
				Else
					Call TRANSPOP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRANSPOP
	dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis

	If .txtTRANSYEARMON1.value <> "" Or Len(.txtTRANSYEARMON1.value) = 6 Then
	strYEARMON = .txtTRANSYEARMON1.value
	End If
	'msgbox strYEARMON
		vntInParams = array(strYEARMON, .txtTRANSNO1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "commi","ELECSPON") '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON1.value = vntRet(0,0) and .txtTRANSNO1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTRANSYEARMON1.value = vntRet(0,0)  ' Code값 저장
			.txtTRANSNO1.value = vntRet(1,0)  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 달력
'-----------------------------------------------------------------------------------------
Sub ImgPRINTDAY_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.ImgPRINTDAY,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtTRANSYEARMON1_onblur
	With frmThis
	If .txtTRANSYEARMON1.value <> "" AND Len(.txtTRANSYEARMON1.value) = 6 Then DateClean
	End With
End Sub
Sub imgCalDemandday_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'청구일
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'발행일
Sub txtPRINTDAY_onchange
	gSetChange
End Sub
Sub imgFROM_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgTO_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub img1_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.img1,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub
Sub txtTO_onchange
	gSetChange
End Sub
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			'mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			'mALLCHECK = TRUE
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
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	Dim strMEDFLAG
	strMEDFLAG = "A"
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If .rdT.checked = True Then
			'msgbox strMEDFLAG
			vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row),strMEDFLAG) '<< 받아오는경우
			gShowModalWindow "MDCMELECCOMMITAXDTL.aspx",vntInParams , 813,565
			SelectRtn
			End IF
		end if	
	end with
end sub

Sub cmbGUBUN_onchange
	with frmThis
	If .cmbGUBUN.value = "taxdiv" Then
	selectRtn
	Elseif  .cmbGUBUN.value = "taxgroup" Then
	mobjSCGLSpr.setTextBinding .sprSht,"SUMM",-1,"공중파 협찬광고 대행수수료"
	End If
	End with
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성
	set mobjMDCMELECSPONCOMMITAX	= gCreateRemoteObject("cMDET.ccMDETELECSPONCOMMITAX")
	set mobjMDCMGET					= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 35, 0, 7, 2
		mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|TAXYEARMON|TAXNO|TAXMANAGE|TRANSYEARMON|TRANSNO|SEQ|DEMANDDAY|CLIENTNAME|CLIENTBISNO|REAL_MED_NAME|REAL_MED_BISNO|MEDNAME|AMT|SUSURATE|SUSU|VAT|SUMAMT|SUMM|DEPT_NAME|PRINTDAY|CLIENTCODE|CLIENTACCODE|REAL_MED_CODE|REAL_MED_ACCODE|MEDCODE|DEPT_CD|MEDFLAG|VOCHNO|RANKTRANS|MEMO|SPONSOR| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetHeader .sprSht,		  "선택|년월|번호|관리번호|년월|번호|순번|청구일|광고주명|광고주사업자등록번호|매체사명|매체사사업자등록번호|매체명|취급고|수수료율|수수료|부가세액|합계금액|적요|부서명|발행일|광고주코드|광고주AC코드|매체사코드|매체사AC코드|매체코드|부서코드|집계구분|전표번호|순위|비고|협찬구분| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|5   |4   |	    11|   5|   4|   4|8     |      19|0                   |19      |17                  |9     |9     |8       |9     |9       |9       |30  |10    |8     |0         |0           |0         |0           |0       |0       |0       |10      |0   |20  |0|0|0|0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT|SUSU", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUSURATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXMANAGE|TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|CLIENTNAME|REAL_MED_NAME|MEDNAME|AMT|SUSURATE|SUSU|VAT|SUMAMT|DEPT_NAME|CLIENTCODE|CLIENTACCODE|CLIENTBISNO|REAL_MED_CODE|REAL_MED_ACCODE|REAL_MED_BISNO|MEDCODE|DEPT_CD|MEDFLAG|SEQ|VOCHNO|RANKTRANS|MEMO"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXMANAGE|SUMM", -1, -1, 100
		mobjSCGLSpr.SetCellAlign2 .sprSht, "TRANSYEARMON|TRANSNO|SEQ|TAXNO|REAL_MED_BISNO|TAXYEARMON|TAXMANAGE",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME|REAL_MED_NAME|MEDNAME|SUMM|DEPT_NAME|MEMO",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "TAXNO|CLIENTCODE|CLIENTACCODE|REAL_MED_CODE|REAL_MED_ACCODE|MEDCODE|DEPT_CD|MEDFLAG|VOCHNO|TAXYEARMON|RANKTRANS|SPONSOR|CLIENTBISNO|REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2", true
		.sprSht.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMELECSPONCOMMITAX = Nothing
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
		.txtTRANSYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		'Sheet초기화
		DateClean
		.sprSht.MaxRows = 0
		
		.txtTRANSYEARMON1.focus()
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'완료/미완료 조회
Sub rdT_onclick
	SelectRtn
End Sub
Sub rdF_onclick
	SelectRtn
End Sub
'전체체크
Sub rdA_onclick
	SelectRtn
End Sub
'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCUSTCODE, strTRANSNO,strGUBUN
	Dim strFROM,strTO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		If .txtTRANSYEARMON1.value = "" Then
			gErrorMsgBox "년월을 입력하십시오","조회안내!"
			Exit Sub
		End If	
		If Len(.txtTRANSYEARMON1.value) <> 6 Then
			gErrorMsgBox "년월의 형식이 아닙니다.","조회안내!"
			Exit Sub
		End If

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strYEARMON	= .txtTRANSYEARMON1.value
		strCUSTCODE	= .txtCLIENTCODE1.value
		strGUBUN = .cmbGUBUN.value
		
		'세금계산서 완료조회
		If .rdT.checked = True Then
			vntData = mobjMDCMELECSPONCOMMITAX.Get_ELECSPON_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strCUSTCODE,strFROM,strTO)
			If not gDoErrorRtn ("Get_ELECSPON_TAX") then
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'초기 상태로 설정
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				mobjSCGLSpr.ColHidden .sprSht, "TAXMANAGE|VOCHNO", False
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|SEQ|MEDCODE|MEDNAME|CLIENTNAME", True
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
			End If
		'미완료 거래명세서 디테일 조회
		ElseIf .rdF.checked = True Then
			vntData = mobjMDCMELECSPONCOMMITAX.Get_ELECSPON_TAXBUILD(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE,strGUBUN)
			If not gDoErrorRtn ("Get_ELECSPON_TAXBUILD") then
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'초기 상태로 설정
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON|TAXNO|TAXMANAGE|VOCHNO", True
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|SEQ|MEDNAME|CLIENTNAME", False
				'Layout_change
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
			End If
		ElseIf .rdA.checked = True Then			
			vntData = mobjMDCMELECSPONCOMMITAX.Get_ELECSPON_TAXALL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE, strGUBUN)
			If not gDoErrorRtn ("Get_ELECSPON_TAXALL") then
				'초기 상태로 설정
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
				mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON|TAXNO|TAXMANAGE|VOCHNO", True
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|SEQ|MEDNAME|CLIENTNAME", False
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
				'Layout_change
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
			End If
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

'Sub Layout_change ()
'	Dim intCnt
'	with frmThis
'	For intCnt = 1 To .sprSht.MaxRows 
'		If mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",intCnt) = "Y" Then
'		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
'		End If
'	Next 
'	End With
'End Sub

'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
    Dim intRtn2
   	Dim vntData, vntData1
	Dim strMasterData
	Dim strTAXYEARMON
	Dim intTAXNO
	Dim strTAXSET
	Dim strSUMM
	'Dim strATTR02FLAG
	Dim intCnt
	Dim strDEMANDDAY,strPRINTDAY
	Dim chkcnt
	Dim intCnt2
	Dim intColFlag
	Dim intMaxCnt
	Dim bsdiv
	Dim strVALIDATION
	with frmThis
		
		
		'발행일은 xml 에서 처리할수 없으므로 반드시 저장체크 필요
		If .rdT.checked = True Then
		gErrorMsgBox "미완료 상태에서 저장이 가능합니다.","저장안내!"
		Exit Sub
		End If
		intRtn2 = gYesNoMsgbox("청구일을 확인하셨습니까?","확인")
		IF intRtn2 <> vbYes then exit Sub
			
		For intCnt = 1 To .sprSht.MaxRows
		strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
		strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
			If strDEMANDDAY  = "" Then
				gErrorMsgBox "청구일은 필수 입니다.","저장안내!"
				Exit Sub
			End If
			If  strPRINTDAY = "" Then
				gErrorMsgBox "청구일은 필수 입니다.","저장안내!"
				Exit Sub
			End If
		Next
		'체크 없을 경우 저장 안되도록
		chkcnt = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "세금계산서를 생성할 데이터를 체크 하십시오","저장안내!"
			exit sub
		end if
		 '저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "상세항목 이 없습니다.",""
   			Exit Sub
   		End If
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TAXYEARMON|TAXNO|TAXMANAGE|TRANSYEARMON|TRANSNO|SEQ|DEMANDDAY|CLIENTNAME|CLIENTBISNO|REAL_MED_NAME|REAL_MED_BISNO|MEDNAME|AMT|SUSURATE|SUSU|VAT|SUMAMT|SUMM|DEPT_NAME|PRINTDAY|CLIENTCODE|CLIENTACCODE|REAL_MED_CODE|REAL_MED_ACCODE|MEDCODE|DEPT_CD|MEDFLAG|VOCHNO|RANKTRANS|MEMO|SPONSOR| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2")
		
		'마스터 데이터를 가져 온다.
		'
		
		'처리 업무객체 호출
		intTAXNO = 0
		If .cmbGUBUN.value = "taxdiv" Then
		intRtn = mobjMDCMELECSPONCOMMITAX.ProcessRtn_Div(gstrConfigXml,vntData, intTAXNO)
		Else
		
			'validation
			If Not TaxGroup(strVALIDATION) Then 
				gErrorMsgBox strVALIDATION & vbCrlf & "단위별 [광고주,매체사] [청구일,작성일,적요] 는 동일 하여야 합니다.","저장안내!"
				Exit Sub
			Else
				'최대값
				intColFlag = 0
				For intMaxCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intMaxCnt) = 1 Then
						bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intMaxCnt))
						IF intColFlag < bsdiv THEN
							intColFlag = bsdiv
							'rowflag = lngCnt
						END IF
					End IF
				Next
				'맥스값만 추가하여 보내기
				intRtn = mobjMDCMELECSPONCOMMITAX.ProcessRtn_Group(gstrConfigXml,vntData, intTAXNO,intColFlag)
			End IF
		End If

		if not gDoErrorRtn ("ProcessRtn_Group") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "저장성공","저장안내!"
			.rdT.checked = True
			selectRtn
   		end if
   	end with
End Sub

Function TaxGroup(ByRef strVALIDATION)
	Dim intCnt
	Dim strCLIENTCODE '광고주 사업자 등록번호
	Dim strREAL_MED_CODE '청구지 사업자 등록번호
	Dim strDEMANDDAY
	Dim strPRINTDAY
	Dim strSUMM
	Dim strStartRank
	TaxGroup = False
	with frmThis
		strStartRank = "0"
		strCLIENTCODE = ""
		strREAL_MED_CODE = ""
		strDEMANDDAY = ""
		strPRINTDAY = ""
		strSUMM = ""
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				
				If strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Then
					'If strCLIENTCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",intCnt) Then
					'	Exit Function
					'End If
					If strREAL_MED_CODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt) Then
						Exit Function
					End If 
					If strDEMANDDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt) Then
						strVALIDATION = "청구일확인 거래명세서번호 " & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " 번"
						Exit Function
					End If 
					If strPRINTDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt) Then
						strVALIDATION = "발행일확인 거래명세서번호" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " 번"
						Exit Function
					End If 
					If strSUMM <> mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt) Then
						strVALIDATION = "적요확인 거래명세서번호" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " 번"
						Exit Function
					End If 
				End If
				
				strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBISNO",intCnt)
				strREAL_MED_CODE = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt)
				strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
				strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
				strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt)
			End If
		Next
	End With
	TaxGroup = True
End Function

Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strDESCRIPTION
	with frmThis
	strDESCRIPTION = ""
		For intCnt2 = 1 To .sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
			If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " 에 대하여" &vbcrlf & "전표가 존재하는 내역은 삭제가 되지 않습니다.","삭제안내!"
				Exit Sub
			End If
		END IF
		Next
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
			
				intRtn = mobjMDCMELECSPONCOMMITAX.DeleteRtn_CommiTax(gstrConfigXml,strTAXYEARMON, strTAXNO)
				IF not gDoErrorRtn ("DeleteRtn_CommmiTax") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"삭제안내!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn_CommiTax") then
			gWriteText lblstatus, intCnt & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="165" background="../../../images/back_p.gIF"
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
											<td class="TITLE">협찬 수수료세금계산서 생성</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 326px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End--></TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, txtTRANSNO1)"
												width="80">년월/번호</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtTRANSYEARMON1" title="거래명세년월" style="WIDTH: 111px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="13" name="txtTRANSYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80">매체사&nbsp;</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 203px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="28" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT class="INPUT" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 65px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
												width="80">청구일
											</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="6" name="txtFROM">&nbsp;<IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="6" name="txtTO">&nbsp;<IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													align="absMiddle" border="0" name="imgTo">
											</TD>
											<TD class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 조회합니다."
													src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, txtTRANSNO1)"
												width="80">발행</TD>
											<TD class="SEARCHDATA" colSpan="6">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="rdT" title="완료내역조회" type="radio" name="rdGBN">&nbsp;완료&nbsp;&nbsp;&nbsp; 
												&nbsp; <INPUT id="rdF" title="미완료 내역조회" type="radio" CHECKED name="rdGBN">&nbsp;미완료&nbsp;&nbsp;&nbsp; 
												&nbsp;<INPUT id="rdA" title="전체 내역조회" type="radio" value="on" name="rdGBN">&nbsp;전체</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="ImgTaxCre" onmouseover="JavaScript:this.src='../../../images/ImgTaxCreOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTaxCre.gif'"
																height="20" alt="선택되어진 방식에 따라 세금계산서를 작성합니다." src="../../../images/ImgTaxCre.gif"
																align="absMiddle" border="0" name="ImgTaxCre"></td>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" width="54" border="0"
																name="imgDelete"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
																name="imgPrint"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
									<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 67px">분할/합산</TD>
											<TD class="DATA" style="WIDTH: 189px"><SELECT id="cmbGUBUN" title="매체구분" style="WIDTH: 96px" name="cmbGUBUN">
													<OPTION value="taxdiv" selected>분할발행</OPTION>
													<OPTION value="taxgroup">합산발행</OPTION>
												</SELECT>
											</TD>
											<TD class="LABEL" style="WIDTH: 73px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY, '')">청구일 
												적용</TD>
											<TD class="DATA" style="WIDTH: 294px"><INPUT class="INPUT" id="txtDEMANDDAY" title="청구일자" style="WIDTH: 112px; HEIGHT: 22px"
													accessKey="date" type="text" maxLength="10" size="13" name="txtDEMANDDAY"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="Img1"><IMG id="btnCOMMISSION" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" height="20" alt="해당 정산일로 정산일이 없는 상세항목을 Setting 합니다"
													src="../../../images/imgApp.gIF" width="54" align="absMiddle" border="0" name="btnCOMMISSION">
											</TD>
											<TD class="LABEL" style="WIDTH: 83px">청구/영수구분</TD>
											<td class="DATA">
												<SELECT id="cmbFLAG" title="영수/청구구분" style="WIDTH: 120px" name="cmbFLAG">
													<OPTION value="receipt" selected>청구</OPTION>
													<OPTION value="demand">영수</OPTION>
												</SELECT>
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								DESIGNTIMEDRAGDROP="213">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31856">
								<PARAM NAME="_ExtentY" VALUE="4339">
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
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림">SDF</FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></form>
	</body>
</HTML>
