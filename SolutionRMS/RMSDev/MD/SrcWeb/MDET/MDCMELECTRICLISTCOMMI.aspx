<%@ Page CodeBehind="MDCMELECTRICLISTCOMMI.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRICLISTCOMMI" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 수수료 승인처리</title> 
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
Dim mobjMDCMELECTRICLISTCOMMI 
Dim mobjMDCMGET
Dim mstrCheck
mstrCheck = True
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
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오",""
		exit Sub
	end if
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSetting_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmCancel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht1
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'----------------------------
'수수료 관리 TAB BUTTON CLICK
'----------------------------
Sub btnTab2_onclick
	pnltab2.style.visibility = "visible"
	mobjSCGLCtl.DoEventQueue
End Sub


'****************************************************************************************
' 쉬트 더불클릭 이벤트
'****************************************************************************************
sub sprSht1_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		end if
	end with
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
	
	frmThis.imgSetting.style.visibility = "hidden"
	frmThis.ImgConfirmCancel.style.visibility = "hidden"
	'서버업무객체 생성	
	set mobjMDCMELECTRICLISTCOMMI = gCreateRemoteObject("cMDET.ccMDETELECTRICLISTCOMMI")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "102px"
	pnlTab2.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
   
    '*********************************
    '수수료시트
    '*********************************
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 13, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht1,   "YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET"
		mobjSCGLSpr.SetHeader .sprSht1,		   "년월|매체사|광고주|INPUT_MEDFLAG|매체구분|대행금액|수수료율|수수료|광고주코드|매체사코드|부서코드|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "  0|    30|    38|            0|      13|      13|      13|    15|0         |0         |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, " YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU", -1, -1, 0
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "AMT|SUSURATE|SUSU", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT|SUSU|SUSURATE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"AMT|SUSURATE"		
		mobjSCGLSpr.ColHidden .sprSht1, "CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET", true
		mobjSCGLSpr.CellGroupingEach .sprSht1, "YEARMON|REAL_MED_NAME|CLIENTNAME"
		
		gSetSheetColor mobjSCGLSpr, .sprShtSum
		mobjSCGLSpr.SpreadLayout .sprShtSum, 13, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprShtSum, "YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET"
		mobjSCGLSpr.SetText .sprShtSum, 1, 1, "        합      계"
	    mobjSCGLSpr.SetScrollBar .sprShtSum, 0
	    mobjSCGLSpr.SetBackColor .sprShtSum,"1|1",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeStatic2 .sprShtSum,  "AMT|SUSURATE|SUSU ", -1, -1, 0
	    mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "AMT|SUSU|SUSURATE", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprShtSum, "CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET", true
		
		mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
    End With
    
    pnlTab2.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'기본값 설정
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : mstrFields = vntInParam(i)
				case 2 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				case 3 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 4 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with
	
End Sub
sub sprSht1_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum	
	End with
end sub
Sub sprSht1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub
Sub EndPage()
	set mobjMDCMELECTRICLISTCOMMI = Nothing
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
		.txtYEARMON.value =  Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'Sheet초기화
		.sprSht1.MaxRows = 0
		
		.txtYEARMON.focus
		
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData, vntData1,vntDataPre
	Dim strYEARMON
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strREAL_MED_NAME
	Dim strGFLAG
	Dim strINPUT_MEDFLAG
	Dim vntDataConfirm
	Dim strCONFIRM
   	Dim i, strCols
   	Dim IngsusuColCnt, IngsusuRowCnt
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	Dim strSEARCHGBN
   	Dim strSETENDFLAG
   	Dim lngCnt
	'on error resume next
	with frmThis
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		.sprSht1.MaxRows = 0
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IngsusuColCnt=clng(0)
		IngsusuRowCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		
		strGFLAG = .cmbGROUP.value
		'공동,이월분조회는 거래처를 공백으로 만들어준다.
		If strGFLAG <> "A" Then
			.txtCLIENTNAME.value = ""
		End IF

		vntDataConfirm = mobjMDCMELECTRICLISTCOMMI.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		'확정분
		If IngCOMMITRowCnt > 0 Then
			strSETENDFLAG = "T"
			.ImgConfirmCancel.style.visibility = "visible"
			.imgSetting.style.visibility = "hidden"
			.btnTab2.value = "수수료조회"
			mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"SUSU"
			'2번탭 조회
			vntData1 = mobjMDCMELECTRICLISTCOMMI.SelectRtn_ENDSUSU(gstrConfigXml,IngsusuRowCnt,IngsusuColCnt,strYEARMON)
			
		Else
		'미확정분
			strSETENDFLAG = "F"
			.imgSetting.style.visibility = "visible"
			.ImgConfirmCancel.style.visibility = "hidden"
			mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"SUSU"
			.btnTab2.value = "수수료수정"
			'2번탭 조회
 			vntData1 = mobjMDCMELECTRICLISTCOMMI.SelectRtn_SUSU(gstrConfigXml,IngsusuRowCnt,IngsusuColCnt,strYEARMON)
		End If
		
		if IngsusuRowCnt > 0 then
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,IngsusuColCnt,IngsusuRowCnt,TRUE)
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			AMT_SUM
			If strSETENDFLAG = "F" Then
			gWriteText lblStatus, "미생성내역 " & mlngRowCnt & " 건,수수료수정 " & IngsusuRowCnt & " 건의 자료가 검색" & mePROC_DONE
			Else
			gWriteText lblStatus, "미생성내역 " & mlngRowCnt & " 건,수수료조회 " & IngsusuRowCnt & " 건의 자료가 검색" & mePROC_DONE
			End IF
		else
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   		
   		REAL_MED_CODE_AMT_SUM
	end with   	
End Sub

Sub PreSearchFiledValue (strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME)
	frmThis.txtYEARMON.value = strYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME.value = strCLIENTNAME
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
End Sub

'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntVAT, IntVATSUM, IntSUMATMVAT, IntSUMATMVATSUM
	Dim IntAMT1, IntSUSU, IntSUSUVAT, IntSUMSUSUVAT, IntAMT1SUM, IntSUSUSUM, IntSUSUVATSUM, IntSUMSUSUVATSUM
	With frmThis
		IntAMTSUM = 0
		
		IntAMT1SUM = 0
		IntSUSUSUM = 0
		IntSUSUVATSUM = 0
		IntSUMSUSUVATSUM = 0
		
		'수수료 그리드 합계그리드 값넣기
		For lngCnt = 1 To .sprSht1.MaxRows
			IntAMT1 = 0
			IntSUSU = 0
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME", lngCnt) <> "" THEN
				IntAMT1 = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				IntSUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntAMT1SUM = IntAMT1SUM + IntAMT1
				IntSUSUSUM = IntSUSUSUM + IntSUSU
			END IF
		Next
		
		if .sprSht1.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprShtSum,"AMT",1, IntAMT1SUM
			mobjSCGLSpr.SetTextBinding .sprShtSum,"SUSU",1, IntSUSUSUM
			
		end if
	End With
End Sub


'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub REAL_MED_CODE_AMT_SUM
	Dim lngCnt
	Dim lntB1AMT, lntB1SUSU, IntB1AMTSUM, IntB1SUSUSUM
	Dim lntB2AMT, lntB2SUSU, IntB2AMTSUM, IntB2SUSUSUM
	Dim lntB3AMT, lntB3SUSU, IntB3AMTSUM, IntB3SUSUSUM
	Dim lntB4AMT, lntB4SUSU, IntB4AMTSUM, IntB4SUSUSUM
	Dim lntB6AMT, lntB6SUSU, IntB6AMTSUM, IntB6SUSUSUM
	Dim lntB7AMT, lntB7SUSU, IntB7AMTSUM, IntB7SUSUSUM

	With frmThis
		IntB1AMTSUM = 0
		IntB1SUSUSUM = 0
		
		IntB2AMTSUM = 0
		IntB2SUSUSUM = 0
		
		IntB3AMTSUM = 0
		IntB3SUSUSUM = 0
		
		IntB4AMTSUM = 0
		IntB4SUSUSUM = 0
		
		IntB6AMTSUM = 0
		IntB6SUSUSUM = 0
		
		IntB7AMTSUM = 0
		IntB7SUSUSUM = 0
		
		'수수료 그리드 합계그리드 값넣기
		For lngCnt = 1 To .sprSht1.MaxRows
			lntB1AMT = 0
			lntB1SUSU = 0
			lntB2AMT = 0
			lntB2SUSU = 0
			lntB3AMT = 0
			lntB3SUSU = 0
			lntB4AMT = 0
			lntB4SUSU = 0
			lntB6AMT = 0
			lntB6SUSU = 0
			lntB7AMT = 0
			lntB7SUSU = 0
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00140" THEN
				lntB1AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB1SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB1AMTSUM = IntB1AMTSUM  + lntB1AMT
				IntB1SUSUSUM = IntB1SUSUSUM + lntB1SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00144" THEN
				lntB2AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB2SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB2AMTSUM = IntB2AMTSUM  + lntB2AMT
				IntB2SUSUSUM = IntB2SUSUSUM + lntB2SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00142" THEN
				lntB3AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB3SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB3AMTSUM = IntB3AMTSUM  + lntB3AMT
				IntB3SUSUSUM = IntB3SUSUSUM + lntB3SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00143" THEN
				lntB4AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB4SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB4AMTSUM = IntB4AMTSUM  + lntB4AMT
				IntB4SUSUSUM = IntB4SUSUSUM + lntB4SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00141" THEN
				lntB6AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB6SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB6AMTSUM = IntB6AMTSUM  + lntB6AMT
				IntB6SUSUSUM = IntB6SUSUSUM + lntB6SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00145" THEN
				lntB7AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB7SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB7AMTSUM = IntB7AMTSUM  + lntB7AMT
				IntB7SUSUSUM = IntB7SUSUSUM + lntB7SUSU
			
			END IF
		Next
		
		if .sprSht1.MaxRows >0 Then
			.txtB1AMT.value = IntB1AMTSUM
			.txtB1SUSU.value = IntB1SUSUSUM
			
			.txtB2AMT.value = IntB2AMTSUM
			.txtB2SUSU.value = IntB2SUSUSUM
			
			.txtB3AMT.value = IntB3AMTSUM
			.txtB3SUSU.value = IntB3SUSUSUM
			
			.txtB4AMT.value = IntB4AMTSUM
			.txtB4SUSU.value = IntB4SUSUSUM
			
			.txtB6AMT.value = IntB6AMTSUM
			.txtB6SUSU.value = IntB6SUSUSUM
			
			.txtB7AMT.value = IntB7AMTSUM
			.txtB7SUSU.value = IntB7SUSUSUM
			
		end if
	End With
End Sub


'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strYEARMON
	Dim intCnt
	with frmThis
		'저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht1.MaxRows = 0 Then
   			gErrorMsgBox "상세항목이 없습니다.","확정오류"
   			Exit Sub
   		End If
		
		'마스터 데이터를 가져 온다.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'처리 업무객체 호출
		strYEARMON = .txtYEARMON.value
		for intCnt = 1 to .sprSht1.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht1,"SAVESET",intCnt, "T"	
			Call sprSht1_Change (13,intCnt)
		next
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET")
		intRtn = mobjMDCMELECTRICLISTCOMMI.ProcessRtn(gstrConfigXml, strMasterData,vntData,strYEARMON)

		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			'InitPageData
			gOkMsgBox "승인처리 되었습니다.","확인"
			SelectRtn
   		end if
   	end with
End Sub
Sub sprSht1_change(ByVal Col,ByVal Row)
AMT_SUM
mobjSCGLSpr.CellChanged frmThis.sprSht1, Col,Row
End Sub	
'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON

	with frmThis
		intSelCnt = 0
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		If .sprSht1.MaxRows = 0 Then
   			gErrorMsgBox "상세항목이 없습니다.","확정취소오류"
   			Exit Sub
   		End If
   		
   		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
   		strYEARMON = .txtYEARMON.value
   		intRtn = mobjMDCMELECTRICLISTCOMMI.SelectRtn_CANCEL(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
   		if mlngRowCnt > 0 then
   			gErrorMsgBox "거래명세서가 생성된 데이터는 확정취소가 안됩니다.","확정취소오류"
   			Exit Sub
   		end if
		
		intRtn = gYesNoMsgbox("확정취소 하시겠습니까?","확정취소 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		strYEARMON = .txtYEARMON.value
	
		intRtn = mobjMDCMELECTRICLISTCOMMI.DeleteRtn(gstrConfigXml,strYEARMON)
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gOkMsgBox  strYEARMON & " 의 자료가 확정취소 되었습니다.","확인"
			SelectRtn
   		End IF
	End with
	err.clear	
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<P dir="ltr" style="MARGIN-RIGHT: 0px">
				<TABLE id="tblForm" style="WIDTH: 1040px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
					width="1040" border="0">
					<TR>
						<TD>
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
								border="0">
								<TR>
									<td align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;공중파&nbsp;수수료 승인처리</td>
											</tr>
										</table>
									</td>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 108px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="108" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
												<TD><!--<IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose">--></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">년 
													월</TD>
												<TD class="SEARCHDATA" style="WIDTH: 181px"><INPUT class="INPUT" id="txtYEARMON" title="년월조회" style="WIDTH: 64px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="5" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 74px; CURSOR: hand">조회구분</TD>
												<TD class="SEARCHDATA" style="WIDTH: 174px"><SELECT class="INPUT" id="cmbGROUP" title="그룹구분" style="WIDTH: 99px" name="cmbGROUP">
														<OPTION value="A" selected>전체</OPTION>
														<OPTION value="G">공동분</OPTION>
														<OPTION value="N">이월분</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 77px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">광고주명</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 384px; HEIGHT: 22px"
														type="text" maxLength="100" size="58" name="txtCLIENTNAME"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px">
							<TABLE id="tblTab" style="WIDTH: 1040px; HEIGHT: 5px" cellSpacing="0" cellPadding="0" width="787"
								border="0">
								<TR>
									<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
											type="button" size="20" value="수수료수정" name="btnTab2">
									</TD>
									<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
											height="20" alt="확정합니다." src="../../../images/imgSetting.gIF" width="54" align="right"
											border="0" name="imgSetting"></TD>
									<TD><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gif'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gif'"
											height="20" alt="확정취소합니다." src="../../../images/ImgConfirmCancel.gIF" border="0"
											name="ImgConfirmCancel"></TD>
								</TR>
								<TR class="TABBAR">
									<TD colSpan="3"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 600px" vAlign="top" align="center">
							<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 1040px; HEIGHT: 576px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="15240">
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
								<OBJECT id="sprShtSum" style="WIDTH: 1040px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
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
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
										<TABLE class="DATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="LABEL" width="90">본사대행금액</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB1AMT" title="한국방송광고공사본사대행금액" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB1AMT"></TD>
												<TD class="LABEL" width="90">본사수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB1SUSU" title="한국방송광고공사본사수수료총액" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB1SUSU"></TD>
												<TD class="LABEL" width="90">부산대행금액</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB2AMT" title="한국방송광고공사부산지사대행금액" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB2AMT"></TD>
												<TD class="LABEL" width="90">부산수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB2SUSU" title="한국방송광고공사부산지사수수료총액" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB2SUSU"></TD>
											</TR>
											<TR>
												<TD class="LABEL" width="90">대구대행금액</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB3AMT" title="한국방송광고공사대구지사대행금액" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB3AMT"></TD>
												<TD class="LABEL" width="90">대구수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB3SUSU" title="한국방송광고공사대구지사수수료총액" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB3SUSU"></TD>
												<TD class="LABEL" width="90">대전대행금액</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB4AMT" title="한국방송광고공사대전지사대행금액" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB4AMT"></TD>
												<TD class="LABEL" width="90">대전수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB4SUSU" title="한국방송광고공사대전지사수수료총액" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB4SUSU"></TD>
											</TR>
											<TR>
												<TD class="LABEL" width="90">광주대행금액</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB6AMT" title="한국방송광고공사광주지사대행금액" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB6AMT"></TD>
												<TD class="LABEL" width="90">광주수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB6SUSU" title="한국방송광고공사광주지사수수료총액" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB6SUSU"></TD>
												<TD class="LABEL" width="90">전북대행금액</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB7AMT" title="한국방송광고공사전북지사대행금액" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB7AMT"></TD>
												<TD class="LABEL" width="90">전북수수료</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB7SUSU" title="한국방송광고공사전북지사수수료총액" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB7SUSU"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
					</TR>
				</TABLE>
			</P>
		</form>
	</body>
</HTML>
