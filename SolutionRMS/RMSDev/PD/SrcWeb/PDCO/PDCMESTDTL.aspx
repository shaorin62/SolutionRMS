<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMESTDTL.aspx.vb" Inherits="PD.PDCMESTDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>견적 내역관리</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 견적내역관리 화면(PDCMESTDTL)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPREESTDTL.aspx
'기      능 : 가견적 내역 등록 및 확정
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/16 By Tae Ho Kim
'			 2) 
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
Dim mobjPDCMPREESTDTL '공통코드, 클래스
Dim mstrPROCESS
Dim mstrPROCESS2 '조회상태이면 true 신규상태이면 false
Dim mstrCheck
Dim mobjMDLOGIN
Dim mobjMDCMEMP
Dim mobjPDCMGET
CONST meTAB = 9
mstrPROCESS = TRUE
mstrPROCESS2 = TRUE
mstrCheck = True

'=============================
' 이벤트 프로시져 
'=============================
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
	DataClean
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

'Sub imgDelete_onclick
'	gFlowWait meWAIT_ON
'	DeleteRtn
'	gFlowWait meWAIT_OFF
'End Sub

Sub imgSave_onclick ()
	with frmThis
		
		if frmThis.txtENDFLAG.value = "T" Then
		
			gErrorMsgBox "거래명세서가 작성되어 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		End If
		'if frmThis.txtENDFLAGEXE.value = "T" Then
		'	gErrorMsgBox "외주비 지출내역이 작성되어 저장이 불가능 합니다.","저장안내!"
		'	Exit Sub
		'End If
			gFlowWait meWAIT_ON
			if .txtPREESTNO.value = "" Then
				ProcessRtn
			Else
				ProcessRtn_OLD
			End If
			gFlowWait meWAIT_OFF
		
		
	End with
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

Sub imgRowDel_onclick
	
	if frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "거래명세서가 작성되어 행삭제가 불가능 합니다.","삭제안내!"
		Exit Sub
	End If
	'if frmThis.txtENDFLAGEXE.value = "T" Then
	'	gErrorMsgBox "외주비 지출내역이 작성되어 행삭제가 불가능 합니다.","저장안내!"
	'	Exit Sub
	'End If
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowAdd_onclick ()
	if frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "거래명세서가 작성되어 행추가가 불가능 합니다.","저장안내!"
		Exit Sub
	End If
	'if frmThis.txtENDFLAGEXE.value = "T" Then
	'	gErrorMsgBox "외주비 지출내역이 작성되어  행추가가 불가능 합니다.","저장안내!"
	'	Exit Sub
	'End If
	call sprSht_Keydown(meINS_ROW, 0)
	
End Sub

Sub ImgExeList_onclick
Dim strJOBNO	
Dim vntInParams
Dim vntRet
	with frmThis
		strJOBNO = Trim(.txtJOBNO.value)
		vntInParams = array(strJOBNO)
		vntRet = gShowModalWindow("PDCMEXELISTPOP.aspx",vntInParams , 1060,780)
		SelectRtn
	End with
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'입력용
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalEndar,"txtPRINTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtPRINTDAY_onchange
	gSetChange
End Sub
Sub imgCalEndarAGREE_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtAGREEYEARMON,frmThis.imgCalEndarAGREE,"txtAGREEYEARMON_onchange()"
		gSetChange
	end with
End Sub
Sub txtAGREEYEARMON_onchange
	gSetChange
End Sub
'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
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
		IF Col = 7 Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",.sprSht.ActiveRow)
			vntData = mobjPDCMGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",strCodeName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntData(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntData(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntData(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntData(4,0)			
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
		ElseIf  Col = 11 Then
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
   		ElseIf Col = 12 Then
   			strQTY		= mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",.sprSht.ActiveRow)
			strPRICE   = mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",.sprSht.ActiveRow)
			strAMT = strQTY * strPRICE
			mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT	
			Call SUSUAMT_CHANGEVALUE(Row)
			BUDGET_AMT_SUM
		'금액로직	
   		ElseIf  Col = 13 Then
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
   		Elseif Col = 10 Then
   			Call SUSUAMT_CHANGEVALUE2
   			BUDGET_AMT_SUM
		END IF
	end with
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
Sub SUSUAMT_CHANGEVALUE(ByVal Row)
Dim strAMT,strCOMMIFLAG
Dim strSUSURATE
	with frmThis
		strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		strSUSURATE = .txtSUSURATE.value
		if strCOMMIFLAG = "1" Then
			if strSUSURATE = "" then
				strSUSURATE = 0
			end if
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * strSUSURATE /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
	End with
End SUb
Sub SUSUAMT_CHANGEVALUE2
Dim intCnt
Dim strAMT,strCOMMIFLAG
Dim strSUSURATE
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
		strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
		strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		strSUSURATE = .txtSUSURATE.value
		if strCOMMIFLAG = "1" Then
			if strSUSURATE = "" then
				strSUSURATE = 0
			end if
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * strSUSURATE /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
	Next
	
	End with
End Sub
Sub txtSUSURATE_onchange
	with frmThis
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSURATE  
	End with
End Sub

Sub BUDGET_AMT_SUM
	'총합계 변수
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	Dim lngSUSU
	'수수료 계산 변수
	Dim intCnt,intSUSU,intSUSUSUM 
	'commition 계산 변수
	Dim intCnt1,intCOM,intCOMSUM 
	'noncommition 계산변수
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
		For intCnt = 1 To .sprSht.MaxRows
		
			intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
			intSUSUSUM = intSUSUSUM + intSUSU
			
		Next
		.txtSUSUAMT.value = intSUSUSUM
		
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		IntAMTSUM = IntAMTSUM + intSUSUSUM
		.txtSUMAMT.value = IntAMTSUM
		
		For intCnt1 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG", intCnt1) = "1" Then
				
				intCOM = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intCOMSUM = intCOMSUM + intCOM
			Else
				
				intNON = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intNONSUM = intNONSUM + intNON
			end if
		Next
		.txtCOMMITION.value = intCOMSUM
		.txtNONCOMMITION.value = intNONSUM
		
		txtSUSUAMT_onblur
		txtCOMMITION_onblur
		txtSUMAMT_onblur
		txtNONCOMMITION_onblur
		
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
		IF Col = 7 Then
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
	dim vntRet, vntInParams
	with frmThis
	
		IF Col = 6 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtSUSURATE.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

'=============================
' UI업무 프로시져 
'=============================
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
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
Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	'서버업무객체 생성	
	dim vntInParam
	dim intNo,i
	
	set mobjPDCMPREESTDTL	= gCreateRemoteObject("cPDCO.ccPDCOPREESTDLT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "260px"
	pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
				case 1 : frmThis.txtJOBNO.value = vntInParam(i)
			end select
		next
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis

		
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN"
		mobjSCGLSpr.SetHeader .sprSht,		  "가견적번호|순번|대분류|중분류|견적항목코드|견적항목명|견적명|내역|커미션|수량|단가|금액|수수료금액|저장구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|   0|     8|    12|        8 |2|        15|12    |  20|     6|  12|  13|13  |10         |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMCODENAME|STD|FAKENAME", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "DIVNAME|CLASSNAME|ITEMCODE"
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|ITEMCODESEQ|ITEMCODESEQ|GBN", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME|CLASSNAME|FAKENAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE|ITEMCODESEQ",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0

	InitPageData	
	SelectRtn
	If .txtENDFLAG.value = "T" Then
	Else
		.txtPREESTNAME.value = .txtJOBNAME.value 
	End If
	
	End With
End Sub



Sub EndPage()
	set mobjPDCMPREESTDTL = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub
'-----------------------------
' 확정 및 확정취소 처리
'-----------------------------	
Sub imgSetting_onclick
Dim intRtnConfirm
Dim intRtn
	intRtnConfirm = gYesNoMsgbox("자료를 확정 하시겠습니까?","자료확정 확인")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_Confirm(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_Confirm") then
				gErrorMsgBox " 자료가 확정 되었습니다.","확정안내" 
			End If
			ESTCONFIRM_Search
	End with
End Sub


Sub ImgConfirmCancel_onclick
Dim intRtnConfirm
Dim intRtn
	intRtnConfirm = gYesNoMsgbox("자료를 확정취소 하시겠습니까?","자료확정취소 확인")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_ConfirmCancel(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_ConfirmCancel") then
				gErrorMsgBox " 자료가 확정취소 되었습니다.","확정취소안내" 
			End If
			ESTCONFIRM_Search
	End with
End Sub
Sub txtPREESTNAME_onchange
	gSetChange
End Sub

'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		
		.txtPRINTDAY.value = gNowDate
		.sprSht.MaxRows = 0
		ESTCONFIRM_Search2
	End with
	'DataNewClean
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'청구확정 조회
Sub ESTCONFIRM_Search
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
				.imgSetting.disabled = true
				.ImgConfirmCancel.disabled = false
				
			Else
				.imgSetting.disabled = false
				.ImgConfirmCancel.disabled = true
				
			End if
   		end if
	end with
End Sub
Sub ESTCONFIRM_Search2
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
			
				.imgExeList.style.visibility = "visible"
			Else
			
				.imgExeList.style.visibility = "hidden"
			End if
   		end if
	end with
End Sub
Sub DataNewClean
	with frmThis
	.txtCREDAY.value = ""
	.cmbGROUPGBN.selectedIndex  = -1
	End with
End Sub
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub
'------------------------------------------
' 가견적 번호가 있는 경우 저장처리
'------------------------------------------
Sub ProcessRtn_OLD ()
    Dim intRtn
  	Dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strDELCODE
	Dim intRtnSave
	Dim vntData2
	Dim strPREESTNO
	Dim intHDR
	with frmThis
	'On error resume next
  		'데이터 Validation
		if DataValidation =false then exit sub
		If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "견적명을 입력하십시오.","저장안내"
			Exit Sub
		End If
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|FAKENAME|SUSUAMT")
		'처리 업무객체 호출
		strMasterData = gXMLGetBindingData (xmlBind)
		
		
		if  not IsArray(vntData) then 
				If gXMLIsDataChanged (xmlBind) Then 
					intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
					if not gDoErrorRtn ("ProcessRtn_HDR") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " 자료가" & intHDR & " 건 저장" & mePROC_DONE,"저장안내" 
						SelectRtn
					End If
				Else
					gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				End If
			Exit Sub
		End If
		
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF

			intRtn = mobjPDCMPREESTDTL.ProcessRtn(gstrConfigXml,vntData,strPREESTNO)
				
		if not gDoErrorRtn ("ProcessRtn") then
			intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
			if not gDoErrorRtn ("ProcessRtn_HDR") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " 자료가" & intHDR & " 건 저장" & mePROC_DONE,"저장안내" 
				SelectRtn
			End If
  		end if
 	end with
End Sub
'------------------------------------------
' 즉시 확정 견적 저장시 (가견적번호가 없는 경우)
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

strMasterData = gXMLGetBindingData (xmlBind)
if DataValidation =false then exit sub
strPREESTNO = ""
	with frmThis
	If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "견적명을 입력하십시오.","저장안내"
			Exit Sub
	End If
	If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
	End IF
	vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|FAKENAME|SUSUAMT")
	if  not IsArray(vntData) then 
		gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
		exit sub
	End If
	If .txtAGREEYEARMON.value = "" then
		gErrorMsgBox "견적확정일을 기입하십시오.","저장안내"
		exit sub
	End If
	strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_HDRLESS(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON)
		if not gDoErrorRtn ("ProcessRtn_HDRLESS") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가" & intRtn & " 건 저장" & mePROC_DONE,"저장안내" 
			strJOBNO = .txtJOBNO.value
			intSearchRtn =  mobjPDCMPREESTDTL.SelectRtn_PREESTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
			
			.txtPREESTNO.value = intSearchRtn(0,1)
			SelectRtn
		End If
	End with

End Sub


Sub DelProc
Dim intHDR
Dim strMasterData
Dim strPREESTNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		strPREESTNO = .txtPREESTNO.value
		intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
				if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
					'SelectRtn
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
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
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

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	Dim strJOBCODE
	With frmThis
		strCODE = .txtPREESTNO.value
		strJOBCODE = .txtJOBNO.value
		IF strCODE = ""  THEN 
			
			
			IF not SelectRtn_HeadLess (strJOBCODE) Then Exit Sub
			
		Else
			IF not SelectRtn_Head (strCODE) Then Exit Sub
			'쉬트 조회
			CALL SelectRtn_Detail (strCODE)
			txtSUSUAMT_onblur
			txtCOMMITION_onblur
			txtSUMAMT_onblur
			txtNONCOMMITION_onblur
		End If
		
		If .txtENDFLAG.value = "T" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
		Else
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
		End If
		'If .txtAGREEYEARMON.value = "" Then
		'	.imgExeList.style.visibility = "hidden"
		'Else
		'	.imgExeList.style.visibility = "visible"
		'End If
	End With
	
	
End Sub
'기존내역이 없을 경우 조회
Function SelectRtn_HeadLess (ByVal strJOBCODE)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_HeadLess = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDRLESS(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBCODE)
	
	IF not gDoErrorRtn ("SelectRtn_HeadLess") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 견적에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_HeadLess = True
		End IF
	End IF
End Function
'기존내역이 있을 경우 조회
Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 가견적에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function


'예산 테이블 조회
Function SelectRtn_Detail (ByVal strCODE)
	dim vntData
	Dim intCnt
	Dim strRows
	
	'on error resume next
	'초기화
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt = 1 To .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "T" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,6,7,true
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,intCnt,6,7,true
					End If
				Next
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		End with
	End IF
End Function



'****************************************************************************************
'이전 검색어를 담아 놓는다.
'****************************************************************************************
Sub PreSearchFiledValue (strTBRDSTDATE,strTBRDEDDATE, strCAMPAIGN_CODE, strCAMPAIGN_NAME, strCLIENTCODE, strCLIENTNAME)
	With frmThis
		.txtTBRDSTDATE1.value = strTBRDSTDATE
		.txtTBRDEDDATE1.value = strTBRDEDDATE
		.txtCAMPAIGN_CODE1.value = strCAMPAIGN_CODE
		.txtCAMPAIGN_NAME1.value = strCAMPAIGN_NAME
		.txtCLIENTCODE1.value = strCLIENTCODE
		.txtCLIENTNAME1.value = strCLIENTNAME
	End With
End Sub







Sub DataClean
	with frmThis
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = ""
		.txtCLIENTCODE.value = ""
		.txtCLIENTNAME.value = ""
		.txtCLIENTSUBCODE.value = ""
		.txtCLIENTSUBNAME.value = ""
		.txtSUBSEQ.value = ""
		.txtSUBSEQNAME.value = ""
		.txtCPDEPTCD.value = ""
		.txtCPDEPTNAME.value = ""
		.txtCPEMPNO.value = ""
		.txtCPEMPNAME.value = ""
		.txtMEMO.value = ""
		.cmbGROUPGBN.value = "1"
		.txtCREDAY.value = gNowDate
		.sprSht.MaxRows = 0
	End With
End Sub
Sub sprSht_Keydown(KeyCode, Shift)
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	Dim intRtn
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 14 AND frmThis.txtENDFLAG.value <> "T" Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
		DefaultValue
		end if
	Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		if intRtn = meINS_ROW then
			'DefaultValue
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PREESTNO",frmThis.sprSht.ActiveRow, frmThis.txtPREESTNO.value 
		elseif intRtn = meDEL_ROW then
			DeleteRtn
		end if
'		Select Case intRtn
'			Case meINS_ROW: DefaultValue
'			Case meDEL_ROW: DeleteRtn
'		End Select
	End if
End Sub
Sub DefaultValue
	with frmThis
	mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
	End With
End Sub
'ProjectNO 조회팝업
Sub ImgPROJECTNO1_onclick
	Call PONO_POP()
End Sub
'실제 데이터List 가져오기
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' 코드명 표시
			.txtCLIENTNAME1.focus()					' 포커스 이동
     	end if
	End with
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'자료삭제
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2,lngCnt
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
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
		'PREESTNO,ITEMCODESEQ
		'선택된 자료를 끝에서 부터 삭제
		lngCnt =0
		intRtn2 = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)) <> ""  Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",vntData(i)) = "T"  Then
					gErrorMsgBox "선택한 자료중 외주정산 처리가 되어있어 삭제가 불가능 합니다.","삭제안내"
				Exit Sub
				End iF
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",vntData(i))
				strITEMCODESEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)))
				intRtn2 = mobjPDCMPREESTDTL.DeleteRtn(gstrConfigXml,strPREESTNO, strITEMCODESEQ)
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

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;청구 견적관리</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">JOB명</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="제작건명" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtJOBNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">매체부문</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="매체부문" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtJOBGUBN"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">매체분류</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="매체분류" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtCREPART"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">광고주</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="광고주" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">사업부</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="사업부" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCLIENTSUBNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">브랜드</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUBSEQNAME"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 72px" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;&nbsp;청구 견적 작성</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgExeList" onmouseover="JavaScript:this.src='../../../images/imgExeListOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExeList.gIF'"
																height="20" alt="외주비정산 프로그램을 호출합니다." src="../../../images/imgExeList.gIF" border="0"
																name="imgExeList"></TD>
														<TD><IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'"
																height="20" alt="자료입력을 위해 행을추가합니다." src="../../../images/imgRowAdd.gIF" border="0"
																name="imgRowAdd"></TD>
														<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'"
																height="20" alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF" border="0" name="imgRowDel"></TD>
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></td>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<tr height="5">
											<td></td>
										</tr>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="1040" border="0"
										align="LEFT">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">견적 코드</TD>
											<TD class="DATA" width="230"><INPUT dataFld="PREESTNO" class="NOINPUT_L" id="txtPREESTNO" title="가견적코드" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtPREESTNO"></TD>
											<TD class="LABEL" style="WIDTH: 92px; CURSOR: hand" width="92">견적명</TD>
											<TD class="DATA" width="260"><INPUT dataFld="PREESTNAME" class="NOINPUT_L" id="txtPREESTNAME" title="가견적명" style="WIDTH: 255px; HEIGHT: 22px"
													accessKey="M" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtPREESTNAME"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAGREEYEARMON,'')">견적확정일</TD>
											<TD class="DATA"><INPUT dataFld="AGREEYEARMON" class="INPUT" id="txtAGREEYEARMON" title="견적합의일" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtAGREEYEARMON"><IMG id="imgCalEndarAGREE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndarAGREE"><INPUT dataFld="ENDFLAG" id="txtENDFLAG" style="WIDTH: 40px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtENDFLAG"><INPUT dataFld="ENDFLAGEXE" id="txtENDFLAGEXE" style="WIDTH: 40px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtENDFLAGEXE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">제작수수료</TD>
											<TD class="DATA" width="230"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="제작수수료" style="WIDTH: 224px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="32" name="txtSUSUAMT">
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand" align="right" width="94">Commition</TD>
											<TD class="DATA" width="260"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="commition 계" style="WIDTH: 256px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand" align="right" width="80">합계</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="총합계금액" style="WIDTH: 272px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUMAMT"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">수수료율</TD>
											<TD class="DATA"><INPUT dataFld="SUSURATE" class="INPUT_R" id="txtSUSURATE" style="WIDTH: 200px; HEIGHT: 22px"
													accessKey=",NUM,M" dataSrc="#xmlBind" type="text" size="28" name="txtSUSURATE">&nbsp;(%)
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px">Non Commition</TD>
											<TD class="DATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="noncommition 계"
													style="WIDTH: 256px; HEIGHT: 22px" accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37"
													name="txtNONCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" width="80">견적서 출력</TD>
											<TD class="DATA"><INPUT class="INPUT" id="txtPRINTDAY" title="견적서발행일" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" type="text" maxLength="10" size="10" name="txtPRINTDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndar">&nbsp;&nbsp; <IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="견적서 를 인쇄합니다." src="../../../images/imgPrint.gIF"
													width="54" align="absMiddle" border="0" name="imgPrint">&nbsp;<INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtJOBNO"><INPUT dataFld="CREDAY" id="txtCREDAY" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCREDAY"><INPUT dataFld="CLIENTSUBCODE" id="txtCLIENTSUBCODE" style="WIDTH: 16px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="hidden" size="1" name="txtCLIENTSUBCODE"><INPUT dataFld="CLIENTCODE" id="txtCLIENTCODE" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCLIENTCODE"><INPUT dataFld="SUBSEQ" id="txtSUBSEQ" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtSUBSEQ"></TD>
										</TR>
										<TR>
											<TD class="LABEL">비고</TD>
											<TD class="DATA" colSpan="5"><INPUT dataFld="MEMO" id="txtMEMO" style="WIDTH: 950px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="text" maxLength="255" size="152" name="txtMEMO"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27464">
											<PARAM NAME="_ExtentY" VALUE="11721">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
