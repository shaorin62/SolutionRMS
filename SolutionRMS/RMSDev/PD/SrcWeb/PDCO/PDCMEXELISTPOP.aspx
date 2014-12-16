<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMEXELISTPOP.aspx.vb" Inherits="PD.PDCMEXELISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>외주정산 등록</title>
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
Const meTAB = 9
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMEXE, mobjPDCMGET
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
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "완료건 처리는 완료일 삭제후 가능합니다.","처리안내!"
		Exit Sub
	End If
	
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "검색된 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
	
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn_ALL
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "완료건 처리는 완료일 삭제후 가능합니다.","처리안내!"
		Exit Sub
	End If
	End with
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgAddRow_onclick ()
with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "완료건 처리는 완료일 삭제후 가능합니다.","처리안내!"
		Exit Sub
	End If
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "검색된 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
End with
call sprSht_Keydown(meINS_ROW, 0)
End Sub
Sub imgDelRow_onclick()
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "완료건 처리는 완료일 삭제후 가능합니다.","처리안내!"
		Exit Sub
	End If
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub ImgAccInput_onclick()
	Dim vntInParams
	Dim vntRet
	Dim vntData
	Dim strGUBN
	with frmThis
	
	
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "제작번호 조회후 입력 가능 합니다.","처리안내!"
		Exit Sub
	End If
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	If .txtENDDAY.value <> "" Then
	strGUBN = "END"
	Else
	strGUBN = ""
	End If
	vntData = mobjPDCMEXE.SelectRtn_ACCEXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		If mlngRowCnt = 0 Then
			ProcessRtn_SUB
		End If
		vntInParams = array(.txtJOBNO.value,strGUBN)
		vntRet = gShowModalWindow("PDCMACCLIST.aspx",vntInParams , 788,540)
		SelectRtn
	End If
	
	
	End with
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
'JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME
Dim strSEQ
Dim strJOBNO
Dim strPREESTNO
Dim strITMECODESEQ
Dim strITEMCODE
Dim strITEMCLASS
Dim strITEMNAME
Dim strADDFLAG
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 12 Then
		
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
		DefaultValue
		end if
	Else
	strSEQ = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"SORTSEQ",frmThis.sprSht.ActiveRow) 
	strSEQ = strSEQ+1
	strJOBNO = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"JOBNO",frmThis.sprSht.ActiveRow) 
	strPREESTNO = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"PREESTNO",frmThis.sprSht.ActiveRow) 
	strITMECODESEQ = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCODESEQ",frmThis.sprSht.ActiveRow)  
	strITEMCODE = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCODE",frmThis.sprSht.ActiveRow)  
	strITEMCLASS = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCLASS",frmThis.sprSht.ActiveRow) 
	strITEMNAME = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMNAME",frmThis.sprSht.ActiveRow) 
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: Call DefaultValue(strSEQ,strJOBNO,strPREESTNO,strITMECODESEQ,strITEMCODE,strITEMCLASS,strITEMNAME)
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub

Sub DefaultValue(ByVal strSEQ,ByVal strJOBNO,ByVal strPREESTNO,ByVal strITMECODESEQ,ByVal strITEMCODE, ByVal strITEMCLASS,ByVal strITEMNAME)
	Dim intCnt
	with frmThis

	
	For intCnt = 1 To .sprSht.MaxRows
		if cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",intCnt)) >= strSEQ Then 
			mobjSCGLSpr.SetTextBinding .sprSht,"SORTSEQ",intCnt,cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",intCnt))+1
		End if
	Next
	mobjSCGLSpr.SetTextBinding .sprSht,"SORTSEQ",.sprSht.ActiveRow, strSEQ
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, strJOBNO
	mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, strPREESTNO
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",.sprSht.ActiveRow, strITMECODESEQ
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",.sprSht.ActiveRow, strITEMCODE
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCLASS",.sprSht.ActiveRow, strITEMCLASS
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMNAME",.sprSht.ActiveRow, strITEMNAME
	mobjSCGLSpr.SetTextBinding .sprSht,"ADDFLAG",.sprSht.ActiveRow, "A"
	mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",.sprSht.ActiveRow,"코드선택"
	mobjSCGLSpr.SetTextBinding .sprSht,"INCOMCODE",.sprSht.ActiveRow,"사업소득(3,3%)"
	mobjSCGLSpr.SetSheetSortUser  .sprSht,3,1
	.txtCLIENTNAME.focus()
	.sprSht.Focus()	
	End with 
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
		intRtn = mobjPDCMEXE.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		
		vntData = mobjPDCMEXE.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
	
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
				vntDataTemp = mobjPDCMEXE.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
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
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		strSPONSOR = "Y"
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELECSPON", strSPONSOR) 
		vntRet = gShowModalWindow("MDCMTRANSCUSTPOP.aspx",vntInParams , 413,425)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = vntRet(1,0)		  ' Code값 저장
			.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			IF vntRet(3,0) = "완료" THEN
				window.event.keyCode = meEnter
				txtTRANSNO_onkeydown
			ELSE
				.txtTRANSNO.value = ""
			END IF
			gSetChangeFlag .txtCLIENTCODE             ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strSPONSOR
   		
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strSPONSOR = "Y"
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtTRANSNO.value,.txtCLIENTNAME1.value,"ALL","trans", "ELECSPON", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(0,0)
					.txtCLIENTNAME1.value = vntData(1,0)
				Else
					Call CLIENTCODE_POP()
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
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELECSPON", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)   ' Code값 저장
					.txtTRANSNO.value = vntData(1,0)		' 코드명 표시
					.txtCLIENTCODE.value = vntData(2,0)     ' 코드명 표시
					.txtCLIENTNAME1.value = vntData(3,0)    ' 코드명 표시
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
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
		strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
		
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "ELECSPON") '<< 받아오는경우
		vntRet = gShowModalWindow("MDCMTRANSPOP.aspx",vntInParams , 423,435	)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code값 저장
			.txtTRANSNO.value = vntRet(1,0)  ' 코드명 표시
			.txtCLIENTCODE.value = vntRet(2,0)  ' 코드명 표시
			.txtCLIENTNAME1.value = vntRet(3,0)  ' 코드명 표시
			'Call SelectRtn ()
     	end if
	End with
	gSetChange
End Sub


'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtENDDAY,frmThis.imgCalEndar,"txtENDDAY_onchange()"
		gSetChange
	end with
End Sub

'완료일자
Sub txtENDDAY_onchange
	gSetChange
End Sub




'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------

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

Sub txtDEMANDAMT_onfocus
	with frmThis
		.txtDEMANDAMT.value = Replace(.txtDEMANDAMT.value,",","")
	end with
End Sub
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub

Sub txtESTAMT_onfocus
	with frmThis
		.txtESTAMT.value = Replace(.txtESTAMT.value,",","")
	end with
End Sub
Sub txtESTAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTAMT,0,true)
	end with
End Sub

Sub txtPAYMENT_onfocus
	with frmThis
		.txtPAYMENT.value = Replace(.txtPAYMENT.value,",","")
	end with
End Sub

Sub txtPAYMENT_onblur
	with frmThis
		call gFormatNumber(.txtPAYMENT,0,true)
	end with
End Sub

Sub txtINCOM_onfocus
	with frmThis
		.txtINCOM.value = Replace(.txtINCOM.value,",","")
	end with
End Sub
Sub txtINCOM_onblur
	with frmThis
		call gFormatNumber(.txtINCOM,0,true)
	end with
End Sub

Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'txtACCAMT
Sub txtACCAMT_onfocus
	with frmThis
		.txtACCAMT.value = Replace(.txtACCAMT.value,",","")
	end with
End Sub
Sub txtACCAMT_onblur
	with frmThis
		call gFormatNumber(.txtACCAMT,0,true)
	end with
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
Dim vntInParams
Dim vntRet
Dim strCONTRACTNO
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If Col = 17 AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row) <> "" Then
				strCONTRACTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row)	
				vntInParams = array(strCONTRACTNO)
				vntRet = gShowModalWindow("PDCMCONTRACTPOP.aspx",vntInParams , 1060,900)
			End If
		end if
	end with
end sub

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
				IF Col = 13 Then
					strCode = ""
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",.sprSht.ActiveRow)
					vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName)
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					Else
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End If
					.txtCLIENTNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
					.sprSht.Focus	
					mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
				ELSEIF Col = 14 then
					Payment_changevalue
				END IF
	end with
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
'외주처 버튼클릭
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	with frmThis
	
		IF Col = 12 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				'SUSUAMT_CHANGEVALUE2
				'BUDGET_AMT_SUM
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtCLIENTNAME.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

'스프레드의 항목이 변할시 어떠한 함수를 태우고자 할때 사용
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = 13 Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtCLIENTNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub

Sub Payment_changevalue
Dim intCnt
Dim lngAMT
Dim lngAMTSUM
Dim lngDEMANDAMT
Dim lngRATE
Dim lngACCAMT
with frmThis
	lngAMT= 0
	lngAMTSUM = 0
	For intCnt = 1 To .sprSht.MaxRows
		lngAMT = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"ADJAMT",intCnt))
		lngAMTSUM = lngAMTSUM + lngAMT
	Next
	lngACCAMT = Replace(.txtACCAMT.value,",","")
	.txtPAYMENT.value = lngAMTSUM
	lngDEMANDAMT = Replace(.txtDEMANDAMT.value,",","")
	If lngACCAMT = "" Then
		lngACCAMT = 0
	End If
	If lngAMTSUM = "" Then
		lngAMTSUM = 0
	End If
	If lngDEMANDAMT = "" Then
		lngDEMANDAMT = 0
	End If
	.txtINCOM.value = lngDEMANDAMT - (lngAMTSUM+lngACCAMT)
	If lngDEMANDAMT = 0 Then
	lngRATE = 0
	Else
	
	lngRATE = gRound(((lngDEMANDAMT-(lngAMTSUM+lngACCAMT))/lngDEMANDAMT)*100,2)
	End if
	.txtRATE.value = lngRATE
	txtPAYMENT_onblur
	txtINCOM_onblur
	
 
End with

End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	Dim strComboList
	Dim strComboList2
	'서버업무객체 생성	
	set mobjPDCMEXE	= gCreateRemoteObject("cPDCO.ccPDCOEXE")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "268px"
	pnlTab1.style.left= "7px"
	

	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		strComboList =  "코드선택" & vbTab & "세금계산서(10%)" & vbTab & "세금계산서불공제" & vbTab & "세금계산서영세율" & vbTab & "계산서" & vbTab & "INVOICE" & vbTab & "사업소득(3,3%)" & vbTab & "기타소득(22%)" & vbTab & "기타소득(필요경비80%)" & vbTab & "비거주자(제한세율)" & vbTab & "비거주자"
		strComboList2 =  "사용안함"
		'******************************************************************
		'거래명세서 생성 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 13
		mobjSCGLSpr.AddCellSpan  .sprSht, 11, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,   "JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|BTN|OUTSNAME|ADJAMT|STD|VOCHNO|CONTRACTNO|VATCODE|INCOMCODE|ADJDAY|ADDFLAG|SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		   "제작번호|견적번호|순번|외주항목순번|외주항목코드|대분류|견적항목|수량|단가|금액|외주처코드|외주처|지급액|내역|전표번호|계약서번호|세무코드|소득구분코드|정산일|삽입구분|번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "       0|       0|   4|           0|           0|    10|14      |7   |9   |11  |       8|2|17    |11    |20  |0       |11        |14        |14          |9     |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SORTSEQ|QTY|PRICE|AMT|ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "ITEMCODESEQ|ITEMCLASS|ITEMNAME", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "OUTSNAME|STD", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SORTSEQ|QTY|PRICE|AMT|ADJDAY|CONTRACTNO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "OUTSCODE|CONTRACTNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "JOBNO|PREESTNO|ITEMCODESEQ|ITEMCODE|ADDFLAG|SEQ|VOCHNO|INCOMCODE", true
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,18,18,-1,-1,strComboList
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,19,19,-1,-1,strComboList2
		
		
		 		
    End With    
	pnlTab1.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData	
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'기본값 설정
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
			case 0 : .txtJOBNO1.value = vntInParam(i)
				
				'case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'조회추가필드
				'case 3 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				'case 4 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				'case 5 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjPDCMEXE = Nothing
	set mobjPDCMGET = Nothing
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
		'.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'DateClean
		'.txtDEMANDDAY.value = gNowDate
		'.txtPRINTDAY.value  = gNowDate
		'.sprSht.MaxRows = 0	
		'.sprSht1.MaxRows = 0
		
		'.txtDEMANDDAY.readOnly = "FALSE"
		'.txtDEMANDDAY.className = "INPUT"
		'.imgCalDemandday.disabled = FALSE
	
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
	
		if .txtJOBNO.value = "" Then
			gErrorMsgBox "조회된 제작관리번호가 없습니다.","저장안내!"
			Exit Sub
			Else
			strCODE = .txtJOBNO.value 
		End If
		
  		'데이터 Validation
		if DataValidation =false then exit sub
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ|VATCODE|INCOMCODE")
		strMasterData = gXMLGetBindingData (xmlBind)
		if  not IsArray(vntData) then 
			If gXMLIsDataChanged (xmlBind) Then 'XML 데이터 중 변경된것이 있다면
			Else
				gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				exit sub
			End If
		End If
		'처리 업무객체 호출
		
		
			intRtn = mobjPDCMEXE.ProcessRtn(gstrConfigXml,strMasterData,vntData,strCODE)
				
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가" & intRtn & " 건 저장" & mePROC_DONE,"저장안내" 
			SelectRtn
  		end if
 	end with
End Sub
'****************************************************************************************
' 데이터 처리 - 헤더 없는 경우 진행비 우선투입 하기 위해 헤더만 우선시 저당하도록 설정
'****************************************************************************************
Sub ProcessRtn_SUB ()
   Dim intRtn
  	dim vntData
	Dim strMasterData
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
		strCODE = .txtJOBNO.value 
	
  		'데이터 Validation
		if DataValidation =false then exit sub
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ|VATCODE|INCOMCODE")
		strMasterData = gXMLGetBindingData (xmlBind)
		
		intRtn = mobjPDCMEXE.ProcessRtn(gstrConfigXml,strMasterData,vntData,strCODE)	
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			SelectRtn
  		end if
 	end with
End Sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
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
			if mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) <> "" AND mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "코드선택" Then 
				gErrorMsgBox intCnt & " 번째 행의 세무코드 를 확인하십시오","입력오류"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	with frmThis
	strCODE = .txtJOBNO1.value 
		if strCODE = "" Or Len(strCODE) <> 7 Then
			gErrorMsgBox "제작번호를확인하십시오.","조회안내!"
			Exit Sub
		End if
	
	IF not SelectRtn_Head (strCODE) Then Exit Sub

	'쉬트 조회
	CALL SelectRtn_Detail (strCODE)
	
	txtSUSUAMT_onblur
	txtCOMMITION_onblur
	txtDEMANDAMT_onblur
	txtPAYMENT_onblur
	txtINCOM_onblur
	txtNONCOMMITION_onblur
	txtACCAMT_onblur
	txtESTAMT_onblur
	End with
	
End Sub

Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	Dim strCODENAME
	SelectRtn_Head = false
	strCODENAME = frmThis.txtJOBNAME1.value 
	'on error resume next
	
	'초기화
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMEXE.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE,strCODENAME)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 JOBNO 에 대하여 확정견적서가 " & meNO_DATA, ""
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
	Dim intCnt2
	'on error resume next
	'초기화
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMEXE.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt2 = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt2) <> "" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt2,-1,-1,true
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,intCnt2,11,19,true
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt2,17,17,true
					End If
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt2) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",intCnt2,"코드선택"
					sprSht_Change 18,intCnt2
					End If
					
				Next
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME)
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************


'****************************************************************************************
' 전체 삭제와 각 쉬트별 삭제
'****************************************************************************************
'자료삭제
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2
	dim strYEARMON
	Dim strSEQ
	Dim strJOBNO
	Dim strSORTSEQ
	Dim lngCnt
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
		intRtn2 = 0
		lngCnt = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)) <> "" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",vntData(i)) <> "" Then
					gErrorMsgbox "정산일 이 있는 매입확정 분은 삭제될수 없습니다.","삭제안내!"
					Exit Sub
				End If
				strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",vntData(i))
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)))
				strSORTSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",vntData(i)))
				intRtn2 = mobjPDCMEXE.DeleteRtn(gstrConfigXml,strJOBNO, strSEQ,strSORTSEQ)
			End IF
			
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			End IF
		next
		If lngCnt <> 0 Then
		gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
		End If
		
   		If intRtn2 = 0 Then
   		Else
			Payment_changevalue
			DelProc
		End If
		mobjSCGLSpr.DeselectBlock .sprSht
		
	End with
	err.clear
End Sub

Sub DelProc
Dim intHDR
Dim strMasterData
Dim strPREESTNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		intHDR = mobjPDCMEXE.ProcessRtn_DELHDR(gstrConfigXml,strMasterData)
				if not gDoErrorRtn ("ProcessRtn_DELHDR") then
					SelectRtn
				End If
	End with
End Sub
'여기부터 사용
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
		vntInParams = array(trim(.txtJOBNO1.value), trim(.txtJOBNAME1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO1.value = vntRet(0,0) and .txtJOBNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtJOBNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtJOBNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
			SelectRtn
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtJOBNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO1.value),trim(.txtJOBNAME1.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO1.value = trim(vntData(0,0))
					.txtJOBNAME1.value = trim(vntData(1,0))
					SelectRtn
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
Sub DeleteRtn_ALL
	Dim intRtn
	dIM strJOBNO
	Dim vntInParams
	Dim vntRet
	Dim vntData
	Dim intCnt
	Dim intCntV
	with frmThis
	


	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMEXE.SelectRtn_ACCEXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		If mlngRowCnt = 0 Then
			gErrorMsgBox "삭제 할 데이터가 없습니다.","전체삭제안내!"
			Exit Sub	
		End If
	End If

	if .txtENDDAY.value <> "" Then
		gErrorMsgBox "완료된 정산건은 삭제될수 없습니다.","전체삭제안내!"
		Exit Sub
	End if
	intCntV = 0
	For intCnt =1 To .sprSht.MaxRows
		intCntV = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJDAY",frmThis.sprSht.ActiveRow)
		If intCntV <> "" Then
			gErrorMsgBox "정산일이 존재하는 건은 전체삭제 될수 없습니다.","전체삭제안내!"
			Exit Sub
		End If
	Next
	
	
	intRtn = gYesNoMsgbox("자료를 전체 삭제하시겠습니까?" & vbcrlf & "전체자료가 삭제됩니다.","자료삭제 확인")
	IF intRtn <> vbYes then exit Sub
	
	strJOBNO = .txtJOBNO.value 
	intRtn = mobjPDCMEXE.DeleteRtn_ALL(gstrConfigXml,strJOBNO)
	if not gDoErrorRtn ("DeleteRtn_ALL") then
					gOkMsgbox "삭제가 되었습니다.","삭제안내"
					SelectRtn
				End If
	End with 
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0" >
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD >
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;정산 관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!---->
										<TABLE id="tblButton1" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="50" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!---->
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
									<TD class="SEARCHLABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME1, txtJOBNO1)"
										width="93">JOB명</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtJOBNAME1" title="제작의뢰명 조회조건" style="WIDTH: 266px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="38" name="txtJOBNAME1"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" alt="제작의뢰번호를 조회합니다" src="../../../images/imgPopup.gIF" width="23"
											align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO1" title="제작의뢰번호 조회조건" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="7" align="left" size="3" name="txtJOBNO1"> <INPUT dataFld="JOBNO" id="txtJOBNO" dataSrc="#xmlBind" type="hidden" name="txtJOBNO"><INPUT dataFld="JOBNOINS" id="txtJOBNOINS" dataSrc="#xmlBind" type="hidden" name="txtJOBNOINS"><INPUT dataFld="PREESTNO" id="txtPREESTNO" dataSrc="#xmlBind" type="hidden" name="txtPREESTNO"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 95px">프로젝트명</TD>
									<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="PROJECTNM" class="NOINPUT_L" id="txtPROJECTNM" title="프로젝트명" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPROJECTNM"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">JOB명</TD>
									<TD class="SEARCHDATA" style="WIDTH: 148px"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="제작건명" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">매체부문</TD>
									<TD class="SEARCHDATA" style="WIDTH: 142px"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="매체부문" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBGUBN"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 103px">매체분류</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="매체부문" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 95px">광고주</TD>
									<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="광고주" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">사업부</TD>
									<TD class="SEARCHDATA" style="WIDTH: 148px"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="사업부" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTSUBNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">브랜드</TD>
									<TD class="SEARCHDATA" style="WIDTH: 142px"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUBSEQNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 103px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtENDDAY, '')">결산일</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="ENDDAY" class="NOINPUT" id="txtENDDAY" title="완료일" style="WIDTH: 152px; HEIGHT: 22px"
											accessKey="date" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtENDDAY"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;외주 정산</td>
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
							<TABLE id="tblBody" style="WIDTH: 100%"  cellSpacing="0" cellPadding="0" border="0" align="left">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="left">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="left" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">Noncommition</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="수수료미지불금액"
														style="WIDTH: 152px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtNONCOMMITION"></TD>
												<TD class="LABEL" style="WIDTH: 106px">Commition</TD>
												<TD class="DATA" style="WIDTH: 148px"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="수수료지불금액" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCOMMITION"></TD>
												<TD class="LABEL" style="WIDTH: 106px">수수료율</TD>
												<TD class="DATA" style="WIDTH: 142px"><INPUT dataFld="SUSURATE" class="NOINPUTB_R" id="txtSUSURATE" title="수수료율" style="WIDTH: 128px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="16" name="txtSUSURATE">&nbsp;(%)</TD>
												<TD class="LABEL" style="WIDTH: 103px">
													수수료</TD>
												<TD class="DATA"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="수수료합계금액" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUSUAMT"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">
													청구금액</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="DEMANDAMT" class="NOINPUTB_R" id="txtDEMANDAMT" title="청구금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDAMT"></TD>
												<TD class="LABEL" style="WIDTH: 106px">외주비</TD>
												<TD class="DATA" style="WIDTH: 148px"><INPUT dataFld="PAYMENT" class="NOINPUTB_R" id="txtPAYMENT" title="외주비 합계" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPAYMENT"></TD>
												<TD class="LABEL" style="WIDTH: 106px">내수율</TD>
												<TD class="DATA" style="WIDTH: 142px"><INPUT dataFld="RATE" class="NOINPUTB_R" id="txtRATE" title="내수율" style="WIDTH: 128px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="16" name="txtRATE">&nbsp;(%)</TD>
												<TD class="LABEL" style="WIDTH: 103px">
													내수액</TD>
												<TD class="DATA"><INPUT dataFld="INCOM" class="NOINPUTB_R" id="txtINCOM" title="내수액" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtINCOM"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">
													견적금액</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="ESTAMT" class="NOINPUTB_R" id="txtESTAMT" title="견적금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtESTAMT"></TD>
												<TD class="LABEL">
													진행비</TD>
												<TD class="DATA"><INPUT dataFld="ACCAMT" class="NOINPUTB_R" id="txtACCAMT" title="비용 합계" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtACCAMT"></TD>
												<TD></TD>
												<TD></TD>
												<TD vAlign="bottom" align="right" colSpan="2"><IMG id="ImgAccInput" onmouseover="JavaScript:this.src='../../../images/ImgAccInputOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAccInput.gIF'" height="20" alt="진행비투입" src="../../../images/ImgAccInput.gIF"
														align="absMiddle" border="0" name="ImgAccInput">&nbsp;<IMG id="imgAddRow" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" height="20" alt="한 행 추가" src="../../../images/imgRowAdd.gIF"
														align="absMiddle" border="0" name="imgAddRow">&nbsp;<IMG id="imgDelRow" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="한 행 삭제" src="../../../images/imgRowDel.gIF"
														align="absMiddle" border="0" name="imgDelRow"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								</TABLE>
								<!--Input End-->
								
						</TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="11536">
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
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
