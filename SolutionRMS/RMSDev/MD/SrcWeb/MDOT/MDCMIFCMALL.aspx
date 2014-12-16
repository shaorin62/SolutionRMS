<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMIFCMALL.aspx.vb" Inherits="MD.MDCMIFCMALL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>IFC MALL 청약관리</title>
		<meta name="vs_snapToGrid" content="False">
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'시스템구분 : 위수탁거래명세서 등록 화면(MDCMPRINTTRANS1.aspx)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMIFCMALL.aspx
'기      능 : IFC MALL 청약  기초 데이터 생성 화면
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/28 By Kim Tae Yub
'			 2) 
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script id="clientEventHandlersVBS" language="vbscript">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDOTIFCMALL, mobjMDCOGET
Dim mstrCheck, mstrCheck1

CONST meTAB = 9
mstrCheck=True
mstrCheck1=True

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

'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

sub ImgAddRow_OUT_onclick ()
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "상단의 계약 정보가 없으면 추가할 수 없습니다.","저장안내"
			Exit Sub
		End If
	
		IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
			gErrorMsgBox "해당 데이터는 거래명세서가 진행된 데이터 입니다. 추가 하실 수 없습니다.","저장안내"
			Exit Sub
		end if
		
		call sprSht_OUT_Keydown(meINS_ROW, 0)
		.txtCLIENTCODE1.focus
		.sprSht_OUT.focus
	End With 
End sub

sub ImgAddRow_SUSU_onclick ()
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "상단의 계약 정보가 없으면 추가할 수 없습니다.","저장안내"
			Exit Sub
		End If
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
			gErrorMsgBox "해당 데이터는 거래명세서가 진행된 데이터 입니다. 추가 하실 수 없습니다.","저장안내"
			Exit Sub
		end if
		
		call sprSht_SUSU_Keydown(meINS_ROW, 0)
		.txtCLIENTCODE1.focus
		.sprSht_SUSU.focus
	End With 
End sub

Sub ImgSave_AMT_onclick
	If frmThis.sprSht_HDR.MaxRows = 0 Then
   		gErrorMsgBox "상세항목 이 없습니다.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_AMT
	gFlowWait meWAIT_OFF
End Sub

'-----------인쇄-----------
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i
	Dim chkcnt
	Dim strCONT_CODE
	
	Dim strYEARMON, strCLIENTCODE, strTITLE
	Dim Con1, Con2, Con3, Con4
	
	with frmThis
		Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" 
		
		if frmThis.sprSht_HDR.MaxRows = 0 then
			gErrorMsgBox "인쇄할 데이터가 없습니다.",""
			Exit Sub
		end if
		
		ModuleDir = "MD"

		ReportName = "MDCMIFCMALL.rpt"
		
		strYEARMON		 = .txtYEARMON.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strTITLE		 = .txtTITLE1.value
		
		If strYEARMON		<> ""	Then Con1  = " AND (YEARMON = '" & strYEARMON & "') "
		If strCLIENTCODE	<> ""	Then Con2  = " AND (CLIENTCODE = '" & strCLIENTCODE & "') "
		If strTITLE			<> ""	Then Con3  = " AND (CONT_NAME = '" & strTITLE & "') " 
		
		chkcnt=0
		For i=1 To .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = "1" Then
				if chkcnt = 0 then
					strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)
				else
					strCONT_CODE = strCONT_CODE & "," & mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)  
				end if 
				chkcnt = chkcnt +1
			End If
		Next
		
		if chkcnt <> 0 then
			Con4 = " AND ( CONT_CODE IN (" & strCONT_CODE &"))"
		end if 
	
		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & Con4 
		
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_OUT_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_OUT
	gFlowWait meWAIT_OFF
End Sub
	
Sub imgDelete_SUSU_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_SUSU
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		'mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
'조회용 이벤트
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
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
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "D")

			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	with frmThis
		if  Row > 0 AND Col > 1 then
			SelectRtn_OUT Col, Row
			SelectRtn_SUSU Col, Row
		end if
	end with
End Sub

Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck1
			if mstrCheck1 = True then 
				mstrCheck1 = False
			elseif mstrCheck1 = False then 
				mstrCheck1 = True
			end if
			for intcnt = 1 to .sprSht_OUT.MaxRows
				sprSht_OUT_Change 1, intcnt
			next
		end if
	end with
End Sub  

Sub sprSht_SUSU_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1, 1, , , "", , , , , mstrCheck1
			if mstrCheck1 = True then 
				mstrCheck1 = False
			elseif mstrCheck1 = False then 
				mstrCheck1 = True
			end if
			for intcnt = 1 to .sprSht_SUSU.MaxRows
				sprSht_SUSU_Change 1, intcnt
			next
		end if
	end with
End Sub  

'상단 시트 더블클릭
sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

'----------------------------------------------------------
' [시트 키업]
'----------------------------------------------------------
Sub sprSht_HDR_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_OUT frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		SelectRtn_SUSU frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		KeyUp_SumAmt .sprSht_HDR
	End With
End Sub

Sub sprSht_OUT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt_OUT .sprSht_OUT
	End With
End Sub

Sub sprSht_SUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt_SUSU .sprSht_SUSU
	End With
End Sub

'---------------------------------------------
'시트 마우스 업
'---------------------------------------------
'청약상단 
Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_HDR
	end with
End Sub

'대행사 시트
Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt_OUT .sprSht_OUT
	end with
End Sub

'SUSU 시트 
Sub sprSht_SUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt_SUSU .sprSht_SUSU
	end with
End Sub

'-----------------------------------------------
'스프레드시트 체인지 이벤트
'-----------------------------------------------
Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	With frmThis
		
		if Row = 1 then
			mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow 
		else
			mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow -1
		end if

		mobjSCGLSpr.CellChanged .sprSht_HDR, Col, Row

		'SelectRtn_OUT .sprSht_HDR.ActiveCol,.sprSht_HDR.ActiveRow
   		'SelectRtn_SUSU .sprSht_HDR.ActiveCol,.sprSht_HDR.ActiveRow
	End With
End Sub

Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
   	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   
	With frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"EXCLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"EXCLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"EXCLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC_CGV(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName, strCodeName)

				If not gDoErrorRtn ("GetCC_CGV") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTCODE",frmThis.sprSht_OUT.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTNAME",frmThis.sprSht_OUT.ActiveRow, trim(vntData(1,1))
						
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"EXCLIENTNAME"), Row
						.sprSht_OUT.focus 
						mobjSCGLSpr.ActiveCell .sprSht_OUT, Col+1, Row
					End If
   				End If
   			End If
		End If

		mobjSCGLSpr.CellChanged .sprSht_OUT, Col, Row
	End With
End Sub

Sub sprSht_SUSU_Change(ByVal Col, ByVal Row)
   	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
		
	With frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"EXCLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_SUSU,"EXCLIENTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"")
		
				If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTCODE",frmThis.sprSht_SUSU.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTNAME",frmThis.sprSht_SUSU.ActiveRow, trim(vntData(2,1))
						
						.txtEXCLIENTNAME.focus
						.sprSht_SUSU.focus
					Else
						mobjSCGLSpr_ClickProc2 mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"EXCLIENTNAME"), Row
						.txtEXCLIENTNAME.focus
						.sprSht_SUSU.focus 
						mobjSCGLSpr.ActiveCell .sprSht_SUSU, Col+1, Row
					End If
   				End If
   			End If
		End If
	
		mobjSCGLSpr.CellChanged .sprSht_SUSU, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"EXCLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"EXCLIENTNAME",Row)))

			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTCODE",frmThis.sprSht_OUT.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTNAME",frmThis.sprSht_OUT.ActiveRow, trim(vntRet(2,0))
				
				mobjSCGLSpr.CellChanged .sprSht_OUT, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_OUT, Col+2,Row
			End If
		End If
		
		.txtCLIENTNAME1.focus
		.sprSht_OUT.Focus
	End With
End Sub

Sub mobjSCGLSpr_ClickProc2(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"EXCLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_SUSU,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTCODE",frmThis.sprSht_SUSU.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTNAME",frmThis.sprSht_SUSU.ActiveRow, trim(vntRet(2,0))
				
				mobjSCGLSpr.CellChanged .sprSht_SUSU, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_SUSU, Col+2,Row
			End If
		End If

		.txtCLIENTNAME1.focus
		.sprSht_SUSU.Focus
	End With
End Sub

'-------------------------------------------
'스프레드 시트 버튼 클릭
'-------------------------------------------
Sub sprSht_OUT_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
	
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BTN_EX") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTCGVPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTCODE",frmThis.sprSht_OUT.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_OUT,"EXCLIENTNAME",frmThis.sprSht_OUT.ActiveRow, trim(vntRet(1,0))
				
				mobjSCGLSpr.CellChanged .sprSht_OUT, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_OUT, Col+2,Row
			End If
		End If
	End with
End Sub

Sub sprSht_SUSU_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"BTN_EX2") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_SUSU,"EXCLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTCODE",frmThis.sprSht_SUSU.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht_SUSU,"EXCLIENTNAME",frmThis.sprSht_SUSU.ActiveRow, trim(vntRet(2,0))
				
				mobjSCGLSpr.CellChanged .sprSht_SUSU, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_SUSU, Col+2,Row
			End If
		end if
	End with
End Sub

'------------------------------------
'시트 키다운 이벤트
'------------------------------------
Sub sprSht_HDR_Keydown(KeyCode, Shift)

	with frmThis
		if keycode = meENTER then
			mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow-1
		end if
	end with
end sub

Sub sprSht_OUT_Keydown(KeyCode, Shift)
	Dim intRtn
	With frmThis
		
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		if KeyCode = meCR  Or KeyCode = meTab Then
			
		ElseIf KeyCode = meINS_ROW  Then
		
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_OUT, cint(KeyCode), cint(Shift), -1, 1)
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"CONT_CODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_CD",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_NAME",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_NAME",.sprSht_HDR.ActiveRow)
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"EXCLIENTCODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTCODE",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"EXCLIENTNAME",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTNAME",.sprSht_HDR.ActiveRow)
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow, "30"
			
			OUTAMT_CALCUL .sprSht_HDR.ActiveCol , .sprSht_HDR.ActiveRow
			
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, False, " YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX | EXCLIENTNAME | RATE | AMT | MEMO "
			
		End if
	End With
End Sub

Sub sprSht_SUSU_Keydown(KeyCode, Shift)
	Dim intRtn
	With frmThis
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		if KeyCode = meCR  Or KeyCode = meTab Then
		
		ElseIf KeyCode = meINS_ROW  Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_SUSU, meINS_ROW, 0, -1, 1)
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"YEARMON",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"CONT_CODE",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"DEPT_CD",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_CD",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"DEPT_NAME",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"DEPT_NAME",.sprSht_HDR.ActiveRow)
			
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"EXCLIENTCODE",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTCODE",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"EXCLIENTNAME",.sprSht_SUSU.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTNAME",.sprSht_HDR.ActiveRow)
			
			mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"RATE",.sprSht_SUSU.ActiveRow, "20"
			
			SUSUAMT_CALCUL .sprSht_HDR.ActiveCol , .sprSht_HDR.ActiveRow
			
		End if
	End With
End Sub


'-----------------------------------
'시트에서 키업 금액 합산 이벤트
'-----------------------------------
SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT") or _
							  mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT") Then
		
			strSUM = 0
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT"))  Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

SUB KeyUp_SumAmt_OUT (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")  Then
		
			strSUM = 0
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and  (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

SUB KeyUp_SumAmt_SUSU (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")  Then
		
			strSUM = 0
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and  (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

'-----------------------------------
'시트에서 마우스를 금액합산 이벤트
'-----------------------------------
sub MouseUp_SumAmt(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or _
								  mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT ") or mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT") Then

				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

sub MouseUp_SumAmt_OUT(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") Then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)

				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

sub MouseUp_SumAmt_SUSU(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") Then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))

							End If
						Next
					end if
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

'광고대행수수료 분배 로직
SUB OUTAMT_CALCUL (ByVal Col, ByVal Row)
	Dim strGBN
	Dim intMC_AMT
	Dim intOUT_AMT
	Dim intOVERAMT
	Dim intAMT
	
	with frmThis
		strGBN = "" : intMC_AMT = 0 : intOUT_AMT = 0 : intAMT = 0
		
		strGBN = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"ENDFLAG",Row)
		
		IF strGBN = "Y" THEN
			'총매출 수익이 181500000 가 초과할경우 MC 수익금의 30%에 초과분에 대한 40% 더 지급
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"AMT",Row)
			intMC_AMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MC_AMT",Row)
			intOUT_AMT  = 108900000 * 0.3
			
			intOVERAMT = Clng(intAMT) - 181500000
			intOVERAMT = (intOVERAMT * 0.4) * 0.6
			
		
		
			intOUT_AMT = intOUT_AMT + intOVERAMT
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow , intOUT_AMT
			
		ELSE
			
			'총매출 수익이 181500000 일경우이거나 이하인경우 MC 수익금의 30%  를 광고 대행 수수료 분배
			intMC_AMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MC_AMT",Row)
			intOUT_AMT  = intMC_AMT * 0.3
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow , intOUT_AMT
			
		END IF 
	end with
end sub


'커미션 정산 관련 로직
SUB SUSUAMT_CALCUL (ByVal Col, ByVal Row)
	Dim intMC_AMT
	Dim intSUSU_AMT
	
	with frmThis
		intMC_AMT = 0 : intSUSU_AMT = 0
		
		intMC_AMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MC_AMT",Row)
		intSUSU_AMT  = intMC_AMT * 0.2
		mobjSCGLSpr.SetTextBinding .sprSht_SUSU,"AMT",.sprSht_SUSU.ActiveRow , intSUSU_AMT
			
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
	'서버업무객체 생성	
	set mobjMDOTIFCMALL = gCreateRemoteObject("cMDOT.ccMDOTIFCMALL")
	set mobjMDCOGET	  = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
		'-------------------------------------
		'계약 그리드
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 21, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | YEARMON | SEQ | DEMANDDAY | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | AMT | REAL_AMT | MC_RATE | MC_AMT | EX_RATE | EX_AMT | COMMI_TRANS_NO | MEMO | ENDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "선택|년월|순번|청구일자|광고주코드|광고주명|계약코드|계약명|대행사코드|대행사명|담당부서코드|담당부서명|금액(MG)|당월청구금액|MC수수료율|MC수수료|대행매입율|대행매입금|거래명세서번호|비고|ENDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1","  4|   6|   0|       8|         0|      15|       8|    12|         0|      15|           0|        12|      15|          10|         8|      10|         8|        10|             0|  10|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "SEQ | AMT | REAL_AMT | MC_AMT | EX_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "MC_RATE | EX_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "YEARMON | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | COMMI_TRANS_NO | MEMO | ENDFLAG", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, " YEARMON | SEQ | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | AMT | MC_RATE | MC_AMT | EX_RATE | EX_AMT | COMMI_TRANS_NO | ENDFLAG"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "CLIENTCODE", True
		
		'-------------------------------------
		'대행수수료 분배 그리드
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 13, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_OUT, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_OUT, "CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX | EXCLIENTNAME | RATE | AMT | MEMO | GUBUN"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		 "선택|년월|순번|계약코드|부서코드|부서명|대행사코드|대행사명|비율|대행수수료분배금액|비고|매출구분"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1", " 4|   6|   4|       8|       0|     0|         8|2|    18|   8|                15|   8|       4"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK | GUBUN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_OUT,"..", "BTN_EX"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, " YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, " SEQ | AMT ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, True, "SEQ | CONT_CODE "
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CHK | YEARMON",-1,-1,2,2,False
		mobjSCGLSpr.ColHidden .sprSht_OUT, "CONT_CODE", True
		
		'-------------------------------------
		'IFC 수수료 정산 관련 
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 12, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_SUSU, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU, "CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX2 | EXCLIENTNAME | RATE | AMT | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		 "선택|년월|순번|계약코드|담당부서코드|담당부서명|제작사코드|제작사명|비율|IFC 커미션|비고"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1", "4|   8|   4|       0|           0|         0|         8|2|    18|   8|		  15|  10"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_SUSU,"..", "BTN_EX2"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, " YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "SEQ | AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU, True, "SEQ | CONT_CODE"
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CHK | YEARMON",-1,-1,2,2,False
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "CONT_CODE ", True
		
		.sprSht_HDR.style.visibility = "visible"
		.sprSht_OUT.style.visibility = "visible"
		.sprSht_SUSU.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDOTIFCMALL = Nothing
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
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)

		.sprSht_HDR.MaxRows = 0	
		.sprSht_OUT.MaxRows = 0
		.sprSht_SUSU.MaxRows = 0
	End with
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON
	Dim strCLIENTCODE, strTITLE
   	Dim i, strCols
   	Dim intCnt , strRows
    
	with frmThis
		intCnt = 1
		
		If .txtYEARMON.value = "" Then
			gErrorMsgBox "조회시 년월은 반드시 넣어야 합니다.",""
			Exit SUb
		End If 
		
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.sprSht_SUSU.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strTITLE		= .txtTITLE1.value
		
		vntData = mobjMDOTIFCMALL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											  strYEARMON, strCLIENTCODE, strTITLE)
													
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				
   				for i =1 to .sprSht_HDR.maxRows
   					
   					If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",i) <> "" THEN  '거래명세서가 끊어진 경우 다잠금
   						If intCnt = 1 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
						intCnt = intCnt + 1
   					END IF 
   				next

   				mobjSCGLSpr.SetCellsLock2 .sprSht_HDR,True,strRows,2,20,True
   				
   				AMT_SUM (.sprSht_HDR)
   				SelectRtn_OUT 1,1
   				SelectRtn_SUSU 1,1
   			else
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   	end with
End Sub

'----------사내수익---------
Sub SelectRtn_OUT (Col, Row)
	Dim vntData
	Dim strCONT_CODE, strYEARMON
   	Dim i, strCols
    Dim intCnt, strRows
    
	with frmThis
		'Sheet초기화
		.sprSht_OUT.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",Row)

		vntData = mobjMDOTIFCMALL.SelectRtn_OUT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
																							
		If not gDoErrorRtn ("SelectRtn_OUT") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_OUT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus_OUT, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				
   				'상단의 청약 내역이 거래명세서가 끊어진 경우 그리드를 잠그고 아닌경우 푼다
   				IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
   	
   					mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, True, "YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX | EXCLIENTNAME | RATE | AMT | MEMO | GUBUN"	
   				else
   					mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, False, " YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX | EXCLIENTNAME | RATE | AMT | MEMO | GUBUN"	
   				END IF 
   			End If
   		End If
   	end with
End Sub

'----------수수료수익---------
Sub SelectRtn_SUSU (Col, Row)
	Dim vntData
	Dim strCONT_CODE, strYEARMON
    Dim intCnt
    
	with frmThis
		'Sheet초기화
		.sprSht_SUSU.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",Row)
		
		'저장된 값이 있다면 저장된 값을 뿌린다.
		vntData = mobjMDOTIFCMALL.SelectRtn_SUSU(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
		
		If not gDoErrorRtn ("SelectRtn_SUSU") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_SUSU,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				gWriteText lblStatus_SUSU, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			End If
   		End If
			
		'상단의 청약 내역이 거래명세서가 끊어진 경우 그리드를 잠그고 아닌경우 푼다
   		IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
   			mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU, True, " YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX2 | EXCLIENTNAME | RATE | AMT | MEMO"	
   		else 
   			mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU, False, "YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX2 | EXCLIENTNAME | RATE | AMT | MEMO"	
   		END IF 
   	end with
End Sub

'-----------------------------------
'조회시 금액 합산과 선택 합산
'-----------------------------------
Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'**********************************************
'----------------------저장--------------------
'**********************************************
'-------------상단 청약 데이터 저장-----------
Sub ProcessRtn_AMT ()
	Dim intRtn
   	Dim vntData
   	Dim intCnt
	Dim strDataCHK
	Dim lngCol, lngRow
	Dim strCONT_CODE

	With frmThis
		'---------------------DATAVALIDATION------------------------
		if .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "저장하실 데이터가 없습니다.","저장안내!"
			Exit Sub
		end if

		mobjSCGLSpr.SetFlag  .sprSht_HDR,meINS_TRANS

		'상단 필수입력사항
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_HDR, "YEARMON ",lngCol, lngRow, False) 
		If strDataCHK = False Then
			gErrorMsgBox "상단 청약데이터의" & lngRow & " 줄의 년월 는 필수 입력사항입니다.","저장안내"
			Exit Sub
		End If

		' 광고 대행 수수료 분배  필수 입력사항
		if .sprSht_OUT.MaxRows <> 0 then
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_OUT, "YEARMON | EXCLIENTCODE",lngCol, lngRow, False) 

			If strDataCHK = False Then
				gErrorMsgBox  "광고 대행수수료 분배 그리드" & lngRow & " 줄의 년월/대행사는 필수 입력사항입니다.","저장안내"
				Exit Sub
			End If
			if DataValidation = false then exit sub 	
		end if 
		
		'SUSU 그리드 데이터가 있는경우 필수 입력사항
		if .sprSht_SUSU.MaxRows <> 0 then
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_SUSU, "YEARMON | EXCLIENTCODE ",lngCol, lngRow, False) 
		
			If strDataCHK = False Then
				gErrorMsgBox "SUSU 그리드" & lngRow & " 줄의 년월/ 제작사명는 필수 입력사항입니다.","저장안내"
				Exit Sub
			end if

			if DataValidation_SUSU = false then exit sub 	
		end if
		
		'---------------------------저장---------------------------				
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
	
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | YEARMON | SEQ | DEMANDDAY | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | EXCLIENTCODE | EXCLIENTNAME | DEPT_CD | DEPT_NAME | AMT | REAL_AMT | MC_RATE | MC_AMT | EX_RATE | EX_AMT | COMMI_TRANS_NO | MEMO | ENDFLAG")
		'처리 업무객체 호출
		
		If isArray(vntData) Then
			intRtn = mobjMDOTIFCMALL.ProcessRtn_AMT(gstrConfigXml, vntData)
		end if
		
		'---------------------------
		'좌측 하단의 팀 대행사 저장
		'---------------------------
		For intCnt = 1 to .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"YEARMON",intCnt) <> "" Then
				mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
			End If
		Next
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX | EXCLIENTNAME | RATE | AMT | MEMO | GUBUN")

		'처리 업무객체 호출
		If isArray(vntData) Then
			intRtn = mobjMDOTIFCMALL.ProcessRtn_OUT(gstrConfigXml, vntData)
		end if 
		
		'---------------------------
		'우측 하단의 수수료 지점 저장
		'---------------------------
		For intCnt = 1 to .sprSht_SUSU.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"YEARMON",intCnt) <> "" Then
				mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, 1, intCnt
			End If
		Next
		
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_SUSU,"CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | DEPT_NAME | EXCLIENTCODE | BTN_EX2 | EXCLIENTNAME | RATE | AMT | MEMO")
				
		'처리 업무객체 호출
		If isArray(vntData) Then
			intRtn = mobjMDOTIFCMALL.ProcessRtn_SUSU(gstrConfigXml, vntData)
		end if 
		'---------------------------------------------------------------------
		if not gDoErrorRtn ("ProcessRtn_AMT") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
			gErrorMsgBox "자료가 저장" & mePROC_DONE,"저장안내" 
			SelectRtn
   		end if
   	end with
end sub

'****************************************************************************************
' 데이터 처리를 위한 데이타 검증
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
	'On error resume Next
	With frmThis
   		
   	End With
	DataValidation = True
End Function

Function DataValidation_SUSU ()
	DataValidation_SUSU = False
	'On error resume Next
	With frmThis
   	End With
	DataValidation_SUSU = True
End Function

'--------------청약내용삭제 -------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON,strCONT_CODE
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	With frmThis
		
		for i = 1 to .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
				
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",i) <> "" then
					gErrorMsgBox i & "번째 로우의 데이터는 거래명세서가 진행된 내역 입니다. 삭제 하실 수 없습니다.","삭제안내!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If

		intRtn = gYesNoMsgbox("청약 내역을 삭제하시면 하단의 대행수수료분배 나 수수료 금액이 모두 삭제 됩니다." & vbcrlf & " 자료를 삭제하시겠습니까? ","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_HDR.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",i)
				strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
				else
					intRtn = mobjMDOTIFCMALL.DeleteRtn_AMT(gstrConfigXml, strYEARMON, strSEQ, strCONT_CODE)
					
					IF not gDoErrorRtn ("DeleteRtn_AMT") then
						mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   					End IF
				end if				
   				intCnt = intCnt + 1
   			END IF
		next

		If not gDoErrorRtn ("DeleteRtn_AMT") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_OUT
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		SelectRtn
	End With
	err.clear
End Sub

'--------------대행수수료 분배수익 삭제-------------
Sub DeleteRtn_OUT ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON
	Dim strSEQFLAG '실제데이터여부 플레
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	strSEQFLAG = False
	With frmThis
	
		for i = 1 to .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" then
					gErrorMsgBox "해당 데이터는 거래명세서가 진행된 내역 입니다. 삭제 하실 수 없습니다.","삭제안내!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If

		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0

		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_OUT.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"YEARMON",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_OUT,i
				else
					intRtn = mobjMDOTIFCMALL.DeleteRtn_OUT(gstrConfigXml, strYEARMON, strSEQ)
					
					IF not gDoErrorRtn ("DeleteRtn_OUT") then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End IF
   					
   					strSEQFLAG = TRUE
				end if				
   				intCnt = intCnt + 1
   			END IF
		next
		
		If not gDoErrorRtn ("DeleteRtn_OUT") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If

		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_OUT
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn_OUT frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		End If
	End With
	err.clear	
End Sub

'--------------수수료수익 삭제-------------
Sub DeleteRtn_SUSU ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON
	Dim strSEQFLAG '실제데이터여부 플레
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	strSEQFLAG = False
	With frmThis
	
		for i = 1 to .sprSht_SUSU.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
	
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" then
					gErrorMsgBox "해당 데이터는 거래명세서가 진행된 내역 입니다. 삭제 하실 수 없습니다.","삭제안내!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_SUSU.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"YEARMON",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
				else
					intRtn = mobjMDOTIFCMALL.DeleteRtn_SUSU(gstrConfigXml, strYEARMON, strSEQ)

					IF not gDoErrorRtn ("DeleteRtn_SUSU") then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
   					End IF

   					strSEQFLAG = TRUE
				end if				
   				intCnt = intCnt + 1
   			END IF
		next

		If not gDoErrorRtn ("DeleteRtn_SUSU") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If

		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_SUSU
		'내역복사 된 데이터삭제시 조회를 안태우고, 실 데이터 삭제시 재조회
		If strSEQFLAG Then
			SelectRtn_SUSU frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		End If
	End With
	err.clear	
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28">
							<TR>
								<TD height="28" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="121" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td class="TITLE">IFC MALL 청약관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="처리중입니다."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="93%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="TOPSPLIT"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" class="KEYFRAME" vAlign="top" align="left">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
										align="left">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="50">년월</TD>
											<TD style="WIDTH: 85px" class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 80px; HEIGHT: 22px" id="txtYEARMON" class="INPUT"
													title="년월" maxLength="10" size="6" name="txtYEARMON"></TD>
											<TD style="WIDTH: 47px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="47">광고주</TD>
											<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 174px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="광고주명"
													maxLength="100" size="23" name="txtCLIENTNAME1"><IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF"><INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT" title="코드입력"
													maxLength="6" size="4" name="txtCLIENTCODE1"></TD>
											<TD style="WIDTH: 42px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTITLE1, '')"
												width="42">계약명</TD>
											<TD class="SEARCHDATA" width="220"><INPUT style="WIDTH: 216px; HEIGHT: 22px" id="txtTITLE1" class="INPUT_L" title="계약명" maxLength="100"
													size="30" name="txtTITLE1"></TD>
											<TD class="SEARCHDATA">
												<TABLE border="0" cellSpacing="0" cellPadding="2" align="right">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery"
																alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							<tr>
								<td>
									<table class="DATA" cellSpacing="0" cellPadding="0" width="100%" height="10">
										<TR>
										</TR>
									</table>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<td style="WIDTH: 1000px" class="DATA">총금액합계:<INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 20px" id="txtSUMAMT" class="NOINPUTB_R"
													title="합계금액" readOnly maxLength="100" size="13" name="txtSUMAMT">합계<INPUT style="WIDTH: 120px; HEIGHT: 20px" id="txtSELECTAMT" class="NOINPUTB_R" title="선택금액"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
											<TD height="20" width="400" align="left"></TD>
											<TD height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgCho"
																alt="화면을 초기화 합니다." src="../../../images/imgCho.gif"></TD>
														<TD><IMG style="CURSOR: hand" id="ImgSave_AMT" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/ImgSave.gIF'" border="0" name="ImgSave"
																alt="청약 자료를 저장합니다.." src="../../../images/ImgSave.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
																alt="청약 내용을 삭제합니다..." src="../../../images/imgDelete.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" border="0" name="imgPrint"
																alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF"></TD>
														<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
																alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 3px" class="BODYSPLIT"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 30%" class="LISTFRAME" vAlign="top" align="center">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_HDR" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="3889">
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
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 3px" class="BODYSPLIT"></TD>
							</TR>
							<!--Input End-->
							<TR>
								<TD>
									<TABLE id="tblTitle3" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="28" width="50%" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="140" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="1" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td class="TITLE">1) 광고 대행수수료분배</td>
														<TD height="22" vAlign="middle" align="right">
															<TABLE style="HEIGHT: 20px" id="tblButton_OUT" border="0" cellSpacing="0" cellPadding="2">
																<TR>
																	<TD><IMG style="CURSOR: hand" id="ImgAddRow_OUT" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" border="0" name="imgAddRow_OUT"
																			alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgDelete_OUT" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete_OUT"
																			alt="사내수릭자료를 삭제합니다..." src="../../../images/imgDelete.gIF" height="20"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgExcel_OUT" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel_OUT"
																			alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
																</TR>
															</TABLE>
														</TD>
													</tr>
												</table>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<TABLE border="0" cellSpacing="1" cellPadding="0" width="100%" align="left" height="98%">
										<TR>
											<td style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab2"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_OUT" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="100%"
														height="100%">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="31802">
														<PARAM NAME="_ExtentY" VALUE="2698">
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
											</td>
										</TR>
										<TR>
											<TD id="lblStatus_OUT" class="BOTTOMSPLIT"></TD>
										</TR>
										<tr>
											<TD height="28" width="50%" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="115" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="1" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td class="TITLE">2) 수수료 정산관련</td>
														<TD height="22" vAlign="middle" align="right">
															<TABLE style="HEIGHT: 20px" id="tblButton_SUSU" border="0" cellSpacing="0" cellPadding="2">
																<TR>
																	<TD><IMG style="CURSOR: hand" id="ImgAddRow_SUSU" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" border="0" name="ImgAddRow_SUSU"
																			alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgDelete_SUSU" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete_SUSU"
																			alt="SUSU 자료를 삭제합니다." src="../../../images/imgDelete.gIF" height="20"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgExcel_SUSU" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel_SUSU"
																			alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
																</TR>
															</TABLE>
														</TD>
													</tr>
												</table>
											</TD>
										</tr>
										<tr>
											<td style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab3"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_SUSU" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="100%"
														height="100%">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="31802">
														<PARAM NAME="_ExtentY" VALUE="2698">
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
											</td>
										</tr>
										<TR>
											<TD id="lblStatus_SUSU" class="BOTTOMSPLIT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
