<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_INDIRECTCOST.aspx.vb" Inherits="PD.PDCMJOBMST_INDIRECTCOST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>간접비관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_SUBITEM.aspx
'기      능 : JOBMST의 두번째 탭 PDCMJOBMST_ESTDTL 의 간접비처리 버튼을 클릭하였을때 처리 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/28 By KimTH
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
		
Dim mlngRowCnt,mlngColCnt
Dim mlngTempRowCnt,mlngTempColCnt
Dim mobjPDCOPREESTINDIRECTCOST
Dim mstrPREESTNO			'견적번호
Dim mstrCheck	
Dim mstrGBN					'가견적과 본견적 구분
Dim mstrSAVEGBN				'청구요청 진향을 위한 플래그
Dim mstrFIRSTPRODUCTIONCHECK'자동 저장을 위한 플래그
Dim mstrProcessData			'최초 조회시 상세내역의 데이터를 임시 테이블에 저장 하기위함.
Dim mstrCHANGEFALG			'변경확인 플래그(삭제를 할경우 전체 삭제일 경우의 예외를 처리한다.)  [ T/F  (T 일반 이벤트시 / F 삭제 이벤트가 발생할경우)]

mstrCheck = True	
mstrCHANGEFALG = "F"
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
Dim vntData
Dim returnAMT

	with frmThis
	
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		'set mobjPDCOPREESTINDIRECTCOST = gCreateRemoteObject("cPDCO.ccPDCOPREESTINDIRECTCOST")
		'HDRINPUT 에 값이 있다는것은 저장 했다는 뜻.
		vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_returnAMT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
	
		'Set mobjPDCOPREESTINDIRECTCOST = Nothing
		
		if mstrGBN = "가견적" then
			if vntData(0,1) = "" then
				'returnAMT = False
				returnAMT = split("False;" & mstrCHANGEFALG ,";")
			else
				'returnAMT = vntData(0,1) 
				returnAMT = split(vntData(0,1) & ";" & mstrCHANGEFALG ,";")
			end if
			
		elseif mstrGBN = "본견적" then 
		
			if vntData(1,1) = "" then
				'returnAMT = False
				returnAMT = split("False;" & mstrCHANGEFALG ,";")
			else
				'returnAMT = vntData(1,1)
				returnAMT = split(vntData(1,1) & ";" & mstrCHANGEFALG ,";")
			end if
		end if 
		window.returnvalue = returnAMT
	
	end with
	Set mobjPDCOPREESTINDIRECTCOST = Nothing
End Sub

Sub imgClose_onclick ()
	EndPage
End Sub

Sub imgSave_onclick ()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 저장이 불가능 합니다.","저장안내!"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub ImgMoveData_onclick
	Dim intCnt
	with frmThis
		If mstrGBN = "가견적" Then 
			gErrorMsgBox "가견적 상태는 실행견적을 입력할수 없습니다.","입력안내!"
			EXIT SUB
		END IF 
		.txtEXECOMMIRATE.value = .txtCOMMIRATE.value
		.txtEXEAMT.value = .txtAMT.value
		
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"EXECHK",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"CHK", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt)
			'mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt,"1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
	End with
End Sub

Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCOPREESTINDIRECTCOST = gCreateRemoteObject("cPDCO.ccPDCOPREESTINDIRECTCOST")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정

		for i = 0 to intNo
			select case i
				case 0 : mstrPREESTNO = vntInParam(i)			'견적번호
				case 1 : mstrSAVEGBN = vntInParam(i)			'F/T
				case 2 : mstrFIRSTPRODUCTIONCHECK = vntInParam(i) '버튼 클릭 이벤트가 발생하면 Y 로 넘어온다 저장이 되면 N 으로 변경 
				case 3 : mstrGBN = vntInParam(i)				'본견적 가견적
				case 4 : mstrProcessData = vntInParam(i)		'팝업 최초 저장 데이터  
			end select
		next
	
		'**************************************************
		'***Sum Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | EXECHK | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | AMT | EXEAMT | PRINT_SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		    "가선택|본선택|견적번호|순번|항목|코드|견적대분류|견적중분류|견적항목|금액|실행금액|저장구분|정렬순번"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "     4|     4|      10|       4|   4|  10|  10|        10|        15|      12|  12|       0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | EXECHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | EXEAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | PRINT_SEQ"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLASSNAME | ITEMCODENAME",-1,-1,0,2,false ' 왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | PRINT_SEQ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.ColHidden .sprSht, "PRINT_SEQ", true
		
		IF mstrGBN = "가견적" THEN
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EXECHK | EXEAMT"
		ELSE
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CHK | AMT"
		END IF 

		pnlTab1.style.visibility = "visible" 
		
		if .txtAMT.value <> "0" then
			mstrCheck= true
		else 
			mstrCheck= false
		end if
		
		if .txtEXEAMT.value <> "0" then
			mstrCheck= true
		else 
			mstrCheck= false
		end if
	End with

	'화면 초기값 설정
	InitPageData
	'최초 간접비 팝업을 띄울시 메인의 모든 데이터를 가져와서 템프에 저장한후 보여줘야 한다.
	initpageProcess
	
	'조회 
	SelectRtn
End Sub	

Sub EndPage
	gEndPage
End Sub

'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitpageData
	with frmThis

		.txtCOMMIRATE.value = 10
		.txtAMT.value = 0
		.txtEXECOMMIRATE.value = 10
		.txtEXEAMT.value = 0
		
		'.txtEXEAMT.value = 0
		'.txtEXECOMMIRATE.value = 10
		.txtPREESTNO.style.visibility = "hidden"
		
		If mstrSAVEGBN = "T" AND mstrGBN = "본견적" Then
			.txtCOMMIRATE.className = "NOINPUT_R"
			.txtCOMMIRATE.readOnly = true
			.txtAMT.className = "NOINPUT_R"
			.txtAMT.readOnly = true
			.txtEXECOMMIRATE.className = "NOINPUT_R"
			.txtEXECOMMIRATE.readOnly = true
			.txtEXEAMT.className = "NOINPUT_R"
			.txtEXEAMT.readOnly = true
			
		ElseIF  mstrSAVEGBN = "F" AND mstrGBN = "본견적" Then
			.txtCOMMIRATE.className = "NOINPUT_R"
			.txtCOMMIRATE.readOnly = true
			.txtAMT.className = "NOINPUT_R"
			.txtAMT.readOnly = true
			
			.txtEXECOMMIRATE.className = "INPUT_R"
			.txtEXECOMMIRATE.readOnly = false
			.txtEXEAMT.className = "INPUT_R"
			.txtEXEAMT.readOnly = false

		ELSEIF mstrSAVEGBN = "F" AND mstrGBN = "가견적" Then
			.txtCOMMIRATE.className = "INPUT_R"
			.txtCOMMIRATE.readOnly = false
			.txtAMT.className = "INPUT_R"
			.txtAMT.readOnly = false
			
			.txtEXECOMMIRATE.className = "NOINPUT_R"
			.txtEXECOMMIRATE.readOnly = true
			.txtEXEAMT.className = "NOINPUT_R"
			.txtEXEAMT.readOnly = true
		End If
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'-----------------------------------------
'최초 팝업 을 열때 메인 페이지 데이터 저장 
'-----------------------------------------
sub initpageProcess
	with frmthis
	'변수 초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
		'템프테이블의 데이터가 저장되어 있는지 확인하여 저장된값이 없을 경우에만 최초 세부견적내역중 간접비대상을 투입한다.
		vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)	
		if not gDoErrorRtn ("SelectRtn_TempCnt") then
			if mlngRowCnt > 0 Then
			'템프에 값이 없고 
   			Else	
   				vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Cnt(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)	
   				if mlngRowCnt > 0 then 
   					'실제 테이블에도 값이 저장되어 있지 않다면 [최초 보여주는 경우라면!]
   				else 
   					intRtn = mobjPDCOPREESTINDIRECTCOST.ProcessRtn_Indirect(gstrConfigXml,mstrProcessData)
   				end if  
   			end If
   		end if	
	end with
end sub

'================================================================
'UI
'================================================================
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

Sub txtCOMMIRATE_onchange
	Dim intCnt
	Dim dblAMT

	with frmThis
		dblAMT = 0
		For intCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
					dblAMT = dblAMT +  (mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * .txtCOMMIRATE.value * 0.01)
			End If
		Next
		.txtAMT.value = dblAMT
	End with
End Sub

Sub txtEXEAMT_onfocus
	with frmThis
		.txtEXEAMT.value = Replace(.txtEXEAMT.value,",","")
	end with
End Sub
Sub txtEXEAMT_onblur
	with frmThis
		call gFormatNumber(.txtEXEAMT,0,true)
	end with
End Sub

Sub txtEXECOMMIRATE_onchange
	Dim intCnt
	Dim dblEXEAMT

	with frmThis
		dblEXEAMT = 0
		For intCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
					dblEXEAMT = dblEXEAMT +  (mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * .txtEXECOMMIRATE.value * 0.01)
			End If
		Next
		.txtEXEAMT.value = dblEXEAMT
	End with
End Sub

'------------------------
'-----최초 금액 합산 ----
'------------------------
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	
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

'=============================================================
'Sheet Event
'=============================================================
'-----------------------------------------
'시트 에서 키를 떼었을때 선택 금액 합산. 
'-----------------------------------------
Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCOLUMN
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	'키 움직일때 바인딩

	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT")) Then
				
				
				
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
END SUB

'-----------------------------------
'시트에서 마우스를 떼었늘때 이벤트
'-----------------------------------
Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
	
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
																			
				
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
				
				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					End If
				Next
				
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if	
		
	End With
End Sub


Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim dblAMT
	
	with frmThis
		If mstrGBN = "가견적" Then
			
			if Row = 0 and Col = 1 then
			
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
				dblAMT = 0
				.txtAMT.value=0
				for intcnt =1 to .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
						dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * 0.01)
					Else
						dblAMT = dblAMT	- (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * 0.01)
					End If
					mobjSCGLSpr.CellChanged .sprSht, Col, intcnt
				next
				'만약에 dblamt 가 마이너스면 0으로 ....
				if dblAMT < 0 then 
					.txtAMT.value = 0
				else
					.txtAMT.value = dblAMT
				end if
				
				'플래그 변경
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
			end if
			
		ELSE
		
			if Row = 0 and Col = 2 then
				
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 2, 2, , , "", , , , , mstrCheck
				dblAMT = 0
				.txtEXEAMT.value=0
				
				for intcnt =1 to .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"EXECHK",intcnt) = 1 THEN
						dblAMT = dblAMT	+ (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * 0.01)
					Else
						dblAMT = dblAMT	- (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * 0.01)
					End If
					mobjSCGLSpr.CellChanged .sprSht, Col, intcnt
				next
				'만약에 dblamt 가 마이너스면 0으로 ....
				if dblAMT < 0 then 
					.txtEXEAMT.value = 0
				else
					.txtEXEAMT.value = dblAMT
				end if
				
				'플래그 변경
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
			end if
		END IF 
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	end with	
End Sub




Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

'-----------------------------------
'----------SprSht change------------
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim dblAMT
	Dim intCnt 
	
	with frmThis
		If mstrGBN = "가견적" Then
			
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				dblAMT = .txtAMT.value 
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 THEN
					dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) * 0.01)
				Else
					dblAMT = dblAMT	- (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) * 0.01)
				End If
				.txtAMT.value = dblAMT
				
			ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") THEN
			
				dblAMT =0
				
				FOR intCnt=1 to .sprSht.Maxrows
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 then
						dblAMT = dblAMT + mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
					end if 
				Next
				
				dblAMT = (.txtCOMMIRATE.value * dblAMT * 0.01)
				.txtAMT.value = dblAMT
			End If
		
		ELSE
			dblAMT = .txtEXEAMT.value 
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXECHK") Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"EXECHK",Row) = 1 THEN
					dblAMT = dblAMT	+ (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) * 0.01)
				Else
					dblAMT = dblAMT	- (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) * 0.01)
				End If
				.txtEXEAMT.value = dblAMT
				
			ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") THEN
				dblAMT =0
				
				FOR intCnt=1 to .sprSht.Maxrows
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 then
						dblAMT = dblAMT + mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intCnt)
					end if 
				Next
				
				dblAMT = (.txtCOMMIRATE.value * dblAMT * 0.01)
				.txtEXEAMT.value = dblAMT
			End If
		End If
		mobjSCGLSpr.CellChanged .sprSht,.sprSht.ActiveCol+1,.sprSht.ActiveRow
		txtAMT_onblur
		txtEXEAMT_onblur
	End with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
   	
End Sub



Sub RATESUM
	Dim dblAMT
	Dim intCnt
	with frmThis
		dblAMT = 0
			For intCnt = 1 To .sprSht.MaxRows
				
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
					dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) * 0.01)
				End If
				.txtAMT.value = dblAMT
			
			Next
	End with
End Sub
'=============================================================
'조회
'=============================================================
Sub SelectRtn
	IF not SelectRtn_Head () Then Exit Sub
	CALL SelectRtn_Detail ()
	RATESUM
	txtAMT_onblur
	txtEXEAMT_onblur
	mstrCHANGEFALG = "F"
End Sub

'=============================================================
'------------------상단의 텍스트박스 조회---------------------
'=============================================================
Function SelectRtn_Head()
	Dim vntData
	Dim vntData_temp
	
	'on error resume next
	'초기화
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'임시 테이블을 조회한다.
		vntData_temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempHDRCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		if mlngRowCnt > 0 then 
		
			vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempHDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		else
			'임시 테이블이 없다면 실제 저장될 테이블의 유무를 조회한다.
			vntData_temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_HDRCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			if mlngRowCnt > 0 then
				'실제 저장되어있는 테이블이 있다면 저장되어있는 테이블을 가져온다
				vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			end if 
			'템프와 실제 저장되어있는 간접비가 없다면 상단의 화면에는 최초화면의 값을 가지고있는다.
		end if 
	
	IF not gDoErrorRtn ("SelectRtn_TempHDR") then
		'조회한 데이터를 바인딩
		If mlngRowCnt > 0 Then
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			
		End If
		SelectRtn_Head = True
	End IF
End Function

'=============================================================
'------------------하단의 그리드 조회---------------------
'=============================================================
Function SelectRtn_Detail()
	Dim vntData
   	Dim vntData_Temp
   	Dim vntData_TempCNT
    
	'On error resume next
	with frmThis
	
	'Long Type의 ByRef 변수의 초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

		'디테일 내역의 데이터 유무 확인
		vntData_TempCNT = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		if mlngRowCnt > 0 then
			vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Temp(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			
			if not gDoErrorRtn ("SelectRtn_Temp") then
				if mlngRowCnt > 0 Then
					call mobjSCGLSpr.SetClipbinding (.sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True)
					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				Else	
   					.sprSht.MaxRows = 0
   					gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   				end If
   			end if
		else
			vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Detail(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			
			if not gDoErrorRtn ("SelectRtn_Detail") then
				if mlngRowCnt > 0 Then
					call mobjSCGLSpr.SetClipbinding (.sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True)
					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   				Else	
   					.sprSht.MaxRows = 0
   					gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   				end If
   			end if		
		end if 
		window.setTimeout "AMT_SUM",1	
		txtAMT_onblur
		txtEXEAMT_onblur
   	end with
End Function

'======================================
'---------------저장-------------------
'======================================
Sub processRtn
	Dim vntData
	Dim intRtn
	with frmThis
		
		'XML데이터 상단의 박스의 데이터를 가져온다.
		strMasterData = gXMLGetBindingData (xmlBind)

		'insert 플래그 변경 [모든 데이터 가져오기]
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | EXECHK | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | AMT | EXEAMT | PRINT_SEQ")
		
		if  not IsArray(vntData)  then
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If

		' 간접비의 저장도 모두 input 에저장된다.본테이블에 저장이 되었다가 최종 저장을 안할경우 금액이 다른부분을 막는다.
		'DELETE INSERT 
		intRtn = mobjPDCOPREESTINDIRECTCOST.ProcessRtn(gstrConfigXml,strMasterData,vntData,mstrPREESTNO)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
			.sprSht.focus()
			mstrCHANGEFALG = "T"
		End If

	End with
	mstrFIRSTPRODUCTIONCHECK = "N"
End Sub

		</script>
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="82" background="../../../images/back_p.gIF"
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
								<td class="TITLE">CF간접비관리</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="SEARCHDATA" width="100%">
				<tr>
					<td class="SEARCHLABEL" width="50">간접비율
					</td>
					<td class="SEARCHDATA" width="70"><INPUT dataFld="COMMIRATE" class="INPUT_R" id="txtCOMMIRATE" title="간접비율" style="WIDTH: 70px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtCOMMIRATE">&nbsp;%</td>
					<td class="SEARCHLABEL" width="40">간접비</td>
					<td class="SEARCHdata" width="112"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="간접비" style="WIDTH: 112px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="13" name="txtAMT"></td>
					<td class="SEARCHLABEL" width="80">실행간접비율
					</td>
					<td class="SEARCHDATA" width="70"><INPUT dataFld="EXECOMMIRATE" class="INPUT_R" id="txtEXECOMMIRATE" title="간접비율" style="WIDTH: 70px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtEXECOMMIRATE">&nbsp;%</td>
					<td class="SEARCHLABEL" width="70">실행간접비</td>
					<td class="SEARCHdata" width="112"><INPUT dataFld="EXEAMT" class="INPUT_R" id="txtEXEAMT" title="간접비" style="WIDTH: 112px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="13" name="txtEXEAMT"></td>
					<td class="SEARCHLABEL" width="80">상세견적비고</td>
					<td class="SEARCHDATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="상세견적비고" style="WIDTH: 200px; HEIGHT: 20px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="65" name="txtMEMO"></td>
					<td class="SEARCHDATA" width="54"><INPUT dataFld="PREESTNO" class="INPUT" id="txtPREESTNO" title="간접비" style="WIDTH: 48px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="2" name="txtPREESTNO"></td>
					<td class="SEARCHDATA" width="54"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="화면을 닫습니다."
							src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
				</tr>
			</table>
			</TABLE><BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">합 계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSUMAMT">
					</td>
					<td class="TITLE">선택합계 : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="합계금액" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<TD align="right" width="600"><IMG id="ImgMoveData" onmouseover="JavaScript:this.src='../../../images/ImgMoveDataOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMoveData.gIF'" height="20" alt="가견적내역 을 실행견적으로 복제합니다."
							src="../../../images/ImgMoveData.gIF" align="absMiddle" border="0" name="ImgMoveData">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
							onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" align="absMiddle" border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</TD>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--내용-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id=sprSht classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 width="100%" height="100%">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="_ExtentX" VALUE="31750">
	<PARAM NAME="_ExtentY" VALUE="21060">
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
	<PARAM NAME="ReDraw" VALUE="-1">
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
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
				</TR>
			</table>
		</form>
	</body>
</HTML>
