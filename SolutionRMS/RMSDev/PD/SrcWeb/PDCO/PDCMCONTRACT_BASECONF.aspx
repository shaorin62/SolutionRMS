<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT_BASECONF.aspx.vb" Inherits="PD.PDCMCONTRACT_BASECONF" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>기본 & 단가 계약서 체결 등록 및 조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 기본계약서 체결 등록 및 조회
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 기본 & 단가 계약서 체결 등록 및 조회/저장/삭제
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/21 By 황덕수
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
Dim mobjPDCMCONTRACT_BASE, mobjPDCMGET
Dim mstrCheck
Dim mstrmode
Dim mstrCHKcheck
Dim mstrCONFIRM

CONST meTAB = 9

mstrCheck = True
mstrmode = True
mstrCHKcheck = True
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

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'신규버튼
Sub imgREG_onclick ()
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

sub imgAgree_onclick ()
	mstrCONFIRM = "1"
	gFlowWait meWAIT_ON
	UpdateRtn_CONFIRM(mstrCONFIRM)
	gFlowWait meWAIT_OFF
end sub

sub imgAgreeCanCel_onclick ()
	mstrCONFIRM = "0"
	gFlowWait meWAIT_ON
	UpdateRtn_CONFIRM(mstrCONFIRM)
	gFlowWait meWAIT_OFF
end sub

'계약서 생성 이벤트
Sub imgSave_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub

'삭제버튼 이벤트
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
	
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim intRtn
	Dim i, j, intCount
	Dim strCONTRACTNO
	Dim strUSERID
	Dim vntDataTemp

		'체크가 된 데이터가 있는지 없는지 체크한다.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
				if mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CONTRACTNO",i) = "" then
					gErrorMsgBox "계약서가 생성되지 않았습니다. 생성한후 출력하세요"," 계약서 출력 안내!"
					Exit Sub	
				end if
				intCount = 1
			end if
		next

		'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
		if intCount = 0 then
			gErrorMsgBox "선택된 데이터가 없습니다. 인쇄할 데이터를 체크하시오",""
			Exit Sub
		End if

		gFlowWait meWAIT_ON
		with frmThis
			'인쇄버튼을 클릭하기 전에 md_tax_temp테이블에 내용을 삭제한다
			'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
			'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
			intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn_TEMP(gstrConfigXml)

			ModuleDir = "PD"
			
			IF .rdCONFLAG.checked THEN 
				'기본 계약서
				if .cmbGUBUN.value = "기본" then
					ReportName = "PDCMCONTRACTBASE_CON_N.rpt"			
				'단가 계약서
				elseif .cmbGUBUN.value = "단가" then
					ReportName = "PDCMCONTRACTPRICE_CON_N.rpt"			
				end if
			End if 

			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strCONTRACTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					strUSERID = ""
					vntDataTemp = mobjPDCMCONTRACT_BASE.ProcessRtn_TEMP(gstrConfigXml,strCONTRACTNO, i, strUSERID)
				END IF
			next

			Params = strUSERID 
			Opt = "A"
			
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
			'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
			window.setTimeout "printSetTimeout", 10000
		End with
		gFlowWait meWAIT_OFF
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn_TEMP(gstrConfigXml)
	end with
End sub

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
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' 코드명 표시
			selectrtn
     	End if
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
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
					selectrtn
				Else
					Call SEARCHOUT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 날자컨트롤 및 달력 /
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgTO_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgTO,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub imgFROM2_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtSTDATE,frmThis.imgFROM,"txtSTDATE_onchange()"
		gSetChange
	end with
End Sub

Sub imgTO2_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtEDDATE,frmThis.imgTO,"txtEDDATE_onchange()"
		gSetChange
	end with
End Sub

Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

Sub txtSTDATE_onchange
	IF frmthis.rdCONFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtSTDATE.value 	
	end if 
	gSetChange
End Sub

Sub txtCONTRACTNAME_onchange
	mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTNAME",frmThis.sprSht.ActiveRow, frmthis.txtCONTRACTNAME.value
	gSetChange
End Sub

Sub txtCONTRACTDAY_onchange
	frmthis.txtSTDATE.value = frmthis.txtCONTRACTDAY.value 
	gSetChange
End Sub

Sub txtDELIVERYDAY_onchange
	gSetChange
End Sub

Sub chkCONFIRMFLAG_onClick
	if frmThis.sprSht.ActiveRow > 0  Then
		if frmThis.chkCONFIRMFLAG.checked = TRUE Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "1"
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "0"
		End if
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub cmbGBN_onchange
	Dim strHTML
	with frmThis
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN",frmThis.sprSht.ActiveRow, frmthis.cmbGBN.value
		
		if .cmbGBN.value = "BTL" then
			document.getElementById("test").innerHTML = ""
			.cmbGUBUN.value = "기본"
			.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
			cmbGUBUN_onchange
		else
			document.getElementById("test").innerHTML = "단가"
			.cmbGUBUN.className				= "INPUT"   : .cmbGUBUN.disabled			= false 
			cmbGUBUN_onchange
		end if
		 
		
	end with
	gSetChange
End Sub

Sub cmbGUBUN_onchange
	with frmThis
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, frmthis.cmbGUBUN.value
		'기본계약서의 경우 계약기간을 정하지 않는다.
		if frmThis.cmbGUBUN.value = "기본" then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, ""
			.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
			.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
		
			frmThis.txtSTDATE.value = ""
			frmThis.txtEDDATE.value = ""
		else
			frmThis.txtSTDATE.value		= gNowDate
			DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
			.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
			.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
		end if
		
	end with
	gSetChange
End Sub

Sub txtCOMENT_onchange
	mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMENT",frmThis.sprSht.ActiveRow, frmthis.txtCOMENT.value
	gSetChange
End Sub


Sub cmbGUBUN1_onchange
	with frmThis
		SelectRtn
	end with
end sub

'-----------------------------------
'필드추가 
'------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		frmThis.sprSht.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)		
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CHK",frmThis.sprSht.ActiveRow, 1
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTNO",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, "기본"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN",frmThis.sprSht.ActiveRow, "ATL"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, "0"
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht, false, "OUTSCODE | BTN  | OUTSNAME"
		
		sprShtToFieldBinding 2, 1
		
		if frmThis.cmbGUBUN.value = "기본" then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, ""
		
			frmThis.txtSTDATE.value = ""
			frmThis.txtEDDATE.value = ""
		else
			frmThis.txtSTDATE.value		= gNowDate
			DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
			
			frmThis.sprSht.focus()
		end if	
	End If
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	with frmThis
		If Row = 0 and Col = 1  then 
			mstrCHKcheck = false
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mstrCHKcheck = true


			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if

			for intcnt = 1 to .sprSht.MaxRows
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intcnt
			next
			
		ELSE
			if Row > 0 then
				sprShtToFieldBinding Col, Row
			end if
   			
		End if		
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	DIM vntData
	Dim dblAMT
	Dim intcnt, intcount
	DIM strYEARMON
	DIM strCode		
	DIM strCodeName
	
	with frmThis
	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",trim(strCodeName))

				If not gDoErrorRtn ("GetEXECUSTNO") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						
						.txtOUTSNAME.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME"), Row
						.txtOUTSNAME.focus()
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
   	End with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row)))
			
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		.sprSht.Focus
	End With
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet
	Dim vntInParams
	Dim dblAMT
	
	with frmThis
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then			
		
				vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row)))
				
				vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
				If isArray(vntRet) Then
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)		
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
					
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				End If
			End if
			
			if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CONFIRMFLAG") then
				if mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "1" then
					.chkCONFIRMFLAG.checked = true
				else
					.chkCONFIRMFLAG.checked = false
				end if
			end if 
		.sprSht.Focus
	End with
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
		
	with frmThis		
	
		If KeyCode = 229 Then Exit Sub
		
		If KeyCode <> meCR and KeyCode <> meTab _
			and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
			and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
			and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

		If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
				sprShtToFieldBinding .sprSht.ActiveCol, .sprSht.ActiveRow
		End If
		
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")  Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"))  Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			'.txtSELECTAMT.value = strSUM
			'Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			'.txtSELECTAMT.value = 0
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
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
						'.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
						End If
					Next
					'.txtSELECTAMT.value = strSUM
				End If
			else
				'.txtSELECTAMT.value = 0
			End If
		else
			'.txtSELECTAMT.value = 0
		End If
		'Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

'화면 초기 시트 이미지 생성 및 설정 초기화 
Sub InitPage()
	'서버업무객체 생성	
	set mobjPDCMCONTRACT_BASE	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT_BASE")
	set mobjPDCMGET				= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | GBN | GUBUN | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG "
		mobjSCGLSpr.SetHeader .sprSht,		 "선택|계약종류|계약서구분|계약서번호|외주처코드|외주처명|계약명|금액|계약일|계약시작일|계약종료일|특약사항|승인"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|       8|         9|        10|         9|2|    15|    16|   8|     8|        10|        10|      12|   4"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | CONFIRMFLAG"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"♡", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONTRACTDAY | STDATE | EDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONTRACTNO | OUTSCODE | OUTSNAME | COMENT ", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "CONTRACTNO | GBN | GUBUN"
		'mobjSCGLSpr.ColHidden .sprSht, "SEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GBN | GUBUN | OUTSCODE",-1,-1,2,2,False
		mobjSCGLSpr.CellGroupingEach .sprSht," OUTSNAME"

		.sprSht.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	

	pnlTab1.style.visibility = "visible"
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT_BASE = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub


'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	With frmThis
		.sprSht.MaxRows = 0
		
		frmThis.txtCONTRACTDAY.value= gNowDate
		frmThis.txtFROM.value		= Mid(gNowDate,1,4) & "-"  & Mid(gNowDate,6,2) & "-" & "01"
		frmThis.txtSTDATE.value		= gNowDate
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		
		.txtCONTRACTDAY.value = gNowDate
		.txtCOMENT.value  = ""
		.cmbCONFIRM.value = ""
		
		Field_Lock
	End With
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	
End Sub


'청구일 조회조건 생성
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtTo.value = date2
	end if
End Sub

Sub DateClean2 (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtEDDATE.value = date2
	end if
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '그리드 데이터가 없으면 나간다.
		.txtCONTRACTNAME.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
		.txtCONTRACTDAY.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
		.txtSTDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
		.txtEDDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
		
		.cmbGUBUN.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",Row)
		.cmbGBN.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",Row)
		
		.txtCOMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMENT",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",Row) = "1" THEN
			.chkCONFIRMFLAG.checked = TRUE
		ELSE
			.chkCONFIRMFLAG.checked = FALSE
		END IF
		
		Field_Lock
	End with
End Function

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	With frmThis
		IntAMTSUM = 0
		
		If .sprSht.MaxRows > 0 Then
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		ELSE
			.txtSUMAMT.value = 0
		END IF
	End With
End Sub

'-----------------------------------------------------------------------------------------
' Field_Lock  상황에 따라서 수정할수 없도록 필드를 ReadOnly처리
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",.sprSht.ActiveRow) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",.sprSht.ActiveRow) = "1" Then
			
				.txtCONTRACTNAME.className		= "NOINPUT_L" : .txtCONTRACTNAME.readOnly	= True 
				.txtCONTRACTDAY.className		= "NOINPUT"	  : .txtCONTRACTDAY.readOnly	= True
				.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
				.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
				.txtCOMENT.className			= "NOINPUT"   : .txtCOMENT.readOnly			= True 
				.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
				.cmbGBN.className				= "NOINPUT"   : .cmbGBN.disabled			= True 
				
				.ImgCONTRACTDAY.disabled = true
				.imgFROM2.disabled = true
				.imgTO2.disabled = true
			
			'계약서 번호가 있고 승인이 되지 않은데이터의 경우 계약구분과 종류만 잠근다.
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",.sprSht.ActiveRow) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",.sprSht.ActiveRow) = "0" then
			'기본계약서의 경우 계약기간을 정하지 않는다.
				if .cmbGUBUN.value = "기본" then
					.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
					.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
					
				else
					.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
					.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
					
				end if
				.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
				.cmbGBN.className				= "NOINPUT"   : .cmbGBN.disabled			= True 
				.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
				.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
				.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
			
				.ImgCONTRACTDAY.disabled = False
				.imgFROM2.disabled = False
				.imgTO2.disabled = False
			Else 
				'기본계약서의 경우 계약기간을 정하지 않는다.
				if .cmbGUBUN.value = "기본" then
					.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
					.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
					
				else
					.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
					.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
					
				end if
				.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
				.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
				.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
				.cmbGUBUN.className				= "INPUT"   : .cmbGUBUN.disabled			= False
				.cmbGBN.className				= "INPUT"   : .cmbGBN.disabled			= False
				
				.ImgCONTRACTDAY.disabled = False
				.imgFROM2.disabled = False
				.imgTO2.disabled = False
			End If
		Else
			'기본계약서의 경우 계약기간을 정하지 않는다.
			if .cmbGUBUN.value = "기본" then
				.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
				.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
			else
				.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
				.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
			end if
			.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
			.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
			.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
			.cmbGUBUN.className				= "INPUT"	: .cmbGUBUN.disabled			= False
			.cmbGBN.className				= "INPUT"	: .cmbGBN.disabled			= False
			

			.ImgCONTRACTDAY.disabled = False
			.imgFROM2.disabled = False
			.imgTO2.disabled = False
		End If
	End With
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strFROM, strTO
	Dim strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN,strcmbGBN
	Dim i,j,strRows

	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		j = 1

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strFROM			= .txtFrom.value
		strTO			= .txtTo.value
		strOUTSCODE		= .txtOUTSCODE.value
		strOUTSNAME		= .txtOUTSNAME.value
		strCONFIRM		= .cmbCONFIRM.value
		strCONTRACTNO	= .txtCONTRACTNO.value
		strCONTRACTNAME1 = .txtCONTRACTNAME1.value
		strcmbGBN		= .cmbGBN1.value
		strcmbGUBUN		= .cmbGUBUN1.value

		vntData = mobjPDCMCONTRACT_BASE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strFROM,strTO, _ 
												  strOUTSCODE,strOUTSNAME,strCONFIRM,strCONTRACTNO, _ 
												  strCONTRACTNAME1,strcmbGBN, strcmbGUBUN)

		If not gDoErrorRtn ("SelectRtn") Then
			InitPageData
   			PreSearchFiledValue strFROM,strTO, strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE

   				for i = 1 to .sprSht.MaxRows
   					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "1" then
   					
						If j = 1 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
						j = j + 1
						
					End If
   				next
   				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,11,True

   				'AMT_SUM
   				sprShtToFieldBinding 2, 1
   			else
   				.sprSht.MaxRows = 0
   			End If
   		End If
   	end With
End Sub

'****************************************************************************************
'이전 검색어를 담아 놓는다.
'****************************************************************************************
Sub PreSearchFiledValue (strFROM,strTO, strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN)
	With frmThis
		.txtFrom.value			= strFROM
		.txtTo.value			= strTO
		.txtOUTSCODE.value		= strOUTSCODE
		.txtOUTSNAME.value		= strOUTSNAME
		.cmbCONFIRM.value		= strCONFIRM
		.txtCONTRACTNO.value	= strCONTRACTNO
		.txtCONTRACTNAME1.value = strCONTRACTNAME1
		.cmbGUBUN1.value		= strcmbGUBUN
	End With
End Sub

'------------------------------------------
' 계약서 생성
'------------------------------------------
Sub ProcessRtn_HDR ()
   	Dim intRtn
   	Dim strMasterData
   	Dim vntData
	Dim lngchkCnt
	Dim i
	Dim strOUTSCODE
	Dim strCONTRACTNO
	Dim strGUBUN
	
	With frmThis
		strMasterData = gXMLGetBindingData (xmlBind)
		lngchkCnt = 0 :  strCONTRACTNO = "" : strGUBUN = ""

		For i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
				strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",i)
				lngchkCnt = lngchkCnt +1
				'데이터 시트의 변경 플래그를 변경한다.
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
			End If 
			
		Next
		
		if strOUTSCODE = "" then
			gErrorMsgBox "외주처 입력은 필수 사항입니다.","계약서 생성 안내!"
			Exit Sub
		end if 

		If lngchkCnt = 0 Then
			gErrorMsgBox "생성할 계약서를 체크해 주세요.","확정안내!"
			EXIT Sub
		End If

		'저장할 데이터 체크
		if DataValidation (strOUTSCODE) = false then exit sub
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | GBN | GUBUN | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG ")

		if  not IsArray(vntData)  then 
			gErrorMsgBox "변경된 입력필드 " & meNO_DATA,"저장안내"
			exit sub
		End If

		intRtn = mobjPDCMCONTRACT_BASE.ProcessRtn(gstrConfigXml, strMasterData, vntData, strCONTRACTNO, strOUTSCODE)

		If not gDoErrorRtn ("ProcessRtn_HDR") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox strCONTRACTNO & " 번호로 확정 되었습니다.","저장안내!"
			SelectRtn
   		End If
   	End With
End Sub

'-------------------------------------------------
'이전에 기본계약서나 단가계약서가 존해 하는지 확인
'-------------------------------------------------
Function DataValidation (strOUTSCODE)
	DataValidation = false
	
	Dim vntData
	Dim intYNRtn
	
	'On error resume next
	with frmThis
		'Master 입력 데이터 Validation : 필수 입력항목 검사 
   		'IF not gDataValidation(frmThis) then exit Function
   		
   		mlngRowCnt=clng(0): mlngColCnt=clng(0)
   		
   		vntData = mobjPDCMCONTRACT_BASE.SelctRtn_validation(gstrConfigXml, mlngRowCnt,mlngColCnt, strOUTSCODE)
   		
   		If not gDoErrorRtn ("SelctRtn_validation") Then
			
			'해당 데이터가 일단있다면 계약서 종류를 확인하고 메시지를 띄운다.
			If mlngRowCnt >0 Then
				
				'기본 계약서일경우
				if vntData(0,1) = "기본" then
					if .cmbGUBUN.value = "기본" then
						gErrorMsgBox "해당 외주처의 기본계약서가 존재 합니다." ,"저장안내"
						exit function
					elseif .cmbGUBUN.value = "단가" then
						intYNRtn = gYesNoMsgbox("기본계약서가 존해 합니다 단가계약서를 추가 생성 하시겠습니까?","확정확인")
						IF intYNRtn <> vbYes then exit function	
					end if
	
				elseif vntData(0,1) = "단가" then
					if .cmbGUBUN.value = "기본" then
						gErrorMsgBox "해당 외주처의 단가계약서가 존재 합니다." ,"저장안내"
						exit function
					elseif .cmbGUBUN.value = "단가" then
						gErrorMsgBox "해당 외주처의 단가계약서가 존재 합니다." ,"저장안내"
						exit function
					end if
				end if
   			End If
   		End If
   		
   	End with
	DataValidation = true
End Function

'------------------------------------------
' 계약서 삭제 데이터 처리
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCONTRACTNO
	Dim lngchkCnt
	
	with frmThis
		lngchkCnt = 0
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
				
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "1" then
					gErrorMsgBox "승인된 데이터는 삭제 하실 수 없습니다..","자료 삭제 안내!"
					EXIT Sub		
				end if
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","확정취소안내!"
			EXIT Sub
		End If

		'선택된 자료를 끝에서 부터 삭제
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub

		For i = .sprSht.MaxRows to 1 step -1
			strCONTRACTNO = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
					strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn(gstrConfigXml, strCONTRACTNO)
					IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				else
					mobjSCGLSpr.DeleteRow .sprSht,i
				End IF
   			End If
		Next
		gOkMsgBox "자료가 삭제 되었습니다.","삭제 안내!"
		gWriteText lblstatus, "자료가 " & intRtn & " 건 삭제되었습니다."
	End with
	err.clear
End Sub

Sub UpdateRtn_CONFIRM (confirm)
   	Dim i, lngchkCnt
   	Dim intRtn
   	Dim vntData
   	Dim strMSG

	With frmThis
		
		If confirm = 1 then
			strMSG = "승인"
		Else 
			strMSG = "승인 취소"
		End if 

		For i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = confirm then
					gErrorMsgBox "이미 " & strMSG & " 된 데이터 입니다..","데이터 승인 안내!"
					EXIT Sub
				end if 
				lngchkCnt = lngchkCnt +1
			End If 
		Next

		If lngchkCnt = 0 Then
			gErrorMsgBox strMSG & "할 계약서를 체크해 주세요.","확정안내!"
			EXIT Sub
		End If

		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG | GUBUN")

		if  not IsArray(vntData)  then 
			gErrorMsgBox "변경된 입력필드 " & meNO_DATA,"저장안내"
			exit sub
		End If

		intRtn = mobjPDCMCONTRACT_BASE.Processrtn_CONFIRM(gstrConfigXml, vntData,confirm)
		If not gDoErrorRtn ("Processrtn_CONFIRM") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox " 데이터가 " & strMSG & " 되었습니다.","저장안내!"
			SelectRtn
   		End If
   	End With
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
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
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="150" background="../../../images/back_p.gIF"
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
												<td class="TITLE">기본 &amp; 단가 계약서 체결</td>
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
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
							<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand; HEIGHT: 9px" onclick="vbscript:Call gCleanField(txtfrom,txtTo)"
										width="56">계약기간</TD>
									<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 9px"><INPUT class="INPUT" id="txtFrom" title="계약검색 시작일자" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtFrom"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="계약검색 종료일자" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtTo"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgTo">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 64px; CURSOR: hand; HEIGHT: 9px">계약서확인</TD>
									<TD class="SEARCHDATA" style="WIDTH: 181px; CURSOR: hand; HEIGHT: 9px"><SELECT id="cmbCONFIRM" style="WIDTH: 120px" name="cmbCONFIRM">
											<OPTION value="" selected>전체</OPTION>
											<OPTION value="0">계약서 미승인</OPTION>
											<OPTION value="1">계약서 승인</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 60px; HEIGHT: 9px" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)">외주처</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 9px" colSpan="3"><INPUT class="INPUT_L" id="txtOUTSNAME" title="외주처명 조회" style="WIDTH: 160px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtOUTSNAME"> <IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
										<INPUT class="INPUT" id="txtOUTSCODE" title="외주처코드조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtOUTSCODE">
									</TD>
									<td><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
											height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" align="right" border="0"
											name="imgQuery">
									</td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')">계약서번호</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCONTRACTNO" title="계약서번호 조회" style="WIDTH: 240px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME1, '')">계약명</TD>
									<TD class="SEARCHDATA" style="WIDTH: 181px"><INPUT class="INPUT_L" id="txtCONTRACTNAME1" title="계약명명 조회" style="WIDTH: 180px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="30" name="txtCONTRACTNAME1"></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField('', '')">계약서종류</TD>
									<TD class="SEARCHDATA" style="WIDTH: 116px"><SELECT id="cmbGBN1" title="계약서 종류" style="WIDTH: 112px" name="cmbGBN1">
											<OPTION value="" selected>전체</OPTION>
											<OPTION value="ATL">ATL</OPTION>
											<OPTION value="BTL">BTL</OPTION>
										</SELECT>
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField('', '')">계약서구분</TD>
									<TD class="SEARCHDATA"><SELECT id="cmbGUBUN1" title="계약서 구분" style="WIDTH: 112px" name="cmbGUBUN1">
											<OPTION value="" selected>전체</OPTION>
											<OPTION value="기본">기본</OPTION>
											<OPTION value="단가">단가</OPTION>
										</SELECT>
									</TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" width="300" height="20">
										<table id="TABLE1" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="180" background="../../../images/back_p.gIF"
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
												<td class="TITLE">기본&amp; 단가 계약서 승인 및 조회</td>
											</tr>
										</table>
									</TD>
									<!--<td>
										<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td class="TITLE" vAlign="middle" align="left" height="20">&nbsp;합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
														accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
													<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
														readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</td>
									-->
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<!--<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="신규자료를 생성합니다."
														src="../../../images/imgNew.gIF" border="0" name="imgREG"></TD> -->
												<td><IMG id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
														height="20" alt="선택한 행을 승인합니다." src="../../../images/imgAgree.gIF" align="absMiddle"
														border="0" name="imgAgree"><IMG id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'" height="20" alt="선택한 행을 승인취소 합니다."
														src="../../../images/imgAgreeCanCel.gIF" align="absMiddle" border="0" name="imgAgreeCanCel">
												</td>
												<!--<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
														src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
												<TD width="15"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD>
														-->
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
							<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 11px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1"
											cellPadding="0" align="left" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')">계약명</TD>
												<TD class="SEARCHDATA" style="WIDTH: 289px" colSpan="3"><INPUT dataFld="CONTRACTNAME" class="INPUT_L" id="txtCONTRACTNAME" title="계약명" style="WIDTH: 240px; HEIGHT: 21px"
														accessKey=",M" dataSrc="#xmlBind" type="text" size="30" name="txtCONTRACTNAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTDAY,'')">계약일</TD>
												<TD class="SEARCHDATA" style="WIDTH: 200px"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="계약일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY">
													<IMG id="ImgCONTRACTDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" alt="ImgCONTRACTDAY" src="../../../images/btnCalEndar.gIF" align="absMiddle"
														border="0" name="ImgCONTRACTDAY">&nbsp;&nbsp;계약서승인<INPUT dataFld="CONFIRMFLAG" id="chkCONFIRMFLAG" title="계약서승인" dataSrc="#xmlBind" type="checkbox"
														value="" name="chkCONFIRMFLAG"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSTDATE,txtEDDATE)">계약 
													기간</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="계약기간 시작일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtSTDATE">
													<IMG id="imgFROM2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgFROM2">&nbsp;~
													<INPUT dataFld="EDDATE" class="INPUT" id="txtEDDATE" title="계약기간 종료일" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtEDDATE">
													<IMG id="imgTO2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgTO2">
												</TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">계약서종류</TD>
												<TD class="SEARCHDATA" style="WIDTH: 100px; HEIGHT: 25px"><SELECT dataFld="GBN" id="cmbGBN" title="계약서종류" style="WIDTH: 100px" dataSrc="#xmlBind"
														name="cmbGBN">
														<OPTION value="ATL" selected>ATL</OPTION>
														<OPTION value="BTL">BTL</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 82px">계약서구분</TD>
												<TD class="SEARCHDATA" style="WIDTH: 99px" HEIGHT:style="WIDTH: 102px"><SELECT dataFld="GUBUN" id="cmbGUBUN" title="계약서 구분" style="WIDTH: 100px" dataSrc="#xmlBind"
														name="cmbGUBUN">
														<OPTION value="기본" selected>기본</OPTION>
														<OPTION value="단가" id="test">단가</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="HEIGHT: 21px">출력사항</TD>
												<TD class="SEARCHDATA" style="HEIGHT: 21px" colSpan="5"><INPUT dataFld="CONFLAG" id="rdCONFLAG" title="광고용역 단가 계약서" dataSrc="#xmlBind" type="radio"
														CHECKED value="rdCONFLAG" name="rdCONFLAG">광고용역 기본 &amp; 단가 계약서
												</TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 53px" onclick="vbscript:Call gCleanField(txtCOMENT, '')">특약사항</TD>
												<TD class="SEARCHDATA" colSpan="7"><TEXTAREA dataFld="COMENT" id="txtCOMENT" style="WIDTH: 778px" dataSrc="#xmlBind" name="txtCOMENT"
														wrap="hard" cols="10"></TEXTAREA></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
					</TR>
					<tr>
						<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="48763">
									<PARAM NAME="_ExtentY" VALUE="11324">
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
					</tr>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
	</body>
</HTML>
