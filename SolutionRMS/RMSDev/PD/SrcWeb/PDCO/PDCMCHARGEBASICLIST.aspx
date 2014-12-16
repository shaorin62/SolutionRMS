<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGEBASICLIST.aspx.vb" Inherits="PD.PDCMCHARGEBASICLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMCHARGEBASICLIST.aspx
'기      능 : JOBMST의 세번째 탭 - 제작리스트 엑셀 업로드
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/24 By 황덕수
'****************************************************************************************
-->
		<meta content="False" name="vs_snapToGrid">
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<script language="vbscript" id="clientEventHandlersVBS">
option explicit


Dim mlngRowCnt, mlngColCnt		
Dim mobjccPDDCCHARGEEXCOM
Dim mobjPDCOGET
Dim mcomecalender
Dim mstrCheck
Dim mstrGrid
Dim strPARENTJOBNO

CONST meTAB = 9
mcomecalender = FALSE
mstrCheck=True
mstrGrid = FALSE


'=============================
' 이벤트 프로시져 
'=============================
'입력 필드 숨기기
Sub Set_TBL_HIDDEN(byVal strmode)
	With frmThis
		If  strmode = "EXTENTION"  Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "60%"
			document.getElementById("tblSheet2").style.height = "30%"
		ELSEIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ELSEIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "60%"
		END IF
	End With
End Sub

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'조회버튼
Sub imgQuery_onclick
	mstrGrid = TRUE
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub		

'신규버튼
Sub imgNEW_onclick ()
	mstrGrid = False
	Call sprSht_HDR_Keydown(meINS_ROW, 0)	
	Call sprSht_DTL_Keydown(meINS_ROW, 0)	
	
end Sub

'저장버튼
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'삭제버튼
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'엑셀버튼
Sub imgExcel_HDR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_DTL_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
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
	SelectRtn
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
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
     	end if
	End with
	SelectRtn
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
			vntData = mobjPDCOGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
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
' 필드 체인지
'-----------------------------------------------------------------------------------------
Sub txtFROM_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	Dim strOLDYEARMON
	strdate = ""
	strFROM =""
	strFROM2 = ""

	With frmThis
		strdate=.txtFROM.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strFROM2 = strdate
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		
		.txtFROM.value = strFROM2
		DateClean strFROM
		txtTo_onchange
	End With

	gSetChange
End Sub

Sub txtTo_onchange
	SelectRtn
	gSetChange
End Sub

Sub txtOUTSNAME_onchange
	SelectRtn
	gSetChange
End Sub


Sub txtJOBNAME_onchange
	SelectRtn
	gSetChange
End Sub


Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		SelectRtn
		gSetChange
	end with
End Sub



'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
'클릭
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		elseif Row > 0 AND Col > 1 then
			if mstrGrid then SelectRtn_DTL Col, Row
		end if
	end with
End Sub


Sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		Else
			'strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"JOBNO",.sprSht_HDR.ActiveRow)
			'parent.jobMst_Call
			'mobjSCGLSpr.ActiveCell .sprSht_HDR, strCol, strRow	
		End If
	End With
End Sub


Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	Dim strRow, strCol
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub


Sub sprSht_HDR_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strINPUTJOBNO
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If frmThis.txtJOBNO.value <> ""Then
		strINPUTJOBNO = frmThis.txtJOBNO.value 
		Else
		strINPUTJOBNO = parent.document.forms("frmThis").txtJOBNO.value
		End If
		frmThis.sprSht_HDR.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_HDR, cint(KeyCode), cint(Shift), -1, 1)
		
		'여기서 부모창에서 받아온 JOBNO를 넣는다.
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"JOBNO",frmThis.sprSht_HDR.ActiveRow,strINPUTJOBNO 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"CONFIRMFLAG",frmThis.sprSht_HDR.ActiveRow, "미확정"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"CREDAY",frmThis.sprSht_HDR.ActiveRow, gNowDate
		
	End If
End Sub


Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
	
		frmThis.sprSht_DTL.MaxRows = 0
		frmThis.sprSht_DTL.MaxRows = 100
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht_DTL, 1,1
		frmThis.sprSht_DTL.focus()
	End If
End Sub


Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_HDR.MaxRows >0 Then
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") Then
				If .sprSht_HDR.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,strCol,vntData_row(j))
						End If
					Next
					
					.txtSELECTAMT.value = strSUM
				End If
				
			else
				.txtSELECTAMT.value = 0
			End If
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub


Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
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
		If mstrGrid Then
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"QTY") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
						FOR i = 0 TO intSelCnt -1
							If vntData_col(i) <> "" Then
								strColFlag = strColFlag + 1
								strCol = vntData_col(i)
							End If 
						Next
						
						If strColFlag <> 1 Then 
							.txtSELECTAMT.value = 0
							exit Sub
						End If
						
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
							End If
						Next
						
						.txtSELECTAMT.value = strSUM
					End If
					
				else
					.txtSELECTAMT.value = 0
				End If
			else
				.txtSELECTAMT.value = 0
			End If
			Call gFormatNumber(.txtSELECTAMT,0,True)
		ELSE
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"QTY") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE")  Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
						FOR i = 0 TO intSelCnt -1
							If vntData_col(i) <> "" Then
								strColFlag = strColFlag + 1
								strCol = vntData_col(i)
							End If 
						Next
						
						If strColFlag <> 1 Then 
							.txtSELECTAMT.value = 0
							exit Sub
						End If
						
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
							End If
						Next
						
						.txtSELECTAMT.value = strSUM
					End If
					
				else
					.txtSELECTAMT.value = 0
				End If
			else
				.txtSELECTAMT.value = 0
			End If
			Call gFormatNumber(.txtSELECTAMT,0,True)
		END IF
		
	End With
End Sub

Sub sprSht_HDR_Keyup(KeyCode, Shift)
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

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub


'쉬트 버튼클릭
Sub sprSht_HDR_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strCUSTCODE , strCUSTNAME

	with frmThis

		'외주처명
		IF Col = 8 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"BTN") then exit Sub
			strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow)
			strCUSTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow)
			
			vntInParams = array(trim(strCUSTCODE), trim(strCUSTNAME)) '<< 받아오는경우
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		
			if isArray(vntRet) then
				if strCUSTCODE = vntRet(0,0) and strCUSTNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit

				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow, trim(vntRet(0,0))   
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow, trim(vntRet(1,0))     
				
				mobjSCGLSpr.CellChanged .sprSht_HDR, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_HDR, Col+3, Row			
			end if
			gSetChange
     	END IF	
	End with
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
Dim vntData
   	Dim i, strCols , vntInParams
   	Dim strCode, strCodeName
   	Dim strCUSTCODE , strCUSTNAME
   	Dim intCnt
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
	
					
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"CUSTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjPDCOGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",trim(strCodeName))
				
				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",Row, trim(vntData(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",Row, trim(vntData(1,0))
						mobjSCGLSpr.CellChanged .sprSht_HDR, .sprSht_HDR.ActiveCol-1,frmThis.sprSht_HDR.ActiveRow
						
						.sprSht_HDR.focus
					Else
					
						strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow)
						strCUSTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow)
						
						vntInParams = array(trim(strCUSTCODE), trim(strCUSTNAME)) '<< 받아오는경우
						vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
					
						if isArray(vntRet) then
							if strCUSTCODE = vntRet(0,0) and strCUSTNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit

							mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow, trim(vntRet(0,0))   
							mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow, trim(vntRet(1,0))     
						end if
						
					End If
					.sprSht_HDR.focus 
					mobjSCGLSpr.ActiveCell .sprSht_HDR, Col+4, Row
   				End If
   			End If
		End If
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row
End Sub

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화
'-----------------------------------------------------------------------------------------	
Sub InitPage()
	'서버업무객체 생성	
	dim vntInParam
	dim intNo,i
	
	set mobjccPDDCCHARGEEXCOM	= gCreateRemoteObject("cPDCO.ccPDDCCHARGEEXCOM")
	set mobjPDCOGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	'탭 위치 설정 및 초기화
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "260px"
	'pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	'JOBNO 받아오는 부분==========================================================
	'vntInParam = window.dialogArguments
	'	intNo = ubound(vntInParam)
	'	'기본값 설정
	'	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'	
	'	for i = 0 to intNo
	''		select case i
	'			case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
	'			case 1 : frmThis.txtJOBNO.value = vntInParam(i)
	'		end select
	'	next
	'==============================================================================
	
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis

		gSetSheetColor mobjSCGLSpr, .sprSht_HDR
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 12, 0, 0, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_HDR,7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK|REVSEQ|JOBNO|OUTSCODE|OUTSNAME|CUSTCODE|CUSTNAME|BTN|AMT|CREDAY|CONFIRMFLAG|BIGO"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "선택|순번|JOBNO|가견적코드|가견적명|외주처코드|외주처명|금액|작성일|확정|비고"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1","  4|   4|   7|         9|      15|         9|      15|2|   9|    9|  9|  20"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_HDR,"..", "BTN"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_HDR, "CONFIRMFLAG", -1, -1, "확정" & vbTab & "미확정" , 10, 70, False, False
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "OUTSNAME | CUSTNAME | BIGO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "REVSEQ|JOBNO|OUTSCODE|CUSTCODE|AMT|CREDAY|"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "REVSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "OUTSNAME | CUSTNAME | BIGO",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "REVSEQ | JOBNO | OUTSCODE | CUSTCODE | CREDAY | CONFIRMFLAG",-1,-1,2,2,false
	
	    .sprSht_HDR.style.visibility  = "visible"
		
		
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 9, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ|ITEMNAME|STD|QTY|PRICE|AMT|BIGO"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "순번|가견적코드|가견적순번|항목|규격|수량|단가|금액|비고"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1","  5|         9|         9|  30|   12|   12|   12|   12|  30"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "QTY|PRICE|AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "ITEMNAME|STD|BIGO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "SEQ|OUTSCODE|REVSEQ"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ",-1,-1,2,2,false
	
	    .sprSht_DTL.style.visibility  = "visible"

	InitPageData	

	'날자관련 전체조회 사용자요청시 취소
	'msgbox parent.document.forms("frmThis").txtJOBNO.value 
	window.setTimeout "call time_data()",1000 
	SelectRtn
	End With
End Sub

Sub time_data
 with frmThis
	.txtJOBNO.value =  parent.document.forms("frmThis").txtJOBNO.value 
	.txtJOBNAME.value =  parent.document.forms("frmThis").txtJOBNAME.value 
	
 End with
End Sub
Sub EndPage()
	'set mobjccPDDCCHARGEEXCOM = Nothing
	'set mobjPDCOGET = Nothing
	gEndPage
End Sub

Sub InitPageData
	'초기 데이터 설정
	with frmThis
		
		'초기값 세팅
		.txtFROM.value = gNowDate
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)	
		
		
		'.sprSht_HDR.focus
	End with
End Sub

Sub DateClean(strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	with frmThis
	
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		.txtTO.value = date2
	End With
End Sub


'-----------------------------------------------------------------------------------------
' 조회
'-----------------------------------------------------------------------------------------
Sub SelectRtn
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		'세금계산서 완료조회
		vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNO.value),Trim(.txtJOBNAME.value),TRIM(.txtOUTSCODE.value),TRIM(.txtOUTSNAME.value))
		
		If not gDoErrorRtn ("SelectRtn_HDR") then
			If mlngRowCnt >0 Then
				mobjSCGLSpr.SetClipBinding .sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True
				mobjSCGLSpr.SetFlag  frmThis.sprSht_HDR,meCLS_FLAG
					
				gWriteText lblstatus1, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE	
		
			ELSE
				.sprSht_HDR.MaxRows = 0
				gWriteText lblstatus1, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE	
				
			End If
		End If	
		
		sprSht_HDR_Click 2, 1
		AMT_SUM
			
	END WITH
End Sub

Sub SelectRtn_DTL (Col , Row)
	Dim vntData
	Dim strOUTSCODE,strREVSEQ
   	Dim i, strCols
   	Dim intCnt
   	Dim strRow
	
	'On error resume next
	
	with frmThis
		'Sheet초기화
		.sprSht_DTL.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strOUTSCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"OUTSCODE",Row)
		strREVSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"REVSEQ",Row)
		
		IF strOUTSCODE <> "" THEN
			vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(strOUTSCODE),Trim(strREVSEQ))
		end if
	
		If not gDoErrorRtn ("SelectRtn_DTL") then
			IF mlngRowCnt > 0 THEN
				mobjSCGLSpr.SetClipBinding .sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True
				mobjSCGLSpr.SetFlag  frmThis.sprSht_DTL,meCLS_FLAG
			ELSE 
				.sprSht_DTL.MaxRows = 0
				
			END IF
		End If	
		gWriteText lblstatus2, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
	
		
	END WITH
End Sub


'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht_HDR.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht_HDR.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub



'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn()
	Dim intRtn ,Cnti , Cntj
  	dim vntData_BASICLIST_HDR , vntData_BASICLIST_DTL
	Dim strOLDSEQ , strRow
	Dim strSEQFlag
	Dim IntCnt ,i
	Dim lngchkCnt
	Dim strJOBNO
	
	with frmThis
 
  		
  		'초기화
  		Cnti=0
  		Cntj=0
  		lngchkCnt = 0
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		'변경된 데이타를 받는다
		vntData_BASICLIST_HDR = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"REVSEQ|JOBNO|OUTSCODE|OUTSNAME|CUSTCODE|BTN|AMT|CREDAY|CONFIRMFLAG|BIGO")
		vntData_BASICLIST_DTL = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"SEQ|OUTSCODE|REVSEQ|ITEMNAME|STD|QTY|PRICE|AMT|BIGO")
		
		' validation 시작
		'------------------------------------------------------------------------------------------------------------------------------------------
			
		'헤더 VALIDATION
		if  not IsArray(vntData_BASICLIST_HDR) then 
			IF not IsArray (vntData_BASICLIST_DTL )THEN
				gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				exit sub
			END IF
		End If
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "저장할 견적리스트가 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
		FOR  i = 1 to .sprSht_HDR.maxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"OUTSNAME",i) = "" THEN
				gErrorMsgBox "가견적명은 필수 입니다.","저장안내"	
				EXIT SUB
			END IF
		Next
		
		'디테일 VALIDATION
		if  not IsArray(vntData_BASICLIST_DTL) and .sprSht_DTL.MaxRows = 0 Then
			gErrorMsgBox "최소 하나의 항목을 입력해야 합니다. " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		
		'부모창에서 받은 STRBJONO를 가지고간다
		strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"JOBNO",.sprSht_HDR.ActiveRow) 
		
		
		'시트1의 SEQ유무로 NEW인지 UPDATE인지에 사용
		strRow = .sprSht_HDR.ActiveRow
		strOLDSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"REVSEQ",strRow)
		IF strOLDSEQ = "" THEN strOLDSEQ = 0
		
		
		'시트2에서 SEQ에따른 NEW ,  UPDATE
		if strOLDSEQ = 0 then
			strSEQFlag = "new"
		else
			strSEQFlag = "update"
		end if
		
		'OLDSEQ (HDR의 SEQ)를 꼭 가져아가야함. 헤더저장을 다 하고나서 DTL저장을 하도록 되어있기 때문이다.
		intRtn = mobjccPDDCCHARGEEXCOM.ProcessRtn(gstrConfigXml,vntData_BASICLIST_HDR,vntData_BASICLIST_DTL,strJOBNO,strOLDSEQ )
		
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
			
			if strSEQFlag = "new" then
				gErrorMsgBox " 자료가 신규저장 " & mePROC_DONE ,"저장안내" 
			else
				gErrorMsgBox " 자료가 수정저장 " & mePROC_DONE , "저장안내"
			end if
  		end if
		
		
		'저장후 쉬트클릭이 되게 하기위해서사용..
		mstrGrid = TRUE
  		SelectRtn
 	end with
End Sub



'자료삭제
Sub DeleteRtn ()

	Dim vntData
	Dim intCnt, intRtn, i
	Dim lngchkCnt
	
	
	with frmThis
		
		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "삭제할 내역이 없습니다.","삭제안내!"
			Exit Sub
		END IF
		
		For i = 1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | REVSEQ | OUTSCODE ")
		intRtn = mobjccPDDCCHARGEEXCOM.DeleteRtn(gstrConfigXml,vntData)
		
		
		IF not gDoErrorRtn ("DeleteRtn") then
			'선택된 자료를 끝에서 부터 삭제
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "외주견적이 삭제되었습니다.","삭제안내!"
			if .sprSht_HDR.MaxRows > 0 then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, 1,1
				mstrGrid = true
				SelectRtn_DTL 1,1
			else
				mstrGrid = FALSE
				SelectRtn
			end if
   		End IF
		
	End with
	err.clear
End Sub


		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="95%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<tr>
					<TD id="Td2" align="left" width="100%" height="20" runat="server">
						<TABLE id="tblTitle1" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="85" background="../../../images/back_p.gIF"
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
											<td class="TITLE">외주견적 현황&nbsp;</td>
										</tr>
									</table>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</tr>
				<TR>
					<TD vAlign="top">
						<TABLE class="SEARCHDATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" align="left"
							border="0">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" width="60"><FONT face="굴림">견적일</FONT></TD>
								<TD class="SEARCHDATA" width="224"><INPUT class="INPUT" id="txtFROM" title="기간검색(FROM)" style="WIDTH: 88px; HEIGHT: 22px"
										accessKey="DATE" type="text" maxLength="10" size="9" name="txtFROM"> <IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"										border="0" name="imgCalEndarFROM1">~<INPUT class="INPUT" id="txtTO" title="기간검색(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										 align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)"
									width="60">외주처</TD>
								<TD class="SEARCHDATA" width="263"><INPUT class="INPUT_L" id="txtOUTSNAME" title="외주처" style="WIDTH: 170px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="37" name="txtOUTSNAME"> <IMG id="imgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="imgOUTSCODE"> <INPUT class="INPUT" id="txtOUTSCODE" title="외주처" style="WIDTH: 65px; HEIGHT: 22px" accessKey=",M"
										type="text" maxLength="6" size="9" name="txtOUTSCODE"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
									width="60">JOB명</TD>
								<TD class="SEARCHDATA" width="263"><INPUT class="INPUT_L" id="txtJOBNAME" title="JOBNO" style="WIDTH: 170px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="23" name="txtJOBNAME"> <IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"  align="absMiddle"
										border="0" name="ImgJOBNO"> <INPUT class="INPUT" id="txtJOBNO" title="JOBNO" style="WIDTH: 65px; HEIGHT: 22px" type="text"
										maxLength="7" align="left" size="5" name="txtJOBNO"></TD>
								<TD class="SEARCHDATA2" align="right" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
										align="right" border="0" name="imgQuery"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" id="spacebar" style="WIDTH: 100%; HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD>
						<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD class="TITLE" style="WIDTH: 100%; HEIGHT: 8px" vAlign="absmiddle"></TD>
							</TR>
							<TR>
								<TD class="TITLE" width="210" vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_normal.gif'
											align='absMiddle' border='0' name='btn_normal'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')">
										<IMG id='btn_multi' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_multi.gif'
											align='absMiddle' border='0' name='btn_multi'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')">
										<IMG id='btn_hide' style='CURSOR: hand' alt='자료를 검색합니다.' src='../../../images/btn_hide.gif'
											align='absMiddle' border='0' name='btn_hide'></span>
								</TD>
							</TR>
						</table>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
												<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgNEW" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" height="20" alt="자료를 추가합니다."
													src="../../../images/imgNew.gIF" border="0" name="imgNEW"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgExcel_HDR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel_HDR"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--테이블이 무너지는것을 막아준다-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR id="tblBody1">
					<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_HDR" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="42545">
								<PARAM NAME="_ExtentY" VALUE="3334">
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
					<TD class="BOTTOMSPLIT" id="lblStatus1" style="WIDTH: 1040px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="64" background="../../../images/back_p.gIF"
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
												<td class="TITLE">제작리스트&nbsp;</td>
											</tr>
										</table>
									</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgExcel_DTL" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel_DTL"></TD>
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
					<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="left">
						<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="42545">
								<PARAM NAME="_ExtentY" VALUE="6826">
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
					<TD class="BOTTOMSPLIT" id="lblStatus2" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
