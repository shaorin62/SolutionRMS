<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMREALMEDCODEMST.aspx.vb" Inherits="MD.MDCMREALMEDCODEMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>매체코드 등록</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/공중파 매체코드 등록 화면(MDCMREALMEDCODEMST)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/08 By 황덕수
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
'전역변수 설정
Dim mobjMDCMREALMEDCODEMST
Dim mobjMDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mstrGUBUN

Dim intSelectRows 'lock을 걸기위해 조회해온 row수를 가지고 있는다.

mstrGUBUN = "KOBACO"

CONST meTAB = 9
intSelectRows = 0
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'---------------------------------
'-----시트 버튼 클릭 이벤트------
'---------------------------------

'신규 - 신규시 광고주팝업
Sub imgNew_onclick
	With frmThis
		IF mstrGUBUN = "KOBACO" THEN
			CALL sprSht_Keydown(meINS_ROW, 0)
		ELSE
			CALL sprSht_SBS_Keydown(meINS_ROW, 0)
		END IF 
		
	End With
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
		SelectRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		IF mstrGUBUN = "KOBACO" THEN
			mobjSCGLSpr.ExportExcelFile .sprSht
		ELSE
			mobjSCGLSpr.ExportExcelFile .sprSht_SBS
		END IF 
	End With
	gFlowWait meWAIT_OFF
End Sub


Sub imgSave_onclick ()
	with frmThis
		gFlowWait meWAIT_ON
			IF mstrGUBUN = "KOBACO" THEN
				ProcessRtn(.sprSht)
			ELSE
				ProcessRtn(.sprSht_SBS)
			END IF
			
		gFlowWait meWAIT_OFF
	end with
End Sub


'텝처리 (코바코)
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
		
	pnlTab_KOBACO.style.visibility = "visible" 
	pnlTab_SBS.style.visibility = "hidden" 	
	
	document.getElementById("strMsgBox").innerHTML = "코바코 광고주 코드"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "KOBACO"
	gridLayOut
	
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'텝처리 (SBS)
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TABON
	
	pnlTab_KOBACO.style.visibility = "hidden" 
	pnlTab_SBS.style.visibility = "visible" 
	
	document.getElementById("strMsgBox").innerHTML = "SBS 광고주 코드"
		
	gFlowWait meWAIT_ON
	mstrGUBUN = "SBS"
	gridLayOut
	
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'-----------------------------------------------------------------------------------------
' 매체사코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE_onclick
	Call CUSTCODE_POP()
End Sub
'실제 데이터List 가져오기
Sub CUSTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(TRIM(.txtCUSTCODE.value), TRIM(.txtCUSTNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP_ALL.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCUSTCODE.value = vntRet(0,0)  ' Code값 저장
			.txtCUSTNAME.value = vntRet(1,0)  ' 코드명 표시
			.cmbMEMO.focus()					' 포커스 이동
			gSetChangeFlag .txtCUSTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	SelectRtn(mstrGUBUN)
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetMEDNO(gstrConfigXml,mlngRowCnt,mlngColCnt,TRIM(.txtCUSTCODE.value),TRIM(.txtCUSTNAME.value))
			if not gDoErrorRtn ("GetMEDNO") then
				If mlngRowCnt = 1 Then
					.txtCUSTCODE.value = vntData(0,0)
					.txtCUSTNAME.value = vntData(1,0)
					.cmbMEMO.focus()
				Else
					Call CUSTCODE_POP()
				End If
   			end if
   		end with
   		SelectRtn(mstrGUBUN)
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' onchange이벤트
'-----------------------------------------------------------------------------------------

Sub cmbMEMO_onchange
	gSetChange
	SelectRtn(strGUBUN)
End Sub

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
'---------------------------------
' 스프레드 쉬트 변경시 체크 
'--------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

sub sprSht_SBS_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser .sprSht_SBS, ""
		end if
	end with
end sub


Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 9 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				'strBCODE = .txtBCODE.value
				strCUSTCODE = TRIM(vntRet(0,0))
				intRtn = mobjMDCMREALMEDCODEMST.SelectRtn_VALIDATION(gstrConfigXml,mlngRowCnt,mlngColCnt,strCUSTCODE)
				If not gDoErrorRtn ("SelectRtn_VALIDATION") then
					If mlngRowCnt > 0 Then
						gErrorMsgBox "현재사용중인코드입니다.","입력안내"
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, "미사용"		
					Else
					mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, "사용"			
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					End If
				End If
			End IF
			.cmbMEMO.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

'SBS탭 팝업
Sub sprSht_SBS_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 9 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_SBS,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_SBS,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_SBS,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				'strBCODE = .txtBCODE.value
				strCUSTCODE = TRIM(vntRet(0,0))
				intRtn = mobjMDCMREALMEDCODEMST.SelectRtn_VALIDATION(gstrConfigXml,mlngRowCnt,mlngColCnt,strCUSTCODE)
				If not gDoErrorRtn ("SelectRtn_VALIDATION") then
					If mlngRowCnt > 0 Then
						gErrorMsgBox "현재사용중인코드입니다.","입력안내"
						mobjSCGLSpr.SetTextBinding .sprSht_SBS,"CUSTCODE",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht_SBS,"CUSTNAME",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht_SBS,"MEMO",Row, "미사용"		
					Else
					mobjSCGLSpr.SetTextBinding .sprSht_SBS,"CUSTCODE",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_SBS,"CUSTNAME",Row, vntRet(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht_SBS,"MEMO",Row, "사용"			
					mobjSCGLSpr.CellChanged .sprSht_SBS, Col,Row
					End If
				End If
			End IF
			.cmbMEMO.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_SBS.Focus
			mobjSCGLSpr.ActiveCell .sprSht_SBS, Col+2, Row
		end if
	End with
End Sub
'-----------------------------------
'----------시트 키다운--------------
'-----------------------------------

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		
		If KeyCode = meINS_ROW Then
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
			
			strRow = .sprSht.ActiveRow
			'조회된건이 없으면 신규이므로 lock을푼다.
			if intSelectRows = 0 Then
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE | CUSTCODE | BTN | CUSTNAME | MEMO",1,strRow,false
			else 
			'조회된건이 있으면 조회시 받아온 intSelectRows만큼 lock을 걸어준다.
				'BCODE만 빼고 LOCK
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | CUSTCODE | BTN | CUSTNAME | MEMO",1,strRow,false
				strRow = intSelectRows
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE|OFFICECODENAME|BCODE",1,strRow,false
			end if
			
			mobjSCGLSpr.SetTextBinding .sprSht,"MEMO", .sprSht.ActiveRow, "미사용"
			mobjSCGLSpr.ActiveCell .sprSht, 2, .sprSht.ActiveRow
		End If
	end with
End Sub

Sub sprSht_SBS_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		
		If KeyCode = meINS_ROW Then
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_SBS, cint(KeyCode), cint(Shift), -1, 1)
			
			strRow = .sprSht_SBS.ActiveRow
			'조회된건이 없으면 신규이므로 lock을푼다.
			if intSelectRows = 0 Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_SBS,false,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE | CUSTCODE | BTN | CUSTNAME | MEMO",1,strRow,false
			else 
			'조회된건이 있으면 조회시 받아온 intSelectRows만큼 lock을 걸어준다.
				'BCODE만 빼고 LOCK
				mobjSCGLSpr.SetCellsLock2 .sprSht_SBS,false,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | CUSTCODE | BTN | CUSTNAME | MEMO",1,strRow,false
				strRow = intSelectRows
				mobjSCGLSpr.SetCellsLock2 .sprSht_SBS,true,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE|OFFICECODENAME|BCODE",1,strRow,false
			end if
			
			mobjSCGLSpr.SetTextBinding .sprSht_SBS,"MEMO", .sprSht_SBS.ActiveRow, "미사용"
			mobjSCGLSpr.ActiveCell .sprSht_SBS, 2, .sprSht_SBS.ActiveRow
		End If
	end with
End Sub


'-----------------------------------------------
'----------------시트 체인지--------------------
'-----------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	With frmThis
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BRODCODENAME") Then
			Call sprSht_SelectCode ("BRODCODENAME", "BRODCODE",  mobjSCGLSpr.GetTextBinding(.sprSht,"BRODCODENAME",Row) ,Col, Row)
		end if
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDIUMCODENAME") Then
			Call sprSht_SelectCode ("MEDIUMCODENAME", "MEDIUMCODE",  mobjSCGLSpr.GetTextBinding(.sprSht,"MEDIUMCODENAME",Row) ,Col, Row)
		end if
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OFFICECODENAME") Then
			Call sprSht_SelectCode ("OFFICECODENAME", "OFFICECODE",  mobjSCGLSpr.GetTextBinding(.sprSht,"OFFICECODENAME",Row) ,Col, Row)
		end if
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub sprSht_SBS_Change(ByVal Col, ByVal Row)
	With frmThis
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_SBS,"BRODCODENAME") Then
			Call sprSht_SBS_SelectCode ("BRODCODENAME", "BRODCODE",  mobjSCGLSpr.GetTextBinding(.sprSht_SBS,"BRODCODENAME",Row) ,Col, Row)
		end if
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_SBS,"MEDIUMCODENAME") Then
			Call sprSht_SBS_SelectCode ("MEDIUMCODENAME", "MEDIUMCODE",  mobjSCGLSpr.GetTextBinding(.sprSht_SBS,"MEDIUMCODENAME",Row) ,Col, Row)
		end if
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_SBS,"OFFICECODENAME") Then
			Call sprSht_SBS_SelectCode ("OFFICECODENAME", "OFFICECODE",  mobjSCGLSpr.GetTextBinding(.sprSht_SBS,"OFFICECODENAME",Row) ,Col, Row)
		end if
	End With
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht_SBS, Col, Row
End Sub

'---------------------------------------------------------------------------------------------------------------------------------------------
'[[ColName '비교할 컬럼의 헤더명]] '[[ByColCode '비교할 컬럼의 헤더코드명]]  [['ColNameData '비교할 컬럼의 내용]  [['Row '비교할 컬럼의 행]]
'---------------------------------------------------------------------------------------------------------------------------------------------
Sub sprSht_SelectCode( ByVal ColName, ByVal ByColCode, ByVal ColNameData, ByVal Col,   ByVal Row )
	Dim i
	Dim intCnt
	
	intCnt = 0
	With frmThis
		'조회된 내역중에서 찾는다.
		For i =1 to intSelectRows
			if ColNameData = mobjSCGLSpr.GetTextBinding(.sprSht,ColName,i) then
				mobjSCGLSpr.SetTextBinding .sprSht,ByColCode,Row, mobjSCGLSpr.GetTextBinding(.sprSht,ByColCode,i)
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				intCnt = intCnt + 1
				exit sub
			end if
		next
		
		'못찾았다면 직접입력 요망.
		if intCnt = 0 Then 
			gErrorMsgBox "일치하는 코드가 없습니다. 직접 입력하십시오.","입력안내"
			mobjSCGLSpr.ActiveCell .sprSht, Col-1, Row
			exit sub
		End if
	end With
End sub

Sub sprSht_SBS_SelectCode( ByVal ColName, ByVal ByColCode, ByVal ColNameData, ByVal Col,   ByVal Row )
	Dim i
	Dim intCnt
	
	intCnt = 0
	With frmThis
		'조회된 내역중에서 찾는다.
		For i =1 to intSelectRows
			if ColNameData = mobjSCGLSpr.GetTextBinding(.sprSht_SBS,ColName,i) then
				mobjSCGLSpr.SetTextBinding .sprSht_SBS,ByColCode,Row, mobjSCGLSpr.GetTextBinding(.sprSht_SBS,ByColCode,i)
				mobjSCGLSpr.ActiveCell .sprSht_SBS, Col+2, Row
				intCnt = intCnt + 1
				exit sub
			end if
		next
		
		'못찾았다면 직접입력 요망.
		if intCnt = 0 Then 
			gErrorMsgBox "일치하는 코드가 없습니다. 직접 입력하십시오.","입력안내"
			mobjSCGLSpr.ActiveCell .sprSht_SBS, Col-1, Row
			exit sub
		End if
	end With
End sub

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjMDCMREALMEDCODEMST = gCreateRemoteObject("cMDET.ccMDETREALMEDCODEMST")
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")

	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
    gSetSheetDefaultColor
    
	InitPageData	
	btnTab1_onclick
End Sub

sub gridLayOut
	Dim strComboList
	
	with frmThis
		if mstrGUBUN = "KOBACO" THEN
			'**************************************************
			'***코바코 코드 입력 시트
			'**************************************************	
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 0
			mobjSCGLSpr.AddCellSpan  .sprSht, 8, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.SpreadDataField .sprSht,    "BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE | CUSTCODE | BTN | CUSTNAME | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,		    "코드|방송국명|코드|매체|코드|본지사명|코바코코드|매체코드|매체명|사용구분"
			mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|      12|   4|  12|   4|      13|        10|       8|2|  30|       10"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTCODE", -1, -1, 6
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTNAME | MEMO", -1, -1, 255
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
			mobjSCGLSpr.SetCellAlign2 .sprSht, "BRODCODENAME | MEDIUMCODENAME | OFFICECODENAME",-1,-1,0,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht, "BCODE | CUSTCODE | BRODCODE | MEDIUMCODE | OFFICECODE",-1,-1,2,2,false
			mobjSCGLSpr.SetCellTypeComboBox .sprSht,11,11,-1,-1,"사용" & vbTab & "미사용",2,140,false
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE"
			'mobjSCGLSpr.ColHidden .sprSht, "", true
		
		else
			
			'**************************************************
			'**SBS 미디어 랩사 코드 입력 시트
			'**************************************************	
			strComboList =  "사용" & vbTab & "미사용"
			
			gSetSheetColor mobjSCGLSpr, .sprSht_SBS
			mobjSCGLSpr.SpreadLayout .sprSht_SBS, 11, 0, 0
			mobjSCGLSpr.AddCellSpan  .sprSht_SBS, 8, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.SpreadDataField .sprSht_SBS,    "BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE | CUSTCODE | BTN | CUSTNAME | MEMO"
			mobjSCGLSpr.SetHeader .sprSht_SBS,		    "코드|방송국명|코드|매체|코드|본지사명|SBS코드|매체코드|매체명|사용구분"
			mobjSCGLSpr.SetColWidth .sprSht_SBS, "-1",  "   4|      12|   4|  12|   4|      13|     10|       8|2|  30|      10"
			mobjSCGLSpr.SetRowHeight .sprSht_SBS, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht_SBS, "-1", "13"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SBS, "CUSTCODE", -1, -1, 6
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SBS, "CUSTNAME | MEMO", -1, -1, 255
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht_SBS,"..", "BTN"
			mobjSCGLSpr.SetCellAlign2 .sprSht_SBS, "BRODCODENAME | MEDIUMCODENAME | OFFICECODENAME",-1,-1,0,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht_SBS, "BCODE | CUSTCODE | BRODCODE | MEDIUMCODE | OFFICECODE",-1,-1,2,2,false
			mobjSCGLSpr.SetCellTypeComboBox .sprSht_SBS,11,11,-1,-1,strComboList
			mobjSCGLSpr.SetCellsLock2 .sprSht_SBS,true,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE"
			'mobjSCGLSpr.ColHidden .sprSht_SBS, "", true
		
		end IF
	end with
end sub

Sub InitPageData
	gClearAllObject frmThis
	with frmThis
		gridLayOut
		.sprSht.MaxRows = 0	
		.sprSht_SBS.MaxRows = 0	
		
		document.getElementById("strMsgBox").innerHTML = "코바코 광고주 코드"	
	END WITH
End Sub

Sub EndPage()
	set mobjMDCMREALMEDCODEMST = Nothing
	set mobjMDCMGET = Nothing
	gEndPage	
End Sub

'------------------------------------
'-----------------조회 --------------
'------------------------------------
Sub SelectRtn (mstrGUBUN)
   	Dim vntData
   	Dim i, strCols
	Dim strSEQ
	Dim strBCODE
	Dim strCUSTCODE
	Dim strMEMO
	Dim strOFFICECODENAME
	'On error resume next
	With frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
'		gridLayOut
		.sprSht.MaxRows = 0
		.sprSht_SBS.MaxRows = 0
		
		strBCODE		= .txtBCODE.value
		strCUSTCODE		= .txtCUSTCODE.value
		strMEMO			= .cmbMEMO.value
		strOFFICECODENAME  = .txtOFFICECODENAME.value
		
		vntData = mobjMDCMREALMEDCODEMST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strBCODE,strCUSTCODE,strMEMO, strOFFICECODENAME, mstrGUBUN)
		intSelectRows = mlngRowCnt
		
		If not gDoErrorRtn ("SelectRtn") then
		
			IF mstrGUBUN = "KOBACO" THEN
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			ELSE
				mobjSCGLSpr.SetClipBinding .sprsht_SBS, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			END IF 
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		End if
   		
   	End With
End Sub

Sub ProcessRtn(sprSht)
	Dim intRtn
   	Dim vntData , vntCheckData
   	Dim strBCODE
   	Dim strBRODCODE
   	Dim strMEDIUMCODE
   	Dim strOFFICECODE
   	
	with frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'조회된 row 다음줄부터 신규데이터만 validation
		for	i=intSelectRows+1 to sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(sprSht,"BRODCODE",i) = "" OR mobjSCGLSpr.GetTextBinding(sprSht,"MEDIUMCODE",i) = "" _
																	 OR mobjSCGLSpr.GetTextBinding(sprSht,"OFFICECODE",i) = "" then 
				
				gErrorMsgBox "신규입력시 [방송국/매체/본지사명]은 필수내역 입니다.","저장안내"
				exit sub
			else
				'다입력이 되었다면 bocde를 만들어서 이미 있는 bcode인지 확인한다.
				strBRODCODE   = mobjSCGLSpr.GetTextBinding(sprSht,"BRODCODE",i)
				strMEDIUMCODE = mobjSCGLSpr.GetTextBinding(sprSht,"MEDIUMCODE",i)
				strOFFICECODE = mobjSCGLSpr.GetTextBinding(sprSht,"OFFICECODE",i)
				strBCODE = strBRODCODE + strMEDIUMCODE + strOFFICECODE
				
				vntCheckData = mobjMDCMREALMEDCODEMST.CHECK_BCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strBCODE)
				if mlngRowCnt > 0 Then
					gErrorMsgBox "중복된 내역은 저장할수 없습니다.","저장안내"
					exit sub
				end if
			end if
		next
		
		vntData = mobjSCGLSpr.GetDataRows(sprSht,"BRODCODE | BRODCODENAME | MEDIUMCODE | MEDIUMCODENAME | OFFICECODE | OFFICECODENAME | BCODE | CUSTCODE | MEMO")
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If
		
		'처리 업무객체 호출
		intRtn = mobjMDCMREALMEDCODEMST.ProcessRtn(gstrConfigXml,vntData,mstrGUBUN)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  sprSht,meCLS_FLAG
			gErrorMsgBox intRtn & " 건 이 저장되었습니다.","저장안내"
			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="98%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="600" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
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
											<td class="TITLE">코바코 매체 등록</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 101; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 95%" cellSpacing="0" cellPadding="0" width="1040"
							border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 120px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtBCODE,'')"
												width="150"><span id="strMsgBox"></span>
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 70px"><INPUT class="INPUT_L" id="txtBCODE" style="WIDTH: 96px; HEIGHT: 22px" maxLength="8" size="10"
													name="txtBCODE" dataFld="BCODE" accessKey="NUM" dataSrc="#xmlBind"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 60px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTCODE,txtCUSTNAME)"
												width="86">매체명
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 250px"><INPUT class="INPUT_L" id="txtCUSTNAME" style="WIDTH: 150px; HEIGHT: 22px" maxLength="255"
													size="46" name="txtCUSTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCUSTCODE" style="WIDTH: 68px; HEIGHT: 22px" maxLength="6"
													size="6" name="txtCUSTCODE"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtOFFICECODENAME,'')"
												width="150">본지사 명
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 70px"><INPUT class="INPUT_L" id="txtOFFICECODENAME" style="WIDTH: 96px; HEIGHT: 22px" maxLength="50"
													size="10" name="txtOFFICECODENAME" dataFld="BCODE" dataSrc="#xmlBind"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px;CURSOR: hand" onclick="vbscript:Call gCleanField(cmbMEMO,'')"
												width="86">사용구분
											</TD>
											<TD class="SEARCHDATA"><SELECT id="cmbMEMO" title="사용구분" style="WIDTH: 104px" name="cmbMEMO">
													<OPTION value="A" selected>전체</OPTION>
													<OPTION value="사용">사용</OPTION>
													<OPTION value="미사용">미사용</OPTION>
												</SELECT>
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 10px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
																type="button" value="KOBACO" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
																type="button" value="SBS" name="btnTab2">
														</TD>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="신규자료를 작성합니다."
																src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<!--테이블이 무너지는것을 막아준다-->
									<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
										<TR>
											<TD align="left" width="100%" height="1"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
							<!--내용 및 그리드-->
							<TR vAlign="top" align="left">
								<!--내용-->
								<TD id="tblSheet" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab_KOBACO" style="POSITION: absolute; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31855">
											<PARAM NAME="_ExtentY" VALUE="15266">
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
									<DIV id="pnlTab_SBS" style="POSITION:relative; WIDTH:100%; HEIGHT:100%; VISIBILITY:hidden; LEFT:7px"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_SBS" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="15266">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR> 
			<!--List End-->
			<!--Bottom Split Start-->
			<!--Bottom Split End--> </TABLE> 
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
