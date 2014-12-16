<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMIFCMALLVOCH.aspx.vb" Inherits="MD.MDCMIFCMALLVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>IFC MALL 전표생성</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : IFCMALL 전표처리 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : IFMALL 전표처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/01/22 By kty
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
Dim mlngRowCnt,mlngColCnt
Dim mobjMDCMIFCMALLVOCH
Dim mobjMDCOVOCH
Dim mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim mstrPROCESS
Dim mstrSTAY

mstrSTAY = TRUE

mstrGUBUN = "S"
mstrPROCESS = ""
mstrCheck = True

CONST meTAB = 9
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'강제삭제 버튼 숨기기
Sub Set_delete(byVal strmode)
	With frmThis
		IF .rdT.checked = TRUE then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'조회버튼
Sub imgQuery_onclick
	If frmThis.txtYEARMON.value = "" Then
		gErrorMsgBox "조회년월을 입력하시오","조회안내"
		exit Sub
	End If

	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'매출
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TABON
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	frmThis.btnTab4.style.backgrounDimage = meURL_TAB
	frmThis.btnTab5.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "visible" 
	pnlTab_out.style.visibility = "hidden" 
	pnlTab_EOM.style.visibility = "hidden"
	pnlTab_EOMSUSU.style.visibility = "hidden"
	pnlTab_OUTSUSU.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "hidden" 
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "S"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF

	mobjSCGLCtl.DoEventQueue
End Sub

'매입
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TABON
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	frmThis.btnTab4.style.backgrounDimage = meURL_TAB
	frmThis.btnTab5.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "visible" 
	pnlTab_EOM.style.visibility = "hidden" 
	pnlTab_EOMSUSU.style.visibility = "hidden"
	pnlTab_OUTSUSU.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "visible" 
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "YO"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'대행매입
Sub btnTab3_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	frmThis.btnTab3.style.backgrounDimage = meURL_TABON
	frmThis.btnTab4.style.backgrounDimage = meURL_TAB
	frmThis.btnTab5.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "hidden"
	pnlTab_EOM.style.visibility = "visible"
	pnlTab_EOMSUSU.style.visibility = "hidden"
	pnlTab_OUTSUSU.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "visible" 
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "YO2"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF

	mobjSCGLCtl.DoEventQueue
End Sub

'대행매출
Sub btnTab4_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	frmThis.btnTab4.style.backgrounDimage = meURL_TABON
	frmThis.btnTab5.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "hidden"
	pnlTab_EOM.style.visibility = "hidden"
	pnlTab_EOMSUSU.style.visibility = "visible"
	pnlTab_OUTSUSU.style.visibility = "hidden" 
	
	pnlFLAG.style.visibility = "hidden" 
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "YO3"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF

	mobjSCGLCtl.DoEventQueue
End Sub

'매입수수료
Sub btnTab5_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	frmThis.btnTab3.style.backgrounDimage = meURL_TAB
	frmThis.btnTab4.style.backgrounDimage = meURL_TAB
	frmThis.btnTab5.style.backgrounDimage = meURL_TABON
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_out.style.visibility = "hidden" 
	pnlTab_EOM.style.visibility = "hidden" 
	pnlTab_EOMSUSU.style.visibility = "hidden"
	pnlTab_OUTSUSU.style.visibility = "visible" 
	
	pnlFLAG.style.visibility = "visible" 
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "YS"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF

	mobjSCGLCtl.DoEventQueue
End Sub

'엑셀버튼 클릭
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis

		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		
		if mstrGUBUN = "S"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_SUSU
		elseif mstrGUBUN = "YO"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_OUT
		elseif mstrGUBUN = "YO2"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_EOM
		elseif mstrGUBUN = "YO3"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_EOMSUSU
		elseif mstrGUBUN = "YS"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_OUTSUSU
		end if
		
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'전표생성 클릭
Sub ImgvochCre_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'전표삭제 클릭
Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'오류전표삭제클릭
Sub ImgErrVochDel_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'전표강제 삭제 클릭
Sub imgVochDelco_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'--적용버튼 클릭
Sub ImgSUMMApp_onclick()
	Dim intRtn

	with frmThis
		if .cmbSETTING.value = "" then
			gErrorMsgBox "적용하실 컬럼 명이 없습니다. ","적용오류"
			exit sub
		end if 
		
		'취급액
		if mstrGUBUN = "S"  then 
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_SUSU)
			settingRowChange (.sprSht_SUSUDTL)
		'일반
		elseif mstrGUBUN = "YO"  then  
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUT)
		elseif mstrGUBUN = "YO2"  then  
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_EOM)
		elseif mstrGUBUN = "YO3"  then  
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_EOMSUSU)
		elseif mstrGUBUN = "YS"  then  
			intRtn = gYesNoMsgbox("체크하신 데이터의 내용이 변경됩니다 적용하시겠습니까? ","처리안내!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUTSUSU)
		end if
	End With
End Sub

sub settingRowChange(sprsht)
	Dim strSETTINGDATA
	Dim intCnt 
	Dim i ,j

	with frmThis
		intCnt = 0
		
		for j = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) <> "DTL" Then
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",j) = "1" Then
					intCnt = intCnt + 1
				End if
			END IF
		next
		
		if right(sprsht.ID,3) <> "DTL" Then
			if intCnt = 0 Then
				gErrorMsgBox "체크된 데이터가 없습니다. 적용하실 데이터를 체크하세요. ","적용오류"
				EXIT SUB
			End if
		End if
		
		strSETTINGDATA = ""
		strSETTINGDATA = .txtSUMM.value
		
		for i = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) = "DTL" Then
				mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
			ELSE 
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",i) = "1" Then
					mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
				End if
			End if 
		next 
	End with
end sub

'-----------------------------------------------------------------------------------------
' 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주팝업버튼
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")

			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))

				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'완료체크
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'미완료체크
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'에러체크
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUMM_onchange
	Dim blnByteCHk
	Dim intRtn
	blnByteCHk =  checkBytes(frmThis.txtSUMM.value)
	
	If blnByteCHk  > 23 Then
		intRtn = gYesNoMsgbox("적요의 크기는 23Byte 를 넘을수 없습니다. 초기화 하시겠습니까?","처리안내!")
		
		IF intRtn <> vbYes then exit Sub		
		frmThis.txtSUMM.value = ""
	End If
End Sub

function checkBytes(expression)
	Dim VLength
	Dim temp
	Dim EscTemp
	Dim i

	VLength=0
	temp = expression
	
	if temp <> "" then
		for i=1 to len(temp) 
			if mid(temp,i,1) <> escape(mid(temp,i,1))  then
				EscTemp=escape(mid(temp,i,1))
				if (len(EscTemp)>=6) then
					VLength = VLength +2
				else
					VLength = VLength +1
				end if
			else
				VLength = VLength +1
			end if
		Next
	end if
	checkBytes = VLength
End Function

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_SUSU_Change(ByVal Col, ByVal Row)
	Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
	Dim strFROMDATE, strTODATE, strSUMMTEXT
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT") Then
			DeleteRtn_SUSUDTL (Row)
			SelectRtn_SUSUDTL Col,Row

			If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"TODATE",Row,Row,false
			End If
		End If
		
		if .sprSht_SUSUDTL.MaxRows > 0 then
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)
			strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMM",Row)
			strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SEMU",Row)
			strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DEMANDDAY",Row)
			strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DUEDATE",Row)
			strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DOCUMENTDATE",Row)
			strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row)
			strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"FROMDATE",Row)
			strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TODATE",Row)
			strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMMTEXT",Row)

			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
					mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) = strTAXNO Then

					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMM",intCnt, strSUMM
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SEMU",intCnt, strSEMU
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DUEDATE",intCnt, strDUEDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"FROMDATE",intCnt, strFROMDATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"TODATE",intCnt, strTODATE
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
				End If
			Next
		end if
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, Col, Row
End Sub

Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
	
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PAYCODE") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404103"
			end if 
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUT,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUT)
		end if 
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,false
		End If
	End With

	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

Sub sprSht_EOM_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"paycode") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DEBTOR",Row, "404103"
			end if 
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_EOM,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_EOM)
		end if
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,True,"TODATE",Row,Row,false
		End If

		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"DOCUMENTDATE") and Row = 1 then
			for i=1 to .sprSht_EOM.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DOCUMENTDATE",i, mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE", Row) 
			Next
		end if

		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"ACCOUNT") and Row = 1 then
			for i=1 to .sprSht_EOM.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"ACCOUNT",i, mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"DEBTOR") and Row = 1 then
			for i=1 to .sprSht_EOM.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DEBTOR",i, mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR", Row) 
			Next
		end if

		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"DEMANDDAY") and Row = 1 then
			for i=1 to .sprSht_EOM.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_EOM,"DEMANDDAY",i, mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY", Row) 
			Next
		end if
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_EOM, Col, Row
End Sub

Sub sprSht_EOMSUSU_Change(ByVal Col, ByVal Row)
	Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
	Dim strFROMDATE, strTODATE, strSUMMTEXT
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSU,"PREPAYMENT") Then
			DeleteRtn_EOMSUSUDTL (Row)
			SelectRtn_EOMSUSUDTL Col,Row

			If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"PREPAYMENT",Row) = "Y" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,false,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,false,"TODATE",Row,Row,false
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,True,"FROMDATE",Row,Row,false
				mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,True,"TODATE",Row,Row,false
			End If
		End If
		
		if .sprSht_EOMSUSUDTL.MaxRows > 0 then
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXNO",Row)
			strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SUMM",Row)
			strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SEMU",Row)
			strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DEMANDDAY",Row)
			strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DUEDATE",Row)
			strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DOCUMENTDATE",Row)
			strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"PREPAYMENT",Row)
			strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"FROMDATE",Row)
			strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TODATE",Row)
			strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SUMMTEXT",Row)

			For intCnt = 1 To .sprSht_EOMSUSUDTL.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
					mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) = strTAXNO Then

					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SUMM",intCnt, strSUMM
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SEMU",intCnt, strSEMU
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DUEDATE",intCnt, strDUEDATE
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"FROMDATE",intCnt, strFROMDATE
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"TODATE",intCnt, strTODATE
					mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
				End If
			Next
		end if
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_EOMSUSU, Col, Row
End Sub


Sub sprSht_OUTSUSU_Change(ByVal Col, ByVal Row)
	Dim strCODE
	with frmThis
	
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"paycode") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",Row, "404103"
			end if 
			
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht_OUTSUSU,"VENDOR",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht_OUTSUSU)
		end if 
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"TODATE",Row,Row,false
		End If
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"DOCUMENTDATE") and Row = 1 then
			for i=1 to .sprSht_OUTSUSU.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DOCUMENTDATE",i, mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"ACCOUNT") and Row = 1 then
			for i=1 to .sprSht_OUTSUSU.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"ACCOUNT",i, mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT", Row) 
			Next
		end if
		
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"DEBTOR") and Row = 1 then
			for i=1 to .sprSht_OUTSUSU.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEBTOR",i, mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR", Row) 
			Next
		end if

		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"DEMANDDAY") and Row = 1 then
			for i=1 to .sprSht_OUTSUSU.MaxRows
				mobjSCGLSpr.SetTextBinding .sprSht_OUTSUSU,"DEMANDDAY",i, mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY", Row) 
			Next
		end if
	End With

	mobjSCGLSpr.CellChanged frmThis.sprSht_OUTSUSU, Col, Row
End Sub

'-----------------------------------
' SpreadSheet 클릭
'-----------------------------------
Sub sprSht_SUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			.sprSht_SUSUDTL.MaxRows = 0
			for intCnt = 1 To .sprSht_SUSU.MaxRows
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1, 1, intCnt, intCnt, "", , , , , mstrCheck
			Next    

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_SUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_EOM_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOM, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_EOM.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_EOM, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_EOMSUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			.sprSht_EOMSUSUDTL.MaxRows = 0
			for intCnt = 1 To .sprSht_EOMSUSU.MaxRows
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOMSUSU, 1, 1, intCnt, intCnt, "", , , , , mstrCheck
			Next    

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_EOMSUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_EOMSUSU, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_OUTSUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUTSUSU, 1, intCnt
				Next
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

'--------------------------
'------시트 버튼 클릭------
'--------------------------
Sub sprSht_SUSU_ButtonClicked (Col,Row,ButtonDown)
	if Col = 1 and Row > 0 then 
		if mobjSCGLSpr.GetTextBinding( frmThis.sprSht_SUSU,"CHK",Row) = 1 THEN
			SelectRtn_SUSUDTL Col,Row
		ELSE
			call DeleteRtn_SUSUDTL(Row)
		END IF
	end if
End Sub

Sub DeleteRtn_SUSUDTL (Row)
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strSEQ	

	With frmThis
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_SUSUDTL.MaxRows to 1 step -1
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)

			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",i) = strTAXYEARMON and _
			   mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",i) = strTAXNO then

				mobjSCGLSpr.DeleteRow .sprSht_SUSUDTL,i
			end if				
		next
	End With
	err.clear	
End Sub

Sub sprSht_EOMSUSU_ButtonClicked (Col,Row,ButtonDown)
	if Col = 1 and Row > 0 then 
		if mobjSCGLSpr.GetTextBinding( frmThis.sprSht_EOMSUSU,"CHK",Row) = 1 THEN
			SelectRtn_EOMSUSUDTL Col,Row
		ELSE
			call DeleteRtn_EOMSUSUDTL(Row)
		END IF
	end if
End Sub

Sub DeleteRtn_EOMSUSUDTL (Row)
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strSEQ	

	With frmThis
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_EOMSUSUDTL.MaxRows to 1 step -1
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXNO",Row)

			if mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",i) = strTAXYEARMON and _
			   mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",i) = strTAXNO then

				mobjSCGLSpr.DeleteRow .sprSht_EOMSUSUDTL,i
			end if				
		next
	End With
	err.clear	
End Sub

'-----------------------------------
' SpreadSheet 더블 클릭
'-----------------------------------
sub sprSht_SUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_SUSU, ""
		end if
	end with
end sub

sub sprSht_OUT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUT, ""
		end if
	end with
end sub

sub sprSht_EOM_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_EOM, ""
		end if
	end with
end sub

sub sprSht_EOMSUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_EOMSUSU, ""
		end if
	end with
end sub

sub sprSht_OUTSUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUTSUSU, ""
		end if
	end with
end sub

'----------------------------------------------------------
'시트 자동 계산 [시트 키업]
'----------------------------------------------------------
'취급액
Sub sprSht_SUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_SUSU
	End With
End Sub

'취급액 ->매입
Sub sprSht_SUSUDTL_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_SUSUDTL
	End With
End Sub

'매입
Sub sprSht_OUT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_OUT
	End With
End Sub

'대행매입
Sub sprSht_EOM_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_EOM
	End With
End Sub

'대행매출
Sub sprSht_EOMSUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_EOMSUSU
	End With
End Sub

'매입수수료
Sub sprSht_OUTSUSU_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht_OUTSUSU
	End With
End Sub

SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"VAT") Then
		
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"VAT")) Then
				
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

'---------------------------------------------
'시트 마우스 업
'---------------------------------------------
'매출
Sub sprSht_SUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSU
	end with
End Sub

'매출
Sub sprSht_SUSUDTL_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_SUSUDTL
	end with
End Sub

'매입
Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_OUT
	end with
End Sub

'대행매입
Sub sprSht_EOM_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_EOM
	end with
End Sub

'대행매출
Sub sprSht_EOMSUSU_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_EOMSUSU
	end with
End Sub

'매입수수료
Sub sprSht_EOM_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_EOM
	end with
End Sub

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
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"AMT") or SprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"VAT") then
				
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

'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjMDCMIFCMALLVOCH = gCreateRemoteObject("cMDOT.ccMDOTIFCMALLVOCH")
	Set mobjMDCOVOCH		= gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	Set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
    mobjSCGLCtl.DoEventQueue
    
    Dim strComboPREPAYMENT
	Dim strSemuComboListB, strSemuComboListA
	Dim strBMORDER
	
	gSetSheetDefaultColor

    with frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strSemuComboListB =  "B0" & vbTab & "B5" & vbTab & "BR"
		strSemuComboListA =  "I0" & vbTab & "A0" & vbTab & "AI" & vbTab & "A8" & vbTab & "AZ"
		strBMORDER = "AD0420" & vbTab & " "

		'**************************************************
		'매출 시트 디자인 hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 25, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "선택|전표일자|거래처코드|거래처|적요|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    15|  20|  10|    10|       7| 5|       8|          10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "CUSTOMERCODE | CUSTNAME | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CUSTOMERCODE | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | BP | DUEDATE  | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "GBN | GFLAG | DEMANDDAY | AMTGBN | MEDFLAG", true 
		
		'**************************************************
		'매출 시트 디자인 dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_SUSUDTL, 30, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSUDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSUDTL,		    "전표일자|거래처코드|거래처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|VENDOR|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSUDTL, "-1",  "       8|        10|    15|  20|       5|         8|  10|    10|       7| 5|       8|          10|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSUDTL, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),-1,-1,strBMORDER,,80

		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "CUSTOMERCODE | BA | SEMU | BP | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | GBN | VOCHNO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true," CUSTOMERCODE | CUSTNAME | AMT | BP | DUEDATE | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN | GFLAG | DEMANDDAY | AMTGBN | VENDOR | MEDFLAG", true 

		'**************************************************
		'매입 시트 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 36, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "선택|전표일자|거래처코드|거래처|외주처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|지급기일|지급방법|BANKTYPE|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|삭제|GFLAG|매출구분|AMTGBN|VENDOR|RMSNO|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|   8|    0|      10|     0|     0|    0|		 0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO", true 

		'**************************************************
		'대행매입 시트 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_EOM
		mobjSCGLSpr.SpreadLayout .sprSht_EOM, 38, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_EOM,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_EOM,		    "선택|전표일자|거래처코드|거래처|외주처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|지급기일|지급방법|BANKTYPE|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|매출구분|AMTGBN|VENDOR|RMSNO|본사사업자번호|본사명|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_EOM, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      10|     0|     0|    0|		      0|     0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht_EOM, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_EOM, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_EOM, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_EOM, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOM, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EOM, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOM,mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOM,mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOM,mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_EOM,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_EOM, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht_EOM, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,true,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_EOM, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME", true 
		mobjSCGLSpr.CellGroupingEach .sprSht_EOM,"VOCHNO | ERRCODE | ERRMSG"
		
		
		'**************************************************
		'대행매출 시트 디자인 hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_EOMSUSU
		mobjSCGLSpr.SpreadLayout .sprSht_EOMSUSU, 25, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_EOMSUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_EOMSUSU,		    "선택|전표일자|거래처코드|거래처|적요|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_EOMSUSU, "-1",  "  4|       8|        10|    15|  20|  10|    10|       7| 5|       8|          10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_EOMSUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_EOMSUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_EOMSUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_EOMSUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSU, "CUSTOMERCODE | CUSTNAME | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EOMSUSU, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOMSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOMSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSU,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_EOMSUSU, "CUSTOMERCODE | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | BP | DUEDATE  | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_EOMSUSU, "GBN | GFLAG | DEMANDDAY | AMTGBN | MEDFLAG", true 
		
		'**************************************************
		'대행매출 시트 디자인 dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_EOMSUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_EOMSUSUDTL, 30, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_EOMSUSUDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_EOMSUSUDTL,		    "전표일자|거래처코드|거래처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|VENDOR|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_EOMSUSUDTL, "-1",  "       8|        10|    15|  20|       5|         8|  10|    10|       7| 5|       8|          10|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_EOMSUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_EOMSUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_EOMSUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSUDTL, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EOMSUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EOMSUSUDTL, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOMSUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOMSUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_EOMSUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_EOMSUSUDTL,"BMORDER"),-1,-1,strBMORDER,,80

		mobjSCGLSpr.SetCellAlign2 .sprSht_EOMSUSUDTL, "CUSTOMERCODE | BA | SEMU | BP | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | GBN | VOCHNO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSUDTL,true," CUSTOMERCODE | CUSTNAME | AMT | BP | DUEDATE | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO", true 
		
		'**************************************************
		'매출 시트 디자인 hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 25, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "선택|전표일자|거래처코드|거래처|적요|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    15|  20|  10|    10|       7| 5|       8|          10|   0|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "CUSTOMERCODE | CUSTNAME | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSU,mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSU,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CUSTOMERCODE | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | BP | DUEDATE  | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "GBN | GFLAG | DEMANDDAY | AMTGBN | MEDFLAG", true 
		
		'**************************************************
		'매출 시트 디자인 dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_SUSUDTL, 30, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSUDTL,    "POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht_SUSUDTL,		    "전표일자|거래처코드|거래처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|지급기일|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|AMTGBN|VENDOR|MEDFLAG"
		mobjSCGLSpr.SetColWidth .sprSht_SUSUDTL, "-1",  "       8|        10|    15|  20|       5|         8|  10|    10|       7| 5|       8|          10|   0|       7|   7|      7|     8|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|     0|     0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSUDTL, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"BMORDER"),-1,-1,strBMORDER,,80

		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "CUSTOMERCODE | BA | SEMU | BP | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | GBN | VOCHNO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true," CUSTOMERCODE | CUSTNAME | AMT | BP | DUEDATE | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN | GFLAG | DEMANDDAY | AMTGBN | VENDOR | MEDFLAG", true 

		'**************************************************
		'매입 수수료 시트 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUTSUSU
		mobjSCGLSpr.SpreadLayout .sprSht_OUTSUSU, 38, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUTSUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUTSUSU,		    "선택|전표일자|거래처코드|거래처|외주처|적요|사업영역|코스트센터|금액|부가세|세무코드|BP|광고주지급일|구분|차변계정|계정|BMORDER|증빙일|지급기일|지급방법|BANKTYPE|선수금구분|선수금(시작일)|선수금(종료일)|본문TEXT|RMS년월|RMS번호|전표번호|에러코드|에러메세지|GFLAG|매출구분|AMTGBN|VENDOR|RMSNO|본사사업자번호|본사명|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUTSUSU, "-1",  "   4|       8|        10|     0|    15|  20|       5|         8|  10|    10|       7| 5|          10|   0|       7|   7|      7|     8|       8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      10|     0|     0|    0|		      0|     0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUTSUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUTSUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUTSUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUTSUSU, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUTSUSU, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUTSUSU, "AMT | VAT", -1, -1, 0 '숫자형
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUTSUSU,mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht_OUTSUSU,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUTSUSU, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUTSUSU, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_OUTSUSU, "GBN | CUSTNAME | GFLAG | MEDFLAG | DUEDATE | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME", true 
		mobjSCGLSpr.CellGroupingEach .sprSht_OUTSUSU,"VOCHNO | ERRCODE | ERRMSG"

	End with
	pnlTab_susu.style.visibility = "visible"

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMIFCMALLVOCH = Nothing
	Set mobjMDCOVOCH = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet초기화
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.sprSht_EOM.MaxRows = 0
		.sprSht_EOMSUSU.MaxRows = 0
		.sprSht_EOMSUSUDTL.MaxRows = 0
		.sprSht_OUTSUSU.MaxRows = 0
		
		.txtYEARMON.focus		
		
		Get_COMBO_VALUE	
		'처음에 강제 삭제 감춤
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
	End with
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt
   	
	With frmThis	
		'Sheet초기화
		.sprSht_SUSU.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.sprSht_EOM.MaxRows = 0
		.sprSht_EOMSUSU.MaxRows = 0
		.sprSht_OUTSUSU.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
	
		vntData = mobjMDCMIFCMALLVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUT, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_EOM, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUTSUSU, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True
   		End If
   	End With
End Sub

'-----------------------------------------------------------------------------------------
' 그리드 서브 콤보 설정
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strCODE, row, sprsht)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCODE = replace(strCODE,"-","")
		
       	vntData = mobjMDCMIFCMALLVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
		If not gDoErrorRtn ("Get_SUBCOMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 sprsht, "BANKTYPE",Row,Row,vntData,,160 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
Sub SelectRtn (strVOCH_TYPE)
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.sprSht_EOM.MaxRows = 0
		.sprSht_EOMSUSU.MaxRows = 0
		.sprSht_EOMSUSUDTL.MaxRows = 0
		.sprSht_OUTSUSU.MaxRows = 0
		
		IF strVOCH_TYPE = "S" THEN
			CALL SelectRtn_SUSU()
		ELSEIF strVOCH_TYPE = "YO" THEN
			CALL SelectRtn_OUT()
		ELSEIF strVOCH_TYPE = "YO2" THEN
			CALL SelectRtn_EOM()
		ELSEIF strVOCH_TYPE = "YO3" THEN
			CALL SelectRtn_EOMSUSU()
		ELSEIF strVOCH_TYPE = "YS" THEN
			CALL SelectRtn_OUTSUSU()
		END IF
		mstrSTAY = TRUE
   	end with
End Sub

Sub SelectRtn_SUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
    
    
	with frmThis
		.sprSht_SUSU.MaxRows = 0

		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value

		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 

		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_SUSU(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, _
													 strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_SUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "S"
				mobjSCGLSpr.SetClipbinding .sprSht_SUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht_SUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"DUEDATE",intCnt,intCnt,false
					End If
				
					'선수금 처리시
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,True,"TODATE",intCnt,intCnt,false
					End If		
				Next
				AMT_SUM .sprSht_SUSU
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_SUSUDTL (Col, Row)
	Dim vntData
   	Dim i, strCols
    Dim intCnt, intCnt2
    Dim strTAXYEARMON
    Dim strTAXNO
    Dim strRow
    Dim strCLIENTGBN
    
    Dim intHDRAMT
    Dim intDTLAMT
    Dim intSUMAMT
    Dim chkTRUE
    Dim intCNT3
    
	with frmThis
		'Sheet초기화
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		intHDRAMT = 0 : intDTLAMT = 0 : intSUMAMT = 0 : chkTRUE = true
		
		Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
		Dim strFROMDATE, strTODATE, strSUMMTEXT, strPOSTINGDATE
		


		strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
		strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)
		strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMM",Row)
		strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SEMU",Row)
		strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DEMANDDAY",Row)
		strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DUEDATE",Row)
		strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"DOCUMENTDATE",Row)
		strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"PREPAYMENT",Row)
		strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"FROMDATE",Row)
		strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TODATE",Row)
		strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"SUMMTEXT",Row)
		strPOSTINGDATE = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"POSTINGDATE",Row)
		
		
		
		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO, strPREPAYMENT)
		
		
		If not gDoErrorRtn ("SelectRtn_SUSUDTL") Then
			If mlngRowCnt >0 Then
				strRow = 0
				strRow = .sprSht_SUSUDTL.MaxRows + 1
				Call mobjSCGLSpr.SetClipBinding (.sprSht_SUSUDTL,vntData, 1, strRow, mlngColCnt, mlngRowCnt,True)
				
				For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) = strTAXNO Then

						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMM",intCnt, strSUMM
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SEMU",intCnt, strSEMU
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DUEDATE",intCnt, strDUEDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"FROMDATE",intCnt, strFROMDATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"TODATE",intCnt, strTODATE
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
						mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"POSTINGDATE",intCnt, strPOSTINGDATE
					End If
				Next
				
				'반복을 돌면서 HDR 값을 넘기면 뿌리지 않는다.
				for intCnt2 = 1 To .sprSht_SUSUDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt2) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt2) = strTAXNO Then
						
						if chkTRUE = true then
							intHDRAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"AMT",Row)
							chkTRUE = false
						else
							intDTLAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt2)
							intSUMAMT = intSUMAMT + intDTLAMT
							
							if intHDRAMT < intSUMAMT then
								for intCNT3 = intCnt2 to .sprSht_SUSUDTL.MaxRows
									mobjSCGLSpr.DeleteRow .sprSht_SUSUDTL,intCNT3 
								next 
								
							end if 
						end if 
					end if
				next 
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_OUT ()
   	Dim vntData
    Dim intCnt
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN

	with frmThis
		.sprSht_OUT.MaxRows = 0

		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value

		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 

		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
													  strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_OUT") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "YO"

				mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"ACCOUNT",-1,-1,false
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"POSTINGDATE",intCnt,intCnt,false
					End If
				
					'선수금 처리시
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,false
					End If
				Next

				AMT_SUM .sprSht_OUT
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_EOM ()
   	Dim vntData
    Dim intCnt
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN

	with frmThis
		.sprSht_EOM.MaxRows = 0

		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value

		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 
		
		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_EOM(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
												  strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_EOM") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "YO2"

				mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"ACCOUNT",-1,-1,false
				mobjSCGLSpr.SetClipbinding .sprSht_EOM, vntData, 1, 1, mlngColCnt, mlngRowCnt, True

				For intCnt = 1 To .sprSht_EOM.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOM, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,true,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOM, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"POSTINGDATE",intCnt,intCnt,false
					End If

					'선수금 처리시
					If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOM,True,"TODATE",intCnt,intCnt,false
					End If
				Next

				AMT_SUM .sprSht_EOM
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

'대행 매출
Sub SelectRtn_EOMSUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
    
    
	with frmThis
		.sprSht_EOMSUSU.MaxRows = 0

		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strYEARMON = .txtYEARMON.value 
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value

		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 

		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_EOMSUSU(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, _
													 strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_EOMSUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "YO3"
				mobjSCGLSpr.SetClipbinding .sprSht_EOMSUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht_EOMSUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOMSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_EOMSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,false,"DUEDATE",intCnt,intCnt,false
					End If
				
					'선수금 처리시
					If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_EOMSUSU,True,"TODATE",intCnt,intCnt,false
					End If		
				Next
				AMT_SUM .sprSht_EOMSUSU
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_EOMSUSUDTL (Col, Row)
	Dim vntData
   	Dim i, strCols
    Dim intCnt, intCnt2
    Dim strTAXYEARMON
    Dim strTAXNO
    Dim strRow
    Dim strCLIENTGBN
    
    Dim intHDRAMT
    Dim intDTLAMT
    Dim intSUMAMT
    Dim chkTRUE
    Dim intCNT3
    
	with frmThis
		'Sheet초기화
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		intHDRAMT = 0 : intDTLAMT = 0 : intSUMAMT = 0 : chkTRUE = true
		
		Dim strSUMM, strSEMU, strDEMANDDAY, strDUEDATE, strDOCUMENTDATE, strPREPAYMENT
		Dim strFROMDATE, strTODATE, strSUMMTEXT, strPOSTINGDATE
		


		strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXYEARMON",Row)
		strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXNO",Row)
		strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SUMM",Row)
		strSEMU = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SEMU",Row)
		strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DEMANDDAY",Row)
		strDUEDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DUEDATE",Row)
		strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"DOCUMENTDATE",Row)
		strPREPAYMENT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"PREPAYMENT",Row)
		strFROMDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"FROMDATE",Row)
		strTODATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TODATE",Row)
		strSUMMTEXT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"SUMMTEXT",Row)
		strPOSTINGDATE = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"POSTINGDATE",Row)
		
		
		
		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_EOMSUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO, strPREPAYMENT)
		
		
		If not gDoErrorRtn ("SelectRtn_EOMSUSUDTL") Then
			If mlngRowCnt >0 Then
				strRow = 0
				strRow = .sprSht_EOMSUSUDTL.MaxRows + 1
				Call mobjSCGLSpr.SetClipBinding (.sprSht_EOMSUSUDTL,vntData, 1, strRow, mlngColCnt, mlngRowCnt,True)
				
				For intCnt = 1 To .sprSht_EOMSUSUDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) = strTAXNO Then

						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SUMM",intCnt, strSUMM
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SEMU",intCnt, strSEMU
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt, strDEMANDDAY
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DUEDATE",intCnt, strDUEDATE
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt, strDOCUMENTDATE
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt, strPREPAYMENT
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"FROMDATE",intCnt, strFROMDATE
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"TODATE",intCnt, strTODATE
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt, strSUMMTEXT
						mobjSCGLSpr.SetTextBinding .sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt, strPOSTINGDATE
					End If
				Next
				
				'반복을 돌면서 HDR 값을 넘기면 뿌리지 않는다.
				for intCnt2 = 1 To .sprSht_EOMSUSUDTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt2) = strTAXYEARMON AND _
						mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt2) = strTAXNO Then
						
						if chkTRUE = true then
							intHDRAMT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"AMT",Row)
							chkTRUE = false
						else
							intDTLAMT = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMT",intCnt2)
							intSUMAMT = intSUMAMT + intDTLAMT
							
							if intHDRAMT < intSUMAMT then
								for intCNT3 = intCnt2 to .sprSht_EOMSUSUDTL.MaxRows
									mobjSCGLSpr.DeleteRow .sprSht_EOMSUSUDTL,intCNT3 
								next 
								
							end if 
						end if 
					end if
				next 
   			End If
   		End If
   	end with
End Sub

'매입 수수료 조회
Sub SelectRtn_OUTSUSU ()
   	Dim vntData
    Dim intCnt
    dIM strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN

	with frmThis
		.sprSht_OUTSUSU.MaxRows = 0
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value

		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF

		vntData = mobjMDCMIFCMALLVOCH.SelectRtn_OUTSUSU(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
														strCLIENTCODE, strCLIENTNAME, strGBN)

		if not gDoErrorRtn ("SelectRtn_OUTSUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "YS"

				mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"ACCOUNT",-1,-1,false
				mobjSCGLSpr.SetClipbinding .sprSht_OUTSUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True

				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"DUEDATE",intCnt,intCnt,false
						'mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,true,"POSTINGDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUTSUSU, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"DUEDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"POSTINGDATE",intCnt,intCnt,false
					End If

					'선수금 처리시
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUTSUSU,True,"TODATE",intCnt,intCnt,false
					End If	
				Next
				AMT_SUM .sprSht_OUTSUSU
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

'--------------------------
'조회된 내역 금액 합산 로직
'-------------------------- 
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

Function DataValidation_SUSU ()
	DataValidation_SUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0

	With frmThis
		For intCnt =1  To .sprSht_SUSUDTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_SUSU = True
End Function

Function DataValidation_OUT ()
	DataValidation_OUT = false	
	Dim intCnt, intCnt2
	Dim chkcnt

	intCnt = 0

	With frmThis
		For intCnt =1  To .sprSht_OUT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUT = True
End Function

Function DataValidation_EOM ()
	DataValidation_EOM = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_EOM.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_EOM = True
End Function

Function DataValidation_EOMSUSU ()
	DataValidation_EOMSUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0

	With frmThis
		For intCnt =1  To .sprSht_EOMSUSUDTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_EOMSUSU = True
End Function

'대행수수료
Function DataValidation_OUTSUSU ()
	DataValidation_OUTSUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_OUTSUSU.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 광고주청구일 을 확인하십시오","저장오류"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUTSUSU = True
End Function

'저장로직
Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn
	gFlowWait meWAIT_ON

	with frmThis
		IF mstrPROCESS = "Create" THEN
			IF NOT .rdF.checked THEN
				gErrorMsgBox "미완료조회시 가능합니다.","생성및삭제"
				exit sub
			end IF 
		end if 
		IF mstrPROCESS = "Delete" THEN
			IF NOT .rdT.checked THEN
				gErrorMsgBox "완료조회시 가능합니다.","생성및삭제"
				exit sub
			end IF
		end if

		IF mstrSTAY THEN 
			mstrSTAY = FALSE
			IF strVOCH_TYPE = "S" THEN
				if DataValidation_SUSU =false then exit sub
				CALL ProcessRtn_SUSU()
			ELSEIF strVOCH_TYPE = "YO" THEN
				if DataValidation_OUT =false then exit sub
				CALL ProcessRtn_OUT()
			ELSEIF strVOCH_TYPE = "YO2" THEN
				if DataValidation_EOM =false then exit sub
				CALL ProcessRtn_EOM()
			ELSEIF strVOCH_TYPE = "YO3" THEN
				if DataValidation_EOMSUSU =false then exit sub
				CALL ProcessRtn_EOMSUSU()
			ELSEIF strVOCH_TYPE = "YS" THEN
			if DataValidation_OUTSUSU =false then exit sub
				CALL ProcessRtn_OUTSUSU()
			END IF
		ELSE
			gErrorMsgBox "전표처리 진행중입니다.","전표처리 안내"
		END IF
   	end with
End Sub

Sub ProcessRtn_SUSU()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	
	'채번을 설정하기 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht_SUSUDTL, meINS_FLAG
		
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_SUSUDTL,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그
		
		strTAXYEARMON = "" : strTAXNO = ""
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0006"
		
		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1
		
				strRMS_DOC_TYPE = "D"
				
				'채번을 설정한다.
				'--------------------------------------------------------------------------------------
					
				'DTL 시트의 같은 로우 묶음은 하나의 전표번호가 채번된다.
				If strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) Then
				ELSE

					strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

					strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","")
					strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"MEDFLAG",intCnt)
					strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
					strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)'
					strTYPE				= "1"

					if strGROUPSEQ = true then
						strGROUP = TRUE
					else 
						strGROUP = FALSE
					END IF 

					If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP,strTYPE) Then 
						gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
						Exit Sub
					END IF 

					strGROUPSEQ = FALSE
					
					'생성 저장한 RMS 채번 가져오기
					vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
					
					strVOCHNORMS =  vntData(0,1)
				END IF
				'---------------------------------------------------------------------------------------
				
				if strIF_CNT = "1" then

					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				else

					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN

						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if

					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)

				end if

				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1

				strRMS_DOC_TYPE = "Z"

				if strIF_CNT = "1" then
					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				else
					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if
					
					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BMORDER",intCnt)
				end if
	
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)

			Next
		end if 

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'저장로직[매입전표생성]
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'전표 채번을 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | TRANSRANK")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If

		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그

		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)

		Dim IF_GUBUN

		IF_GUBUN = "RMS_0007"

		'시트 전체를 돌면서 체크된 값의 TRANSRANK 의 최대 값을 담는다.
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next

		'합산에 사용할 변수
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		'분할체크 상태일 경우
		IF .rdDIV.checked THEN
		
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then	
						'채번을 설정한다.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt)'
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'생성 저장한 RMS 채번 가져오기
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)
						'---------------------------------------------------------------------------------------
							
						strIF_CNT = strIF_CNT + 1

						strRMS_DOC_TYPE = "O"

						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_OUT.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ+1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",intCnt)
						end if												
					end if 
				Next
			end if
		'합산 체크일 경우
		ELSE
			if mstrPROCESS = "Create" then
				'1부터 TRANRANK 의 최고값만큼 돌면서 전표처리
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'전체 시트를 돌면서 체크가 된 데이터의 TRANSRANK 가 같으면 금액을 합산한다.
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'청구합계
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprSht_OUT.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = "1" then	
						'체크가 되어있는 데이터가 합산할수있는 데이터라면 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'같은 합산 데이터는 한데이터만 생성한다.
								Else
									'채번을 설정한다.(합산전표의 채번 설정)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'생성 저장한 RMS 채번 가져오기
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
									
									IF strIF_CNT = "1" THEN
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									ELSE
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									END IF 

									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then
												'합산 데이터라면 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",j)
											END IF 
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								END IF
							end if
						end if 
					Next
				 Next
			'합산일 경우 전표삭제
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
					
					For i = 1 To .sprSht_OUT.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
							'TRANSRANK 가 같으면 금액과 부가세를 합산한다.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUT.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'금액합계 부가세 합계는 헤더에 변수로 저장
								Else
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)	
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)

									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
																			
									For j = 1 To .sprSht_OUT.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",j) = 1 Then
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BMORDER",i)
															
											END IF
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								END IF
								intCnt2 = intCnt
							END IF
						end if 
					Next
				NEXT
			end if
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)

   	end with
End Sub

'저장로직[대행매입전표생성]
Sub ProcessRtn_EOM()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'전표 채번을 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_EOM,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		
		IF_GUBUN = "RMS_0007"
		
		'시트 전체를 돌면서 체크된 값의 TRANSRANK 의 최대 값을 담는다.
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_EOM.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		'합산에 사용할 변수
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		'분할체크 상태일 경우
		IF .rdDIV.checked THEN
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_EOM.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"chk",intCnt) = "1" then	
						'채번을 설정한다.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",intCnt)'
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'생성 저장한 RMS 채번 가져오기
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)
						'---------------------------------------------------------------------------------------
						
						strIF_CNT = strIF_CNT + 1

						strRMS_DOC_TYPE = "O"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",intCnt) 
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",intCnt)
						end if
					end if
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_EOM.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ+1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",intCnt)
						end if												
					end if 
				Next
			end if
		'합산 체크일 경우
		ELSE
			if mstrPROCESS = "Create" then
				'1부터 TRANRANK 의 최고값만큼 돌면서 전표처리
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'전체 시트를 돌면서 체크가 된 데이터의 TRANSRANK 가 같으면 금액을 합산한다.
					For i = 1 To .sprSht_EOM.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",i) = 1 Then
							'청구합계
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_EOM.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"chk",i) = "1" then	
						'체크가 되어있는 데이터가 합산할수있는 데이터라면 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'같은 합산 데이터는 한데이터만 생성한다.
								Else
									'채번을 설정한다.(합산전표의 채번 설정)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",i)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'생성 저장한 RMS 채번 가져오기
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",i)
									
									IF strIF_CNT = "1" THEN
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									ELSE
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									END IF 
									
									For j = 1 To .sprSht_EOM.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",j) = 1 Then
												'합산 데이터라면 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",j),"-","") + "|" + _
															strVOCHNO + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",j)
											END IF 
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								END IF
							end if
						end if 
					Next
				 Next
			'합산일 경우 전표삭제
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
					
					For i = 1 To .sprSht_EOM.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",i) = 1 Then
							'TRANSRANK 가 같으면 금액과 부가세를 합산한다.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_EOM.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'금액합계 부가세 합계는 헤더에 변수로 저장
								Else
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VOCHNO",i)	
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",i)
									
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
																			
									For j = 1 To .sprSht_EOM.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",j) = 1 Then
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"BMORDER",i)
											END IF
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								END IF
								intCnt2 = intCnt
							END IF
						end if 
					Next
				NEXT
			end if
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'대행 매출 전표 
Sub ProcessRtn_EOMSUSU()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	
	'채번을 설정하기 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht_EOMSUSUDTL, meINS_FLAG
		
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_EOMSUSUDTL,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | VENDOR | MEDFLAG")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그
		
		strTAXYEARMON = "" : strTAXNO = ""
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0006"
		
		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_EOMSUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1
		
				strRMS_DOC_TYPE = "D"
				
				'채번을 설정한다.
				'--------------------------------------------------------------------------------------
					
				'DTL 시트의 같은 로우 묶음은 하나의 전표번호가 채번된다.
				If strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) Then
				ELSE

					strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

					strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt),"-","")
					strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"MEDFLAG",intCnt)
					strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt)
					strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt)'
					strTYPE				= "1"

					if strGROUPSEQ = true then
						strGROUP = TRUE
					else 
						strGROUP = FALSE
					END IF 

					If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP,strTYPE) Then 
						gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
						Exit Sub
					END IF 

					strGROUPSEQ = FALSE
					
					'생성 저장한 RMS 채번 가져오기
					vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
					
					strVOCHNORMS =  vntData(0,1)
				END IF
				'---------------------------------------------------------------------------------------
				
				if strIF_CNT = "1" then

					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BMORDER",intCnt)
				else

					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) THEN

						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if

					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									strVOCHNORMS + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BMORDER",intCnt)

				end if

				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt)
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_EOMSUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1

				strRMS_DOC_TYPE = "Z"

				if strIF_CNT = "1" then
					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BMORDER",intCnt)
				else
					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if
					
					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VENDOR",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"VOCHNO",intCnt) + "|" + _
									"" + "|" + _  
									mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"BMORDER",intCnt)
				end if
	
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSUDTL,"TAXNO",intCnt)

			Next
		end if 

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub


'저장로직[매입수수료 전표생성]
Sub ProcessRtn_OUTSUSU()
	Dim intRtn
	Dim intColFlag, bsdiv, intMaxCnt
	
	'전표 채번을 위한 변수
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUTSUSU,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DUEDATE  | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | DEMANDDAY | PAYCODE | BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | VENDOR | RMSNO | REAL_MED_BUSINO | REAL_MED_NAME | TRANSRANK")
		'처리 업무객체 호출
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '임시전표 삭제 플래그
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		
		IF_GUBUN = "RMS_0007"
		
		'시트 전체를 돌면서 체크된 값의 TRANSRANK 의 최대 값을 담는다.
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUTSUSU.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		'합산에 사용할 변수
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		'분할체크 상태일 경우
		IF .rdDIV.checked THEN
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",intCnt) = "1" then	
						'채번을 설정한다.
						'--------------------------------------------------------------------------------------
						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt)'
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'생성 저장한 RMS 채번 가져오기
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)
						'---------------------------------------------------------------------------------------
						
						strIF_CNT = strIF_CNT + 1

						strRMS_DOC_TYPE = "O"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						end if
					end if
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprSht_OUTSUSU.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ+1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",intCnt)
						end if												
					end if 
				Next
			end if
		'합산 체크일 경우
		ELSE
			if mstrPROCESS = "Create" then
				'1부터 TRANRANK 의 최고값만큼 돌면서 전표처리
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M"
					'전체 시트를 돌면서 체크가 된 데이터의 TRANSRANK 가 같으면 금액을 합산한다.
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
							'청구합계
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",i) = "1" then	
						'체크가 되어있는 데이터가 합산할수있는 데이터라면 
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then	
								If intCnt2 = intCnt Then
								'같은 합산 데이터는 한데이터만 생성한다.
								Else
								
									'채번을 설정한다.(합산전표의 채번 설정)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"MEDFLAG",i)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "전표 번호가 제대로 생성되지 않았습니다. 개발자에게 문의하세요 ","전표 생성 취소"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'생성 저장한 RMS 채번 가져오기
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i)
									strRMS_DOC_TYPE	= "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
									
									IF strIF_CNT = "1" THEN
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									ELSE
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									END IF 
									
									For j = 1 To .sprSht_OUTSUSU.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",j) = 1 Then
												'합산 데이터라면 
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",j)) = intCnt Then
												strIF_CNT = strIF_CNT + 1
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",j)
											END IF 
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								END IF
							end if
						end if 
					Next
				 Next
			'합산일 경우 전표삭제
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
							'TRANSRANK 가 같으면 금액과 부가세를 합산한다.
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next
					
					For i = 1 To .sprSht_OUTSUSU.MaxRows
						if mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"chk",i) = "1" then	
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",i)) = intCnt Then
								If intCnt2 = intCnt Then
								'금액합계 부가세 합계는 헤더에 변수로 저장
								Else
									strIF_CNT = strIF_CNT + 1
									'각각의 값을 변수에 담아 한데이터로 만들고 금액은 합산한 금액을 입력한다.
									strPOSTINGDATE  = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"REAL_MED_BUSINO",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i)
									strCUSTOMERCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"REAL_MED_BUSINO",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i)
									strRMS_DOC_TYPE	= "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEBTOR",i)
									strDOCUMENTDATE	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i)	
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
									
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
																			
									For j = 1 To .sprSht_OUTSUSU.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",j) = 1 Then
											If CDbl(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"POSTINGDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VENDOR",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMM",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BA",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"COSTCENTER",i) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMT",i)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SEMU",i) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BP",i) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DEMANDDAY",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CUSTOMERCODE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"GFLAG",i) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"ACCOUNT",i) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DOCUMENTDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PREPAYMENT",i) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"FROMDATE",i),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TODATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"SUMMTEXT",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"AMTGBN",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"PAYCODE",i) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"DUEDATE",i),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BANKTYPE",i) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"BMORDER",i)
											END IF
										END IF 
									NEXT
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								END IF
								intCnt2 = intCnt
							END IF
						end if 
					Next
				NEXT
			end if
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'****************************************************************************************
' 채번 설정처리
'****************************************************************************************
Function InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
	InsertRtn_VOCHNO = false
   	Dim strVOCHNO
	With frmThis
		
		'채번을 설정& 저장한다 (저장은 중복을 막고 SAP 쪽에서 에러가 날 경우에도 같은 번호로 생성되는 것을 막는다.).
		intRtn = mobjMDCOVOCH.InsertRtn_VOCHNO(gstrConfigXml,strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
		If not gDoErrorRtn ("InsertRtn_VOCHNO") Then
		
			If intRtn = 0 Then
				Exit Function
			End If		
   		End If
   	end With
   	InsertRtn_VOCHNO = true
End Function

'--------------------------------------------------
' 전표상태 및 전표번호 받아오기 및 실제 RMS업데이트
'--------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	gFlowWait meWAIT_ON
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
	
		if mstrPROCESS ="Create" then
			if mstrGUBUN = "S" then
				intRtn = mobjMDCMIFCMALLVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "Y")
			elseif mstrGUBUN = "YO" then
				intRtn = mobjMDCMIFCMALLVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "YO")
			elseif mstrGUBUN = "YO2" then
				intRtn = mobjMDCMIFCMALLVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "YO2")
			elseif mstrGUBUN = "YO3" then
				intRtn = mobjMDCMIFCMALLVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "YO3")
			elseif mstrGUBUN = "YS" then
				intRtn = mobjMDCMIFCMALLVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "YS")
			end if 

			if not gDoErrorRtn ("ProcessRtn") then
				'모든 플래그 클리어
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "YO" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT, meCLS_FLAG
				ELSEIF mstrGUBUN = "YO2" THEN
					mobjSCGLSpr.SetFlag  .sprSht_EOM, meCLS_FLAG
				ELSEIF mstrGUBUN = "YO3" THEN
					mobjSCGLSpr.SetFlag  .sprSht_EOMSUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "YS" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU, meCLS_FLAG
				END IF

				if intRtn > 0 Then
					gErrorMsgBox "전표가 생성되었습니다.","저장안내"
				else
					gErrorMsgBox "에러가 발생했습니다.","저장안내"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   			
   		elseif mstrPROCESS ="Delete" then
   			if mstrGUBUN = "S" or mstrGUBUN = "YO3" then
				intRtn = mobjMDCMIFCMALLVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "IFCMALL")
			else
				intRtn = mobjMDCMIFCMALLVOCH.VOCHDELL_BUY(gstrConfigXml, strRETURNLIST, mstrGUBUN, "IFCMALL", mstrGUBUN)
			end if 
			
   			if not gDoErrorRtn ("VOCHDELL") then
				'모든 플래그 클리어
				IF mstrGUBUN = "S" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "YO" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				ELSEIF mstrGUBUN = "YO2" THEN
					mobjSCGLSpr.SetFlag  .sprSht_EOM,meCLS_FLAG
				ELSEIF mstrGUBUN = "YO3" THEN
					mobjSCGLSpr.SetFlag  .sprSht_EOMSUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "YS" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "전표가 삭제되었습니다.","저장안내"
				else
					gErrorMsgBox "에러가 발생했습니다.","저장안내"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   		end if

   		IF mstrGUBUN = "S" THEN
			.sprSht_SUSU.focus()
		ELSEIF mstrGUBUN = "YO" THEN
			.sprSht_OUT.focus()
		ELSEIF mstrGUBUN = "YO2" THEN
			.sprSht_EOM.focus()
		ELSEIF mstrGUBUN = "YO3" THEN
			.sprSht_EOMSUSU.focus()
		ELSEIF mstrGUBUN = "YS" THEN
			.sprSht_OUTSUSU.focus()
		END IF
	End With
	gFlowWait meWAIT_OFF
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   		
   		IF NOT .rdE.checked THEN
			gErrorMsgBox "오류조회시 가능합니다.","생성및삭제"
			exit sub
		end if

		IF mstrGUBUN = "S" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_SUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "YO" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "YO2" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_EOM,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "YO3" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_EOMSUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		ELSEIF mstrGUBUN = "YS" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUTSUSU,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF

		'처리 업무객체 호출
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"삭제취소"
			exit sub
		End If
		
		intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'모든 플래그 클리어
			IF mstrGUBUN = "S" THEN
				mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "YO" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
			ELSEIF mstrGUBUN = "YO2" THEN
				mobjSCGLSpr.SetFlag  .sprSht_EOM,meCLS_FLAG
			ELSEIF mstrGUBUN = "YO3" THEN
				mobjSCGLSpr.SetFlag  .sprSht_EOMSUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "YS" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUTSUSU,meCLS_FLAG
			END IF
			
			if intRtn > 0 Then
				gErrorMsgBox "오류 전표가 삭제되었습니다.","저장안내"
			End If

			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

'-----------------------------------------
'전표 강제 삭제
'-----------------------------------------
Sub DeleteRtn (strGUBUN)
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strVOCHNO
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
	
		If mstrGUBUN = "S"  then  
			If .sprSht_SUSU.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If

			For i = 1 To .sprSht_SUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "YO"  then
			If .sprSht_OUT.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUT.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "YO2"  then
			If .sprSht_EOM.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If

			For i = 1 To .sprSht_EOM.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "YO3"  then
			If .sprSht_EOMSUSU.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If

			For i = 1 To .sprSht_EOMSUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "YS"  then
			If .sprSht_OUTSUSU.MaxRows = 0 then
				gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUTSUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
				exit sub
			end if			
		END IF
	
		intRtn = gYesNoMsgbox("강제삭제는 SAP에서 승인된 전표를 SAP에서 취소하여 RMS쪽에서 삭제할 수 없을때 RMS쪽 전표를 강제로 삭제할때 사용합니다. " & vbCrlf & "  " & vbCrlf & " 전표를 강제로 삭제하시겠습니까?","강제삭제 확인")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0

		'선택된 자료를 끝에서 부터 삭제
		If mstrGUBUN = "S"  then
			for i = .sprSht_SUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"VOCHNO",i)

					intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "IFCMALL" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
   					End If

   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "YO"  then
			for i = .sprSht_OUT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
					
					intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "IFCMALL", "YO" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "YO2"  then
			for i = .sprSht_EOM.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOM,"VOCHNO",i)
					
					intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "IFCMALL", "YO2")
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_EOM,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "YO3"  then
			for i = .sprSht_EOMSUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_EOMSUSU,"VOCHNO",i)
					
					intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "IFCMALL", "YO3")
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_EOMSUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "YS"  then
			for i = .sprSht_OUTSUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUTSUSU,"VOCHNO",i)

					intRtn = mobjMDCMIFCMALLVOCH.DeleteRtn_GANG_BUY(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "IFCMALL", "YS")

					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUTSUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		END IF
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
   		
		SelectRtn (strGUBUN)
	End With
	err.clear	
End Sub

		</script>
		<script language="javascript">
		//##########################################################################################################################################
		//***************************************주1) frmSapCon 아이 프레임 을 이용하여 Submit 하는 함수********************************************
		//##########################################################################################################################################
		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {
			//헤더
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			//dtl 
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			
			window.frames[0].document.forms[0].submit();
		}

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind" />
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD id="TD1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="131" background="../../../images/back_p.gIF"
													border="0">
													<TR>
														<TD align="left" width="100%" height="2" />
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3" />
										<tr>
											<td class="TITLE">IFC MALL 전표관리&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1" />
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%" height="93%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2" />
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')"
													width="50">&nbsp;년월
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 60px"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 60px; HEIGHT: 22px" accessKey="NUM"
														maxLength="6" size="9" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
													width="50">&nbsp;광고주
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 112px; HEIGHT: 22px"
														maxLength="100" size="13" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 53px; HEIGHT: 22px"
														accessKey=",M" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
												<td class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery">
															</TD>
														</TR>
													</TABLE>
												</td>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">발행
												</TD>
												<TD class="SEARCHDATA" colSpan="5"><INPUT id="rdT" title="완료내역조회" onclick="vbscript:Call Set_delete('imgVochDelco')" type="radio"
														value="rdT" name="rdGBN">&nbsp;완료&nbsp; <INPUT id="rdF" title="미완료 내역조회" onclick="vbscript:Call Set_delete('imgVochDelco')" type="radio"
														CHECKED value="rdF" name="rdGBN">&nbsp;미완료&nbsp; <INPUT id="rdE" title="오류전표 내역조회" onclick="vbscript:Call Set_delete('imgVochDelco')" type="radio"
														value="rdE" name="rdGBN">&nbsp;오류&nbsp;
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<td class="DATA">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 20px"
											accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"><INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 20px"
											readOnly maxLength="100" size="16" name="txtSELECTAMT"></td>
								</TR>
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE height="15" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
														type="button" value="매출" name="btnTab1"><INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" value="매입" name="btnTab2"><INPUT class="BTNTAB" id="btnTab3" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" value="대행 매입" name="btnTab3"><INPUT class="BTNTAB" id="btnTab4" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" value="대행 매출" name="btnTab4"><INPUT class="BTNTAB" id="btnTab5" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
														type="button" value="매입수수료" name="btnTab5"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
														<TR>
															<td><IMG id="ImgvochCre" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																	height="20" alt="전표를 저장합니다." src="../../../images/ImgvochCre.gIF" border="0" name="ImgvochCre"></td>
															<td><IMG id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'"
																	height="20" alt="전표를 삭제합니다." src="../../../images/imgVochDel.gIF" border="0" name="imgVochDel"></td>
															<td><IMG id="ImgErrVochDel" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																	height="20" alt="오류전표 를 삭제합니다." src="../../../images/ImgErrVochDel.gIF" border="0"
																	name="ImgErrVochDel"></td>
															<td><IMG id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																	title="SAP에서 직접삭제하여 RMS에서 삭제할 수 없을때 RMS전표를 강제로 삭제한다." style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																	height="20" alt="전표를 강제로 삭제합니다." src="../../../images/imgVochDelco.gIF" border="0"
																	name="imgVochDelco"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%" />
								<TR>
									<TD>
										<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHDATA" onclick="vbscript:Call gCleanField(txtSUMM,'')" width="90"><select id="cmbSETTING" style="WIDTH: 90px"><OPTION value="" selected>선택</OPTION>
														<OPTION value="POSTINGDATE">전표일자</OPTION>
														<OPTION value="SUMM">적요</OPTION>
														<OPTION value="BA">사업영역</OPTION>
														<OPTION value="COSTCENTER">코스트센터</OPTION>
														<OPTION value="SEMU">세무코드</OPTION>
														<OPTION value="DEMANDDAY">지급기일</OPTION>
														<OPTION value="DUEDATE">입금기일</OPTION>
														<OPTION value="DEBTOR">차변계정</OPTION>
														<OPTION value="ACCOUNT">계정</OPTION>
														<OPTION value="PREPAYMENT">선수금구분</OPTION>
														<OPTION value="SUMMTEXT">본문TEXT</OPTION>
													</select></TD>
												<TD class="DATA"><INPUT class="INPUT_L" id="txtSUMM" title="적요적용" style="WIDTH: 368px; HEIGHT: 21px" size="56"
														name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
														title="적요를 일괄 적용합니다" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
														height="20" alt="적요를 일괄 적용합니다" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0"
														name="ImgSUMMApp">
													<DIV id="pnlFLAG" style="VISIBILITY: hidden; WIDTH: 250px; POSITION: absolute; HEIGHT: 24px"
														align="center" ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="rdDIV" title="분할" type="radio" CHECKED value="rdDIV" name="rdDIVGUBUN">&nbsp;분할&nbsp;&nbsp;&nbsp; 
														&nbsp; <INPUT id="rdSUM" title="합산" type="radio" value="rdSUM" name="rdDIVGUBUN">&nbsp;합산</DIV>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px" />
								<!--내용 및 그리드-->
								<TR vAlign="top" align="left">
									<!--내용-->
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 80%" vAlign="top" align="center">
										<DIV id="pnlTab_susu" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_SUSU" style="WIDTH: 100%; HEIGHT: 70%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="9260">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
											<OBJECT id="sprSht_SUSUDTL" style="WIDTH: 100%; HEIGHT: 30%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="3969">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
										<DIV id="pnlTab_out" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_OUT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="13229">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
										<DIV id="pnlTab_EOM" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_EOM" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="13229">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
										<DIV id="pnlTab_EOMSUSU" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_EOMSUSU" style="WIDTH: 100%; HEIGHT: 70%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="9260">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
											<OBJECT id="sprSht_EOMSUSUDTL" style="WIDTH: 100%; HEIGHT: 30%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="3969">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
										<DIV id="pnlTab_OUTSUSU" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_OUTSUSU" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27543">
												<PARAM NAME="_ExtentY" VALUE="13229">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
							</TBODY>
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%" />
				</TR>
			</TABLE>
		</FORM>
		<iframe id="frmSapCon" style="DISPLAY: block; WIDTH: 100%; HEIGHT: 300px" name="frmSapCon"
			src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx" /><!--style="DISPLAY: none"-->
	</body>
</HTML>
