<%@ Page CodeBehind="MDCMPRINTTRUTAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMPRINTTRUTAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>인쇄광고 위수탁 세금계산서 발행</title> 
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMPRINTTRUTAX , mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
CONST meTAB = 9
mstrGUBUN = ""
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
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "인쇄는 완료상태일때 가능합니다..","처리안내!"
		Exit Sub
	end if
	
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
			ReportName = "TRANSTAXNO_BLACK.rpt"
						
			Params = ""
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
			intRtn = mobjMDCMPRINTTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp삭제 끝
			
			ModuleDir = "MD"
			'공급자/공급받는자 보관용을 한장에 다보여주거나 공급받는자 보관용만 보여주는 구
			IF .chkPRINT.value THEN
				ReportName = "TRANSTAX_BLACK_NEW.rpt"
			ELSE
				ReportName = "TRANSTAX_BLACKONE_NEW.rpt"
			END IF
			
			
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
					
					vntDataTemp = mobjMDCMPRINTTRUTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			
			Params = strUSERID & ":" & "MD_TAX_TEMP"
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
			'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
			window.setTimeout "printSetTimeout", 20000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMPRINTTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	'gFlowWait meWAIT_ON
	with frmThis
		'mobjSCGLSpr.ExportMerge = true
		'mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	'gFlowWait meWAIT_OFF
End Sub

Sub ImgTaxCre_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "세금계산서 생성할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
	
	If frmThis.rdF.checked <> true then
		gErrorMsgBox "세금계산서생성은 미완료상태일때 가능합니다..","처리안내!"
		Exit Sub
	end if
	
	For i = 1 To frmThis.sprSht.MaxRows
		IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = 1 THEN
			chkcnt = chkcnt + 1
		END IF
	next
	if chkcnt = 0 then
		gErrorMsgBox "선택하신 자료가 없습니다.","저장안내!"
		exit sub
	end if
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "삭제는 완료상태일때 가능합니다..","처리안내!"
		Exit Sub
	end if
	
	For i = 1 To frmThis.sprSht.MaxRows
		IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = 1 THEN
			chkcnt = chkcnt + 1
		END IF
	next
	if chkcnt = 0 then
		gErrorMsgBox "선택하신 자료가 없습니다.","삭제안내!"
		exit sub
	end if
	
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
		If .rdT.checked = True OR .rdA.checked = TRUE Then
			gErrorMsgBox "청구일 적용은 미완료상태 에서 적용됩니다.","처리안내!"
			Exit Sub
		End If
		
		if .txtDEMANDDAY.value = "" then
			gErrorMsgBox "적용할 청구일을 입력하시오.","처리안내!"
			Exit Sub
		end if
		
		strDEMANDDAY = .txtDEMANDDAY.value
		strTAXYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		intRtn = gYesNoMsgbox("선택된 항목의 청구일을 변경 하시겠습니까?","변경 확인")
		IF intRtn <> vbYes then exit Sub
			
		If .cmbGUBUN.value = "taxdiv" Then
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
					mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
					mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
					mobjSCGLSpr.setTextBinding .sprSht,"SUMM",intCnt,"인쇄광고비 - (" & mobjSCGLSpr.GetTextBinding(.sprSht,"MEDNAME",intCnt) & ")"
				End If
			Next
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
					mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
					mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
					mobjSCGLSpr.setTextBinding .sprSht,"SUMM",intCnt,"인쇄광고비"
				End If
			Next
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtTRANSYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "PRINT") 
		vntRet = gShowModalWindow("../MDCO/MDCMTAXCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(1,0) and .txtCLIENTNAME1.value = vntRet(2,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code값 저장
			.txtCLIENTNAME1.value = vntRet(2,0)       ' 코드명 표시
			if .txtTRANSYEARMON1.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			End if
		end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetTAXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "PRINT")
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					if .txtTRANSYEARMON1.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					End if
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
' 사업부코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'실제 데이터List 가져오기
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams 
	
	with frmThis
		vntInParams = array(trim(.txtTRANSYEARMON1.value), trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "PRINT")
		
		vntRet = gShowModalWindow("../MDCO/MDCMTAXTIMPOP.aspx",vntInParams , 413,455)
		if isArray(vntRet) then
			.txtTRANSYEARMON1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtTIMNAME1.value = trim(vntRet(1,0))  ' Code값 저장
			.txtTIMCODE1.value = trim(vntRet(2,0))  ' 코드명 표시
			.txtCLIENTNAME1.value = trim(vntRet(3,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			if .txtTRANSYEARMON1.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			End if
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtTIMNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetTAXTIMNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtTRANSYEARMON1.value), _
											  trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
											  trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "PRINT")
			
			if not gDoErrorRtn ("GetTAXTIMNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON1.value = trim(vntData(0,1))  ' Code값 저장
					.txtTIMNAME1.value = trim(vntData(1,1))  ' Code값 저장
					.txtTIMCODE1.value = trim(vntData(2,1))  ' 코드명 표시
					.txtCLIENTNAME1.value = trim(vntData(3,1))
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					if .txtTRANSYEARMON1.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					End if
				Else
					Call TIMCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' 달력
'-----------------------------------------------------------------------------------------
Sub imgDEMANDDAY_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgDEMANDDAY,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'청구일
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

Sub txtTRANSYEARMON1_onblur
	With frmThis
	If .txtTRANSYEARMON1.value <> "" AND Len(.txtTRANSYEARMON1.value) = 6 Then DateClean
	End With
End Sub

Sub imgFrom_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgTO_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTO_onchange
	gSetChange
End Sub

Sub cmbGUBUN_onchange
	with frmThis
		If .cmbGUBUN.value = "taxdiv" Then
			pnlFLAG.style.display = "none"
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			pnlFLAG.style.display = "inline"
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
		End If
	End with
End Sub

'****************************************************************************************
' 조회필드 체인지 이벤트
'****************************************************************************************

Sub cmbMED_FLAG_onChange ()
	if frmThis.txtTRANSYEARMON1.value <> "" then
		gFlowWait meWAIT_ON
			SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE0_onClick ()
	if frmThis.txtTRANSYEARMON1.value <> "" then
		gFlowWait meWAIT_ON
			SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE1_onClick ()
	if frmThis.txtTRANSYEARMON1.value <> "" then
		gFlowWait meWAIT_ON
			SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE2_onClick ()
	if frmThis.txtTRANSYEARMON1.value <> "" then
		gFlowWait meWAIT_ON
			SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE3_onClick ()
	if frmThis.txtTRANSYEARMON1.value <> "" then
		gFlowWait meWAIT_ON
			SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub


'완료체크
Sub rdT_onclick
	SelectRtn
End Sub
'미완료체크
Sub rdF_onclick
	SelectRtn
End Sub

'전체체크
Sub rdA_onclick
	SelectRtn
End Sub

Sub rdMED_onclick
	SelectRtn	
End Sub

Sub rdREAL_onclick
	SelectRtn
End Sub

Sub txtTRANSYEARMON1_onkeydown
	'or window.event.keyCode = meTAB 탭일때는 아님 엔터일때만 조회
	If window.event.keyCode = meEnter Then
		txtTRANSYEARMON1_onblur
		SELECTRTN
		frmThis.txtCLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------
' SpreadSheet 이벤트	
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), , , "", , , , , mstrCheck
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
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If .rdT.checked = True Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row)) '<< 받아오는경우
				gShowModalWindow "../MDPT/MDCMPRINTTRUTAXDTL.aspx",vntInParams , 898,680
				'SelectRtn
			End IF
		end if	
	end with
end sub

Sub sprSht_Keyup(KeyCode, Shift)
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

	With frmThis
		If mstrGUBUN = "TAX" Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
					strCOLUMN = "SUMAMT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) Then
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
		else
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
					strCOLUMN = "SUMAMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
					strCOLUMN = "COMMISSION"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) Then
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
		end if
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
		If mstrGUBUN = "TAX" Then
			If .sprSht.MaxRows >0 Then
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
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
							.txtSELECTAMT.value = 0
							exit Sub
						End If
						
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
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
			If .sprSht.MaxRows >0 Then
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
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
							.txtSELECTAMT.value = 0
							exit Sub
						End If
						
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
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

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDCMPRINTTRUTAX	= gCreateRemoteObject("cMDPT.ccMDPTPRINTTRUTAX")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCMPRINTTRUTAX = Nothing
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
		.txtTRANSYEARMON1.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet초기화
		DateClean
		.sprSht.MaxRows = 0
		.chkVOCH_TYPE0.checked = true
		.chkVOCH_TYPE1.checked = true
		.chkVOCH_TYPE2.checked = true
		.chkVOCH_TYPE3.checked = true
		
		CALL Grid_Setting ("TRANS")
		
		.txtCLIENTNAME1.focus
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
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

Sub Grid_Setting (strGUBUN)
	With frmThis
		'Sheet 기본Color 지정
		.sprSht.MaxRows = 0
		.sprSht.style.visibility = "hidden"
		Call Grid_init()
		gSetSheetDefaultColor() 
		If strGUBUN = "TAX" Then
			'Sheet 기본Color 지정
			 gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 28, 0, 3, 3
			mobjSCGLSpr.SpreadDataField .sprSht, "VOCH_TYPE_OLD | VOCH_TYPE | CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | MEDFLAG"
			mobjSCGLSpr.SetHeader .sprSht,		           "구분|발행|선택|관리번호|청구년월|광고주|광고주사업자번호|팀|브랜드|매체사|매체사사업자번호|매체명|금액|부가세|합계금액|적요|발행일|담당부서|광고주대표자명|광고주주소1|광고주주소2|매체사대표자명|매체사주소1|매체사주소2|전표번호|세금계산서년월|세금계산서번호|매체구분"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "    	   7|   7|	 4|      10|       8|    15|	          13|10|	12|	   15|              13|     9|  10|    10|	    11|  20|     9| 	  8|             0|          0|          0|             0|          0|          0|      10|             0|             0|     8"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXMANAGE | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | SUMM | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | MEDFLAG", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "VOCH_TYPE_OLD | VOCH_TYPE | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | MEDFLAG "
			mobjSCGLSpr.ColHidden .sprSht, "CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE | CLIENTBISNO | REAL_MED_BISNO | VOCHNO",-1,-1,2,2,False
			mstrGUBUN = "TAX"
		Else
			'Sheet 기본Color 지정
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 48, 0, 1, 2
			mobjSCGLSpr.SpreadDataField .sprSht, "VOCH_TYPE_OLD | VOCH_TYPE | CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | COMMI_RATE | COMMISSION | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MATTERNAME | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | TIMCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | SUBSEQ | MATTERCODE | MEDFLAG | RANKTRANS | GENERALAMT | OUTLISTAMT"
			mobjSCGLSpr.SetHeader .sprSht,		           "구분|발행|선택|거래번호|청구년월|광고주|광고주사업자번호|팀|브랜드|매체사|매체사사업자번호|매체명|금액|부가세|합계금액|적요|발행일|수수료율|수수료|담당부서|광고주대표자명|광고주주소1|광고주주소2|매체사대표자명|매체사주소1|매체사주소2|소재명|전표번호|세금계산서년월|세금계산서번호|거래명세서년월|거래명세서번호|거래명세서순번|신탁년월|신탁순번|광고주코드|팀코드|광고주AC코드|매체사코드|매체사AC코드|매체코드|부서코드|브랜드코드|소재코드|매체구분|순위"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "     	   7|   7|	 4|      10|       8|    15|	          13|10|	12|	   15|              13|     9|  10|    10|	    11|  20|     9|       5|    10| 	  8|             0|          0|          0|             0|          0|          0|    10|      10|             0|             0|             0|             0|             5|       0|       0|         0|     0|           0|         0|           0|       0|       0|         0|       0|       0|  0"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | PRINTDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMT | COMMISSION | SEQ", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXMANAGE | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | SUMM | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MATTERNAME | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | SUBSEQ | MATTERCODE | MEDFLAG | RANKTRANS ", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | COMMI_RATE | COMMISSION | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MATTERNAME | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | SUBSEQ | MATTERCODE | MEDFLAG | RANKTRANS "
			mobjSCGLSpr.ColHidden .sprSht, "CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MATTERNAME | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | TIMCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | SUBSEQ | MATTERCODE | MEDFLAG | RANKTRANS | GENERALAMT | OUTLISTAMT", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE | CLIENTBISNO | REAL_MED_BISNO | VOCHNO",-1,-1,2,2,False
			mstrGUBUN = "TRANS"
		End If
		Get_COMBO_VALUE
		.sprSht.style.visibility = "visible"
		
	End With
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, ""
		mobjSCGLSpr.SetHeader .sprSht,		 ""
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 그리드 콤보박스 설정
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDCMPRINTTRUTAX.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE_OLD | VOCH_TYPE",,,vntData,,60 
			mobjSCGLSpr.TypeComboBox = True 
   		End If    
   	End With
End Sub


'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCLIENTCODE
	Dim strTIMCODE
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim strGUBUN
   	Dim strGROUPGBN
   	Dim strMED_FLAG
   	Dim strVOCH_TYPE_TEMP
   
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
		
		IF .rdT.checked = TRUE THEN
			CALL Grid_Setting ("TAX")
		ELSE
			CALL Grid_Setting ("TRANS")
		END IF
		
		strYEARMON		= .txtTRANSYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strTIMCODE		= .txtTIMCODE1.value
		strFROM			=  MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO			=  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)		
		strMED_FLAG		= .cmbMED_FLAG.value
		strGUBUN = .cmbGUBUN.value
		
		IF strGUBUN = "taxgroup" then
			IF .rdMED.checked = TRUE THEN
				strGROUPGBN = "MED"
			ELSEIF .rdREAL.checked = TRUE THEN
				strGROUPGBN = "REAL"
			END IF 
		else
			strGROUPGBN = ""
		end if
		
		strVOCH_TYPE_TEMP = ""
		
		IF .chkVOCH_TYPE0.checked THEN
			strVOCH_TYPE_TEMP = "0"
		END IF
		
		IF .chkVOCH_TYPE1.checked THEN
			IF strVOCH_TYPE_TEMP = "" THEN
				strVOCH_TYPE_TEMP = "1"
			ELSE
				strVOCH_TYPE_TEMP = strVOCH_TYPE_TEMP & ",1"
			END IF
		END IF
		
		IF .chkVOCH_TYPE2.checked THEN
			IF strVOCH_TYPE_TEMP = "" THEN
				strVOCH_TYPE_TEMP = "2"
			ELSE
				strVOCH_TYPE_TEMP = strVOCH_TYPE_TEMP & ",2"
			END IF
		END IF
		
		IF .chkVOCH_TYPE3.checked THEN
			IF strVOCH_TYPE_TEMP = "" THEN
				strVOCH_TYPE_TEMP = "3"
			ELSE
				strVOCH_TYPE_TEMP = strVOCH_TYPE_TEMP & ",3"
			END IF
		END IF
		
		'세금계산서 완료조회
		If .rdT.checked = True Then
			vntData = mobjMDCMPRINTTRUTAX.Get_PRINT_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strFROM, strTO, strMED_FLAG, strVOCH_TYPE_TEMP)
			If not gDoErrorRtn ("Get_PRINT_TAX") then
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'초기 상태로 설정
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SUMM",-1,-1,false	
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				Layout_change
				AMT_SUM
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				if .sprSht.MaxRows = 0 then
					.imgDelete.style.display = "none"
				else
					.imgDelete.style.display = "inline"
					
					For i = 1 to .sprSht.MaxRows
						IF mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",i) = "3" THEN
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HCCFFFF, &H000000,False
							'mobjSCGLSpr.SetCellsLock2 .sprSht,TRUE,i,1,-1,true
						END IF 
					next
				end if
			End If
		'미완료 거래명세서 디테일 조회
		ElseIf .rdF.checked = True Then			
			vntData = mobjMDCMPRINTTRUTAX.Get_PRINT_TAXBUILD(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strCLIENTCODE,strTIMCODE, strFROM, strTO, strGUBUN, strMED_FLAG, strVOCH_TYPE_TEMP, strGROUPGBN)
			If not gDoErrorRtn ("Get_PRINT_TAXBUILD") then
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'초기 상태로 설정
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"SUMM",-1,-1,false	
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				'Layout_change
				AMT_SUM
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				if .sprSht.MaxRows = 0 then
					.ImgTaxCre.style.display = "none"
				else
					.ImgTaxCre.style.display = "inline"
					
					For i = 1 to .sprSht.MaxRows
						IF mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",i) = "3" THEN
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HCCFFFF, &H000000,False
							'mobjSCGLSpr.SetCellsLock2 .sprSht,TRUE,i,1,-1,true
						END IF 
					next
				end if
			End If
		ElseIf .rdA.checked = True Then
			vntData = mobjMDCMPRINTTRUTAX.Get_PRINT_TAXALL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strCLIENTCODE,strTIMCODE, strFROM, strTO, strGUBUN, strMED_FLAG, strVOCH_TYPE_TEMP)
			If not gDoErrorRtn ("Get_PRINT_TAXALL") then
				'초기 상태로 설정
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SUMM",-1,-1,false	
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'Layout_change
				AMT_SUM
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				.ImgTaxCre.style.display = "none"
				.imgDelete.style.display = "none"
			End If
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
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

Sub Layout_change ()
	Dim intCnt
	with frmThis
		For intCnt = 1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE_OLD",intCnt) = "" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HAAF290, &H000000,False 
			End If 
		Next 
	End With
End Sub

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
		
		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "상세항목 이 없습니다.",""
   			Exit Sub
   		End If
   		
		intRtn2 = gYesNoMsgbox("청구일을 확인하셨습니까?","확인")
		IF intRtn2 <> vbYes then exit Sub
		
		'체크 없을 경우 저장 안되도록
		chkcnt = 0
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
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		Next
		
		
		if chkcnt = 0 then
			gErrorMsgBox "세금계산서를 생성할 데이터를 체크 하십시오","저장안내!"
			exit sub
		end if
		 '저장플레그 설정
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
   		
		'if DataValidation =false then exit sub
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"VOCH_TYPE_OLD | VOCH_TYPE | CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | SUBSEQNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | COMMI_RATE | COMMISSION | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MATTERNAME | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | TIMCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | SUBSEQ | MATTERCODE | MEDFLAG | RANKTRANS | GENERALAMT | OUTLISTAMT ")
		
		'마스터 데이터를 가져 온다.
		'처리 업무객체 호출
		intTAXNO = 0
		If .cmbGUBUN.value = "taxdiv" Then
		intRtn = mobjMDCMPRINTTRUTAX.ProcessRtn_Div(gstrConfigXml,vntData, intTAXNO)
		Else
			If Not TaxGroup(strVALIDATION) Then 
				gErrorMsgBox strVALIDATION & vbCrlf & "단위별 [광고주,매체사] [청구일,작성일,적요,발행] 는 동일 하여야 합니다.","저장안내!"
				Exit Sub
			Else
				'최대값
				intColFlag = 0
				For intMaxCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intMaxCnt) = 1 Then
						bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intMaxCnt))
						IF intColFlag < bsdiv THEN
							intColFlag = bsdiv
						END IF
					End IF
				Next
				'맥스값만 추가하여 보내기
				intRtn = mobjMDCMPRINTTRUTAX.ProcessRtn_Group(gstrConfigXml,vntData, intTAXNO,intColFlag)
			End IF
		End If

		If not gDoErrorRtn ("ProcessRtn") Then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "세금계산서가 생성되었습니다.","저장안내!"
			.rdT.checked = True
			selectRtn
   		End If
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
	Dim strVOCH_TYPE
	
	TaxGroup = False
	with frmThis
		strStartRank = "0"
		strCLIENTCODE = ""
		strREAL_MED_CODE = ""
		strDEMANDDAY = ""
		strPRINTDAY = ""
		strSUMM = ""
		strVOCH_TYPE = ""
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				
				If strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Then
					If strCLIENTCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBISNO",intCnt) Then
						Exit Function
					End If
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
					If strVOCH_TYPE <> mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",intCnt) Then
						strVALIDATION = "발행확인 거래명세서번호" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " 번"
						Exit Function
					End If 
				End If
				
				strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBISNO",intCnt)
				strREAL_MED_CODE = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt)
				strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
				strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
				strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt)
				strVOCH_TYPE = mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",intCnt)
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
	Dim strVOCH_TYPE
	
	with frmThis
		strDESCRIPTION = ""
		For intCnt2 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
				If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt2) <> "" THEN
					gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " 에 대하여" &vbcrlf & "전표가 존재하는 내역은 삭제가 되지 않습니다.","삭제안내!"
					Exit Sub
				ELSE
					if NOT VOCHNO_CHECKED (mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2), mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2)) then
						gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " 에 대하여" &vbcrlf & "전표처리 진행중인 내역은 삭제가 되지 않습니다.","삭제안내!"
						Exit Sub
					END IF
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
				strVOCH_TYPE = mobjSCGLSpr.GetTextBinding(.sprSht,"VOCH_TYPE",i)
				intRtn = mobjMDCMPRINTTRUTAX.DeleteRtn_TruTax(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCH_TYPE)
				
				IF not gDoErrorRtn ("DeleteRtn_TruTax") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"삭제안내!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn_TruTax") then
			gWriteText lblstatus, intCnt & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub

Function VOCHNO_CHECKED (ByRef strTAXYEARMON, ByRef strTAXNO)
	Dim vntData
	Dim intCnt
	Dim strCOUNT
	'on error resume next

	'초기화
	VOCHNO_CHECKED = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCOGET.TRUVOCHNO_CHECKED(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON,strTAXNO)
	
	IF mlngRowCnt >0 THEN
		VOCHNO_CHECKED = false
	ELSE
		VOCHNO_CHECKED = TRUE	
	End IF
End Function

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="144" background="../../../images/back_p.gIF"
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
											<td class="TITLE">위수탁 세금계산서 관리</td>
										</tr>
									</table>
								</td>
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
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, '')"
												width="50">년월</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtTRANSYEARMON1" title="거래명세년월" style="WIDTH: 89px; HEIGHT: 22px"
													accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="50">광고주
											</TD>
											<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="코드명" style="WIDTH: 143px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="14" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
												width="50">팀
											</TD>
											<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtTIMNAME1" title="팀명" style="WIDTH: 143px; HEIGHT: 22px" type="text"
													maxLength="100" size="14" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="팀코드" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" size="6" name="txtTIMCODE1">
											</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
												width="50">청구일
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="2" name="txtFROM">&nbsp;<IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													 align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="6" name="txtTO">&nbsp;<IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													 align="absMiddle" border="0" name="imgTo">
											</TD>
											<TD class="SEARCHDATA" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" width="50">매체구분</TD>
											<TD class="SEARCHDATA" width="90"><SELECT id="cmbMED_FLAG" title="매체구분" style="WIDTH: 89px" name="cmbMED_FLAG">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="B">신문</OPTION>
													<OPTION value="C">잡지</OPTION>
												</SELECT></TD>
											<TD class="SEARCHLABEL">발행
											</TD>
											<TD class="SEARCHDATA"><INPUT id="rdT" title="완료내역조회" type="radio" value="rdT" name="rdGBN">
												&nbsp;완료&nbsp; <INPUT id="rdF" title="미완료 내역조회" type="radio" CHECKED value="rdF" name="rdGBN">
												&nbsp;미완료&nbsp;&nbsp;<INPUT id="rdA" title="전체 내역조회" type="radio" value="rdA" name="rdGBN">&nbsp;전체</TD>
											<TD class="SEARCHLABEL">구분
											</TD>
											<TD class="SEARCHDATA" colSpan="4">&nbsp;위수탁 <INPUT id="chkVOCH_TYPE0" title="위수탁" type="checkbox" name="chkVOCH_TYPE0">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
												협찬 <INPUT id="chkVOCH_TYPE1" title="협찬" type="checkbox" name="chkVOCH_TYPE1">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
												일반 <INPUT id="chkVOCH_TYPE2" title="일반" type="checkbox" name="chkVOCH_TYPE2">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
												AOR  <INPUT id="chkVOCH_TYPE3" title="AOR" type="checkbox" name="chkVOCH_TYPE3">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="absmiddle" align="center">
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 20px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD height="20" colspan="4">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 20px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 20px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
										</TR>
										<TR>
											<TD height="4" colspan="4"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="WIDTH: 67px">청구일적용</TD>
											<TD class="DATA" style="WIDTH: 400px"><INPUT class="INPUT" id="txtDEMANDDAY" title="청구일자" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="date" type="text" maxLength="10" size="14" name="txtDEMANDDAY">&nbsp;<IMG id="imgDEMANDDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle"
													border="0" name="imgDEMANDDAY">&nbsp;<IMG id="btnCOMMISSION" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" height="20" alt="해당 정산일로 정산일이 없는 상세항목을 Setting 합니다"
													src="../../../images/imgApp.gIF" width="54" align="absMiddle" border="0" name="btnCOMMISSION">
													<DIV id="pnlFLAG" style="DISPLAY: none; WIDTH: 170px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdMED" title="매체명합산" type="radio" CHECKED value="MED" name="rdGROUP">&nbsp;매체명&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdREAL" title="매체사합산" type="radio" value="REAL" name="rdGROUP">&nbsp;매체사</DIV>
											</TD>
											<TD class="DATA" style="WIDTH: 250px"><SELECT id="cmbGUBUN" title="매체구분" style="WIDTH: 80px" name="cmbGUBUN">
													<OPTION value="taxdiv" selected>분할발행</OPTION>
													<OPTION value="taxgroup">합산발행</OPTION>
												</SELECT>&nbsp;<SELECT id="chkPRINT" title="출력물구분" style="WIDTH: 80px" name="chkPRINT">
													<OPTION value="1" selected>양자용</OPTION>
													<OPTION value="0">공급받는자용</OPTION>
												</SELECT>&nbsp;<SELECT id="cmbFLAG" title="영수/청구구분" style="WIDTH: 80px" name="cmbFLAG">
													<OPTION value="receipt" selected>청구</OPTION>
													<OPTION value="demand">영수</OPTION>
												</SELECT></TD>
											<TD class="DATA_RIGHT" vAlign="middle" align="right" height="20">
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
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31856">
										<PARAM NAME="_ExtentY" VALUE="13547">
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
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></form>
	</body>
</HTML>
