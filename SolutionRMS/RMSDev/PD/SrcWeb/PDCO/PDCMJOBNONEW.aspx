<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBNONEW.aspx.vb" Inherits="PD.PDCMJOBNONEW" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>제작관리번호 등록</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/제작관리번호 등록 화면
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBNO.aspx
'기      능 : 제작관리번호 C/D/U/R
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/19 By Kim Tae Ho
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt		'조회시 로우,컬럼 갯수를 반환
Dim mobjPDCOJOBNO, mobjPDCOACTUALRATE, mobjSCCOGET, mobjPDCOGET '모듈(JOBNO-CRUD, ACTUALRATE-CRUD, 전체공통, 제작공통)
Dim mstrHIDDEN					'입력필드의 숨기기
Dim mstrSTATUS					'상세내역 인지 신규투입인지 구분
Dim mstrJOBNO
Dim mstrCheck

Const meTab = 9
mstrHIDDEN = 0
mstrJOBNO = ""
mstrCheck = true

'상위입력 필드 숨기기
Sub Set_SELECTTBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblSelectBody").style.display = "inline"
		Else
			document.getElementById("tblSelectBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

' 하위입력 필드 숨기기
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='입력필드 숨김' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblBody2").style.display = "inline"
			document.getElementById("spacebar1").style.display = "inline"
			document.getElementById("spacebar2").style.display = "inline"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='입력필드 노출' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblBody2").style.display = "none"
			document.getElementById("spacebar1").style.display = "none"
			document.getElementById("spacebar2").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'=========================================================================================
' 명령버튼
'=========================================================================================
'조회버튼
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'저장버튼
Sub imgSave_onclick ()
	If frmThis.cmbENDFLAG.value <> "PF01" and  frmThis.cmbENDFLAG.value <> "PF02" Then
		gErrorMsgBox "진행상태가 의뢰 및 진행 이 아닌건은 수정될수 없습니다.","저장안내"
		SelectRtn
		exit Sub
	End If
			
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'닫기버튼
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'부서실적 추가버튼
sub imgAddRow_sprSht_JOBNODEPT_onclick ()
	With frmThis
		call sprSht_JOBNODEPT_Keydown(meINS_ROW, 0)
	End With 
end sub

Sub sprSht_JOBNODEPT_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	with  frmThis
		if KeyCode <> meINS_ROW then exit sub
		
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_JOBNODEPT, cint(KeyCode), cint(Shift), -1, 1)
		'mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"CHK",.sprSht_JOBNODEPT.ActiveRow, "1"
	end with
End Sub

'부서실적 삭제버튼
Sub imgDelete_sprSht_JOBNODEPT_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn_DTL_JOBNODEPT
	gFlowWait meWAIT_OFF
End Sub

Sub imgEndChange_onclick
	Dim vntData2
	Dim strCODE
	Dim intRtnSave
	Dim intRtn
	Dim intCnt
	Dim intCode
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCODE = .txtJOBNO.value

		If strCODE = "" Then
			gErrorMsgBox "우선 JOBNO 를 조회하십시오.","처리안내"
			Exit Sub
		End If

		If .cmbENDFLAG.value <> "PF02" Then
			gErrorMsgBox "완료구분 변경은 '진행' 일경우만 가능합니다.","처리안내"
			Exit Sub
		End If
		
		vntData2 = mobjPDCOJOBNO.GetJOBNOSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
		
		If mlngRowCnt = 0 Then
			intRtnSave = gYesNoMsgbox("완료구분을 '의뢰'상태 로 변경하시겠습니까?","처리안내")
			IF intRtnSave <> vbYes then exit Sub
			
			intRtn = mobjPDCOJOBNO.ProcessRtn_ENDFLAG(gstrConfigXml,strCODE)
			
			if not gDoErrorRtn ("ProcessRtn_ENDFLAG") then
				gErrorMsgBox "JOBNO [" & strCODE & " ]완료구분이 '의뢰' 상태로 변경되었습니다.","처리안내" 
				SelectRtn
			end if
		Else
			gErrorMsgBox "해당 JOBNO 의 외주정산내역을 확인하십시오","처리안내"
		End If
	End with
End Sub

'-----------------------------
' 담당부서 조회 
'-----------------------------
Sub ImgDEPTCD_onclick
	Call DEPT_POP()
End Sub

Sub DEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC 코드/명,optional(현재사용여부,사용검사일,추가조회 필드,Key Like여부)
		vntInParams = array(trim(.txtDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = trim(vntRet(0,0))	'Code값 저장
			.txtDEPTNAME.value = trim(vntRet(1,0))	'코드명 표시
			
			.txtEMPNAME.focus()
			gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub

Sub txtDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTNAME.value)
			' mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = trim(vntData(0,0))
					.txtDEPTNAME.value = trim(vntData(1,0))
					
					.txtEMPNAME.focus()
				Else
					Call DEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'실제 데이터List 가져오기
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtDEPTCD.value), trim(.txtDEPTNAME.value), trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtDEPTCD.value = trim(vntRet(2,0))  ' Code값 저장
			.txtDEPTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			
			
			
			.txtEXCLIENTNAME.focus()
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
			gSetChangeFlag .txtDEPTCD
			gSetChangeFlag .txtDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A",.txtDEPTCD.value,.txtDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					.txtDEPTCD.value = trim(vntData(2,1))
					.txtDEPTNAME.value = trim(vntData(3,1))
					
					
					.txtEXCLIENTNAME.focus()
					'.txtMEMO.focus()
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------
' 제작부서 조회 
'-----------------------------
Sub ImgCREDEPTCD_onclick
	Call CREDEPT_POP()
End Sub

Sub CREDEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC 코드/명,optional(현재사용여부,사용검사일,추가조회 필드,Key Like여부)
		vntInParams = array(trim(.txtCREDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCREDEPTCD.value = trim(vntRet(0,0))	'Code값 저장
			.txtCREDEPTNAME.value = trim(vntRet(1,0))	'코드명 표시
		
			.txtCREEMPNAME.focus()
			gSetChangeFlag .txtCREDEPTCD
		end if
	end with
End Sub

Sub txtCREDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREDEPTNAME.value)
			' mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = trim(vntData(0,0))
					.txtCREDEPTNAME.value = trim(vntData(1,0))
					
					.txtCREEMPNAME.focus()
				Else
					Call CREDEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 제작사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
Sub ImgCREEMPNO_onclick
	Call CREEMP_POP()
End Sub

'실제 데이터List 가져오기
Sub CREEMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCREDEPTCD.value), trim(.txtCREDEPTNAME.value), trim(.txtCREEMPNO.value), trim(.txtCREEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCREEMPNO.value = vntRet(0,0) and .txtCREEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			
			.txtCREDEPTCD.value = trim(vntRet(2,0))  ' Code값 저장
			.txtCREDEPTNAME.value = trim(vntRet(3,0))  ' 코드명 표시
			.txtCREEMPNO.value = trim(vntRet(0,0))
			.txtCREEMPNAME.value = trim(vntRet(1,0))
		
			.txtBUDGETAMT.focus()					' 포커스 이동
			gSetChangeFlag .txtCREEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtCREEMPNAME
			gSetChangeFlag .txtCREDEPTCD
			gSetChangeFlag .txtCREDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCREEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		
		On error resume next
		
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREEMPNO.value, .txtCREEMPNAME.value,"A",.txtCREDEPTCD.value,.txtCREDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCREEMPNO.value = trim(vntData(0,1))
					.txtCREEMPNAME.value = trim(vntData(1,1))
					.txtCREDEPTCD.value = trim(vntData(2,1))
					.txtCREDEPTNAME.value = trim(vntData(3,1))
					
					.txtBUDGETAMT.focus()
					gSetChangeFlag .txtCREEMPNO
				Else
					Call CREEMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 대행사 코드팝업 버튼
'-----------------------------------------------------------------------------------------
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value), "") '<< 받아오는경우
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEXCLIENTCODE.value = trim(vntRet(1,0))  ' Code값 저장
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))  ' 코드명 표시
			
			.txtBIGO.focus() 
			gSetChangeFlag .txtEXCLIENTCODE	
     	End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
   		
		On error resume next
		
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value), "")
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))
					.txtEXCLIENTNAME.value = trim(vntData(2,1))
					
					.txtBIGO.focus()
					gSetChangeFlag .txtEXCLIENTCODE
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'------------------------------------------
' 달력 이벤트
'------------------------------------------
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtHOPEENDDAY,frmThis.imgCalEndar,"txtHOPEENDDAY_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarREQ_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtREQDAY,frmThis.imgCalEndar,"txtREQDAY_onchange()"
		gSetChange
	end with
End Sub

'------------------------------------------
' 데이터 입력시 SHEET BINDING onchange EVENT
'------------------------------------------
'생략부분
Sub txtJOBNAME_onchange
	gSetChange
End Sub
Sub txtJOBNO_onchange
	gSetChange
End Sub
Sub cmbJOBGUBN_onchange
	SUBCOMBO_TYPE
	gSetChange
End Sub
Sub txtDEPTNAME_onchange
	gSetChange
End Sub
Sub txtDEPTCD_onchange
	gSetChange
End Sub
Sub txtREQDAY_onchange
	gSetChange
End Sub
Sub cmbCREPART_onchange
	gSetChange
End Sub
Sub txtEMPNAME_onchange
	gSetChange
End Sub
Sub txtEMPNO_onchange
	gSetChange
End Sub
Sub txtHOPEENDDAY_onchange
	gSetChange
End Sub
'여기부터
Sub cmbCREGUBN_onchange
	gSetChange
End Sub
'cmbJOBBASE
Sub cmbJOBBASE_onchange
	gSetChange
End Sub
'txtCREDEPTNAME
Sub txtCREDEPTNAME_onchange
	gSetChange
End Sub
Sub txtCREDEPTCD_onchange
	gSetChange
End Sub
Sub cmbENDFLAG_onchange
	gSetChange
End Sub
'txtCREEMPNAME
Sub txtCREEMPNAME_onchange
	gSetChange
End Sub
Sub txtCREEMPNO_onchange
	gSetChange
End Sub
Sub txtBUDGETAMT_onchange
	gSetChange
End Sub
'txtBIGO
Sub txtBIGO_onchange
	gSetChange
End Sub
'PROJECT,JOBNO 선택


Sub txtEXCLIENTNAME_onchange
	gSetChange
End Sub

Sub txtEXCLIENTCODE_onchange
	gSetChange
End Sub

'onblur, onfocus 이벤트
Sub txtBUDGETAMT_onfocus
	with frmThis
		.txtBUDGETAMT.value = Replace(.txtBUDGETAMT.value,",","")
	end with
End Sub
Sub txtBUDGETAMT_onblur
	with frmThis
		call gFormatNumber(.txtBUDGETAMT,0,true)
	end with
End Sub

'--------------------------------------------------
'쉬트 이벤트
'--------------------------------------------------
Sub sprSht_JOBNODEPT_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_JOBNODEPT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_JOBNODEPT.MaxRows
				sprSht_JOBNODEPT_Change 1, intcnt
			next
		end if
	end with
End Sub

sub sprSht_JOBNODEPT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_JOBNODEPT, ""
		end if
	end with
end sub

Sub sprSht_JOBNODEPT_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim strDeptCodeName
	
	with frmThis
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"DEPTNAME") Then
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTNAME",Row)
				
			If strCode = "" AND strCodeName <> "" Then		
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
					
				vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT,Col,.sprSht_JOBNODEPT.ActiveRow			
						
						.txtJOBNAME.focus
						.sprSht_JOBNODEPT.focus
					Else
						mobjSCGLSpr_ClickProc "sprSht_JOBNODEPT", Col, .sprSht_JOBNODEPT.ActiveRow
						
						.txtJOBNAME.focus
						.sprSht_JOBNODEPT.focus 
					End If
   				End If
   			End If
	   	Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"EMPNAME") Then
			strCode = ""
			strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTNAME",.sprSht_JOBNODEPT.ActiveRow)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"EMPNAME",.sprSht_JOBNODEPT.ActiveRow)
			
			vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntData(2,1)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntData(3,1)
				
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT,Col,frmThis.sprSht_JOBNODEPT.ActiveRow
			Else
				mobjSCGLSpr_ClickProc "sprSht_JOBNODEPT", Col, .sprSht_JOBNODEPT.ActiveRow
			End If
			
			.txtJOBNAME.focus
			.sprSht_JOBNODEPT.Focus	   			
   		End If
	End with
	mobjSCGLSpr.CellChanged frmThis.sprSht_JOBNODEPT, Col, Row
End Sub

'--------------------------------------------------
'쉬트 팝업버튼클릭후 프로세스
'--------------------------------------------------
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	Dim vntRet, vntInParams
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"DEPTNAME") Then
	
			vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
			End IF
			
			.txtJOBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_JOBNODEPT.Focus	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"EMPNAME") Then
			
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"EMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
			End IF
			
			.txtJOBNAME.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht_JOBNODEPT.Focus	
		End if
	End With
End Sub

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_JOBNODEPT_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN") Then
			vntInParams = array("","","")
			vntRet = gShowModalWindow("../PDCO/PDCMDEPTPOP.aspx",vntInParams , 413,440)
				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row
				
				.txtJOBNAME.focus()
				.sprSht_JOBNODEPT.focus 
				
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_JOBNODEPT,"BTN2") Then
			vntInParams = array("","","","") '<< 받아오는경우
			vntRet = gShowModalWindow("../PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)
					
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",Row, vntRet(2,0)			
				mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_JOBNODEPT, Col,Row		
								
				.txtJOBNAME.focus()
				.sprSht_JOBNODEPT.focus 
				
				mobjSCGLSpr.ActiveCell .sprSht_JOBNODEPT, Col+2,Row
			End IF
		end if
		.sprSht_JOBNODEPT.focus 
	End with
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
Sub InitPage()
	gClearAllObject frmThis
	
	Dim vntInParam
	Dim intNo,i
	
	'서버업무객체 생성	
	set mobjPDCOJOBNO		= gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjPDCOACTUALRATE	= gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		
		'기본값 설정
		for i = 0 to intNo
			Select case i
				case 0: mstrSTATUS = vntInParam(i)
			End Select
		Next
		
		'신규일 경우 프로젝트 의 내역을 받아온다.
		for i = 0 to intNo
			select case i
				case 1 : .txtPROJECTNO.value = vntInParam(i)	
				case 2 : .txtPROJECTNM.value = vntInParam(i)
				case 3 : .txtCLIENTNAME.value = vntInParam(i)
				case 4 : .txtSUBSEQNAME.value = vntInParam(i)
				case 5 : .txtGROUPGBN.value = vntInParam(i)
				case 6 : .txtCREDAY.value = vntInParam(i)
				case 7 : .txtCPDEPTNAME.value = vntInParam(i)
				case 8 : .txtCPEMPNAME.value = vntInParam(i)
				case 9 : .txtCLIENTTEAMNAME.value = vntInParam(i)
				case 10: .txtMEMO.value = vntInParam(i)
			end select
		next
		
		gSetSheetColor mobjSCGLSpr, .sprSht_JOBNODEPT
		mobjSCGLSpr.SpreadLayout .sprSht_JOBNODEPT, 10, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_JOBNODEPT, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_JOBNODEPT, "CHK | SEQ | EMPNAME | BTN2 | EMPNO | DEPTNAME | BTN | DEPTCODE | JOBNOSEQ | ACTRATE"
		mobjSCGLSpr.SetHeader .sprSht_JOBNODEPT,        "선택|순번|담당자|담당자사번|담당부서|담당부서코드|JOBSEQ|부서실적입력"
		mobjSCGLSpr.SetColWidth .sprSht_JOBNODEPT, "-1","   4|   5|  10|2|        10|    28|2|          10|6     |12" 
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_JOBNODEPT, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_JOBNODEPT, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_JOBNODEPT,"..", "BTN | BTN2"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_JOBNODEPT, "ACTRATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "EMPNAME | BTN2 | EMPNO | DEPTNAME | BTN | DEPTCODE | JOBNOSEQ",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_JOBNODEPT, "",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht_JOBNODEPT, true, "SEQ | JOBNOSEQ"
		mobjSCGLSpr.SetScrollBar .sprSht_JOBNODEPT, 2, True, 0, -1
		mobjSCGLSpr.colhidden .sprSht_JOBNODEPT, "JOBNOSEQ | SEQ",true
		
		.sprSht_JOBNODEPT.style.visibility = "visible"
		
		InitPageData
    End With
End Sub

'초기값 세팅
Sub InitPageData
	with frmThis
		.txtHOPEENDDAY.value = gNowDate
		.txtREQDAY.value = gNowDate
		.txtJOBNO.focus
		.txtEMPNO.value = gstrUsrID
		
		Call EMPNAME_SETTING()
		Call COMBO_TYPE()
		Call SUBCOMBO_TYPE()
	
		.sprSht_JOBNODEPT.MaxRows = 1
		
		'실적부서 입력
		mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNO",1, .txtEMPNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"EMPNAME",1, .txtEMPNAME.value 
		mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTCODE",1, .txtDEPTCD.value 
		mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"DEPTNAME",1, .txtDEPTNAME.value 
		mobjSCGLSpr.SetTextBinding .sprSht_JOBNODEPT,"ACTRATE",1, "100"
		sprSht_JOBNODEPT_Change 1,1
		
		.txtJOBNAME.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub EndPage()
	set mobjPDCOJOBNO = Nothing
	set mobjPDCOGET = Nothing
	set mobjPDCOACTUALRATE = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'화면클리어
Sub DataClean
	with frmThis
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = "" 
		.txtCLIENTNAME.value =  ""
		.txtCPDEPTNAME.value =  ""
		.txtCREDAY.value = ""
		.txtCPEMPNAME.value =  ""
		.txtGROUPGBN.value = ""
		.txtSUBSEQNAME.value =  ""
		.txtCLIENTTEAMNAME.value = ""
		.txtMEMO.value =  ""
		.txtJOBNAME.value =  ""
		.txtJOBNO.value =  ""
		.cmbJOBGUBN.selectedIndex = 0 
		.txtDEPTNAME.value =  ""
		.txtDEPTCD.value =  ""
		.txtREQDAY.value = gNowDate
		SUBCOMBO_TYPE
		.txtEMPNAME.value =  "" 
		.txtEMPNO.value =  ""
		.txtHOPEENDDAY.value = gNowDate 
		.cmbCREGUBN.selectedIndex = 0 
		.cmbJOBBASE.selectedIndex = 0 
		.cmbENDFLAG.selectedIndex = 0 
		.txtCREEMPNAME.value =  ""
		.txtCREEMPNO.value =  ""
		.txtAGREEYEARMON.value = "" 
		.txtDEMANDYEARMON.value =  ""
		.txtSETYEARMON.value =  ""
		.txtBUDGETAMT.value =  ""
		.txtBIGO.value =  ""
		.txtEXCLIENTCODE.value = ""
		.txtEXCLIENTNAME.value = ""
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 사용자 정보가져오기
'-----------------------------------------------------------------------------------------
Sub EMPNAME_SETTING
	Dim vntData
   	Dim i, strCols
	
	On error resume next
	
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml, mlngRowCnt, mlngColCnt, .txtEMPNO.value, .txtEMPNAME.value,"A", .txtDEPTCD.value, .txtDEPTNAME.value)
		if not gDoErrorRtn ("GetCUSTNO") then
			If mlngRowCnt = 1 Then
				.txtEMPNO.value = trim(vntData(0,1))
				.txtEMPNAME.value = trim(vntData(1,1))
				.txtDEPTCD.value = trim(vntData(2,1))
				.txtDEPTNAME.value = trim(vntData(3,1))
				
				.txtEXCLIENTNAME.focus()
				gSetChangeFlag .txtEMPNO
			Else
				Call EMP_POP()
			End If
   		end if
   	end with
	window.event.returnValue = false
	window.event.cancelBubble = true
End Sub

'-----------------------------------------------------------------------------------------
' COMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
	Dim strCODE
    
    With frmThis   
		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCOJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB종류 호출
		vntJOBBASE = mobjPDCOJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBBASE")  '청구기준 호출	
		vntCREGUBN = mobjPDCOJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREGUBN")  '신규/기존 호출
		vntENDFLAG = mobjPDCOJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '제작상태 호출
		vntCREPART = mobjPDCOJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREPART")  
		
		if not gDoErrorRtn ("COMBO_TYPE") then 
			 gLoadComboBox .cmbENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbJOBGUBN, vntJOBGUBN, False
			 gLoadComboBox .cmbJOBBASE, vntJOBBASE, False
			 gLoadComboBox .cmbCREGUBN, vntCREGUBN, False 
			 gLoadComboBox .cmbCREPART, vntCREPART, False 
   		end if    
   	end with     	
End Sub

'-----------------------------------------------------------------------------------------
' SUBCOMBO TYPE 설정
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()
	Dim vntCREPART
   	Dim vntCREGUBN
   
	With frmThis   
		'On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
       	
	    vntCREPART = mobjPDCOJOBNO.GetDataTypeChange(gstrConfigXml, mlngRowCnt, mlngColCnt,.cmbJOBGUBN.value,"K")  '제작종류 호출	

		if not gDoErrorRtn ("SUBCOMBO_TYPE") then 
			 gLoadComboBox .cmbCREPART, vntCREPART, False
   		end if  
   		.cmbCREPART.selectedIndex = 1
   		cmbCREPART_onchange		   		
   	end with   
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	IF not SelectRtn_Head Then 
		gErrorMsgBox "조회값이 없습니다. 관리자에게 문의하세요.","조회안내"
		Exit Sub
	End If
	
	CALL SelectRtn_Detail ()
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

Function SelectRtn_Head
	Dim vntData
	'On error resume next
	
	SelectRtn_Head = False
	
	with frmThis
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		vntData = mobjPDCOJOBNO.SelectRtn_PROJECTORJOB_XML(gstrConfigXml,mlngRowCnt,mlngColCnt,"","","","","","","","",mstrJOBNO,Trim(.txtJOBNAME.value),"2")
		
		If not gDoErrorRtn ("SelectRtn") then
			If mlngRowCnt > 0  Then
				call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
				
				If .txtBUDGETAMT.value <> "" Then
					txtBUDGETAMT_onblur
				End If
				SelectRtn_Head = True
			End If				
		End If		
		.cmbJOBGUBN.disabled = true
		.cmbCREPART.disabled = true
	End With
End Function

Function SelectRtn_Detail
	Dim vntData
   	Dim strRow,strJOBNO , strJOBNOSEQ
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'시트1의 JOB번호를 가지고오는데 사용
		
		strJOBNO = .txtJOBNO.value 
		
		'시트1의 JOBNOSEQ번호를 가지고오는데 사용
		vntData = mobjPDCOACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,"1")
		
		If not gDoErrorRtn ("SelectRtn_DTL_JOBNODEPT") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_JOBNODEPT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_JOBNODEPT,meCLS_FLAG
			
			If mlngRowCnt < 1 Then
				.sprSht_JOBNODEPT.MaxRows = 0	
			End If
		End If		
	END WITH
End Function

'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
  	Dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strJOBNO
	Dim intCnt
	Dim intCode,intEDITCODE
	Dim strEDITJOBNO
	Dim intRtnSave
	Dim intRtn2
	with frmThis
	'On error resume next
		strJOBNO = ""
		
  		'데이터 Validation
		if DataValidation =false then exit sub
		'strCODE = .txtPROJECTNO.value
		
	
		'처리 업무객체 호출
		strMasterData = gXMLGetBindingData (xmlBind)
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_JOBNODEPT,"CHK | SEQ | EMPNAME | BTN2 | EMPNO | DEPTNAME | BTN | DEPTCODE | JOBNOSEQ | ACTRATE")
		
		If .txtJOBNO.value = "" Then
			If  Not gXMLIsDataChanged (xmlBind) Then
				gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				exit sub
			End If
		End If
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		intRtn2 = gYesNoMsgbox("실적부서를 확인하셨습니까?" & vbCrlf & " " & vbCrlf & " " & vbCrlf & "실적부서는 추후에 부서별 실적을 집계하는데 중대한 영향을 미칠수 있습니다. "& vbCrlf & " " & vbCrlf & "반드시 확인 바랍니다. ","확인")
		IF intRtn2 <> vbYes then exit Sub
		
		
		'실제저장 시작
		if .txtJOBNO.value = "" then
			strSEQFlag = "new"
			intRtn = mobjPDCOJOBNO.ProcessRtn(gstrConfigXml,strMasterData, "new",strJOBNO)
			mstrJOBNO = strJOBNO
		Else
			'변경된 자료가 있을경우에만 처리
			If  gXMLIsDataChanged (xmlBind) Then
				intRtn = mobjPDCOJOBNO.ProcessRtn(gstrConfigXml,strMasterData, "Edit",mstrJOBNO)
			End If
		end if
		
		if not gDoErrorRtn ("ProcessRtn") then
			'실적분배저장
			gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내"
			intRtn = mobjPDCOACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData,mstrJOBNO,"1")
			SelectRtn 
  		end if
 	end with
End Sub

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols,intCnt,intCnt2,intCnt3
   	Dim intColSum
   	Dim strDupACTNAME,strDupACTNAME_CHECK
   	Dim intDupCnt,intSubCnt
   	Dim strTotalRATE
   	Dim strRATE
	'On error resume next
	with frmThis
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
		If .cmbCREPART.value = "PC01" Or .cmbCREPART.value = "PR01" Or .cmbJOBGUBN.value = "PA01" Then
			If .txtEXCLIENTCODE.value = "" Then
				gErrorMsgBox "매체구분 이 '인쇄' 일경우 및 " & vbcrlf & "매체분류가 TV-CF 또는 Radio-CM 일때 크리조직 입력은 필수 입니다.","입력안내"
				.txtEXCLIENTNAME.focus()
				Exit Function
			End If
		End If
   	
   		'Sheet Validation
   		'intColSum = 0
  		'for intCnt = 1 to .sprSht_JOBNODEPT.MaxRows
		'		if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"CHK",intCnt) = 1  Then 
		'				intColSum = intColSum + 1
		'		End if
		'next
		
		'If intColSum = 0 Then 
		'	gErrormsgbox "선택된 데이터가 없습니다.","처리안내"
		'	exit Function
		'End If
		
		'필수항목검사
		for intCnt2 = 1 to .sprSht_JOBNODEPT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"DEPTCODE",intCnt2) = "" Then 
				gErrorMsgBox intCnt2 & " 번째 행의 부서코드를 선택 하지 않으셨습니다. 팝업버튼을 눌러 정확한 부서를 선택하시기 바랍니다.","입력오류"
				'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
				Exit Function
			End if
		next
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'부서중복체크
		For intDupCnt = 1 To .sprSht_JOBNODEPT.MaxRows	
			
			strDupACTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTCODE",intDupCnt)

			For intSubCnt = 1 To .sprSht_JOBNODEPT.MaxRows 
				If intDupCnt <> intSubCnt Then 
					strDupACTNAME_CHECK = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"DEPTCODE",intSubCnt)
					If strDupACTNAME = strDupACTNAME_CHECK Then
						gErrorMsgBox "담당자" & intDupCnt & "행에 중복된 부서가 있습니다.","중복안내"
						mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,intSubCnt
						'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
						Exit Function
					End IF
				End IF
			Next
		Next
		
		
		'실적비율 합계 100% 처리...
		strTotalRATE = 0
		For intCnt3 = 1 To .sprSht_JOBNODEPT.MaxRows 
			strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"ACTRATE",intCnt3)
			
			If strRATE = 0 Then
				gErrorMsgBox "실적분배비율을 반드시 입력하여야 합니다.","저장안내"
				'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
				Exit Function
			End IF
			
			strTotalRATE= strTotalRATE+strRATE
		Next
		
		'저장한 내역을 처음부터 끝까지 돌면서 합이 100인지 확인
		If strTotalRATE <> 100 Then 
			gErrorMsgBox "실적분배비율의 총합은 100 이어야 합니다.","저장안내"
			'mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
			Exit Function
		End If
   	End with
	DataValidation = true
End Function

'------------------------------------------
' JOBNODEPT 삭제
'------------------------------------------
Sub DeleteRtn_DTL_JOBNODEPT
	Dim vntData
	Dim vntDataDEL
	Dim intSelCnt, intRtn, i , intCnt
	Dim strYEARMON
	Dim strSEQ
	Dim strRATE , strTotalRATE
	Dim strRow
	Dim strJOBNO
	Dim strJOBNOSEQ
	Dim intValCnt
	Dim strCHKCnt
	Dim intDelCount
	Dim intColSum
	
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht_JOBNODEPT,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		strCHKCnt = 0
		
		'삭제할 내역을뺀 나머지를 돌면서 실적분배비율의 합이 100인지 확인
		For intCnt = 1 To .sprSht_JOBNODEPT.MaxRows
			If mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"CHK",intCnt) <> "1" And mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"JOBNOSEQ",intCnt) <> "" Then
				strRATE = mobjSCGLSpr.GetTextBinding( .sprSht_JOBNODEPT,"ACTRATE",intCnt)
				strTotalRATE= strTotalRATE+strRATE
			Else 
				strCHKCnt = 1
			End if
		Next
		
		'저장시 체크된것만 저장될수 있도록
		intColSum = 0
  		for intCnt = 1 to .sprSht_JOBNODEPT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"CHK",intCnt) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "선택된 데이터가 없습니다.","삭제안내"
			exit Sub
		End If
		
		
		If strTotalRATE <> 100 Then
			if strTotalRATE <> 0 Then
				gErrorMsgBox "삭제할데이터를 제외한 나머지 실적분배비율을 반드시 100% 또는 0%(전체삭제) 이여야 합니다.","삭제안내"
				Exit Sub
			End If
		End IF
			
		intRtn = gYesNoMsgbox("선택된 자료가 삭제 됩니다." & vbcrlf & "자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		strJOBNO = .txtJOBNO.value
		
		'선택된 자료를 끝에서 부터 삭제
		intDelCount = 0
		for i = .sprSht_JOBNODEPT.MaxRows to 1 step -1	
			if mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"CHK",i) = 1 THEN
				If mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"SEQ",i) <> "" Then 
					strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"SEQ",i)
					strJOBNOSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_JOBNODEPT,"JOBNOSEQ",i) 
					intRtn = mobjPDCOACTUALRATE.DeleteRtn_DTL_JOBNODEPT_JOBNO(gstrConfigXml, strSEQ,strJOBNO,strJOBNOSEQ)
				End If
				
				IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht_JOBNODEPT,i
   				End IF
   				
   				intDelCount = intDelCount + 1
   				gWriteText "",intDelCount & " 건이 삭제" & mePROC_DONE
   			END IF
		next
		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_JOBNODEPT
		'변경된 내역 재저장
		ProcessRtn_DTL_ACTUALRATE_DEL
		SelectRtn
	End with
	err.clear
End Sub


Sub ProcessRtn_DTL_ACTUALRATE_DEL()
    Dim intRtn
  	Dim vntData
	Dim strJOBNO ,strSEQFlag , strRATE ,strTotalRATE, strJOBNOSEQ
	Dim strRow , strOLDSEQ
	Dim intCnt , intCode ,intEDITCODE 
	Dim dblACTLRATE
	
	with frmThis
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht_JOBNODEPT,"CHK | SEQ | EMPNAME | EMPNO | DEPTNAME | DEPTCODE | JOBNOSEQ | ACTRATE")
		if  not IsArray(vntData) then 
			exit sub
		End If
		intRtn = mobjPDCOACTUALRATE.ProcessRtn_DTL_JOBNODEPT(gstrConfigXml,vntData,.txtJOBNO.value,"1")
		'건수저장 
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_JOBNODEPT,meCLS_FLAG
  		end if
  		
 	end with
End Sub

'클리어
Sub CleanField (objField1, objField2)
	if isobject(objField1) then 
		objField1.value = ""
	end if
	if isobject(objField2) then 
		objField2.value = ""
	End If
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<TBODY>
					<tr>
						<td>
							<TABLE style=" HEIGHT: 8px" height="8" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" height="20">
										<table style="WIDTH: 640px; HEIGHT: 26px" cellSpacing="0" cellPadding="0" width="640" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="128" background="../../../images/back_p.gIF"
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
												<td class="TITLE"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id="imgTableUp" style="CURSOR: hand" alt="자료를 검색합니다." src="../../../images/imgTableUp.gif"
															align="absMiddle" border="0" name="imgTableUp"></span>&nbsp;JOBNO 신규등록
												</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="화면을 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
						</td>
					</tr>
					<TR>
						<TD class="BODYSPLIT" id="spacebar1" style="WIDTH: 1030px; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 1030px" vAlign="top" align="left">
							<TABLE class="SEARCHDATA" id="tblBody1" cellSpacing="1" cellPadding="0" width="1040" border="0">
								<TR>
									<TD class="GROUP" width="18" rowSpan="3" style="WIDTH: 18px">기<BR>
										본<BR>
										정<BR>
										보
									</TD>
									<TD class="SEARCHLABEL" width="101" style="WIDTH: 101px; HEIGHT: 17px">프로젝트명</TD>
									<TD class="SEARCHDATA" width="230" style="HEIGHT: 17px"><INPUT dataFld="PROJECTNM" class="NOINPUTB_L" id="txtPROJECTNM" title="프로젝트명" style="WIDTH: 160px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="21" name="txtPROJECTNM">&nbsp;<INPUT dataFld="PROJECTNO" class="NOINPUTB" id="txtPROJECTNO" title="프로젝트건" style="WIDTH: 65px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="6" name="txtPROJECTNO"></TD>
									<TD class="SEARCHLABEL" width="90" style="HEIGHT: 17px">브랜드</TD>
									<TD class="SEARCHDATA" width="230" style="HEIGHT: 17px"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_L" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="24" name="txtSUBSEQNAME"></TD>
									<TD class="SEARCHLABEL" width="90" style="HEIGHT: 17px">담당부서 [CP]</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 17px"><INPUT dataFld="CPDEPTNAME" class="NOINPUTB_L" id="txtCPDEPTNAME" title="담당부서 CP" style="WIDTH: 248px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="36" name="txtCPDEPTNAME"></TD>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 101px">등록일</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CREDAY" class="NOINPUTB" id="txtCREDAY" title="등록일" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCREDAY"></TD>
									<TD class="SEARCHLABEL">팀</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CLIENTTEAMNAME" class="NOINPUTB_L" id="txtCLIENTTEAMNAME" title="팀" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTTEAMNAME"></TD>
									<TD class="SEARCHLABEL">담당자 [CP]</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CPEMPNAME" class="NOINPUTB_L" id="txtCPEMPNAME" title="담당자CP" style="WIDTH: 248px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="36" name="txtCPEMPNAME"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 101px">그룹구분</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="GROUPGBN" class="NOINPUTB_L" id="txtGROUPGBN" title="그룹구분" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtGROUPGBN"></TD>
									<TD class="SEARCHLABEL">광고주</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
									<TD class="SEARCHLABEL">비고</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="MEMO" class="NOINPUTB_L" id="txtMEMO" title="사업부" style="WIDTH: 248px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="36" name="txtMEMO"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" id="spacebar2" style="WIDTH: 1030px; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 1040px" vAlign="top" align="left">
							<TABLE class="SEARCHDATA" id="tblBody2" cellSpacing="1" cellPadding="0" width="1040" border="0">
								<TR>
									<TD class="GROUP" width="19" rowSpan="7" style="WIDTH: 19px"><BR>
										부<BR>
										가<BR>
										정<BR>
										보<BR>
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 111px; CURSOR: hand" onclick="vbscript:Call CleanField(txtJOBNAME, '')"
										width="111">JOB명</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="JOBNAME" id="txtJOBNAME" title="제작건명" style="WIDTH: 171px; HEIGHT: 21px"
											accessKey=",M" dataSrc="#xmlBind" type="text" size="23" name="txtJOBNAME" class="INPUT_L"><INPUT dataFld="JOBNO" class="NOINPUT" id="txtJOBNO" title="제작관번호코드" style="WIDTH: 62px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBNO"></TD>
									<TD class="SEARCHLABEL" width="100">매체부문</TD>
									<TD class="SEARCHDATA" width="200"><SELECT dataFld="JOBGUBN" id="cmbJOBGUBN" title="매체구분" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBGUBN"></SELECT></TD>
									<TD class="SEARCHLABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPTNAME, txtDEPTCD)">담당팀</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="DEPTNAME" class="INPUT_L" id="txtDEPTNAME" title="담당부서명" style="WIDTH: 160px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtDEPTNAME">
										<IMG id="ImgDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
											src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgDEPTCD">
										<INPUT dataFld="DEPTCD" class="INPUT_L" id="txtDEPTCD" title="담당부서코드" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="6" name="txtDEPTCD"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 111px; CURSOR: hand" onclick="vbscript:Call CleanField(txtREQDAY, '')">의뢰일</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="REQDAY" class="INPUT" id="txtREQDAY" title="의뢰일" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtREQDAY"><IMG id="imgCalEndarREQ" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
											name="imgCalEndarREQ"></TD>
									<TD class="SEARCHLABEL">매체분류</TD>
									<TD class="SEARCHDATA"><SELECT dataFld="CREPART" id="cmbCREPART" title="매체분류" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREPART"></SELECT></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEMPNAME, txtEMPNO)">담당자</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="EMPNAME" class="INPUT_L" id="txtEMPNAME" title="담당자명" style="WIDTH: 160px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEMPNO">
										<INPUT dataFld="EMPNO" class="INPUT_L" id="txtEMPNO" title="담당자사번" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtEMPNO"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 111px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtHOPEENDDAY, '')">완료예정일</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 18.24pt"><INPUT dataFld="HOPEENDDAY" class="INPUT" id="txtHOPEENDDAY" title="완료예정일" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtHOPEENDDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndar"></TD>
									<TD class="SEARCHLABEL" style="HEIGHT: 24px">신규/수정 구분</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 18.24pt"><SELECT dataFld="CREGUBN" id="cmbCREGUBN" title="신규/수정 구분" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREGUBN"></SELECT></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)">크리조직</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 18.24pt"><INPUT dataFld="EXCLIENTNAME" class="INPUT_L" id="txtEXCLIENTNAME" title="대대행사명" style="WIDTH: 160px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="21" name="txtEXCLIENTNAME">
										<IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
											src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">
										<INPUT dataFld="EXCLIENTCODE" class="INPUT_L" id="txtEXCLIENTCODE" title="대대행사코드" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="8" size="6" name="txtEXCLIENTCODE"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 111px">완료구분</TD>
									<TD class="SEARCHDATA"><SELECT dataFld="ENDFLAG" id="cmbENDFLAG" title="완료구분" style="WIDTH: 112px" dataSrc="#xmlBind"
											name="cmbENDFLAG"></SELECT>&nbsp;&nbsp;<IMG id="imgEndChange" onmouseover="JavaScript:this.src='../../../images/imgEndChangeOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgEndChange.gIF'" height="20" alt="진행상태 를 의뢰상태로 변경합니다."
											src="../../../images/imgEndChange.gIF" align="absMiddle" border="0" name="imgEndChange"></TD>
									<TD class="SEARCHLABEL">정산대상</TD>
									<TD class="SEARCHDATA"><SELECT dataFld="JOBBASE" id="cmbJOBBASE" title="정산대상" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBBASE"></SELECT></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtBUDGETAMT, '')">예산</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="BUDGETAMT" class="INPUT_R" id="txtBUDGETAMT" title="예산금액" style="WIDTH: 248px; HEIGHT: 21px"
											accessKey="NUM" dataSrc="#xmlBind" type="text" size="36" name="txtBUDGETAMT"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 111px">합의월</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="AGREEYEARMON" class="NOINPUTB" id="txtAGREEYEARMON" title="합의월" style="WIDTH: 232px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="33" name="txtAGREEYEARMON"></TD>
									<TD class="SEARCHLABEL">청구월</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="DEMANDYEARMON" class="NOINPUTB" id="txtDEMANDYEARMON" title="청구월" style="WIDTH: 224px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="32" name="txtDEMANDYEARMON"></TD>
									<TD class="SEARCHLABEL">결산월</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="SETYEARMON" class="NOINPUTB" id="txtSETYEARMON" title="결산월" style="WIDTH: 248px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="36" name="txtSETYEARMON"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" onclick="vbscript:Call CleanField(txtBIGO, '')" style="WIDTH: 111px">비고</TD>
									<TD class="SEARCHDATA" colSpan="5"><INPUT dataFld="BIGO" id="txtBIGO" title="부가정보 비고" style="WIDTH: 906px; HEIGHT: 21px" dataSrc="#xmlBind"
											type="text" size="145" name="txtBIGO" class="INPUT_L"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--Input End-->
					<!--BodySplit Start-->
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<td>
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="103" background="../../../images/back_p.gIF"
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
												<td class="TITLE">실적 분배율 입력</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgAddRow_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
														alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
												<td><IMG id="imgDelete_sprSht_JOBNODEPT" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete_sprSht_JOBNODEPT"></td>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
						</td>
					</TR>
					<tr>
						<td>
							<table height="200" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<!--첫번째-->
									<TD style="WIDTH: 100%; HEIGHT: 200px" vAlign="top" align="center" colSpan="2">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 200px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_JOBNODEPT" style="WIDTH: 100%; HEIGHT: 200px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="5292">
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
							</table>
						</td>
					</tr>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1030px"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
