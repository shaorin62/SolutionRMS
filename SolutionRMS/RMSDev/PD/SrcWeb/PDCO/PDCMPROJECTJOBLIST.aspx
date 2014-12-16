<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPROJECTJOBLIST.aspx.vb" Inherits="PD.PDCMPROJECTJOBLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>프로젝트/JOB 관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD
'실행  환경 : ASP.NET, VB.NET, COM+
'프로그램명 : PDCMPROJECTJOBLIST.aspx
'기      능 : 프로젝트 및 JOB 을 동시에 관리한다.
'파라  메터 : 
'특이  사항 : 개발 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT id="clientEventHandlersVBS" language="vbscript">
<!--
Option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjPDCOGET , mobjPDCOPONO , mobjPDCOJOBNO, mobjSCCOGET
Dim mstrCheck

CONST meTAB = 9
mstrCheck = true

Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
End Sub

'=============================
' 이벤트프로시져 
'=============================
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn_PROJECT
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_PROJECT
	gFlowWait meWAIT_OFF
End Sub

Sub imgJobDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_DTL
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_PROJECT
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgJobExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectNew_onclick
	call sprSht_PROJECT_Keydown(meINS_ROW, 0)
End Sub

Sub imgDTLNew_onclick
	Dim vntInParams
	Dim vntRet
	Dim strRow, strCol
	'신규 값넘기기
	Dim strPROJECTNO, strPROJECTNM
	Dim strCLIENTNAME
	Dim strSUBSEQNAME
	Dim strGROUPGBN
	Dim strCREDAY
	Dim strCPDEPTNAME
	Dim strCPEMPNAME
	Dim strCLIENTTEAMNAME
	Dim strMEMO
	
	with frmThis
		IF .sprSht_PROJECT.MaxRows > 0  then
			strPROJECTNO	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNO",.sprSht_PROJECT.ActiveRow)
			strPROJECTNM	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNM",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			strSUBSEQNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
			strGROUPGBN		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow)
			
			If strGROUPGBN = "2" Then
				strGROUPGBN = "그룹"
			Elseif strGROUPGBN = "1" Then
				strGROUPGBN = "비그룹"
			End If
			
			strCREDAY			= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CREDAY",.sprSht_PROJECT.ActiveRow)
			strCPDEPTNAME		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
			strCPEMPNAME		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
			strCLIENTTEAMNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
			strMEMO				= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"MEMO",.sprSht_PROJECT.ActiveRow)
			
			If strPROJECTNO = "" Then
				gErrorMsgBox "프로젝트를 저장후 JOB을 등록 하십시오.","입력안내"
				Exit Sub
			Else
				vntInParams = array("New", strPROJECTNO, strPROJECTNM, strCLIENTNAME, strSUBSEQNAME, _
									strGROUPGBN, strCREDAY, strCPDEPTNAME, strCPEMPNAME, strCLIENTTEAMNAME, strMEMO)
				
				vntRet = gShowModalWindow("PDCMJOBNONEW.aspx",vntInParams , "1060", "600")
			End If

			sprSht_PROJECT_click 2,.sprSht_PROJECT.ActiveRow
			.txtCLIENTCODE1.focus()
			.sprSht_DTL.Focus
		else
			gErrorMsgBox "JOB등록할 프로젝트를 조회하세요.","입력안내"
		end if	
	end with
End Sub

Sub imgJobDetail_onclick()
	Dim strJOBNO, strPRONO
	Dim vntInParams
	Dim vntRet
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	strWith =  Screen.width
	strHeight =  Screen.height - 100
	with frmThis
		IF .sprSht_DTL.MaxRows >0  then
			If mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"SEQ",.sprSht_DTL.ActiveRow) = "1" Then
				strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNO",.sprSht_DTL.ActiveRow)
				strPRONO = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"PROJECTNO",.sprSht_DTL.ActiveRow)
				
				vntInParams = array("Detail",mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNO",.sprSht_DTL.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNAME",.sprSht_DTL.ActiveRow))
				vntRet = gShowModalWindow("PDCMJOBNODETAIL.aspx",vntInParams , strWith,strHeight)
				strRow = .sprSht_DTL.ActiveRow
				strCol = .sprSht_DTL.ActiveCol
				'여기서 부터 실제 견적 화면 호출
				.txtCLIENTCODE1.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_DTL.Focus
				
				SelectRtn_DTL(strPRONO)
				mobjSCGLSpr.ActiveCell .sprSht_DTL, strCol, strRow		
			Else
				msgbox "대표JOBNO 가 아닙니다.SUBNO 가 1 인 항목을 선택하여 주십시오"
			End If
		end if	
	end with
End Sub

'=============================
' 명령버튼클릭이벤트
'=============================
Sub imgClose_onclick()
    Window_OnUnload()
End Sub

'-----------------------------------------------------------------------------------------
' Project 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'ProjectNO 조회팝업
Sub ImgPROJECTNO1_onclick
	with frmThis
		'1은 PROJECT 조회   2는 JOBNO조회
		IF .cmbPOPUPTYPE.value = "1" then
			Call PONO_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	
	End with
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
			'.txtCLIENTNAME1.focus()					' 포커스 이동
     	end if
	End with
	gSetChange
End Sub

Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array( trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' 코드명 표시
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
		if .cmbPOPUPTYPE.value = "1" Then '프로젝트 코드 라면
			vntData = mobjPDCOGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,1))
					.txtPROJECTNM1.value = trim(vntData(1,1))
				Else
					Call PONO_POP()
				End If
   			end if
		Else
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		End If
   		
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub




'****************************************************************************************
' 팝업 이벤트, 광고주, 매체명, 매체사
'****************************************************************************************
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
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))       ' Code값 저장
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' 코드명 표시
			gSetChangeFlag .txtCLIENTCODE1                  ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
	
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value),"A")
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' 달력
'****************************************************************************************
'조회용
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
		gSetChange
	end with
End Sub


Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub

Sub cmbPOPUPTYPE_onchange
	with frmThis
		.txtPROJECTNM1.value = ""
		.txtPROJECTNO1.value = ""
	End with
	gSetChange
End Sub
'=============================
'SheetEvent
'=============================
'더블클릭
sub sprSht_PROJECT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_PROJECT, ""
		end if
	end with
end sub

'쉬트데이터 변경시
Sub sprSht_PROJECT_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim strDeptCodeName
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strCLIENTSUBCODE
	Dim strCLIENTSUBNAME
	Dim strTIMCODE
	Dim strCLIENTTEAMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	
	With frmThis
		'담당자
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNAME")  Then
					strCode = ""
					strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
				
					vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(3,1)
						
						mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNO"),frmThis.sprSht_PROJECT.ActiveRow
					Else
						mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
					End If
					.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다 이거수
					.sprSht_PROJECT.Focus	
					If Row <> .sprSht_PROJECT.MaxRows Then
						mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
					Else
						mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
					End IF
		'담당부서
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTNAME")  Then
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
				vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTCD"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		'광고주
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A")
				
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(4,1)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTCODE"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTTEAMNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE,strCLIENTNAME,"",strCodeName)
				
		
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(4,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(5,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(6,1)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"TIMCODE"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		'브랜드
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntData = mobjSCCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,strCLIENTCODE,strCLIENTNAME)
				
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",Row, vntData(4,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",Row, vntData(5,1)
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(8,1)
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(9,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(10,1)
					
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQ"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		End If
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_PROJECT, Col, Row
End Sub

'버튼연계처리
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	Dim vntRet, vntInParams
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strCLIENTSUBCODE
	Dim strCLIENTSUBNAME
	Dim strTIMCODE
	Dim strCLIENTTEAMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	
	
	With frmThis
		'PROJECT 그리드
		If sprSht = "sprSht_PROJECT" Then
			
			'담당자
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNAME") Then
			
				vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPEMPNAME",Row))
				
				vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
				
				'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntRet(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntRet(3,0)		
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			
			'담당부서
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTNAME") Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTNAME",Row))
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			'광고주
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTNAME") Then
			
				vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CLIENTNAME",Row))
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntRet(4,0)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If

			'팀
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTTEAMNAME") Then
			
				strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTTEAMNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
					
				vntInParams = array("", trim(strCLIENTNAME),"", trim(strCLIENTTEAMNAME) )  '<< 받아오는경우
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(5,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow,  trim(vntRet(6,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			
			'브랜드
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQNAME") Then
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow)
				strSUBSEQNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
					
				vntInParams = array("", trim(strSUBSEQNAME),"", trim(strCLIENTNAME))  '<< 받아오는경우
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 520,430)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(5,0))
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(8,0))
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(9,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(10,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			End If
		'JOB 그리드 변동시
		Elseif sprSht = "sprSht_DTL" Then
		
		
		End If	
	
	End With
End Sub

'더블클릭
sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	imgJobDetail_onclick
end sub
'프로젝트 리스트 방향키 누를때
Sub sprSht_PROJECT_Keyup(KeyCode, Shift)
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
	
	'키가 움질일때 바인딩
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprSht_PROJECT_Click frmThis.sprSht_PROJECT.ActiveCol,frmThis.sprSht_PROJECT.ActiveRow
	End If
End Sub

'프로젝트 리스트 클릭시
Sub sprSht_PROJECT_Click(ByVal Col, ByVal Row)
	Dim intcnt,intCnt2
	Dim strPROJECTNO
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_PROJECT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_PROJECT.MaxRows
				sprSht_PROJECT_Change 1, intcnt
			next
		Else
			'쉬트바인딩 프로젝트-JOB 
			strPROJECTNO = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNO",.sprSht_PROJECT.ActiveRow)
			
			IF strPROJECTNO <> "" Then
				SelectRtn_DTL(strPROJECTNO)
				'JOBNO 등록이 있는 경우 프로젝트 비고 및 프로젝트 명 외, 수정불가
				if .sprSht_DTL.MaxRows = 0 Then
					mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT,false,"CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO",Row,Row,false
				Else
					mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT,true,"CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO",Row,Row,false
				End If
			Else
				.sprSht_DTL.MaxRows = 0	
			End If
		end if
	end with
End Sub

'행 신규
Sub sprSht_PROJECT_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim vntData
	
	On error resume Next
	
	if KeyCode <> meINS_ROW then exit sub
	
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_PROJECT, cint(KeyCode), cint(Shift), -1, 1)
	
	with frmThis
		'사용자 정보 가져오기
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOGET.GetSCEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,gstrUsrID,"","A","","")
		
		if not gDoErrorRtn ("GetSCEMP") then
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CREDAY",.sprSht_PROJECT.ActiveRow,gNowDate
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow,gstrUsrID
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow,vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow,vntData(2,1)
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow,vntData(3,1)
			Else
				gErrorMsgBox "사용자 정보를 얻지 못하였습니다." & vbcrlf & "재로그인 하여 주십시오.","입력안내" 
			End If
		End If
	End with
End Sub

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_PROJECT_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	Dim strCLIENTSUBCODE , strCLIENTSUBNAME , strCLIENTCODE , strCLIENTNAME,strTIMCODE,strCLIENTTEAMNAME
	Dim strSUBSEQ , strSUBSEQNM
	Dim strCPDEPTCD , strCPDEPTNAME
	Dim strCPEMPNO , strCPEMPNAME
	
	with frmThis

		'광고주
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CLIENT") Then
		
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CLIENT") then exit Sub
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
			if isArray(vntRet) then
				if strCLIENTCODE = vntRet(0,0) and strCLIENTNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
			end if
			.txtFrom.focus()
			.sprSht_PROJECT.focus()	
			gSetChange
     	'팀
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_TEAM") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_TEAM") then exit Sub
			strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTTEAMNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME),"", trim(strCLIENTTEAMNAME) ) '<< 받아오는경우
			
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if strTIMCODE = vntRet(0,0) and strCLIENTTEAMNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				
				if .sprSht_PROJECT.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(5,0))
					
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow,  trim(vntRet(6,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
				end if
				.txtFrom.focus()
				.sprSht_PROJECT.focus()					' 포커스 이동
				gSetChange 
     		end if
     	'브랜드
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_BRAND") Then
		
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_BRAND") then exit Sub
				
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow)
				strSUBSEQNM = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
			
				
				vntInParams = array("", trim(strSUBSEQNM),"", trim(strCLIENTNAME)) '<< 받아오는경우
		
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 520,430)
				if isArray(vntRet) then
					if strSUBSEQ = vntRet(0,0) and strSUBSEQNM = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit

					if .sprSht_PROJECT.ActiveRow >0 Then
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(5,0))
								'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(8,0))
								'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(9,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(10,0))
						
								mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
					end if
					.txtFrom.focus()
					.sprSht_PROJECT.focus()
					gSetChange	
     			end if
     	'담당부서
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPDEPT") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPDEPT") then exit Sub
				
				strCPDEPTCD = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow)
				strCPDEPTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntInParams = array(trim(strCPDEPTNAME))
				
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
				if isArray(vntRet) then
			
					if .sprSht_PROJECT.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
						
						mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
					end if
					.txtFrom.focus()
					.sprSht_PROJECT.focus()
					gSetChange	
				end if
		'담당자
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPEMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPEMP") then exit Sub
		
			strCPDEPTCD = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow)
			strCPDEPTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
			strCPEMPNO = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow)
			strCPEMPNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
			
			vntInParams = array("", trim(strCPDEPTNAME), "", trim(strCPEMPNAME)) '<< 받아오는경우
		
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if strCPEMPNO = vntRet(0,0) and strCPEMPNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				
				if .sprSht_PROJECT.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
					
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
				end if
				.txtFrom.focus()
				.sprSht_PROJECT.focus()
				gSetChange
     		end if
     	END IF	
	End with
End Sub

Sub InitPage()
    '서버업무객체 생성	
    set mobjPDCOPONO	= gCreateRemoteObject("cPDCO.ccPDCOPONO")
    set mobjPDCOJOBNO	= gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
    set mobjPDCOGET		= gCreateRemoteObject("cPDCO.ccPDCOGET")
    set mobjSCCOGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
   '권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	gSetSheetDefaultColor() 
	with frmThis
		'프로젝트 리스트 시트세팅
		gSetSheetColor mobjSCGLSpr, .sprSht_PROJECT
		mobjSCGLSpr.SpreadLayout .sprSht_PROJECT, 21, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,9, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,11, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_PROJECT, "CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO"
		mobjSCGLSpr.SetHeader .sprSht_PROJECT,		"선택|프로젝트코드|등록일|프로젝트명|광고주|팀|브랜드|담당부서|담당자|그룹구분|비고|광고주코드|팀코드|브랜드코드|부서코드|사번"
		mobjSCGLSpr.SetColWidth .sprSht_PROJECT, "-1","4 |          10|     8|        25|  20|2|18|2|18|2|    15|2|  10|2|      10|  25|         0|     10|  0       |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht_PROJECT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_PROJECT, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_PROJECT, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CLIENT"

		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_BRAND"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CPDEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CPEMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_TEAM"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_PROJECT, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT, true, "TIMCODE | PROJECTNO | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP"
		mobjSCGLSpr.ColHidden .sprSht_PROJECT, "CLIENTCODE | CPDEPTCD | CPEMPNO | SUBSEQ | TIMCODE", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_PROJECT, "PROJECTNM | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | GROUPGBN | CREDAY | CPDEPTCD | CPDEPTNAME | CPEMPNO | CPEMPNAME | MEMO | CLIENTTEAMNAME",-1,-1,0,2,false
        mobjSCGLSpr.SetCellAlign2 .sprSht_PROJECT, "PROJECTNO | CPEMPNAME",-1,-1,2,2,false '가운데
        
   
        '견적 확정이되고 청구요청 승인된 JOB LIST
        gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 24, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | CREDAY | JOBNAME | PREESTNO | JOBNO | SEQ | JOBGUBN | CREPART | BUDGETAMT | ENDFLAG | CREGUBN | JOBBASE | EMPNO | EMPNAME | DEPTCD | DEPTNAME | EXCLIENTCODE | EXCLIENTNAME | BIGO | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | PROJECTNO | RANKJOB"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		   "선택|의뢰일|JOB명|확정견적번호|JOBNO|SUBNO|매체부문|매체분류|예산|상태|신규|정산|담당자코드|담당자|담당팀코드|담당팀|제작사코드|제작사|비고|합의월|청구월|결산월|프로젝트번호|그룹지정"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "  4|      8|   25|           0|    8|    5|       8|      10|  10|   6|   6|   6|         0|     6|         0|    15|         0|    15|  17|    8 |    8 |   8  |0           |10"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "BUDGETAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CREDAY | JOBNAME | PREESTNO | JOBNO | SEQ | JOBGUBN | CREPART | BUDGETAMT | ENDFLAG | CREGUBN | JOBBASE | EMPNO | EMPNAME | DEPTCD | DEPTNAME | EXCLIENTCODE | EXCLIENTNAME | BIGO | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | PROJECTNO | RANKJOB"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "PREESTNO | EMPNO | DEPTCD | EXCLIENTCODE | PROJECTNO | RANKJOB", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "JOBNAME|BIGO",-1,-1,0,2,false ' 왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CHK | CREDAY | JOBNO | ENDFLAG | CREGUBN | JOBBASE | EMPNAME | DEPTNAME | EXCLIENTNAME | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | JOBGUBN | CREPART | SEQ",-1,-1,2,2,false '가운데
				
        .cmbPOPUPTYPE.value=1	
        
        .sprSht_PROJECT.style.visibility = "visible"
        .sprSht_DTL.style.visibility = "visible"
	end with
	InitPageData
end Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		
		.sprSht_PROJECT.MaxRows = 0
		.sprSht_DTL.maxRows = 0
		DateClean
		
		call COMBO_TYPE()
	End with
	'새로운 XML 바인딩을 생성
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		'.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub

sub COMBO_TYPE()
   	Dim vntGROUPGUBN
    With frmThis   
		On error resume next
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntGROUPGUBN = mobjPDCOPONO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"PONOGUBN")  'JOB종류 호출

		if not gDoErrorRtn ("COMBO_TYPE") then 

			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_PROJECT, "GROUPGBN",,,vntGROUPGUBN,,60 
			mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",-1, "1"
			mobjSCGLSpr.TypeComboBox = True 
			 gLoadComboBox .cmbGROUPGUBN, vntGROUPGUBN, False
   		end if    
   	end with     	
End Sub	

Sub EndPage()
	set mobjPDCOPONO = Nothing
	set mobjPDCOGET = Nothing
	set mobjPDCOJOBNO = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht_PROJECT.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		'세금계산서 완료조회
		vntData = mobjPDCOPONO.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtPROJECTNM1.value),Trim(.txtPROJECTNO1.value),Trim(.txtCLIENTNAME1.value),Trim(.txtCLIENTCODE1.value),"AA",Trim(.cmbPOPUPTYPE.value))
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_PROJECT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht_PROJECT,meCLS_FLAG
			gWriteText lblstatus_hdr, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			If mlngRowCnt = 0 Then
				.sprSht_PROJECT.MaxRows = 0	
				.sprSht_DTL.maxRows= 0
			else
				Call sprSht_PROJECT_Click(1,1)
			End If

		End If		
	END WITH
End Sub

Sub SelectRtn_DTL (ByVal strPONO)
	Dim vntData
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		vntData = mobjPDCOJOBNO.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strPONO)
		
		If not gDoErrorRtn ("SelectRtn_DTL") then
			If mlngRowCnt > 0 Then
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				mobjSCGLSpr.SetFlag  frmThis.sprSht_DTL,meCLS_FLAG
				
				gWriteText lblstatus_dtl, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
				If mlngRowCnt < 1 Then  '조회된값 없으면
					frmThis.sprSht_DTL.MaxRows = 0   '로우를 0으로 하고
				Else     '조회된값 있으면
					For intCnt = 1 To .sprSht_DTL.MaxRows '조회된 내역을 처음부터 끝까지 돌면서
						'JOB별 컬러 통일
						If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"RANKJOB",intCnt) Mod 2 = "0" Then
							mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
							mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
						'의뢰인 경우 CHK Lock 풀기
						if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ENDFLAG",intCnt)  = "의뢰" Then
							mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,false,"CHK",intCnt,intCnt,false
						Else
							mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,"CHK",intCnt,intCnt,false
						End If
					Next
			  End If
			ELSE
				.sprSht_DTL.MaxRows = 0
			END IF
		END IF
	End with
End SUB

'------------------------------------------
' 데이터 처리
'------------------------------------------
Sub ProcessRtn_PROJECT ()
    Dim intRtn
  	Dim vntData
  	Dim vntData1
	Dim intRtnSave
	Dim strPROJECTNO
	Dim intCnt
	Dim intEDITCODE
	Dim strPROJECTLIST
	Dim strDataCHK
	Dim lngCol, lngRow
	
	with frmThis
		If .sprSht_PROJECT.MaxRows = 0 Then
			gErrorMsgBox "저장할 내역이 존재 하지 않습니다.","저장안내"
			Exit Sub
		End IF
   		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_PROJECT, "PROJECTNM | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO | GROUPGBN",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 프로젝트명/광고주/팀/브랜드/담당부서/담당사원/그룹구분은 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		End If
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_PROJECT,"CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | PROJECTNO | CLIENTCODE | SUBSEQ | CPDEPTCD | CPEMPNO | TIMCODE")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		'if PROJECT_DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strPROJECTNO = ""
		strPROJECTLIST = ""
		
		intRtn = mobjPDCOPONO.ProcessRtnSheet_Insert(gstrConfigXml,vntData, strPROJECTNO, strPROJECTLIST)
		
		if not gDoErrorRtn ("ProcessRtnSheet_Insert") then
			mobjSCGLSpr.SetFlag  .sprSht_PROJECT,meCLS_FLAG
			gErrorMsgBox " 자료가" & intRtn & " 건 저장" & mePROC_DONE,"저장안내" 
			SelectRtn

			For intCnt = 1 To .sprSht_PROJECT.MaxRows 
				If strPROJECTNO = mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",intCnt) Then
					intEDITCODE = intCnt 
					Exit For
				End If
			Next
			
			.txtFROM.focus()
			.sprSht_PROJECT.focus()
			mobjSCGLSpr.ActiveCell .sprSht_PROJECT, 2,intEDITCODE
			sprSht_PROJECT_Click 2,intEDITCODE
				
  		end if
  		
  		vntData1 = mobjPDCOPONO.SelectRtn_PROJECTLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strPROJECTLIST)
  		
  		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1		
		Dim IF_GUBUN : IF_GUBUN = "RMS_0011"
		Dim intCol, intRow, i
		
		intCol = ubound(vntData1, 1)
		intRow = ubound(vntData1, 2)
		
		
		For i = 1 To intRow
			strIF_CNT = strIF_CNT + 1
			
			if strIF_CNT = "1" then
				strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
								vntData1(0,i) + "|" + _
								vntData1(1,i) + "|" + _
								vntData1(2,i) + "|" + _
								vntData1(3,i) + "|" + _
								vntData1(4,i) + "|" + _
								vntData1(5,i) + "|" + _
								vntData1(6,i) + "|" + _
								vntData1(7,i) + "|" + _
								vntData1(8,i)
			else
				strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
								vntData1(0,i) + "|" + _
								vntData1(1,i) + "|" + _
								vntData1(2,i) + "|" + _
								vntData1(3,i) + "|" + _
								vntData1(4,i) + "|" + _
								vntData1(5,i) + "|" + _
								vntData1(6,i) + "|" + _
								vntData1(7,i) + "|" + _
								vntData1(8,i)
			end if
		
			strHSEQ = strHSEQ+1
		Next
		
		
		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
		
 	end with
End Sub

Function PROJECT_DataValidation ()
	PROJECT_DataValidation = false
   	Dim intCnt
   	Dim intCntChk
   	Dim intChk
   	
	On error resume next
	with frmThis
		intChk= 0
		For intCntChk = 1 To .sprSht_PROJECT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",intCnt) = "" Then
			Else
				intChk = intChk +1
			End If
		Next
		If intChk = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
  			Exit Function
		End If
		
  		for intCnt = 1 to .sprSht_PROJECT.MaxRows
  			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",intCnt) = "" Then
  				'필수 항목 체크
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNM",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 프로젝트명은 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CLIENTCODE",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 광고주는 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"TIMCODE",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 팀은 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"SUBSEQ",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 브랜드는 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTCD",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 담당부서는 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPEMPNO",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 담당사원은 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"GROUPGBN",intCnt) = "" Then
  					gErrorMsgBox intCnt & " 행의 그룹구분은 필수 입력 사항입니다.","저장안내!"
  					Exit Function
  				End If
  			End If
		next
	End with

	PROJECT_DataValidation = true
End Function


'자료삭제
Sub DeleteRtn_PROJECT ()
	Dim intSelCnt, intRtn, i , intCnt,intCnt2
	Dim vntData
	Dim strPROJECTNO , strCODE
	Dim intDelCount
	Dim intColSum
	
	with frmThis
	

		'저장시 체크된것만 저장될수 있도록
		intColSum = 0
  		for intCnt2 = 1 to .sprSht_PROJECT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "선택된 데이터가 없습니다.","삭제안내"
			exit Sub
		End If

		'JOB이 등록되어 있는지 확인
		If .sprSht_DTL.MaxRows <> 0  Then
			gErrorMsgBox "등록된JOBNO 가 있습니다.","삭제안내"
		End If
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		

		intDelCount = 0
		'실제삭제 ; 선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_PROJECT.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",i) = 1 then
				strPROJECTNO = mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",i)
				'자료 삭제
				intRtn = mobjPDCOPONO.DeleteRtn(gstrConfigXml,strPROJECTNO)
				
				IF not gDoErrorRtn ("DeleteRtn") then
					mobjSCGLSpr.DeleteRow .sprSht_PROJECT,i
   				End IF
   			
			End If
   			intDelCount = intDelCount + 1
   			gWriteText lblstatus_hdr, "선택한 자료에 대해서 " & intDelCount & " 건이 삭제" & mePROC_DONE	
   		next
			
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht_PROJECT
		SelectRtn
	End with
	err.clear
End Sub


'자료삭제
Sub DeleteRtn_DTL ()
    Dim intSelCnt, intRtn, i , intCnt,intCnt2
	Dim vntData
	Dim strCODE
	Dim intDelCount
	Dim intColSum
	Dim strENDFLAG
	with frmThis
	
		'저장시 체크된것만 저장될수 있도록
		intColSum = 0
  		for intCnt2 = 1 to .sprSht_DTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "선택된 데이터가 없습니다.","삭제안내"
			exit Sub
		End If
			
		for i = .sprSht_DTL.MaxRows to 1 step -1
			strCODE = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 then
				strENDFLAG = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ENDFLAG",i)
				If strENDFLAG <> "의뢰" Then
					gErrorMsgBox "[" & i & "행] 의진행상태가 의뢰가 아닌건은 삭제하실수 없습니다.","삭제안내!"
					Exit Sub
				ELSE
					mlngRowCnt=clng(0) : mlngColCnt=clng(0)
					
					strCODE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"JOBNO",i)
					
					vntData = mobjPDCOJOBNO.GetJOBNOSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
					If mlngRowCnt <> 0 Then
						gErrorMsgBox "해당 JOBNO 의 견적내역 또는 외주정산내역을 확인하십시오","처리안내"
						Exit Sub
					End If
				End If
			End If
   		next
			
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht_DTL.MaxRows to 1 step -1
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			strCODE = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 then
				strCODE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"JOBNO",i)
				
				intRtn = mobjPDCOJOBNO.DeleteRtn(gstrConfigXml,strCODE)
			End IF
		next	
		'선택 블럭을 해제
		
		mobjSCGLSpr.DeselectBlock .sprSht_DTL
		'SelectRtn
		sprSht_PROJECT_click 2,.sprSht_PROJECT.ActiveRow
		.txtCLIENTCODE1.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.sprSht_DTL.Focus
	End with
	err.clear
	
End Sub

-->
		</SCRIPT>
		<script language="javascript">
		//##########################################################################################################################################
		//******************************************주1) frmSapCon 아이 프레임 을 이용하여 Submit 하는 함수
		//##########################################################################################################################################

		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {
		
			//헤더
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			
			//dtl
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			//EAI 서비스 종료 되서 더이상 보내지 않음.
			//window.frames[0].document.forms[0].submit();
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<td style="WIDTH: 400px" height="28" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="110" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td id="tblTitleName" class="TITLE">프로젝트/JOB 관리</td>
										</tr>
									</table>
								</td>
								<TD height="28" vAlign="middle" width="640" align="right">
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 350px"
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
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
							<TR>
								<TD style="HEIGHT: 10px" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="left">
									<TABLE id="tblKey0" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
										align="left">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call DateClean()" width="80">등록일</TD>
											<TD class="SEARCHDATA" width="230"><INPUT accessKey="DATE" style="WIDTH: 80px; HEIGHT: 22px" id="txtFROM" class="INPUT" title="기간검색(FROM)"
													maxLength="10" size="6" name="txtFROM">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarFROM1" align="absMiddle" src="../../../images/btnCalEndar.gIF"
													height="15">&nbsp;~ <INPUT accessKey="DATE" style="WIDTH: 80px; HEIGHT: 22px" id="txtTO" class="INPUT" title="기간검색(TO)"
													maxLength="10" size="7" name="txtTO">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarTO1" align="absMiddle" src="../../../images/btnCalEndar.gIF"
													height="15"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80"><FONT face="굴림">광고주</FONT></TD>
											<TD class="SEARCHDATA" width="260"><FONT face="굴림"><FONT face="굴림"><INPUT style="WIDTH: 179px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="코드명"
															maxLength="100" size="24" name="txtCLIENTNAME1"></FONT> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
													<INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT" title="코드입력"
														maxLength="6" size="4" name="txtCLIENTCODE1"></FONT></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtPROJECTNO1, txtPROJECTNM1)"
												width="80"><SELECT style="WIDTH: 88px" id="cmbPOPUPTYPE" title="프로젝트,JOBNO선택" name="cmbPOPUPTYPE">
													<OPTION selected value="1">PROJECT</OPTION>
													<OPTION value="2">JOBNO</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA"><FONT face="굴림"><INPUT style="WIDTH: 142px; HEIGHT: 22px" id="txtPROJECTNM1" class="INPUT_L" title="코드명"
														maxLength="100" size="18" name="txtPROJECTNM1"> <IMG style="CURSOR: hand" id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgPROJECTNO1" align="absMiddle" src="../../../images/imgPopup.gIF">
													<INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtPROJECTNO1" class="INPUT" title="코드" maxLength="7"
														size="4" name="txtPROJECTNO1"></FONT></TD>
											<td class="SEARCHDATA" width="53"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="자료를 검색합니다." align="right"
													src="../../../images/imgQuery.gIF" height="20"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<tr>
								<td>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="13">
										<TR>
											<TD style="WIDTH: 1040px; HEIGHT: 25px" class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="20" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="97" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="2" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td height="3"></td>
													</tr>
													<tr>
														<td class="TITLE">프로젝트 리스트</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgProjectNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" border="0" name="imgProjectNew"
																alt="신규자료를 작성합니다." src="../../../images/imgNew.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgProjectSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgProjectSave"
																alt="자료를 저장합니다." src="../../../images/imgSave.gIF" width="54" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgProjectDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" border="0" name="imgProjectDelete"
																alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" height="20"></td>
														<TD><IMG style="CURSOR: hand" id="imgProjectExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgProjectExcel"
																alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							<!--BodySplit Start-->
							<TR>
								<TD style="WIDTH: 1040px" class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 40%" vAlign="top" align="left">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: visible" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 95%" id="sprSht_PROJECT" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="4709">
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
										<!--/DIV--></DIV>
								</TD>
							</TR>
							<TR>
								<TD id="lblstatus_hdr" class="BODYSPLIT"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="20" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="67" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="2" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td height="3"></td>
													</tr>
													<tr>
														<td class="TITLE">JOB 리스트</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton1" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgDTLNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" border="0" name="imgDTLNew"
																alt="신규자료를 작성합니다." src="../../../images/imgNew.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgJobDetail" onmouseover="JavaScript:this.src='../../../images/imgDetailOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgDetail.gif'" border="0" name="imgJobDetail"
																alt="자료를 상세보기합니다." src="../../../images/imgDetail.gIF" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgJobDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" border="0" name="imgJobDelete"
																alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" height="20"></td>
														<TD><IMG style="CURSOR: hand" id="imgJobExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgJobExcel"
																alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px" class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 55%" vAlign="top" align="left">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 95%; VISIBILITY: visible" id="pnlTab2"
										ms_positioning="GridLayout">
										<!--DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 1038px; POSITION: relative" ms_positioning="GridLayout"-->
										<OBJECT style="WIDTH: 100%; HEIGHT: 93.13%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="6217">
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
										<!--/DIV--></DIV>
								</TD>
							</TR>
							<!--Bottom Split Start-->
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<TR>
					<TD id="lblStatus_dtl" class="BOTTOMSPLIT"></TD>
				</TR>
				<!--Top TR End--></TABLE>
			<!--Main End--></form>
		</TR></TBODY></TABLE> <iframe id="frmSapCon" style="WIDTH: 100%; DISPLAY: none; HEIGHT: 300px" src="../../../PD/WebService/PROJECTWEBSERVICE.aspx"
			name="frmSapCon"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
