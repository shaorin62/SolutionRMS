<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOBRANDDTLLIST_SRC.aspx.vb" Inherits="SC.SCCOBRANDDTLLIST_SRC" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>브랜드관리-등록</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOCUSTEXELIST.aspx
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By KTY
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
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOBRANDLIST '공통코드, 클래스
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' 명령 버튼 클릭 이벤트
'---------------------------------------------------
'초기화버튼
Sub imgCho_onclick
	InitPageData
End Sub

'-----------------------------------
'조회
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'행추가
'-----------------------------
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtCLIENTNAME.focus
		.sprSht.focus
	End With 
end sub

'-----------------------------------
' 저장   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'수정요청
'-----------------------------
Sub imgReg_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "승인요청할 데이터가 없습니다.","저장안내"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn_Conf
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'삭제
'-----------------------------------
Sub imgDelete_onclick ()
	Dim i
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 닫기
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub



Sub txtCLIENTNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtSEQNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtHIGHSEQNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt

	With frmThis
		'팀 이벤트
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME", Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, trim(vntData(4,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, trim(vntData(5,1))
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		'사업부 이벤트
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  "", "", "",  strCodeName)

				If not gDoErrorRtn ("GetTIMCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, trim(vntData(3,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, trim(vntData(4,1))
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		'광고주이벤트
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row))
			'명이 수정되면 코드를 지운다.
			mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		'부서 이벤트
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						IF .txtUSERDEPT_CD.value = trim(vntData(0,1)) THEN
							mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
							mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						ELSE
							if (.txtUSERDEPT_CD.value = "11002739" and trim(vntData(0,1)) = "11001435") or _
								(.txtUSERDEPT_CD.value = "11002284" and trim(vntData(0,1))= "11001438") or _ 
								(.txtUSERDEPT_CD.value = "11002740" and trim(vntData(0,1)) = "10001210") or _
								(.txtUSERDEPT_CD.value = "11002549" and trim(vntData(0,1)) = "11000533") or _
								(.txtUSERDEPT_CD.value = "11002286" and trim(vntData(0,1)) = "10000017") or _
								(.txtUSERDEPT_CD.value = "11003433" and trim(vntData(0,1)) = "11001312") then 
								
									
								mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
								mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
							else
								gErrorMsgBox "로그인 사용자의 부서와 입력한 실적 부서가 동일해야 합니다." & vbCrlf & " " & vbCrlf & "문의사항은 김노연 팀장에게 문의하세요.","입력안내!"
								mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
								mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, ""
								EXIT Sub
							end if
						END IF
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
	
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TIMNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			vntRet = gShowModalWindow("SCCOTIMPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(5,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTSUBNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row)))
								
			vntRet = gShowModalWindow("SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			
		End If	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			vntRet = gShowModalWindow("SCCODEPTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				IF .txtUSERDEPT_CD.value = trim(vntRet(0,0)) THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntRet(1,0))
				ELSE
					if (.txtUSERDEPT_CD.value = "11002739" and trim(vntRet(0,0)) = "11001435") or _
						(.txtUSERDEPT_CD.value = "11002284" and trim(vntRet(0,0))= "11001438") or _ 
						(.txtUSERDEPT_CD.value = "11002740" and trim(vntRet(0,0)) = "10001210") or _
						(.txtUSERDEPT_CD.value = "11002549" and trim(vntRet(0,0)) = "11000533") or _
						(.txtUSERDEPT_CD.value = "11002286" and trim(vntRet(0,0)) = "10000017") or _
						(.txtUSERDEPT_CD.value = "11003433" and trim(vntRet(0,0)) = "11001312")   then 
							
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntRet(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntRet(1,0))
					else
						gErrorMsgBox "로그인 사용자의 부서와 입력한 실적 부서가 동일해야 합니다." & vbCrlf & " " & vbCrlf & "문의사항은 김노연 팀장에게 문의하세요.","입력안내!"
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, ""
						EXIT Sub
					end if
				END IF
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.txtCLIENTNAME.focus
		.sprSht.Focus
	End With
End Sub

'-----------------------------------
'쉬트 더블클릭
'-----------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End if
	End With
End Sub

'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,true,frmThis.sprSht.ActiveRow,5,5,true
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ATTR01",frmThis.sprSht.ActiveRow, "등록"
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME.focus
		frmThis.sprSht.focus
	End If
End Sub

'--------------------------------------------------
'쉬트 버튼클릭
'--------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNTIM") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row)))
			vntRet = gShowModalWindow("SCCOTIMPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"TIMNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(5,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUB") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row)))
								
			vntRet = gShowModalWindow("SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			
		End If	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNDEPT") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			vntRet = gShowModalWindow("SCCODEPTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				IF .txtUSERDEPT_CD.value = trim(vntRet(0,0)) THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntRet(1,0))
				ELSE
					if (.txtUSERDEPT_CD.value = "11002739" and trim(vntRet(0,0)) = "11001435") or _
						(.txtUSERDEPT_CD.value = "11002284" and trim(vntRet(0,0))= "11001438") or _ 
						(.txtUSERDEPT_CD.value = "11002740" and trim(vntRet(0,0)) = "10001210") or _
						(.txtUSERDEPT_CD.value = "11002549" and trim(vntRet(0,0)) = "11000533") or _
						(.txtUSERDEPT_CD.value = "11002286" and trim(vntRet(0,0)) = "10000017") or _
						(.txtUSERDEPT_CD.value = "11003433" and trim(vntRet(0,0)) = "11001312")   then 
							
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntRet(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntRet(1,0))
					else
						gErrorMsgBox "로그인 사용자의 부서와 입력한 실적 부서가 동일해야 합니다." & vbCrlf & " " & vbCrlf & "문의사항은 김노연 팀장에게 문의하세요.","입력안내!"
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, ""
						EXIT Sub
					end if
				END IF
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
		End If	
		.txtCLIENTNAME.focus
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	End With
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjSCCOBRANDLIST = gCreateRemoteObject("cSCCO.ccSCCOBRANDLIST")
	set mobjSCCOGET		  = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
	
	gSetSheetColor mobjSCGLSpr, .sprSht	
	mobjSCGLSpr.SpreadLayout .sprSht, 20, 0, 0, 0,0
	mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.AddCellSpan  .sprSht, 9, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.AddCellSpan  .sprSht, 12, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.AddCellSpan  .sprSht, 15, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.SpreadDataField .sprSht, "CHK | ATTR01 | SEQNO | SEQNAME | HIGHSEQNO | TIMCODE | BTNTIM | TIMNAME | CLIENTSUBCODE | BTNSUB | CLIENTSUBNAME | CUSTCODE | BTN | CUSTNAME | DEPT_CD | BTNDEPT | DEPT_NAME | CUSER | CDATE | MEMO"
	mobjSCGLSpr.SetHeader .sprSht,		 "선택|사용구분|코드|브랜드명|대표브랜드|팀코드|팀명|CIC코드|CIC/사업부|광고주코드|광고주명|부서코드|부서명|등록자|등록일|비고"
	mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|      10|   7|      15|        12|     7|2|10|      9|2|       9|         8|2|    15|       8|2|  10|     8|    12|  15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
	mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "ATTR01", -1, -1, "사용" & vbTab & "미사용" & vbTab & "등록" & vbTab & "승인요청" , 10, 60, FALSE, FALSE
	mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTNTIM | BTNSUB | BTN | BTNDEPT"
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SEQNAME | TIMNAME | CLIENTSUBNAME | CUSTNAME | DEPT_NAME | CUSER | CDATE | MEMO", -1, -1, 200
	mobjSCGLSpr.SetCellsLock2 .sprSht, True, "SEQNO | ATTR01 | CUSER | CDATE"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "HIGHSEQNO | TIMCODE | CLIENTSUBCODE | CUSTCODE | DEPT_CD" ,-1,-1,2,2,false
	
	.sprSht.style.visibility = "visible"

    End With

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOBRANDLIST = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis

	'초기 데이터 설정
	With frmThis
		.sprSht.MaxRows = 0
		Set_COMBO
		Get_SESSION_DEPT_CD
	End With
End Sub

Sub Set_COMBO ()
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOBRANDLIST.GET_HighSeq_COMBO(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "HIGHSEQNO",,,vntData,,100 
		mobjSCGLSpr.TypeComboBox = true 
		
   	End With
End Sub


Sub Set_RowCOMBO (strCLIENTCODE, Row)
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim RowCnt, ColCnt
   	
	With frmThis
		'Long Type의 ByRef 변수의 초기화
		RowCnt=clng(0)
		ColCnt=clng(0)
		
		vntData = mobjSCCOBRANDLIST.GET_HighSeq_COMBO_ROW(gstrConfigXml,RowCnt,ColCnt, strCLIENTCODE)
		
		mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "HIGHSEQNO",Row,Row, vntData,10,100, false, false
		mobjSCGLSpr.TypeComboBox = true 
		
   	End With
End Sub


Sub Get_SESSION_DEPT_CD ()
	Dim strDEPT_CD
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strDEPT_CD = mobjSCCOBRANDLIST.Get_SESSION_DEPT_CD(gstrConfigXml,mlngRowCnt,mlngColCnt, gstrEmpNo)
		
		If not gDoErrorRtn ("Get_SESSION_DEPT_CD") Then 
			.txtUSERDEPT_CD.value = strDEPT_CD
   		End If
   		.txtUSERDEPT_CD.style.visibility = "HIDDEN"
   	End With
End Sub
'------------------------------------------
' HDR 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTNAME, strSEQNAME, strHIGHSEQNAME, strUSE_YN
   	Dim intCnt, intCnt2, strRows
   	Dim dblcnt
   	
	With frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'변수 초기화
		strCLIENTNAME = ""
		strSEQNAME = ""
		strHIGHSEQNAME = ""
		strUSE_YN = ""
		dblcnt = true
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value 
		strSEQNAME		= .txtSEQNAME.value
		strHIGHSEQNAME  = .txtHIGHSEQNAME.value
		strUSE_YN		= .cmbUSE_YNSEARCH.value
		
		vntData = mobjSCCOBRANDLIST.SelectRtn_SUBSEQ_SRC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCLIENTNAME, strSEQNAME, strHIGHSEQNAME, strUSE_YN)

		If not gDoErrorRtn ("SelectRtn_SUBSEQ") Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			FOR i =1 TO .sprSht.MaxRows
				Call Set_RowCOMBO (mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",i), i)
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",i) = "사용" or mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",i) = "미사용" Then
					If dblcnt Then
						strRows = i
						dblcnt = false
					Else
						strRows = strRows & "|" & i
					End If
				End If
			Next
			
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,18,True
			mobjSCGLSpr.SetCellsLock2 .sprSht, false, "DEPT_CD | BTNDEPT | DEPT_NAME"
			
   			gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
   		End if
   	End With
End Sub

'------------------------------------------
' HDR 수정/저장 처리 
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	Dim strYEAR
	Dim returnvalue
	Dim strRETURNSEQNO
	Dim i
	Dim intRtn2
	
	With frmThis
   		'데이터 Validation
		'if DataValidation =false then exit sub
		'On error resume next
		
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "SEQNAME | CUSTCODE | CUSTNAME | DEPT_CD",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 브랜드명/광고주/부서는 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | SEQNO | SEQNAME | HIGHSEQNO | TIMCODE | BTNTIM | TIMNAME | CLIENTSUBCODE | BTNSUB | CLIENTSUBNAME | CUSTCODE | BTN | CUSTNAME | DEPT_CD | BTNDEPT | DEPT_NAME | MEMO | ATTR01")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If
		
		intRtn2 = gYesNoMsgbox("실적부서를 확인하셨습니까?" & vbCrlf & " " & vbCrlf & " " & vbCrlf & "브랜드에 잘못 매칭된 부서로 청약될 경우 " & vbCrlf & " " & vbCrlf & "실적부서는 추후에 부서별 실적을 집계하는데 중대한 영향을 미칠수 있습니다. "& vbCrlf & " " & vbCrlf & "반드시 확인 바랍니다. ","확인")
		IF intRtn2 <> vbYes then exit Sub
		
		strYEAR = Mid(gNowDate,3,2)
		
		intRtn = mobjSCCOBRANDLIST.ProcessRtn_SUBSEQ(gstrConfigXml,vntData, strYEAR)
		
		returnvalue = Split(intRtn, "-")
		
		If not gDoErrorRtn ("ProcessRtn_SUBSEQ") Then
			If isArray(returnvalue) Then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				
				gOkMsgBox  "자료가 저장" & mePROC_DONE,"저장안내!"
				
				SelectRtn
				for i=1 to .sprSht.MaxRows
					strRETURNSEQNO = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNO",i)
					IF strRETURNSEQNO = returnvalue(1) THEN
						mobjSCGLSpr.ActiveCell .sprSht, 1, i
						EXIT FOR
					END IF
				Next
			End if
   		End If
   		.txtCLIENTNAME.focus()
		
   	End With
End Sub


Sub ProcessRtn_CONF ()
	Dim vntData
	Dim intCnt, intCnt2, intCnt3, intRtn, i
	Dim strSEQNO
	Dim strMsg
	Dim strMstMsg
	'SMS 정보
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim intMsgCnt
	Dim vntData_info
	Dim strMstEmail
		
	With frmThis
		intMsgCnt = 0
		intCnt3 = true
		strMstMsg = ""
		strMstEmail = ""
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",intCnt2) <> "등록" then
					gErrorMsgBox "체크한 데이터 중 " +  cstr(intCnt2) + " 번째 줄의 상태는 등록상태가 아닙니다. 등록상태인 데이터만 승인요청 할 수 있습니다.","승인요청안내!"
					Exit Sub
				end if 
				
				if mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNO",intCnt2) = "" then
					gErrorMsgBox "체크한 데이터 중 " +  cstr(intCnt2) + " 번째 줄의 상태는 저장되지 않은 데이터 입니다. 저장을 먼저 하신 후에 해당 데이터를 승인요청할 수 있습니다..","승인요청안내!"
					Exit Sub
				end if 
				
				If intCnt3 Then
					strMsg = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNAME",intCnt2)
					intCnt3 = false
				End If
				intMsgCnt = intMsgCnt +1
			end if
		Next
	
		If intMsgCnt = 0 Then
			gErrorMsgBox "승인요청할 데이터를 체크해 주세요.","승인요청안내!"
			EXIT Sub
		End If

		If intMsgCnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] 승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 승인요청이있습니다"
			End If
			
			strMstEmail = "[ " & strMsg & " ]"
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] 외" & intMsgCnt-1 & "건의승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 외" & intMsgCnt-1 & "건의승인요청이있습니다"
			End If
			
			strMstEmail = "[ " & strMsg & " ] 외"
		End If
		
		
		intRtn = gYesNoMsgbox("자료를 승인요청 하시겠습니까?","승인요청 확인")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strSEQNO = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNO",i)
				
				If strSEQNO = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjSCCOBRANDLIST.ProcessRtn_CONF(gstrConfigXml,strSEQNO)
				End If				
  				intCnt = intCnt + 1
 			End If
		Next
		
		If not gDoErrorRtn ("ProcessRtn_CONF") Then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			
			'승인을 수락하였으므로 SMS 발송
			'보내는 사람의 정보 가져오기 'gstrEmpNo, gstrUsrName
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData_info = mobjSCCOGET.Get_SENDINFO2(gstrConfigXml,mlngRowCnt,mlngColCnt, "1001499")
			
			'보내는사람정보
			strFromUserName		= vntData_info(0,2)
			strFromUserEmail	= vntData_info(1,2)
			strFromUserPhone	= vntData_info(2,2)
			
			'받는사람 정보
			strToUserName		=  vntData_info(0,1)
			strToUserEmail		=  vntData_info(1,1)
			strToUserPhone		=  vntData_info(2,1)
			
		
			call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)

			call EMAIL_SEND(strMstEmail, "브랜드", strFromUserName,strFromUserEmail,strToUserEmail)
			
			gErrorMsgBox "자료가 승인요청 되었습니다.","승인요청안내!"
   		End If

		SelectRtn
	End With
	
	err.clear	
End Sub

'------------------------------------------
'데이터 삭제
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strSUBSEQ
	Dim strSUBSEQ2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"SEQNO",i)
				If strSUBSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					vntData = mobjSCCOBRANDLIST.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strSUBSEQ, "S") 
					If mlngRowCnt > 0 Then
						strMSG = ""
						For intCnt = 0 To mlngRowCnt-1
							If vntData(0,intCnt) = "B" Then
								strMSG = strMSG & " 인쇄: " & vntData(1,intCnt) & "건" 
							ElseIf vntData(0,intCnt) = "A2" Then
								strMSG = strMSG & " 케이블: " & vntData(1,intCnt) & "건" 
							ElseIf vntData(0,intCnt) = "A" Then
								strMSG = strMSG & " 공중파: " & vntData(1,intCnt) & "건" 
							ElseIf vntData(0,intCnt) = "O" Then
								strMSG = strMSG & " 인터넷: " & vntData(1,intCnt) & "건" 
							ElseIf vntData(0,intCnt) = "D" Then
								strMSG = strMSG & " 옥외: " & vntData(1,intCnt) & "건" 
							End If
						Next
						gErrorMsgBox i & "행의 코드는 " & strMSG & " 이 청약데이터로 저장되어있습니다.","삭제안내!"
						Exit Sub
					End If
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "삭제할 데이터를 체크해 주세요.","삭제안내!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		For i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strSUBSEQ2 = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNO",i)
			
				If strSUBSEQ2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					intRtn = mobjSCCOBRANDLIST.DeleteRtn_DTL(gstrConfigXml, strSUBSEQ2)
					
					IF not gDoErrorRtn ("DeleteRtn_DTL") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn_DTL") Then
   			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End If
		SelectRtn
	End With
	err.clear
End Sub
-->
		</script>
		<script language="javascript">
		//SMS 발송
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
		function EMAIL_SEND(strMstEmail, strGBN, strFromUserName,strFromUserEmail,strToUserEmail){
			frmEMAIL.location.href = "../../../SC/SrcWeb/SCCO/SENDEMAIL.asp?NAME="+ strMstEmail + "&GBN=" + strGBN + "&FromUserName=" + strFromUserName + "&FromUserEmail=" + strFromUserEmail + "&ToUserEmail=" + strToUserEmail; 
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
											<td class="TITLE">
												<P>브랜드 관리-등록</P>
											</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
								<TD align="left" width="100%" height="1">
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,'')"
												width="60">광고주명</TD>
											<TD class="SEARCHDATA" width="150"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 144px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" name="txtCLIENTNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSEQNAME,'')"
												width="60">브랜드명</TD>
											<TD class="SEARCHDATA" width="150"><INPUT class="INPUT_L" id="txtSEQNAME" title="브랜드명" style="WIDTH: 144px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="18" name="txtSEQNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtHIGHSEQNAME,'')"
												width="80">대표브랜드명</TD>
											<TD class="SEARCHDATA" width="150"><INPUT class="INPUT_L" id="txtHIGHSEQNAME" title="대표브랜드명" style="WIDTH: 144px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="16" name="txtHIGHSEQNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtHIGHSEQNAME,'')"
												width="60">사용구분</TD>
											<TD class="SEARCHDATA"><SELECT id="cmbUSE_YNSEARCH" title="사용구분" style="WIDTH: 104px" name="cmbUSE_YNSEARCH">
													<OPTION value="">전체</OPTION>
													<OPTION value="Y" selected>사용</OPTION>
													<OPTION value="N">미사용</OPTION>
													<OPTION value="R">등록</OPTION>
													<OPTION value="S">승인요청</OPTION>
												</SELECT><INPUT class="INPUT_L" id="txtUSERDEPT_CD" title="사용자부서" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="20" name="txtUSERDEPT_CD"></TD>
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
									</TABLE>
								</TD>
							<tr>
								<td>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20"></TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="자료를 인쇄합니다." src="../../../images/imgCho.gIF"
																border="0" name="imgCho"></TD>
														<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgReg" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'"
																height="20" alt="자료를 등록요청합니다." src="../../../images/ImgConfirmRequest.gIF" border="0"
																name="imgReg"></TD>
														<!--<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="자료를 인쇄합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>-->
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="16378">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 1000px;HEIGHT: 1000px" name="frmSMS">
		</iframe><!--DISPLAY: none; -->
		<iframe id="frmEMAIL" style="DISPLAY: none;WIDTH: 1000px;HEIGHT: 1000px" name="frmEMAIL">
		</iframe><!--DISPLAY: none; -->
	</body>
</HTML>
