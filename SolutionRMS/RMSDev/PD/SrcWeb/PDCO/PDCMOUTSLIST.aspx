<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMOUTSLIST.aspx.vb" Inherits="PD.PDCMOUTSLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>정산관리</title> 
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
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mcomecalender, mcomecalender2
Dim mobjPDCMOUTSLIST
Dim mobjPDCMGET
Dim mstrCheck

Dim mstrMANAGER '슈퍼관리자 체크
mstrCheck = True
mcomecalender = FALSE
mcomecalender2 = FALSE
Const meTab = 9

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
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSetting_onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmCancel_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Cancel
	gFlowWait meWAIT_OFF
End Sub

Sub imgConfirmFlag_onclick 
	gFlowWait meWAIT_ON
	ProcessRtn_Confirm
	gFlowWait meWAIT_OFF
End SUb

'출력 인쇄버튼 클릭시 이벤트
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j,k
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
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
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDCMELECCOMMILIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECCOMMI.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_ELECCOMMI_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					
					datacnt = strcntsum
					for k=1 to 2
						strUSERID = ""
						vntDataTemp = mobjMDCMELECCOMMILIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
					next
				End IF
			END IF
		next
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "printSetTimeout", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMELECCOMMILIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgVoch_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Vochno
	gFlowWait meWAIT_OFF
End Sub

'정산일적용
Sub ImgSUMMApp_onclick
	Dim intCnt
	Dim intCnt2
	Dim lngCHK
	With frmThis
		lngCHK = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1"  Then
				lngCHK = lngCHK + 1
			End If
		Next
		
		If lngCHK = 0  Then 
			gErrorMsgBox "선택된건이 없습니다.","처리안내"
			Exit Sub
		End If
		
		For intCnt = 1 To .sprSht.MaxRows
			If .cmbCHK.value = "ADJ" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht,"PURCHASENO",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"ADJDAY",intCnt,.txtADJDAY.value 
					sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"ADJDAY"),intCnt
				End if
			ElseIf .cmbCHK.value = "TAX" Then 
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht,"PURCHASENO",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"TAXDATE",intCnt,.txtADJDAY.value 
					sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"TAXDATE"),intCnt
				End if
			End If
		Next
	End With
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
	
	vntInParams = array(.txtREAL_MED_CODE.value, .txtREAL_MED_NAME.value)
		
	vntRet = gShowModalWindow("MDCMREALMEDPOP.aspx",vntInParams , 413,425)
		
	if isArray(vntRet) then
		if .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		.txtREAL_MED_CODE.value = vntRet(0,0)		        ' Code값 저장
		.txtREAL_MED_NAME.value = vntRet(1,0)             ' 코드명 표시
		gSetChangeFlag .txtREAL_MED_CODE                  ' gSetChangeFlag objectID	 Flag 변경 알림
    end if
			
	End with
	
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_CODE_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetREALMEDNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtREAL_MED_CODE.value,.txtREAL_MED_NAME.value)
		
			if not gDoErrorRtn ("GetREALMEDNO") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = vntData(0,0)
					.txtREAL_MED_NAME.value = vntData(1,0)
				Else
					Call REAL_MED_CODE_POP()
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
' 날자컨트롤 및 달력 / Onchange Event
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgTO,"txtTo_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

'ImgADJDAY
Sub ImgADJDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtADJDAY,frmThis.ImgADJDAY,"txtADJDAY_onchange()"
		gSetChange
	end with
End Sub


'****************************************************************************************
' ONCHANGE
'****************************************************************************************
Sub txtSUM_onfocus
	with frmThis
		.txtSUM.value = Replace(.txtSUM.value,",","")
	end with
End Sub
Sub txtSELECTAMT_onfocus
	with frmThis
		.txtSELECTAMT.value = Replace(.txtSELECTAMT.value,",","")
	end with
End Sub
Sub txtSUM_onblur
	with frmThis
		call gFormatNumber(.txtSUM,0,true)
	end with
End Sub
Sub txtSELECTAMT_onblur
	with frmThis
		call gFormatNumber(.txtSELECTAMT,0,true)
	end with
End Sub
Sub txtADJDAY_onchange
	gSetChange
End Sub

Sub txtTRANSYEARMON_onchange
	gSetChange
End Sub

Sub txtTRANSNO_onchange
	gSetChange
End Sub

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
				strFROM = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		DateClean strFROM
	
	End With
	gSetChange
End Sub

Sub txtTO_onchange
	Dim strdate 
	Dim strTO, strTO2
	Dim strOLDYEARMON
	strdate = ""
	strTO =""
	strTO2 = ""
	With frmThis
		strdate=.txtTO.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender Then
			strTO = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strTO2 = strdate
		else
			If len(strdate) = 4 Then
				strTO = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strTO2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strTO = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strTO2 = strdate
			elseif len(strdate) = 3 Then
				strTO = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strTO2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strTO = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strTO2 = strdate
			End If
		End If
		'DateClean strTO
	End With
	gSetChange
End Sub

Sub cmbGUBN_onchange
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


'****************************************************************************************
' 쉬트 더불클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	dim intcnt
	with frmThis
		If Row = 0 and Col = 1 then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,,, , , , , , mstrCheck
				
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			'for intcnt = 1 to .sprSht.MaxRows
			'	sprSht_Change 1, intcnt
			'next
			
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ENDDAY", intCnt) <> "" Then
					'스태틱			
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If	
			Next
		end if
		
		If Row = 0 and Col = 12 then 
			if .cmbGUBN.value = "" then exit sub
			
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 12,12,,, , , , , , mstrCheck
				
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
			
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ENDDAY", intCnt) <> "" Then
					'스태틱			
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 12,12, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CONFIRMFLAG",intCnt," "
				End If		
			Next
		end if
		
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim intCnt
	Dim lngAMT
	Dim lngSUMAMT
	
	if Col = 1 Then
		lngAMT = 0
		lngSUMAMT = 0
		
		For intCnt = 1 To frmThis.sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1" And frmThis.cmbGUBN.value <> " " Then
				lngAMT = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJAMT", intCnt)		
				lngSUMAMT = lngSUMAMT + lngAMT
			End if
		Next
		frmThis.txtSELECTAMT.value = lngSUMAMT
		txtSELECTAMT_onblur
	End if
	
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

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
	
	'키가 움질일때 바인딩
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		'sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")  Then
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
	End With
End Sub


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
	set mobjPDCMOUTSLIST  = gCreateRemoteObject("cPDCO.ccPDCOOUTSLIST") '조회
	set mobjPDCMGET =  gCreateRemoteObject("cPDCO.ccPDCOGET")	  '코드
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
   gSetSheetDefaultColor
    with frmThis
		'화면의 깜박임을 방지하기 위함(Tab의 경우는 처음에 표시되는 것만 함)
		'.sprSht.style.visibility = "hidden"
		
		'**************************************************
		'***첫번째 Sheet 디자인
		'**************************************************
		
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout 디자인
		mobjSCGLSpr.SpreadLayout .sprSht, 18, 0,7
	    mobjSCGLSpr.SpreadDataField .sprSht, "CHK | SEQ | PROJECTNM | JOBNO | JOBNAME  | CLIENTNAME | ITEMNAME | OUTSNAME | ADJAMT | PURCHASENO | VOCHNO | CONFIRMFLAG | REQDAY | DEMANDDAY | ADJDAY | TAXDATE | ENDDAY | OUTSRANK"
		mobjSCGLSpr.SetHeader .sprSht,        "선택|순번|프로젝트|JOBNO|JOB명|광고주명|외주항목|외주처명|정산금액|정산번호|전표|승인|의뢰일|청구일|정산/전표일|증빙일|결산일|랭크"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  4|  0|       10|    7|   20|      15|      12|      15|      10|       8|  10|   4|     8|     8|          9|     9|     8|  0"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | CONFIRMFLAG"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY | DEMANDDAY | ENDDAY | REQDAY|TAXDATE", , , ,3
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SEQ | PROJECTNM | JOBNO | JOBNAME | CLIENTNAME | ITEMNAME | ADJAMT | PURCHASENO | VOCHNO | REQDAY | ENDDAY"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO | SEQ | PURCHASENO | VOCHNO",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PROJECTNM | JOBNAME | CLIENTNAME | OUTSNAME | ITEMNAME",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "SEQ | OUTSRANK|REQDAY | DEMANDDAY|CONFIRMFLAG", true
		'.imgConfirmFlag.style.visibility = "hidden"
		.txtYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		.txtADJDAY.value = gNowDate2
		
	End with
	pnlTab1.style.visibility = "visible" 

	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjPDCMGET = Nothing
	set mobjPDCMOUTSLIST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	with frmThis
		Dim vntData
		DateClean Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		vntData = mobjPDCMOUTSLIST.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		if not gDoErrorRtn ("SelectRtn_USER") then	
			if mlngRowCnt > 0 Then
				mstrMANAGER = vntData(1,1)
				
				If mstrMANAGER = "Y" Then
					mobjSCGLSpr.ColHidden .sprSht, "CONFIRMFLAG", false
					'.imgConfirmFlag.style.visibility = "visible"
				End if
			end if
   		end if	
	End with
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
		frmThis.txtFrom.value = date1
		frmThis.txtTO.value = date2  
	end if
End Sub


'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim i, strCols
	Dim strOUTSCODE
	Dim strOUTSNAME
	Dim strFROM
	Dim strTO
	Dim strGUBN
	Dim strPOPUPTYPE
	Dim intCnt
	Dim lngAMT
	Dim lngSUMAMT

	with frmThis
			call change_Active()
			
			.sprSht.MaxRows = 0
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strOUTSCODE = TRIM(.txtOUTSCODE.value)
			strOUTSNAME =  replace(TRIM(.txtOUTSNAME.value),"'","''")

			strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
			strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
			
			strGUBN = .cmbGUBN.value 
			strPOPUPTYPE =  .cmbPOPUPTYPE.value
			
			vntData = mobjPDCMOUTSLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strOUTSCODE,strOUTSNAME,strFROM,strTO,strGUBN, strPOPUPTYPE)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE			
   					
   					If mlngRowCnt > 0 Then
   					lngSUMAMT = 0
   					lngAMT = 0
   						
   						For intCnt = 1 To .sprSht.MaxRows
								lngAMT = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJAMT", intCnt)		
								lngSUMAMT = lngSUMAMT + lngAMT
								
								If .cmbGUBN.value = "F" Then
									'mobjSCGLSpr.SetCellTypeStatic .sprSht, 11,11, intCnt, intCnt,0,2
									'mobjSCGLSpr.SetTextBinding .sprSht,"VOCHNO",intCnt, " "

								Elseif  .cmbGUBN.value  = ""  then
									If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt) = "1" Then
										mobjSCGLSpr.SetTextBinding .sprSht,"VOCHNO",intCnt,"완료"
									Else
										'mobjSCGLSpr.SetTextBinding .sprSht,"VOCHNO",intCnt," "
									End If
								End If
								
								
								'ENDDAY 확정날짜가 있으면 CHK를 숨긴다.
								If   mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ENDDAY", intCnt) <> "" Then
									
									mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
									mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
									mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,2,2,true
									
									mobjSCGLSpr.SetCellTypeStatic .sprSht, 12,12, intCnt, intCnt,0,2
									mobjSCGLSpr.SetTextBinding .sprSht,"CONFIRMFLAG",intCnt," "
									mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,12,12,true
								Else
									'확정날짜가 없는것은 체크를 보여준다.
									mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
									'확정된데이타면 LOCK을건다
									If .cmbGUBN.value  = "T" Then
										mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,2,2,true
									End If
								End If	
   						Next
   						.txtSUM.value = lngSUMAMT
   						.txtSELECTAMT.value = 0
   					Else
   						.sprSht.MaxRows = 0
   						.txtSELECTAMT.value = 0
   					End If
   			end if
   			txtSELECTAMT_onblur
   			txtSUM_onblur
	End With
End Sub

Sub change_Active
	with frmThis
		if  .cmbGUBN.value  = "" Then
			'.imgSetting.disabled =  true
			'.ImgConfirmCancel.disabled = true
			'.imgVoch.disabled = true
			'.imgConfirmFlag.disabled = true
			mobjSCGLSpr.ColHidden .sprSht, "VOCHNO | CONFIRMFLAG", FALSE
		End If
		if  .cmbGUBN.value  = "F" Then
			'.imgSetting.disabled =  false
			'.ImgConfirmCancel.disabled = true
			'.imgVoch.disabled = true
			'.imgConfirmFlag.disabled = true
			mobjSCGLSpr.ColHidden .sprSht, "VOCHNO | CONFIRMFLAG", true
		end if
		if .cmbGUBN.value = "T" Then
			'.imgSetting.disabled = true
			'.ImgConfirmCancel.disabled = false
			'.imgVoch.disabled = false
			'.imgConfirmFlag.disabled = false
			mobjSCGLSpr.ColHidden .sprSht, "VOCHNO | CONFIRMFLAG", FALSE
		end if	
		if  .cmbGUBN.value  = "V" OR  .cmbGUBN.value  = "C" Then
			'.imgSetting.disabled =  true
			'.ImgConfirmCancel.disabled = true
			'.imgVoch.disabled = false
			'.imgConfirmFlag.disabled = false
			mobjSCGLSpr.ColHidden .sprSht, "VOCHNO | CONFIRMFLAG", FALSE
		End If
	End with
End Sub

'-----------------------------------------------------------------------------------------
' 전표처리 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn_Vochno ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim intCnt2
	Dim lngCHK
	Dim intSaveRtn
	with frmThis
	'On error resume next
		
		if DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SEQ | JOBNO | VOCHNO")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"매입전표처리 확인안내!"
			exit sub
		End If
	
		'처리 업무객체 호출
		intSaveRtn = gYesNoMsgbox("매입전표처리를 하시겠습니까?" & vbcrlf & "체크상태:SAP 전표처리완료확인" & vbcrlf & "체크해지상태:전표 미처리 상태","매입전표처리 확인안내!")
		IF intSaveRtn <> vbYes then exit Sub
		intRtn = mobjPDCMOUTSLIST.ProcessRtn_Voch(gstrConfigXml,vntData)
	
		if not gDoErrorRtn ("ProcessRtn_Voch") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox "매입전표 확인처리 가 되었습니다.","매입전표처리 확인안내!" 
			SelectRtn
  		end if
 	end with
End Sub


'-----------------------------------------------------------------------------------------
' 승인처리 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn_Confirm ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim intCnt2
	Dim lngCHK, intChkCnt
	Dim intSaveRtn
	with frmThis
	'On error resume next
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'For intChkCnt = 1 To .sprSht.MaxRows
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",intChkCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intChkCnt) <> "1" Then
		'		gErrorMsgBox intChkCnt & " 번째 데이터는 매입전표처리가 되어 있습니다." & vbcrlf & "매입전표처리를 취소하시려면 전표 항목의 체크를 해지한후" & vbcrlf & "확정취소 처리를 하십시오.","확정취소안내"
		'		selectrtn
		'		Exit Sub
		'	End If
		'Next
		
		'데이터 Validation
	
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SEQ | JOBNO | CONFIRMFLAG")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장 확인안내!"
			exit sub
		End If
	
		'처리 업무객체 호출
		intRtn = mobjPDCMOUTSLIST.ProcessRtn_Confirm(gstrConfigXml,vntData)
	
		if not gDoErrorRtn ("ProcessRtn_Confirm") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox "저장 되었습니다.","저장안내!" 
			SelectRtn
  		end if
 	end with
End Sub


'-----------------------------------------------------------------------------------------
' 확정 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strYEARMON
	Dim intCnt2
	Dim lngCHK
	Dim intMaxCnt
	Dim intColFlag
	Dim bsdiv
	
	with frmThis
	'On error resume next
		IF .cmbGUBN.value <> "F" THEN
			gErrorMsgBox "확정은 미정산 조회시 가능합니다.","확정안내"
			Exit Sub
		End if
		
		lngCHK = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1"  Then
				lngCHK = lngCHK + 1
			End If
		Next
		If lngCHK = 0  Then 
		gErrorMsgBox "선택된건이 없습니다.","확정안내"
		Exit Sub
		End If
		
  		'데이터 Validation
		if DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | SEQ | PURCHASENO | ADJDAY | JOBNO | OUTSRANK | TAXDATE")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"확정안내!"
			exit sub
		End If
	
		'처리 업무객체 호출
		strYEARMON = Trim(.txtYEARMON.value)
		
		If Len(strYEARMON) <> 6 Or strYEARMON = "" Then
			gErrorMsgBox "정산년월을 확인하십시오","확정안내!"
			Exit Sub
		End if
		
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intMaxCnt) = "1" Then
				bsdiv = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSRANK",intMaxCnt)
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		intRtn = mobjPDCMOUTSLIST.ProcessRtn(gstrConfigXml,vntData,strYEARMON,intColFlag)

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " 확정처리가" & mePROC_DONE,"확정안내" 
			SelectRtn
  		end if
 	end with
End Sub


'-----------------------------------------------------------------------------------------
' 확정 취소 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn_Cancel ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strYEARMON
	Dim intCnt2
	Dim lngCHK
	Dim intChkCnt
	
	with frmThis
	
	'On error resume next
		IF .cmbGUBN.value <> "T" THEN
			gErrorMsgBox "확정취소는 정산 조회시 가능합니다.","확정취소안내"
			Exit Sub
		End if
		
		lngCHK = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1"  Then
				lngCHK = lngCHK + 1
				sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),intCnt2
			End If
		Next
		
		If lngCHK = 0  Then 
			gErrorMsgBox "선택된건이 없습니다.","확정취소안내"
		Exit Sub
		End If
		
  		'데이터 Validation
		'if DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|SEQ|PURCHASENO|ADJDAY|JOBNO|VOCHNO")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"확정취소안내!"
			exit sub
		End If
		
		'For intChkCnt = 1 To .sprSht.MaxRows
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intChkCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intChkCnt) = "1" Then
		'		gErrorMsgBox intChkCnt & " 번째 데이터는 매입전표처리가 되어 있습니다." & vbcrlf & "매입전표처리를 취소하시려면 전표 항목의 체크를 해지한후" & vbcrlf & "확정취소 처리를 하십시오.","확정취소안내"
		'		Exit Sub
		'	End If
		'Next
		
		strYEARMON = Trim(.txtYEARMON.value)
		
		If Len(strYEARMON) <> 6 Or strYEARMON = "" Then
			gErrorMsgBox "정산년월을 확인하십시오","확정취소안내!"
			Exit Sub
		End if
		
		intRtn = mobjPDCMOUTSLIST.ProcessRtn_Cancel(gstrConfigXml,vntData,strYEARMON)
		
		if not gDoErrorRtn ("ProcessRtn_Cancle") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " 확정취소가" & mePROC_DONE,"확정취소안내" 
			SelectRtn
  		end if
 	end with
End Sub


Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" And mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 정산일자를 확인하십시오.","입력오류"
				Exit Function
			End if
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" And mobjSCGLSpr.GetTextBinding(.sprSht,"TAXDATE",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 증빙일자를 확인하십시오.","입력오류"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="58" background="../../../images/back_p.gIF"
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
											<td class="TITLE">정산 관리</td>
										</tr>
									</table>
								</td>
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
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="100"><SELECT id="cmbPOPUPTYPE" title="프로젝트,JOBNO선택" style="WIDTH: 100px" name="cmbPOPUPTYPE">
														<OPTION value="REG" selected>등록일자</OPTION>
														<OPTION value="ADJ">전표일자</OPTION>
														<OPTION value="TAX">증빙일자</OPTION>
													</SELECT></TD>
												<TD class="SEARCHDATA" style="WIDTH: 220px"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="8" size="6" name="txtFROM"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
														align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="8" size="6" name="txtTO"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
														align="absMiddle" border="0" name="imgTo"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)"
													width="70">
												정산구분
												<TD class="SEARCHDATA" style="WIDTH: 140px"><SELECT id="cmbGUBN" style="WIDTH: 120px" name="cmbGUBN">
														<OPTION value="" selected>전체</OPTION>
														<OPTION value="F">미정산</OPTION>
														<OPTION value="T">정산</OPTION>
														<OPTION value="V">전표발행내역</OPTION>
														<OPTION value="C">승인완료내역</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)"
													width="70">외주처
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtOUTSNAME" title="코드명" style="WIDTH: 155px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="38" name="txtOUTSNAME"> <IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgOUTSCODE">
													<INPUT class="INPUT_L" id="txtOUTSCODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtOUTSCODE"></TD>
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
									</FONT>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="98" background="../../../images/back_p.gIF"
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
											<td class="TITLE">매입집계 리스트</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<!--<TD><IMG id="imgConfirmFlag" onmouseover="JavaScript:this.src='../../../images/imgConfirmFlagOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirmFlag.gIF'"
													height="20" alt="승인저장합니다." src="../../../images/imgConfirmFlag.gIF" width="78" border="0"
													name="imgConfirmFlag"></TD>
											-->
											<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
													height="20" alt="확정합니다." src="../../../images/imgSetting.gIF" width="54" border="0"
													name="imgSetting"></TD>
											<TD><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gif'"
													height="20" alt="확정취소합니다." src="../../../images/ImgConfirmCancel.gIF" border="0"
													name="ImgConfirmCancel"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--테이블이 무너지는것을 막아준다-->
						<TABLE cellSpacing="0" cellPadding="0" width="1024" border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody1" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">정산년월</TD>
												<TD class="SEARCHDATA" style="WIDTH: 127px"><INPUT class="INPUT" id="txtYEARMON" title="코드명" style="WIDTH: 120px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="100" align="left" size="14" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL"><SELECT id="cmbCHK" title="프로젝트,JOBNO선택" style="WIDTH: 100px" name="cmbCHK">
														<OPTION value="ADJ" selected>정산일</OPTION>
														<OPTION value="TAX">증빙일</OPTION>
													</SELECT></TD>
												<TD class="SEARCHDATA" style="WIDTH: 208px"><INPUT class="INPUT_L" id="txtADJDAY" title="코드명" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="100" align="left" size="6" name="txtADJDAY"> <IMG id="ImgADJDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
														border="0" name="ImgADJDAY">&nbsp;<IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
														title="적요를 일괄 적용합니다" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="적요를 일괄 적용합니다"
														src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0" name="ImgSUMMApp"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 96px">선택합계</TD>
												<TD class="SEARCHDATA" style="WIDTH: 162px"><INPUT class="INPUT_L" id="txtSELECTAMT" title="코드명" style="WIDTH: 160px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="21" name="txtSELECTAMT"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 80px">총합계</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUM" title="코드명" style="WIDTH: 152px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="20" name="txtSUM"></TD>
											</TR>
										</TABLE>
									</FONT>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 8px"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						</FONT></TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="42413">
								<PARAM NAME="_ExtentY" VALUE="11086">
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
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></form>
	</body>
</HTML>
