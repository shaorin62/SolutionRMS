<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGEDIVLIST.aspx.vb" Inherits="PD.PDCMCHARGEDIVLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_ESTDTL.aspx
'기      능 : JOBMST의 두번째 탭 - 가/본 견적서를 저장 및 수정 한다. 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/18 By KimTH
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
'=============================
' 이벤트 프로시져 
'=============================
option explicit
Const meTAB = 9
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMCHARGEDIV, mobjPDCMGET
'Dim mobjPDCMCONTRACT
'선택체크용
Dim mstrCheck
Dim mALLCHECK
' 벨리데이션에 걸렸을시에 체크 mstrValiCHECK   pub_processrtn에서 사용
Dim mstrValiCHECK
'본견적을 가져왔을때 true   아니고 exe_hdr 에 있다면  초기값인 false
Dim strACTUALFLAG
'헤더의 변경내용 여부    기본 false 변경 true
Dim mstrHEADERFLAG 
Dim mstrPROCESS

Dim strJOBNO 
Dim strPREESTNO

mALLCHECK = TRUE
mstrCheck=TRUE
mstrValiCHECK = TRUE
strACTUALFLAG = FALSE
mstrPROCESS = False
mstrHEADERFLAG = false
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

Sub imgConfirm_onclick ()	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgConfirmCancel_onclick ()	
	gFlowWait meWAIT_ON
	ProcessRtn_ConfirmCancel
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
	
		gErrorMsgBox "선택된 데이터가 없습니다.",""
		Exit Sub
		
	
		'체크가 된 데이터가 있는지 없는지 체크한다.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1"   THEN
				intCount = 1
			end if
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = ""   THEN
				gErrorMsgBox i & " 번째 행의 계약서가 없습니다.","인쇄안내"
				Exit Sub
			End If
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
			'intRtn = mobjPDCMCONTRACT.DeleteRtn_TEMP(gstrConfigXml)
		
			ModuleDir = "PD"
			ReportName = "PDCMCONTRACT.rpt"
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strCONTRACTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					strUSERID = ""
					'vntDataTemp = mobjPDCMCONTRACT.ProcessRtn_TEMP(gstrConfigXml,strCONTRACTNO, i, strUSERID)
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
		'intRtn = mobjPDCMCONTRACT.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'청구일 조회조건 생성
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub


'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------

Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub

Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		CALL gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub

Sub txtDEMANDAMT_onfocus
	with frmThis
		.txtDEMANDAMT.value = Replace(.txtDEMANDAMT.value,",","")
	end with
End Sub
Sub txtDEMANDAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub

Sub txtESTAMT_onfocus
	with frmThis
		.txtESTAMT.value = Replace(.txtESTAMT.value,",","")
	end with
End Sub
Sub txtESTAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtESTAMT,0,true)
	end with
End Sub

Sub txtPAYMENT_onfocus
	with frmThis
		.txtPAYMENT.value = Replace(.txtPAYMENT.value,",","")
	end with
End Sub

Sub txtPAYMENT_onblur
	with frmThis
		CALL gFormatNumber(.txtPAYMENT,0,true)
	end with
End Sub

Sub txtINCOM_onfocus
	with frmThis
		.txtINCOM.value = Replace(.txtINCOM.value,",","")
	end with
End Sub
Sub txtINCOM_onblur
	with frmThis
		CALL gFormatNumber(.txtINCOM,0,true)
	end with
End Sub

Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		CALL gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub

Sub txtACCAMT_onfocus
	with frmThis
		.txtACCAMT.value = Replace(.txtACCAMT.value,",","")
	end with
End Sub
Sub txtACCAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtACCAMT,0,true)
	end with
End Sub

'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
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

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _ 
			or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_CONFIRM")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_NOCONFIRM")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
				strCOLUMN = "DIVAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
				strCOLUMN = "CHARGE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_CONFIRM") Then
				strCOLUMN = "OUTAMT_CONFIRM"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_NOCONFIRM") Then
				strCOLUMN = "OUTAMT_NOCONFIRM"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
				strCOLUMN = "EXEAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) _ 
					OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_CONFIRM")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_NOCONFIRM")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			CALL gFormatNumber(.txtSELECTAMT,0,True)
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
				OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_CONFIRM") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTAMT_NOCONFIRM") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
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
		CALL gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub


Sub sprSht_Change(ByVal Col, ByVal Row)

	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
			 mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "1"
		End if
	End	With
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
	Dim strComboList
	Dim strComboList2
	Dim strMSG
	
	'서버업무객체 생성	
	set mobjPDCMCHARGEDIV	= gCreateRemoteObject("cPDCO.ccPDCOCHARGEDIV")
	set mobjPDCMGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	'set mobjPDCMCONTRACT = gCreateRemoteObject("cPDCO.ccPDCOCONTRACT")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis
		
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht,   "CHK | EXE_FLAGNAME |CLIENTNAME| JOBNO | JOBNOSEQ | DIVRATE | DIVAMT | ADJAMT | CHARGE | OUTAMT_CONFIRM | OUTAMT_NOCONFIRM | EXEAMT | EXE_FLAG "
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|구분|광고주|JOBNO|순번|분담비율|분할금액|청구금액|잔액|외주비분담금(확정)|외주비분담금(미확정)|확정금액|확정구분"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|   5|12    |    9|   4|      12|      12|      12|  12|                16|                  17|      12|       6"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "JOBNOSEQ | DIVAMT | ADJAMT | CHARGE | OUTAMT_CONFIRM | OUTAMT_NOCONFIRM | EXEAMT ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVRATE", -1, -1, 2
		'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "EXE_FLAGNAME|JOBNO |CLIENTNAME| JOBNOSEQ | DIVRATE | DIVAMT | ADJAMT | CHARGE | OUTAMT_CONFIRM | OUTAMT_NOCONFIRM | EXEAMT | EXE_FLAG"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EXE_FLAGNAME | JOBNO | JOBNOSEQ",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "EXE_FLAG|JOBNO", true
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0


		'부모창의 데이터 가져오기  (전역변수에담기)
		
		.txtJOBNO.value = parent.document.forms("frmThis").txtJOBNO.value 
		strJOBNO = parent.document.forms("frmThis").txtJOBNO.value 
		
		.txtPREESTNO.value = parent.document.forms("frmThis").txtPREESTNO.value 
		strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value 
		
		SelectRtn
	End With
End Sub

Sub EndPage()
	'set mobjPDCMCHARGEDIV = Nothing
	'set mobjPDCMGET = Nothing
	'set mobjPDCMCONTRACT = Nothing
	gEndPage
End Sub

'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	with frmThis
		if strJOBNO = "" Or Len(strJOBNO) <> 7 Then
			gErrorMsgBox "제작번호를확인하십시오.","조회안내!"
			Exit Sub
		End if
		
		'JOBNO로 정산데이타를 가져온다. 업으면FALSE
		IF SelectRtn_Head Then 
			CALL SelectRtn_Detail ()
		else
			CALL SelectRtn_Actual_Head ()
			CALL SelectRtn_Detail ()
		END IF
		
		txtSUSUAMT_onblur
		txtCOMMITION_onblur
		txtDEMANDAMT_onblur
		txtPAYMENT_onblur
		txtINCOM_onblur
		txtNONCOMMITION_onblur
		txtACCAMT_onblur
		txtESTAMT_onblur
		AMT_SUM
		mstrHEADERFLAG = false
	End with
End Sub

Function SelectRtn_Head
	Dim vntData
	SelectRtn_Head = false
	'on error resume next
	'초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMCHARGEDIV.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt <=0 then
			'gErrorMsgBox "확정견적서가 " & meNO_DATA ,""
			SelectRtn_Head = FALSE
			strACTUALFLAG = TRUE
			gClearAllObject frmThis
		else
			'조회한 데이터를 바인딩
			SelectRtn_Head = True
			CALL gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
		End IF
	End IF
End Function



Function SelectRtn_Actual_Head
	Dim vntData
	'on error resume next
	
	'초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'
	vntData	= mobjPDCMCHARGEDIV.SelectRtn_Actual_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
	IF not gDoErrorRtn ("SelectRtn_Actual_HDR") then
		IF mlngRowCnt > 0 then
			'조회한 데이터를 바인딩
			CALL gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			
			'바인딩한 후에는 본견적 jobno와 preestno 로 다시 전역변수에 넣어준다.
			'strJOBNO	= frmThis.txtJOBNO.value
			'strPREESTNO = frmThis.txtPREESTNO.value
		Else
		gClearAllObject frmThis
		End IF
	End IF
End Function


'divamt 테이블 조회
Function SelectRtn_Detail
	dim vntData
	Dim strRows
	Dim intCnt
	Dim lngRowCnt
	'on error resume next
	'초기화
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMCHARGEDIV.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		'조회한 데이터를 바인딩
		CALL mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		
		lngRowCnt = mlngRowCnt
		SelectRtn_Detail = True
		
		with frmThis
			IF mlngRowCnt > 0 THEN
				'확정된거
				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht, "EXE_FLAG",intCnt) = "0" THEN '노랑
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False '이게 흰색
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"EXEAMT",intCnt,intCnt,false
					ELSE
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					END IF
			
				Next
				gWriteText lblStatus, lngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function




'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"adjamt", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			CALL gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub


'****************************************************************************************
' 데이터 처리
'****************************************************************************************
Sub ProcessRtn ()
    Dim intRtn , intCnt
  	dim vntData
	Dim intCHK
	Dim intConRtn
	with frmThis
	
	'On error resume next
		if strJOBNO = "" Then
			gErrorMsgBox "조회된 제작관리번호가 없습니다.","저장안내!"
			Exit Sub
		End If
		
		for intCnt	=1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht, "CHK",intCnt) = "1" and mobjSCGLSpr.GetTextBinding(.sprSht, "EXE_FLAG",intCnt) = "1" then
				gErrorMsgBox intCnt & "행은 확정된 상태입니다.","처리안내!"
				exit sub
			End if
		next
		
  		'데이터 Validation
		'if DataValidation = false then exit sub
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | JOBNO | JOBNOSEQ | EXEAMT | EXE_FLAG")
		
		if  not IsArray(vntData)  Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		'처리 업무객체 호출
		intRtn = mobjPDCMCHARGEDIV.ProcessRtn(gstrConfigXml,vntData)
				
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가 확정" & mePROC_DONE,"저장안내" 
			SelectRtn
  		end if
 	end with
End Sub

Sub ProcessRtn_ConfirmCancel ()
    Dim intRtn , intCnt
  	dim vntData
	Dim intCHK
	Dim intConRtn
	with frmThis
	
	'On error resume next
		if strJOBNO = "" Then
			gErrorMsgBox "조회된 제작관리번호가 없습니다.","저장안내!"
			Exit Sub
		End If
		
		for intCnt=1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht, "CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht, "EXE_FLAG",intCnt) = "0" then
				gErrorMsgBox intCnt & "행은 미확정된 상태입니다.","처리안내!"
				exit sub
			End if
		next
		
  		'데이터 Validation
		'if DataValidation = false then exit sub
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | JOBNO | JOBNOSEQ | EXEAMT | EXE_FLAG")
		
		if  not IsArray(vntData)  Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		'처리 업무객체 호출
		intRtn = mobjPDCMCHARGEDIV.ProcessRtn_ConfirmCancel(gstrConfigXml,vntData)
				
		if not gDoErrorRtn ("ProcessRtn_ConfirmCancel") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가 미확정" & mePROC_DONE,"저장안내" 
			SelectRtn
  		end if
 	end with
End Sub




		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="Td2" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="54" background="../../../images/back_p.gIF"
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
											<td class="TITLE">정산관리&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton2" style=" HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<td><INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 20px" dataSrc="#xmlBind" 
													size="1" name="txtJOBNO" type=hidden ><INPUT dataFld="JOBNOINS" id="txtJOBNOINS" style="WIDTH: 20px" dataSrc="#xmlBind" 
													size="1" name="txtJOBNOINS" type=hidden ><INPUT dataFld="PREESTNO" id="txtPREESTNO" style="WIDTH: 20px" dataSrc="#xmlBind" 
													size="1" name="txtPREESTNO" type=hidden ><INPUT dataFld="ENDDAY" id="txtENDDAY" style="WIDTH: 20px" dataSrc="#xmlBind" 
													size="1" name="txtENDDAY" type=hidden ></td>
											<!--<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></TD>-->
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top" width="100%">
						<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
							align="right" border="0">
							<TR>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">프로젝트명</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="PROJECTNM" class="NOINPUTB_R" id="txtPROJECTNM" title="프로젝트명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPROJECTNM"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">광고주</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_R" id="txtCLIENTNAME" title="광고주" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTNAME"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">견적금액</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="ESTAMT" class="NOINPUTB_R" id="txtESTAMT" title="견적금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtESTAMT"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">Noncommition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="수수료미지불금액"
										style="WIDTH: 152px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtNONCOMMITION"></TD>
							</TR>
							<TR>
								<TD class="SEARCHLABEL">JOB명</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="JOBNAME" class="NOINPUTB_R" id="txtJOBNAME" title="JOB명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBNAME"></TD>
								<TD class="SEARCHLABEL">팀</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="TIMNAME" class="NOINPUTB_R" id="txtTIMNAME" title="팀명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtTIMNAME"></TD>
								<TD class="SEARCHLABEL">청구금액</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDAMT" class="NOINPUTB_R" id="txtDEMANDAMT" title="청구금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDAMT"></TD>
								<TD class="SEARCHLABEL">Commition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="수수료지불금액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCOMMITION"></TD>
							</TR>
							<tr>
								<TD class="SEARCHLABEL">매체부문</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="JOBGUBN" class="NOINPUTB_R" id="txtJOBGUBN" title="매체부문" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBGUBN"></TD>
								<TD class="SEARCHLABEL">브랜드</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_R" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUBSEQNAME"></TD>
								<TD class="SEARCHLABEL">외주비</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="PAYMENT" class="NOINPUTB_R" id="txtPAYMENT" title="외주비 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPAYMENT"></TD>
								<TD class="SEARCHLABEL">수수료</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="수수료합계금액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUSUAMT"></TD>
							</tr>
							<tr>
								<TD class="SEARCHLABEL">매체분류</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUTB_R" id="txtCREPART" title="매체분류" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="6" name="txtCREPART"></TD>
								<TD class="SEARCHLABEL">청구일</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDDAY" class="NOINPUTB_R" id="txtDEMANDDAY" title="청구일" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDDAY"></TD>
								<TD class="SEARCHLABEL">진행비</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ACCAMT" class="NOINPUTB_R" id="txtACCAMT" title="비용 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtACCAMT"></TD>
								<TD class="SEARCHLABEL">수수료율</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSURATE" class="NOINPUTB_R" id="txtSUSURATE" title="수수료율" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtSUSURATE">&nbsp;(%)</TD>
							</tr>
							<TR>
								<TD class="SEARCHLABEL">상태</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ENDFLAG" class="NOINPUTB_R" id="cmbENDFLAG" title="상태" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="cmbENDFLAG"></TD>
								<TD class="SEARCHLABEL">결산일</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CLOSEDAY" class="NOINPUTB_R" id="txtClOSEDAY" title="결산일" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtClOSEDAY"></TD>
								<TD class="SEARCHLABEL">내수액</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="INCOM" class="NOINPUTB_R" id="txtINCOM" title="내수액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtINCOM"></TD>
								<TD class="SEARCHLABEL">내수율</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="RATE" class="NOINPUTB_R" id="txtRATE" title="내수율" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtRATE">&nbsp;(%)</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" id="spacebar" style="WIDTH: 100%; HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="80" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="68" background="../../../images/back_p.gIF"
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
											<td class="TITLE">청구지분할&nbsp;</td>
										</tr>
									</table>
								</TD>
								<td class="TITLE"> 합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
										<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 24px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
													height="20" alt="자료를 확정합니다." src="../../../images/imgSetting.gIF" border="0" name="imgConfirm"></TD>
											<TD><IMG id="imgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgConfirmCancelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirmCancel.gIF'"
													height="20" alt="자료를 확정취소합니다." src="../../../images/imgConfirmCancel.gIF" border="0"
													name="imgConfirmCancel"></TD>
											<!--<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>-->
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1075" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
							<TR>
								<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="16219">
											<PARAM NAME="_ExtentY" VALUE="11880">
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
								</td>
							</tr>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</table>
					</td>
				</tr>
			</TABLE>
		</FORM>
	</body>
</HTML>
