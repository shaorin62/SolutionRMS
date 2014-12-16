<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMEXMAIN01.aspx.vb" Inherits="MD.MDCMEXMAIN01" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>시스템 공통</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/시스템공통/EXCEL업로더
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCEXMAIN0.aspx
'기      능 : 정의테이블에 EXCELUPLOAD
'파라  메터 : 
'특이  사항 : 개발 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/07/03 By ParkJS(박종세)
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
    Option explicit
    Dim mlngRowCnt, mlngColCnt
    Dim sprSht_DataFields
    Dim vntData_DataFields	
    Dim sprSht_DisplayFields
    Dim sprSht_ColWidth
    Dim sprSht_NotNull
    Dim vntData_Nullable
    Dim sprSht_DefualtValueFields
    Dim vntData_DefaultValue
    Dim vntData_DataType
    Dim mdblTAB_ID, mstrTAB_NAME, mstrTAB_USER_NAME, mstrTAB_TYPE, mstrTAB_DESC 
    Dim mobjccSCEXCOM  , mobjccSCEXBrowse
    Dim mobjPDCMJOBNOREG
    Dim mInsOKFlag 'Insert Flag 
    Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode '팝업사용시
    Dim mstrdeletetemp
    mstrdeletetemp = false
'=============================
' 이벤트프로시져 
'=============================
Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
    
End Sub

Sub InitPage()
    '서버업무객체 생성	
    Set mobjccSCEXCOM = gCreateRemoteObject("cMDPT.ccSCEXCOM")
    Set mobjccSCEXBrowse = gCreateRemoteObject("cMDPT.ccSCEXBrowse")

   '권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC" 
   'InsOKFlag 를 false 값으로 설정한다.
	mInsOKFlag   =  false

	gSetSheetDefaultColor
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* 초기화면 입니다. "& vbcrlf & vbcrlf &"* 도움말: 처리버튼을 누르시고, 안내에 따라 자료를 붙여넣으십시오."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "70"
	end with
	
	Call imgFind_onclick
end Sub

Sub EndPage()
	set mobjccSCEXCOM = Nothing
	set mobjccSCEXBrowse = Nothing
	'PopUp Window 일때 mInsOKFlag 를 넘겨준다.
	If gIsPopupWindow then
 	  window.returnvalue = mInsOKFlag
	End if
	gEndPage
End Sub
'=============================
' 명령버튼클릭이벤트
'=============================
Sub imgFind_onclick
    Dim vntRet, vntInParams, dblTAB_ID
    Dim vntData
    Dim intRtn
		
	mdblTAB_ID        = 30
	mstrTAB_NAME      = "MD_BOOKING_MEDIUM"
	mstrTAB_USER_NAME = "인쇄매체부킹등록"
	mstrTAB_TYPE      = "TABLE"
	mstrTAB_DESC      = "인쇄매체부팅 UPLOAD"
	
	gFlowWait meWAIT_ON
	makePageData
	gFlowWait meWAIT_OFF
	
	WITH frmThis
		Dim i, RowNum, intRows

	'	Insert OK Flag 를 True 로 설정한다.
	   	mInsOKFlag = true
		mstrdeletetemp = false
		'추가부분
		RowNum = 301
		
		mobjSCGLSpr.SetMaxRows .sprSht, RowNum 
		intRows = Ubound(vntData_DefaultValue,1) +1
		
		For i=1 To intRows
			mobjSCGLSpr.SetText .sprSht, i , -1, vntData_DefaultValue(i-1) 
		Next 
		gOKMsgbox "데이터를 입력할 준비가 되었습니다. Excel Data를 붙여넣어 주십시요.", ""
	end with
	mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,1
	frmThis.sprSht.focus()
End Sub

Sub imgSave_onclick()
	Dim intRtn
	Dim strPUBYEARMON
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
			Exit Sub
		end if
		
		strPUBYEARMON = MID(REPLACE(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",1),"-",""),1,6)
		
		gFlowWait(meWAIT_ON)
		if mstrdeletetemp then 
			intRtn = mobjccSCEXCOM.Delete_Temp_Rtn(gstrConfigXml, strPUBYEARMON)
		end if
		
		gFlowWait(meWAIT_ON)
		ProcessRtn()
		gFlowWait(meWAIT_OFF)
    END WITH
End Sub


Sub imgDelete_onclick
    gFlowWait(meWAIT_ON)
    DeleteRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgClose_onclick()
    Window_OnUnload()
End Sub

Sub cmbMED_FLAG_onchange
	Dim strMED_FLAGNAME
	with frmThis
		gFlowWait meWAIT_ON
		imgFind_onclick
		gFlowWait meWAIT_OFF
	end with
	gSetChange
End Sub

'=============================
'SheetEvent
'=============================
Sub sprSht_KeyDown(KeyCode, Shift)
	If KeyCode = 86 Then
		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
	end if
End Sub

Sub sprSht_KeyUp(KeyCode, shift)
	If KeyCode = 86 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,1,301) <> "" then
			gErrorMsgbox "일괄청약시 한번에 투입가능한 데이터는 300건입니다. 다시 올려주십시오.",""
			mobjSCGLSpr.ClearText frmThis.sprSht , -1, -1, -1, -1 
			exit sub
		End If
	end if
end Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub

'==================================================
'데이터를 처리
'==================================================
Sub ProcessRtn ()
	Dim intRtn   'Return 값
   	Dim vntData  'Insert 할 데이터
   	Dim intCnt
   	Dim lngAMT
   	Dim lngPRICE
   	Dim lngCOMMISSION
   	Dim lngCOMMI_RATE
   	Dim strSPONSOR
   	Dim strCOMMISSION
	with frmThis
	
		'여분 Rows 삭제처리
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,3,intCnt) = "" and mobjSCGLSpr.GetTextBinding(.sprSht,4,intCnt) = "" and mobjSCGLSpr.GetTextBinding(.sprSht,5,intCnt) = "" then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			ELSE
				CALL SetTrim (intCnt) ' 공백문자열 제거
			END IF
		NEXT
		
		'==================오류검증
		if DataValidation =false then exit sub
		FOR intCnt = 1 to .sprSht.MaxRows
			if mstrdeletetemp then 
				mobjSCGLSpr.CellChanged .sprSht, 1, intCnt
			end if
			
			lngAMT = ""
			lngPRICE = ""
			lngCOMMI_RATE = ""
			lngCOMMISSION = ""
			strSPONSOR = ""
			strCOMMISSION = ""
			
			lngAMT =  mobjSCGLSpr.GetTextBinding(.sprSht,"AMOUNT",intCnt)  
			lngPRICE =  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt)
			lngCOMMI_RATE =  mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",intCnt)
			lngCOMMISSION =  mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",intCnt)
			
			IF lngAMT = "" THEN
				mobjSCGLSpr.SetTextBinding .sprSht,"AMOUNT",intCnt,0
			END IF
			
			IF lngPRICE = "" THEN
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",intCnt,0
			END IF
			
			IF .cmbMED_FLAG.value THEN
				strSPONSOR =  mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",intCnt)				
				IF lngCOMMISSION = "" THEN
					IF lngCOMMI_RATE = "" THEN
						IF strSPONSOR = "Y" THEN
							mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",intCnt,10
							strCOMMISSION = CDbl(lngAMT) * 10 / 100
							mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
						ELSE
							mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",intCnt,15
							strCOMMISSION = CDbl(lngAMT) * 15 / 100
							mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
						END IF
					ELSE
						strCOMMISSION = CDbl(lngAMT) * CDbl(lngCOMMI_RATE) / 100
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
					END IF
				END IF
			ELSE
				IF lngCOMMISSION = "" THEN
					IF lngCOMMI_RATE = "" THEN
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",intCnt,15
						strCOMMISSION = CDbl(lngAMT) * 15 / 100
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
					ELSE
						strCOMMISSION = CDbl(lngAMT) * CDbl(lngCOMMI_RATE) / 100
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
					END IF
				END IF
			END IF
		Next
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
 	    if  not IsArray(vntData) then 
		    gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
		    exit sub
        end if
        
  	    Dim STime, ETime
	    STime = Time
			intRtn = mobjccSCEXCOM.ProcessRtn(gstrConfigXML, vntData, mstrTAB_NAME, sprSht_DataFields, vntData_DataType,  false, .cmbMED_FLAG.value)
		ETime = Time
        'MsgBox FormatDateTime(STime,vbLongTime) & " ~ " & FormatDateTime(ETime,vbLongTime) & " = " & DateDiff("S",STime,ETime)

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
	   	    mobjSCGLSpr.SetMaxRows frmThis.sprSht, 0 
	   	    gOKMsgbox "데이터를 성공적으로 UPLOAD 하였습니다.", ""
   		
	   	    mInsOKFlag = true
   		end if

   	end with
End Sub

Sub SetTrim (Row) 
	With frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PUB_DATE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM_NAME",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM_NAME",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"COL_DEG",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"COL_DEG",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PUB_FACE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"STD_STEP",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"STD_STEP",Row))
		IF .cmbMED_FLAG.value THEN
			mobjSCGLSpr.SetTextBinding .sprSht,"STD_CM",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"STD_CM",Row))
			
		ELSE
			mobjSCGLSpr.SetTextBinding .sprSht,"STD",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row))
		END IF
		mobjSCGLSpr.SetTextBinding .sprSht,"SPONSOR",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"AMOUNT",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"AMOUNT",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTION",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTION",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"NOTE",Row,TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"NOTE",Row))
	End With
End Sub

Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i
	On error resume next
	with frmThis
		'한 건씩 삭제할 경우
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		if gDoErrorRtn ("DeleteRtn") then exit sub
		if intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit sub
		end if
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			if not gDoErrorRtn ("DeleteRtn_SC_USER_COL") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			end if
		next
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	end with
End Sub
'======================================
'기타함수
'======================================
Function DataValidation ()
	dim i,j
	DataValidation = false
	with frmThis
		'데이터가 변경되었는지 검사
		if not mobjSCGLSpr.IsDataChanged(.sprSht) then
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit function
		end if
		
		Dim intCnt
   		Dim strArray
   		Dim Rowcnt
   		Dim Colcnt
   		Dim strMedAndReal
   		Dim strERR
   		Dim strMEDCODENAME
   		Dim strMEDCODE
   		Dim strREALMEDCODE
   		Dim strCLIENTNAME
   		Dim intVal
   		Dim intRtn
   		Dim vntData
   		Dim vntData2
   		Dim vntData3
   		Dim strCLIENTCODE
   		Dim strDEPTCODE
   		Dim strSEQCODE
   		Dim lngAMT
   		Dim lngREAL_AMT
   		Dim strPUB_FACENAME
   		Dim strPUB_FACECODE
   		Dim strCLIENTSUBCODE
   		Dim strDEPT_NAME, strDEPT_CODE
   		intVal = 0
   		'오류체크부분을 일단 공백으로 만든다.
   		 For intCnt = 1 To .sprSht.MaxRows
   			mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,""
   		 Next
   		 '광고주코드체크
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt),1,1) = "A" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt)) = 6 Then
				vntData = mobjccSCEXCOM.SelectRtn_CLIENTCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
				if not gDoErrorRtn ("SelectRtn_CODE") then
					IF mlngRowCnt <> 1 Then
						strERR = "광고주코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strCLIENTNAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
   				vntData = mobjccSCEXCOM.SelectRtn_CLIENTNAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTNAME)
	   			
				if not gDoErrorRtn ("SelectRtn_CLIENTNAME") then
					If mlngRowCnt = 1 Then
						strCLIENTCODE = vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",intCnt,strCLIENTCODE
					Else
						strERR = "광고주코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		 
   		  '사업부 매쳉작업
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt),1,1) = "A" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt)) = 6 Then
				vntData = mobjccSCEXCOM.SelectRtn_CLIENTCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt))
				if not gDoErrorRtn ("SelectRtn_CODE") then
					IF mlngRowCnt <> 1 Then
						strERR = "사업부코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strCLIENTCODE = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
   				strCLIENTSUBCODE = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt))
   				
   				vntData = mobjccSCEXCOM.SelectRtn_CLIENTSUBCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTCODE, strCLIENTSUBCODE)
	   			
				if not gDoErrorRtn ("SelectRtn_CLIENTSUBCODE") then
					If mlngRowCnt = 1 Then
						IF strCLIENTCODE <> vntData(1,0) then
							strERR = "해당광고주의 사업부코드확인"
							mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
							intVal = 1
						else
							strCLIENTSUBCODE = vntData(0,0)
							mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",intCnt,strCLIENTSUBCODE
						end if
					Else
						strERR = "사업부코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		 
   		 
   		 
   		 '채널코드체크
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt),1,1) = "B" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt)) = 6 Then
   				strMEDCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt)
				vntData2 = mobjccSCEXCOM.SelectRtn_REALMEDCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,strMEDCODE)
				If mlngRowCnt = 1 Then
					strREALMEDCODE = vntData2(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",intCnt,strREALMEDCODE
				Else
					strERR = "채널코드오류"
					mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
					intVal = 1
				End If
   			Else 
   				strMEDCODENAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt))
   				vntData = mobjccSCEXCOM.SelectRtn_MEDCODENAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strMEDCODENAME)
				if not gDoErrorRtn ("SelectRtn_MEDCODENAME") then
					If mlngRowCnt = 1 Then
						strMEDCODE = vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",intCnt,strMEDCODE
						vntData2 = mobjccSCEXCOM.SelectRtn_REALMEDCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,strMEDCODE)
						strREALMEDCODE = vntData2(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",intCnt,strREALMEDCODE
						
					Else
						strERR = "채널코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		 
   		 '담당부서코드체크
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",intCnt),1,1) = "1" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",intCnt)) = 8 Then
				vntData = mobjccSCEXCOM.SelectRtn_DEPT_CD(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",intCnt))
				if not gDoErrorRtn ("SelectRtn_DEPT_CD") then
					IF mlngRowCnt <> 1 Then
						strERR = "부서코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strDEPT_NAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",intCnt))
   				vntData = mobjccSCEXCOM.SelectRtn_DEPT_NAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strDEPT_NAME)
	   			
				if not gDoErrorRtn ("SelectRtn_DEPT_NAME") then
					If mlngRowCnt = 1 Then
						strDEPT_CODE = vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",intCnt,strDEPT_CODE
					Else
						strERR = "부서코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		 
   		 '게재면코드체크
'   		 For intCnt = 1 To .sprSht.MaxRows
'   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",intCnt),1,2) = "MP" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",intCnt)) = 5 Then
'   				strPUB_FACENAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",intCnt))
'   				vntData = mobjccSCEXCOM.SelectRtn_PUBFACECODE(gstrConfigXML,mlngRowCnt,mlngColCnt,strPUB_FACENAME, .cmbMED_FLAG.value)
'   				If mlngRowCnt = 1 Then
'					strPUB_FACECODE = vntData(0,0)
'					mobjSCGLSpr.SetTextBinding .sprSht,"PUB_FACE",intCnt,strPUB_FACECODE
'				Else
'					strERR = "게재코드오류"
'					mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
'					intVal = 1
'				End If
'   			Else 
'   				strPUB_FACENAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_FACE",intCnt))
'   				vntData = mobjccSCEXCOM.SelectRtn_PUBFACENAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strPUB_FACENAME, .cmbMED_FLAG.value)
'				if not gDoErrorRtn ("SelectRtn_PUBFACENAME") then
'					If mlngRowCnt = 1 Then
'						strPUB_FACECODE = vntData(0,0)
'						mobjSCGLSpr.SetTextBinding .sprSht,"PUB_FACE",intCnt,strPUB_FACECODE
'					Else
'						strERR = "게재면코드오류"
'						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
'						intVal = 1
'					End If
'				End If
'			End If
'   		 Next
   		 
   		 '소재명체크(자리수,싱글쿼테이션)
   		 For intCnt = 1 To .sprSht.MaxRows
            If mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM_NAME",intCnt) <> "" Then
                If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM_NAME",intCnt)) < 255 Then
                mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM_NAME",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM_NAME",intCnt),"'","") 
                mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM_NAME",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM_NAME",intCnt),",","") 
                Else
                    strERR = "프로그램글자길이"
                    mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                    intVal = 1
                End If
            End If
         Next
   		         
         '게재일 길이 체크
   		 For intCnt = 1 To .sprSht.MaxRows
            If mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",intCnt) <> "" Then
                If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"PUB_DATE",intCnt)) <> 8 Then
                    strERR = "게재일일길이"
                    mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                    intVal = 1
                End If
            End If
         Next    
	   	 If intVal Then Exit Function
	end with
	DataValidation = true
End Function

Sub makePageData
     Dim vntData
     
     With frmThis
        mlngRowCnt=Clng(0): mlngColCnt=Clng(0)
        
        vntData = mobjccSCEXCOM.getTABCOLINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,mdblTAB_ID,.cmbMED_FLAG.value)

        sprSht_DataFields    = mChangeData (vntData,2,"|")
        vntData_DataFields   = gArray2Single(vntData,1)	  
        
        sprSht_DisplayFields = mChangeData (vntData,2,"|")
         
        sprSht_DefualtValueFields = mDefaultValueField(vntData,1,4,"|")

        vntData_DefaultValue = gArray2Single (vntdata,4)
                 
        vntData_DataType     = gArray2Single (vntData,5)
        vntData_Nullable	 = gArray2Single (vntData,8)

        sprSht_NotNull = fNotNull(vntData_Nullable, vntData_DataFields, mlngRowCnt)
        
        IF .cmbMED_FLAG.value THEN
			gSetSheetDefaultColor
			gSetSheetColor mobjSCGLSpr,     .sprSht
			mobjSCGLSpr.SpreadLayout        .sprSht, 20, 0
			mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
			mobjSCGLSpr.SetHeader           .sprSht, "게재일|광고주|사업부|매체명|소재명|담당부서|색도|게재면코드|단|CM|단가|금액|수수료율|수수료|돌출|협찬|비고|오류내용|매체사코드|브랜드코드"
			mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
			mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "AMOUNT|PRICE|COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellsLock2		.sprSht, true, sprSht_DefualtValueFields
			mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"
			mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
			mobjSCGLSpr.SetColWidth         .sprSht, "-1", 10
			mobjSCGLSpr.ColHidden .sprSht,"REAL_MED_CODE|SUBSEQ",TRUE
		ELSE
			gSetSheetDefaultColor
			gSetSheetColor mobjSCGLSpr,     .sprSht
			mobjSCGLSpr.SpreadLayout        .sprSht, 20, 0
			mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
			mobjSCGLSpr.SetHeader           .sprSht, "게재일|광고주|사업부|매체명|소재명|담당부서|색도|게재면코드|규격|페이지|단가|금액|수수료율|수수료|돌출|협찬|비고|오류내용|매체사코드|브랜드코드"
			mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
			mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "AMOUNT|PRICE|COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellsLock2		.sprSht, true, sprSht_DefualtValueFields
			mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"
			mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
			mobjSCGLSpr.SetColWidth         .sprSht, "-1", 10
			mobjSCGLSpr.ColHidden .sprSht,"REAL_MED_CODE|SUBSEQ",TRUE
        END IF
        'mobjSCGLSpr.ColHidden .sprSht,"YEARMON",true
        
        'gOkMsgBox  "테이블: "& mstrTAB_USER_NAME & "입니다." & vbcrlf & _
        '          "엑셀 데이터가 준비가 되셨으면 Control+C(복사) 하신 후" & vbcrlf & _
        '          "입력할 데이터 로우(행)의 숫자를 입력하여 주시기 바랍니다.", ""
        '.txtROWNUM.focus()
    End With
End Sub

Function mChangeData(vntData, colidx, splitMark)
    Dim strRtn, lngRowCnt, i
     
    lngRowCnt=-1: lngRowCnt=ubound(vntData,2)
    for i = 0 to lngRowCnt
      if i=lngRowCnt then splitMark = ""
        strRtn = strRtn & vntData(colidx,i) & splitMark
    Next
    mChangeData = strRtn
End Function

Function mDefaultValueField(vntData, colidx, CheckColidx, splitMark)
	Dim strRtn, lngRowCnt, i
	
	lngRowCnt=-1: lngRowCnt=Ubound(vntData,2)
	
	for i=0 to lngRowCnt
		if i=lngRowCnt then splitMark = ""
		if vntData(CheckColidx, i) <> "" then
			strRtn = strRtn & vntData(colidx,i) & splitMark
		end if
	next
	mDefaultValueField = strRtn
End Function

Function mHaveID()
	mHaveID = false
	Dim i, intRowCnt, temp
	temp = "ID"
	intRowCnt = Ubound(vntData_DefaultValue,1)
	For i=0 To intRowCnt
	  if temp = vntData_DefaultValue(i) then 
	     mHaveID = true
	     exit Function		
	  end if	
	Next
End Function

Function fNotNull(vntData_Nullable, vntData_DataFields, intRows)
	Dim i, sprSht_NotNull
	sprSht_NotNull = ""
	For i=0 to intRows-1
		if vntData_Nullable(i) = "N" then
		  sprSht_NotNull = sprSht_NotNull & vntData_DataFields(i) & "|"
		end if  
	Next
	fNotNull = sprSht_NotNull
End Function
-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE" id="tblTitleName"><FONT face="굴림">&nbsp;인쇄 청약관리 (일괄 청약)</FONT></td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 114px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="114" border="0">
										<TR>
											<TD width="3"><IMG id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imginitOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imginit.gif'" height="20" alt="자료를 저장합니다."
													src="../../../images/imginit.gif" border="0" name="imgFind"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gif" width="54" border="0"
													name="imgDelete"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"></TD>
							</TR>
							<!--TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD style="BORDER-RIGHT: lightsteelblue 1px solid; BORDER-TOP: lightsteelblue 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: lightsteelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: lightsteelblue 1px solid; FONT-FAMILY: 굴림; HEIGHT: 80px; BACKGROUND-COLOR: #eeeeee">
												<P align="left">&nbsp;* 사용방법<BR>
													&nbsp;&nbsp;&nbsp;1. 처리 버튼을 누르시면 기본 1000건의 데이터를 입력할 수 있습니다.<BR>
													&nbsp;&nbsp; 2. 복사한 EXCEL 데이터를 복사하여(CONTROL+C) 첫번째 행에 포커스를 지정해 주신후 (CONTROL+V) 
													로 데이터를 붙여 넣어 주십시요.<BR>
													&nbsp;&nbsp; 3. 잠시 기다리시면(데이터의 양에 따라 속도가 오래 걸릴 수 있습니다.) 화면에 데이터가 보입니다.<BR>
													&nbsp;&nbsp; 4. 데이터를 확인하신 후에 저장 버튼을 누르시면 됩니다.</P>
											</TD>
										</TR>
									</TABLE-->
							<!--
									<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD style="BORDER-RIGHT: lightsteelblue 1px solid; BORDER-TOP: lightsteelblue 1px solid; FONT-SIZE: 9pt; BORDER-LEFT: lightsteelblue 1px solid; WIDTH: 100%; BORDER-BOTTOM: lightsteelblue 1px solid; FONT-FAMILY: 굴림; HEIGHT: 126px; BACKGROUND-COLOR: #eeeeee">
												<P align="left">&nbsp;* 사용방법<BR>
													&nbsp;&nbsp;&nbsp;1. 년도,월,담당부서,담당사원 을 입력하시고&nbsp;처리 버튼을 눌러주십시오.<BR>
													&nbsp;&nbsp; 2. 테이블로 데이터를 올리실 데이터를 EXCEL 에서 만들어 주십시요.<BR>
													&nbsp;&nbsp; 3. EXCEL 로 데이터를 다 만드셨으면 데이터를 올리실 영역을 선택하여 복사하여(CONTROL+C) 주십시요.<BR>
													&nbsp;&nbsp; 4. 올리실 데이터의 로우(행)수만큼 업로드 로우(행) 지정 부분에 숫자를 넣고 확인 버튼을 눌러 주십시요.<BR>
													&nbsp;&nbsp; 5. 첫번째 행에 첫번째&nbsp; 셀에 포커스가 이동되면 (CONTROL+V) 로 데이터를 붙여 넣어 주십시요.<BR>
													&nbsp;&nbsp; 6. 잠시 기다리시면(데이터의 양에 따라 속도가 오래 걸릴 수 있습니다.) 화면에 데이터가 보입니다.<BR>
													&nbsp;&nbsp; 7. 데이터를 확인하신 후에 저장 버튼을 누르시면 됩니다.</P>
											</TD>
										</TR>
									</TABLE>
									
									<TABLE class="DATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 110px">년월선택</TD>
											<TD class="DATA" style="WIDTH: 90px"><INPUT class="INPUT" id="txtYEARMON" title="해당년도" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM,M"
													type="text" maxLength="6" size="9" name="txtYEARMON"><FONT face="굴림"></FONT></TD>
											<TD class="LABEL" style="WIDTH: 62px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPT_CD,txtDEPT_NAME)">담당부서</TD>
											<TD class="DATA" style="WIDTH: 207px"><INPUT class="INPUT" id="txtDEPT_CD" title="해당월" style="WIDTH: 74px; HEIGHT: 22px" accessKey=",M"
													type="text" size="7" name="txtDEPT_CD"><FONT face="굴림">&nbsp;</FONT><IMG id="ImgCRE_DEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgSEQNO"><FONT face="굴림">&nbsp;</FONT><INPUT class="INPUT" id="txtDEPT_NAME" title="해당월" style="WIDTH: 97px; HEIGHT: 22px" dataSrc="#xmlBind"
													type="text" size="10" name="txtDEPT_NAME"></TD>
											<TD class="LABEL" style="WIDTH: 61px;CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMP_NO,txtEMP_NAME)">담당자</TD>
											<TD class="DATA" style="WIDTH: 233px"><INPUT class="INPUT" id="txtEMP_NO" title="해당월" style="WIDTH: 74px; HEIGHT: 22px" accessKey=",M"
													type="text" size="7" name="txtEMP_NO"><FONT face="굴림">&nbsp;</FONT><IMG id="ImgCREEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgSEQNO"><FONT face="굴림">&nbsp;</FONT><INPUT class="INPUT" id="txtEMP_NAME" title="해당월" style="WIDTH: 120px; HEIGHT: 22px" type="text"
													size="14" name="txtEMP_NAME"></TD>
										</TR>	
									</TABLE>
									-->
						</TABLE>
						<TABLE class="DATA" id="tblKey2" cellSpacing="1" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="SEARCHLABEL" style="WIDTH: 67px"><FONT face="굴림">매체구분</FONT></TD>
								<TD class="SEARCHDATA"><SELECT id="cmbMED_FLAG" title="신문/잡지선택" style="WIDTH: 96px; HEIGHT: 22px" name="cmbMED_FLAG">
										<OPTION value="1" selected>신문</OPTION>
										<OPTION value="0">잡지</OPTION>
									</SELECT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<td>
						<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 780px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
							VIEWASTEXT>
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="27517">
							<PARAM NAME="_ExtentY" VALUE="16616">
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
					</td>
				</tr>
			</TABLE>
			</TD></TR></TABLE>
		</form>
	</body>
</HTML>
