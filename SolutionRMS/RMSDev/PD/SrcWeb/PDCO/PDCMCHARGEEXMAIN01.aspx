<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGEEXMAIN01.aspx.vb" Inherits="PD.PDCMCHARGEEXMAIN01" %>
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
    Dim mobjccPDDCCHARGEEXCOM , mobjPDCMGET
    Dim mInsOKFlag 'Insert Flag 
    Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode '팝업사용시
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
	dim vntInParam
	dim intNo,i
	
    '서버업무객체 생성	
    Set mobjccPDDCCHARGEEXCOM = gCreateRemoteObject("cPDCO.ccPDDCCHARGEEXCOM")
    set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")

   '권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

   'InsOKFlag 를 false 값으로 설정한다.
	mInsOKFlag   =  false
	
	gSetSheetDefaultColor
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* 초기화면 입니다. "& vbcrlf & vbcrlf &"* 도움말: JOBNO을 선택하여 주시고, 반드시 처리버튼을 누르십시오."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "70"
		
	end with
	pnlTab1.style.visibility = "visible" 
	'화면 초기값 설정
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'기본값 설정
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)	
				case 1 : .txtJOBNAME.value = vntInParam(i)
				case 2 : .txtOUTSCODE.value = vntInParam(i)		'현재 사용중인 것만
				case 3 : .txtOUTSNAME.value = vntInParam(i)		'코드 사용 시점
				case 4 : mstrFields = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with
	Call imgFind_onclick
end Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.sprSht.MaxRows = 0
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

Sub EndPage()
	set mobjccPDDCCHARGEEXCOM = Nothing
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
	gFlowWait meWAIT_ON
	makePageData
	gFlowWait meWAIT_OFF
	
	'추가부분
	Dim i, RowNum, intRows
	RowNum = 101
	
	mobjSCGLSpr.SetMaxRows frmThis.sprSht, RowNum
	gOKMsgbox "데이터를 입력할 준비가 되었습니다. Excel Data를 붙여넣어 주십시요.", ""
				
	mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,1
	frmThis.sprSht.focus()
End Sub

Sub imgSave_onclick()
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	end if
	
    gFlowWait(meWAIT_ON)
    ProcessRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgDelete_onclick
    gFlowWait(meWAIT_ON)
    DeleteRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgClose_onclick()
    Window_OnUnload()
End Sub


Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtBUDGETDATE,frmThis.imgCalEndar,"txtBUDGETDATE_onchange()"
		gSetChange 
	end with
End Sub

Sub txtBUDGETDATE_onchange
	gSetChange
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
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
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
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'=============================
'SheetEvent
'=============================
Sub sprSht_Change(ByVal Col, ByVal Row)
   mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row

End Sub

Sub sprSht_KeyDown(KeyCode, Shift)
	mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
	'IF KeyCode = 86 THEN
	'	CALL TEST(KeyCode, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow)
	'END IF
End Sub

Sub sprSht_KeyUp(KeyCode, shift)
	If KeyCode = 86 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,1,100) <> "" then
			gErrorMsgbox "일괄투입시 한번에 투입가능한 데이터는 100건입니다. 다시 올려주십시오.",""
			mobjSCGLSpr.ClearText frmThis.sprSht , -1, -1, -1, -1 
			exit sub
		End If
	end if
end Sub

Sub ProcessRtn ()
	Dim intRtn   'Return 값
   	Dim vntData  'Insert 할 데이터
   	Dim vntData2
   	Dim strMasterData
   	Dim intCnt
   	Dim lngAMT
   	Dim lngCOMMI_RATE
   	Dim strCOMMISSION
   	Dim strYEARMON
   	Dim strJOBNO
   	dIM strOUTSCODE
   	dim strREVSEQ
   	'데이터 Validation
   	with frmThis
   		If trim(.txtJOBNO.value) = "" or trim(.txtOUTSCODE.value) = "" or .txtBUDGETDATE.value = ""   Then
			gErrorMsgBox "제작의뢰번호와 외주처코드, 견적일자는 필수 입니다.",""
			exit sub
		End If
		
		'여분 Rows 삭제처리
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMNAME",intCnt) = ""  then 
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			else
				CALL SetTrim (intCnt) ' 공백문자열 제거
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,0
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",intCnt,0
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",intCnt,0
				End If
				
			End If
		Next
		
		
		'==================오류검증
		'if DataValidation =false then exit sub
		'Exit SUb
		'==================수수료계산
		 For intCnt = 1 To .sprSht.MaxRows
            lngAMT =  mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)  
            if lngAMT = "" or lngAMT ="0" then
				if  mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt) <> "" _
					and mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) <> "0" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt)<> "0" THEN
					lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) *  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt)
					 mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,lngAMT
				else
					 mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,0
				END IF 
            end if
         Next
 	
 		strMasterData = gXMLGetBindingData (xmlBind)
 		
		strJOBNO = .txtJOBNO.value
		strOUTSCODE =.txtOUTSCODE.value
		strREVSEQ = 0
		'On error resume next
		'변경된 데이터를 가져온다.
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, "ITEMNAME|STD|QTY|PRICE|AMT|BIGO|ATTR01")
 	    if  not IsArray(vntData) then 
		    gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
		    exit sub
        end if
  	    Dim STime, ETime
  	   
  	    STime = Time
			intRtn = mobjccPDDCCHARGEEXCOM.ProcessRtn(gstrConfigXML, strMasterData, vntData, strJOBNO, strOUTSCODE, strREVSEQ, replace(.txtBUDGETDATE.value,"-",""), .txtMEMO.value)
		ETime = Time

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
		mobjSCGLSpr.SetTextBinding .sprSht,"ITEMNAME",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMNAME",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"STD",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"BIGO",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row))
	End With
End Sub

Function DataValidation ()
	dim i,j
	DataValidation = false
	with frmThis
		'데이터가 변경되었는지 검사
		if not mobjSCGLSpr.IsDataChanged(.sprSht) then
			gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
			exit function
		end if

   		'=================== 오류체크시작
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
   		Dim lngBONUS
   		Dim strCLIENTSUBNAME
   		Dim strCLIENTSUBCODE
   		Dim strDEPT_CD
   		Dim strSUBSEQ
   		Dim strMPPNAME, strMPPCODE
   		
   		
   		intVal = 0
   		'오류체크부분을 일단 공백으로 만든다.
   		 For intCnt = 1 To .sprSht.MaxRows
   			mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,""
   		 Next
   		 
   		 '광고주코드체크
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt),1,1) = "A" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt)) = 6 Then
				vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_CLIENTCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
				if not gDoErrorRtn ("SelectRtn_CODE") then
					IF mlngRowCnt <> 1 Then
						strERR = "광고주코드오류"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strCLIENTNAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
   				vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_CLIENTNAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTNAME)
	   			
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
   		
	   	 If intVal Then Exit Function
	   	 	
   		'=================================
	end with
	
	DataValidation = true
	
End Function

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
Sub makePageData
     Dim vntData
     
     With frmThis        
        
        gSetSheetDefaultColor() 
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 7, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, "ITEMNAME|STD|QTY|PRICE|AMT|BIGO|ATTR01"
        mobjSCGLSpr.SetHeader           .sprSht, "제작항목|규격|수량|단가|금액|비고|오류사항"
        mobjSCGLSpr.SetColWidth .sprSht, "-1","         20|  14|  14|  14|  16|  28|15"
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, "ITEMNAME|STD|BIGO"     , , ,200
        mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "QTY|PRICE|AMT", -1, -1, 0
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"        
       
    End With
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" height="100%" width="100%" >
				<TR>
					<TD>
						<TABLE id="tblTitle" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE" id="tblTitleName"><FONT face="굴림">&nbsp;견적 엑셀업로드</FONT></td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right"  height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD width="3"><IMG id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imginitOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imginit.gif'" height="20" alt="자료를 저장합니다."
													src="../../../images/imginit.gif" border="0" name="imgFind"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gif" width="54" border="0"
													name="imgDelete"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="HEIGHT: 17px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="1040" border="0" align="LEFT">
										<TR>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtJOBNO, txtJOBNAME)"><FONT face="굴림">Job&nbsp;No</FONT></TD>
											<TD class="DATA" width="420"><INPUT class="INPUT_L" id="txtJOBNAME" title="코드명" style="WIDTH: 256px; HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="37" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23"
													align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 88px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="6" size="9" name="txtJOBNO"></TD>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtOUTSCODE, txtOUTSNAME)"><FONT face="굴림">외주처</FONT></TD>
											<TD class="DATA"><INPUT class="INPUT_L" id="txtOUTSNAME" title="코드명" style="WIDTH: 256px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="37" name="txtOUTSNAME"><IMG id="imgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="imgOUTSCODE"><INPUT class="INPUT" id="txtOUTSCODE" title="jobno" style="WIDTH: 88px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="6" size="9" name="txtOUTSCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtBUDGETDATE, '')"><FONT face="굴림">견적일</FONT></TD>
											<TD class="DATA" width="420"><INPUT class="INPUT" id="txtBUDGETDATE" title="견적일" style="WIDTH: 128px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="100" align="left" size="16" name="txtBUDGETDATE"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
													name="imgCalEndar"></TD>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtMEMO, '')"><FONT face="굴림">비&nbsp;&nbsp; 
													고 </FONT>
											</TD>
											<TD class="DATA" colSpan="2"><FONT face="굴림"><INPUT class="INPUT_L" id="txtMEMO" title="비고" style="WIDTH: 336px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="50" name="txtMEMO"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<tr>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%;height:95%; POSITION: relative" 
									ms_positioning="GridLayout">
										<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 95%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5>
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="_ExtentX" VALUE="25321">
	<PARAM NAME="_ExtentY" VALUE="18680">
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
	<PARAM NAME="ReDraw" VALUE="-1">
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
									</div>
								</td>
							</tr>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
