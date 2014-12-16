<%@ Page CodeBehind="MDCMELECTRANSLIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRANSLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 위수탁 거래명세 조회</title> 
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMGET 
Dim mobjMDCMELECTRANSLIST
Dim mstrCheck

mstrCheck = True

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgSetting_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_ConfirmOK
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmCancel_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_ConfirmCancel
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "년월을 입력하시오",""
		exit Sub
	end if
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

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
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
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjMDCMELECTRANSLIST.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_ELETRANS_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					datacnt = strcntsum + mlngRowCnt
					strUSERID = ""
					vntDataTemp = mobjMDCMELECTRANSLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
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
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

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
			SELECTRTN
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
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					SELECTRTN
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'****************************************************************************************
' 쉬트 더불클릭 이벤트
'****************************************************************************************

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
	
		If Row = 0 and Col = 1  then 'AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "N"
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) <> "" Then
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
				
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
				
				for intcnt = 1 to .sprSht.MaxRows
					sprSht_Change 1, intcnt
					
				next
				For intCnt = 1 To .sprSht.MaxRows
					If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
						'스태틱
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					'Else
						'체크
					'	mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
					End If			
				Next
			End IF
		end if
	end with
End Sub  	


sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		elseif Row = 0 and Col =1 then
		else
			strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",Row)
			strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",Row)
			
			vntInParams = array(strTRANSYEARMON, strTRANSNO) '<< 받아오는경우
			vntRet = gShowModalWindow("MDCMELECTRANSGUNLIST.aspx",vntInParams , 813,545)
			if isArray(vntRet) then
     		end if
		end if
	end with
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
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
	set mobjMDCMELECTRANSLIST = gCreateRemoteObject("cMDET.ccMDETELECTRANSLIST")
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONFIRMFLAG | TAXNO | TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | VAT | SUMAMTVAT"
		mobjSCGLSpr.SetHeader .sprSht,		"선택|승인|계산서|년월|번호|광고주코드|광고주|매체사코드|매체사|대행금액 |부가세|계"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|   6|    12|   8|	  6|	     0|	   25|	       0|    19|       13|    12|15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT |VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME|CONFIRMFLAG", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CONFIRMFLAG|TAXNO|TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME|AMT |VAT|SUMAMTVAT" 
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMFLAG | TAXNO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|CLIENTCODE|REAL_MED_CODE", true
		.sprSht.style.visibility = "visible"
    End With

	InitPageData
	'SelectRtn	
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECTRANSLIST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	with frmThis
		.txtYEARMON.value =  Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	dim vntData
	Dim strYEARMON, strCUSTCODE
   	Dim i, strCols
   	Dim intCnt
	on error resume next
	with frmThis
		strYEARMON	= .txtYEARMON.value
		strCUSTCODE	= .txtCLIENTCODE.value

		'초기화
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		'vntData = mobjMDCMELECTRANSLIST.Get_ELECTRANS_ALLLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, .txtCLIENTCODE.value , .txtCLIENTNAME.value)
		vntData = mobjMDCMELECTRANSLIST.Get_ELECTRANS_ALLLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, .txtCLIENTCODE.value , "")
		
		IF not gDoErrorRtn ("Get_ELECTRANS_ALLLIST") then
			'조회한 데이터를 바인딩
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				Else
					'체크
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
				End If			
			Next
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If 	mlngRowCnt < 1 Then
			.sprSht.MaxRows= 0
			End If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE	
		End IF
	end with
End Sub

'------------------------------------------
' 승인 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG 
	
	strFLAG = "CONFIRM"
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		'여기서 부터 문제
		'if DataValidation =false then exit sub
	    '데이터 Validation End
		On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG|TAXNO")
		
		intRtn = mobjMDCMELECTRANSLIST.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인" & mePROC_DONE
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 승인취소 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmCancel

    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG
	strFLAG = "CANCEL"
	with frmThis
   		'데이터 Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
			 IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		If lngCHKSUM = 0 Then
			gErrorMsgBox "저장할 데이터를 선택 하십시오.","저장안내!"
			Exit Sub
		End If
		
		
		'if DataValidation =false then exit sub
	    '데이터 Validation End
		On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG|TAXNO")
		
		intRtn = mobjMDCMELECTRANSLIST.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
	
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " 건의 자료가 승인취소" & mePROC_DONE
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			SelectRtn
   		end if
   		
   	end with
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 98%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
											<td class="TITLE">거래명세서 검증</td>
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
									<!--Wait Button End-->
									<!--Common Button Start-->
									
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--테이블이 무너지는것을 막아준다-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
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
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 95px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">년 
												월</TD>
											<TD class="SEARCHDATA" style="WIDTH: 375px"><INPUT class="INPUT" id="txtYEARMON" title="년월조회" accessKey="NUM" type="text" maxLength="6"
													size="10" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 95px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">광고주
											</TD>
											<TD class="SEARCHDATA" width="313"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" size="32" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 64px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" size="5" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>	
						</table>			
					</TD>
				</TR>				
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%;HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%">
						<!--테스트 시작-->
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0"
										width="50">
										<TR>
											<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
													height="20" alt="자료를승인처리합니다." src="../../../images/imgAgree.gIF" border="0" name="imgSetting">
											</TD>
											<td><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgAgreeCancelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAgreeCancel.gIF'"
													height="20" alt="승인처리를 취소합니다." src="../../../images/ImgAgreeCancel.gif" border="0"
													name="ImgConfirmCancel">
											</td>
											<td><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint">
											</td>
											<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel">
											</td>
										</TR>
									</TABLE>
									</TD>
							</TR>
						</TABLE>
									<!-- 추가 디자인끝-->
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="WIDTH: 100%; POSITION: relative; HEIGHT: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31829">
									<PARAM NAME="_ExtentY" VALUE="13520">
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
						<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
					</TR>
				</TABLE>
		</form>
	</body>
</HTML>
