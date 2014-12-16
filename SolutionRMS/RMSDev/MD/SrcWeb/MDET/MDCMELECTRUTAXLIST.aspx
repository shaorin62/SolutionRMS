<%@ Page CodeBehind="MDCMELECTRUTAXLIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRUTAXLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 위수탁 세금계산서 조회</title> 
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
Dim mobjMDCMELECTRUTAXLIST , mobjMDCMGET
Dim mstrCheck
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
Sub imgFind_onclick
	TRANSPOP
End Sub

Sub imgQuery_onclick
	if frmThis.txtTAXYEARMON.value = "" then
	    gErrorMsgBox "년월 입력하시오",""
		exit Sub
	end if
	If LEN(frmThis.txtTAXYEARMON.value) <> 6 Then
		 gErrorMsgBox "년월은 6자리 입니다",""
		exit Sub
	End If
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
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
		intRtn = mobjMDCMELECTRUTAXLIST.DeleteRtn_TEMP(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "TRANSTAX_BLACK_NEW.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)

				strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
				strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
				
				strUSERID = ""
				
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
				
				vntDataTemp = mobjMDCMELECTRUTAXLIST.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO,  VATFLAG, FLAG, i, strUSERID)
			END IF
		next
		Params = strUSERID & ":" & "MD_TAX_TEMP"
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
		intRtn = mobjMDCMELECTRUTAXLIST.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ImgTAX_onclick ()
	TAXPOP
End Sub

Sub TAXPOP
	Dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis
		vntInParams = array(.txtTAXYEARMON.value, .txtTAXNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME.value, "trans","ELEC") '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMTAXPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			.txtTAXYEARMON.value = vntRet(0,0)  ' Code값 저장
			.txtTAXNO.value = vntRet(1,0)  ' 코드명 표시
			.txtCLIENTCODE.value = vntRet(2,0)  ' 코드명 표시
			.txtCLIENTNAME.value = vntRet(3,0)  ' 코드명 표시
			DateClean
			selectRtn
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
'청구일 조회조건 생성
Sub DateClean
Dim date1
Dim date2
Dim strDATE
	strDATE = MID(frmThis.txtTAXYEARMON.value,1,4) & "-" & MID(frmThis.txtTAXYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
		
	End With
End Sub
'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME.value) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = vntRet(0,0)  ' Code값 저장
			.txtCLIENTNAME.value = vntRet(1,0)  ' 코드명 표시
			
			gSetChangeFlag .txtCLIENTCODE             ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub


'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value,.txtCLIENTNAME.value)
											  
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(0,0)
					.txtCLIENTNAME.value = vntData(1,0)
				Else
					Call CLIENTCODE_POP()
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
Sub txtTAXYEARMON_onblur
	With frmThis
	If .txtTAXYEARMON.value <> "" AND Len(.txtTAXYEARMON.value) = 6 Then DateClean
	End With
End Sub

Sub imgFROM_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
End Sub

Sub imgTO_onclick
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
End Sub
Sub txtFROM_onchange
	gSetChange
End Sub
Sub txtTO_onchange
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
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
	strMEDFLAG = "A"
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If .rdT.checked = True Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row),strMEDFLAG) '<< 받아오는경우
				gShowModalWindow "MDCMCATVCOMMITAXDTL.aspx",vntInParams , 813,565
				SelectRtn
			End IF
		end if	
	end with
end sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDCMELECTRUTAXLIST	= gCreateRemoteObject("cMDET.ccMDETELECTRUTAXLIST")
	set mobjMDCMGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 36, 0, 1, 2
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXYEARMON | TAXNO | TAXMANAGE | TRANSYEARMON | TRANSNO | SEQ | DEMANDDAY | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | VAT | SUMAMT | SUMM | DEPT_NAME | PRINTDAY  | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPTCODE | MEDFLAG | VOCHNO | RANKTRANS | PROGRAM | SPONSOR| CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|년월|번호|관리번호|년월|번호|순번|청구일|광고주명|광고주사업자등록번호|매체사명|매체사사업자등록번호|매체명|대행금액|부가세액|합계금액|적요|부서명|발행일|광고주코드|광고주AC코드|매체사코드|매체사AC코드|MEDCODE|부서코드|집계구분|전표번호|순위|소재명|협찬구분|CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|	5|   4|	     11|   5|   4|	 4|     8|	    19|                  17|	  19|                  17|     0|       9|       9|	      9|  30|    10| 	 8|         0|           0|         0|           0|      0|       0|      10|      10|   0|     0|       0|          0|           0|           0|             0|             0|           0"
			
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "AMT|VAT|SUMAMT"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|CLIENTNAME|REAL_MED_NAME|MEDNAME|AMT|VAT|SUMAMT|DEPT_NAME|CLIENTCODE|CLIENTACCODE|CLIENTBISNO|REAL_MED_CODE|REAL_MED_ACCODE|REAL_MED_BISNO|MEDCODE|DEPTCODE|MEDFLAG|SEQ|VOCHNO|RANKTRANS|PROGRAM"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MEDNAME|REAL_MED_NAME|PROGRAM|TAXMANAGE", -1, -1, 20
		mobjSCGLSpr.ColHidden .sprSht, "TAXNO|TRANSYEARMON|TRANSNO|SEQ | CLIENTCODE|CLIENTACCODE|REAL_MED_CODE|REAL_MED_ACCODE|MEDCODE|DEPTCODE|MEDFLAG|TAXYEARMON|RANKTRANS|SPONSOR|CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2", true
		.sprSht.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMELECTRUTAXLIST = Nothing
	set mobjMDCMGET = Nothing
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
		.txtTAXYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		.txtTAXNO.focus()
	End with

	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCLIENTCODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		If .txtTAXYEARMON.value = "" Then
			gErrorMsgBox "년월을 입력하십시오","조회안내!"
			Exit Sub
		End If	
		If Len(.txtTAXYEARMON.value) <> 6 Then
			gErrorMsgBox "년월의 형식이 아닙니다.","조회안내!"
			Exit Sub
		End If
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strYEARMON	= .txtTAXYEARMON.value
		strTAXNO = .txtTAXNO.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		
		'세금계산서 완료조회
		vntData = mobjMDCMELECTRUTAXLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTAXNO,strCLIENTCODE, strFROM,strTO)
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 684px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
				width="684" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../images/TitleIcon.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">공중파&nbsp;위수탁 세금계산서&nbsp;조회</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 375px" vAlign="middle" align="right" height="28">
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
									<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="203" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<!--TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD-->
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="굴림">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTAXYEARMON, txtTAXNO)"
													width="90">년월/번호</TD>
												<TD class="DATA" style="WIDTH: 291px"><INPUT class="INPUT" id="txtTAXYEARMON" title="거래명세년월" style="WIDTH: 72px; HEIGHT: 22px"
														accessKey="MON" type="text" maxLength="6" size="6" name="txtTAXYEARMON"><IMG id="ImgTAX" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTAX"><INPUT class="INPUT" id="txtTAXNO" title="거래명세번호" style="WIDTH: 72px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTAXNO"></TD>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
													width="90">계산서청구일
												</TD>
												<TD class="DATA"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
														type="text" maxLength="6" size="6" name="txtFROM"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
														type="text" maxLength="6" size="6" name="txtTO"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgTo"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)"
													width="90">광고주&nbsp;</TD>
												<TD class="DATA" colSpan="3"><INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 72px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="6" name="txtCLIENTCODE"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 192px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="26" name="txtCLIENTNAME">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<SELECT id="chkPRINT" title="출력물구분" style="WIDTH: 88px" name="chkPRINT">
														<OPTION value="1" selected>일반용</OPTION>
														<OPTION value="0">공급받는자용</OPTION>
													</SELECT><SELECT id="cmbFLAG" title="영수/청구구분" style="WIDTH: 56px" name="cmbFLAG">
														<OPTION value="receipt" selected>청구</OPTION>
														<OPTION value="demand">영수</OPTION>
													</SELECT></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 791px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD align="center">
						<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 465px" height="101">
									<OBJECT id="sprSht" style="WIDTH: 790px; HEIGHT: 465px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="20902">
										<PARAM NAME="_ExtentY" VALUE="12303">
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
			</TD></TR></TBODY></TABLE></form>
	</body>
</HTML>
