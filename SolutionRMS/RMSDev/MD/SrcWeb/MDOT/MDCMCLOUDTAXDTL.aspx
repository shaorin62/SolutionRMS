<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUDTAXDTL.aspx.vb" Inherits="MD.MDCMCLOUDTAXDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>세금계산서 상세내역</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 세금계산서 상세내역
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/24 By Kim tae yub
'           :2) 2009/09/11 By HWANG DUCK SU
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
Dim mobjMDOTCLOUDTAX
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
	'
End Sub
Sub imgClose_onclick()
	EndPage
End Sub
Sub imgQuery_onclick()
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
	
	IF frmThis.sprSht.MaxRows = 0 then
		gFlowWait meWAIT_ON
		with frmThis		
			ModuleDir = "MD"
			ReportName = "TAXNO_BLACK.rpt"
						
			IF .cmbFLAG.value = "receipt" THEN
				FLAG = "Y"
			ELSE
				FLAG = "N"
			END IF
						
			Params = FLAG
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
		end with
		gFlowWait meWAIT_OFF
	else
		
		gFlowWait meWAIT_ON
		with frmThis
			'인쇄버튼을 클릭하기 전에 md_tax_temp테이블에 내용을 삭제한다
			'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
			'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
			'md_trans_temp삭제 시작
			intRtn = mobjMDOTCLOUDTAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp삭제 끝
			
			ModuleDir = "MD"
			ReportName = "TAX_OUTDOORNEW.rpt"
			
			mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
			strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",1)
			strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",1)
			IF .txtVAT.value = 0 OR .txtVAT.value = "" THEN
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
			
			vntDataTemp = mobjMDOTCLOUDTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
			
			Params = "MD_TAXOUTDOOR_TEMP" & ":" & strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
			'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDOTCLOUDTAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'시트 변경 
Sub sprSht_Change(ByVal Col, ByVal Row)
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
end sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intNo,i,vntInParam
	'서버업무객체 생성	
	set mobjMDOTCLOUDTAX	= gCreateRemoteObject("cMDOT.ccMDOTCLOUDTAX")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "156px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	
    gSetSheetDefaultColor
    with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		
		for i = 0 to intNo
			select case i
				case 0 : .txtTAXYEARMON.value = vntInParam(i)	
				case 1 : .txtTAXNO.value = vntInParam(i)
			end select
		next
		
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0, 1, 2
	    mobjSCGLSpr.SpreadDataField .sprSht, "TAXYEARMON | TAXNO | TAXNOSEQ | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | AMT | VAT | SUMAMT"
		mobjSCGLSpr.SetHeader .sprSht,       "년월|번호|상세순번|광고주|광고주사업자번호|외주처|외주처사업자번호|금액|부가세|계",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "0 |   0|	   0|	 20|              15|     0|               0|  12|    12|12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXYEARMON | TAXNO | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"TAXYEARMON | TAXNO | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | AMT | SUMAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTBISNO | REAL_MED_BISNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON | TAXNO | TAXNOSEQ",true
	
	End with
	pnlTab1.style.visibility = "visible" 
	
	'일단조회
	SelectRtn
End Sub

Sub EndPage()
	set mobjMDOTCLOUDTAX = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	gClearAllObject frmThis
	
	'새로운 XML 바인딩을 생성
	frmThis.sprSht.MaxRows = 0
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'=================================합계쉬트 처리 끝
'부가세 금액처리
Sub txtVAT_onfocus
	with frmThis
		.txtVAT.value = Replace(.txtVAT.value,",","")
	end with
End Sub

Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

'합계금액 처리
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
	end with
End Sub

Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub

'-----------------------------------------------------------------------------------------
' 세금계산서조회MASTER
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim intCnt
	Dim strCNT
	With frmThis
		strTAXYEARMON	= .txtTAXYEARMON.value
		strTAXNO		= .txtTAXNO.value
		
		IF strTAXYEARMON = "" OR strTAXNO = ""  THEN
			gErrorMsgBox "검색조건에 세금계산서 번호를 반드시 넣으셔야 합니다.","조회안내!"
			If strTAXYEARMON = "" Then
				.txtTAXYEARMON.focus
			Elseif strTAXYEARMON <> "" And strTAXNO = "" Then
				.txtTAXNO.focus
			End If
			Exit Sub
		End If
	
		IF not SelectRtn_HDR (strTAXYEARMON, strTAXNO) Then Exit Sub
		
		'쉬트 조회
		If not SelectRtn_DTL(strTAXYEARMON, strTAXNO) Then
			gErrorMsgBox "상세조회내역 조회실패","조회안내!"
			.sprSht.MaxRows = 0
			Exit Sub
		End If
	End With 
	'SHEET1_SUM
	gWriteText lblStatus, "선택하신 세금계산세서 에 대하여 자료가 검색" & mePROC_DONE
End Sub
'-----------------------------------------------------------------------------------------
' 세금계산서조회HEADER
'-----------------------------------------------------------------------------------------
Function SelectRtn_HDR(ByVal strTAXYEARMON, ByVal strTAXNO)
	SelectRtn_HDR = false
	Dim vntData
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
	vntData = mobjMDOTCLOUDTAX.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strTAXYEARMON,strTAXNO)

	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 세금계산세서 번호 에 대하여" & meNO_DATA, ""
			InitPageData
			frmThis.txtTAXYEARMON.value = strTAXYEARMON
			frmThis.txtTAXNO.value = strTAXNO
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			call gFormatNumber(frmThis.txtAMT,0,true)
			call gFormatNumber(frmThis.txtVAT,0,true)
			call gFormatNumber(frmThis.txtSUMAMT,0,true)
			SelectRtn_HDR = True 
		End IF
	End IF
End Function
'-----------------------------------------------------------------------------------------
' 세금계산서조회DETAIL
'-----------------------------------------------------------------------------------------
Function SelectRtn_DTL (ByVal strTAXYEARMON, ByVal strTAXNO)
	SelectRtn_DTL = false
	Dim vntData
	Dim lngCnt
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDOTCLOUDTAX.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strTAXYEARMON,strTAXNO)
	
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		mobjSCGLSpr.SetClipbinding frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_DTL = True
	End IF
End Function


'------------------------------------------------
'부가세 적요 변경 저장 
'------------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim vntData
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strSUMM
	Dim i 
	
	with frmThis
		If .txtTAXYEARMON.value = "" Or .txtTAXNO.value = "" Or .txtSUMM.value = "" Then
			gErrorMsgBox "세금계산서 년월 및 번호를 입력하여주세요","저장안내!"
			Exit Sub
		End If
		
		strTAXYEARMON = .txtTAXYEARMON.value
		strTAXNO = .txtTAXNO.value
		strSUMM = .txtSUMM.value

		'시트의 변경된 데이터를 가져온다
		for i = 1 to .sprSht.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i			
		next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"TAXYEARMON | TAXNO | TAXNOSEQ | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | AMT | VAT | SUMAMT")
						
		intRtn = mobjMDOTCLOUDTAX.ProcessRtn_VAT(gstrConfigXml,vntData,strTAXYEARMON,strTAXNO,strSUMM)
				 
		if not gDoErrorRtn ("ProcessRtn_VAT") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "세금계산서 [" & strTAXYEARMON & "-" & strTAXNO & "] 의 부가세 및 적요가 수정" & mePROC_DONE,"저장안내" 
			SelectRtn
   		end if
	end with
End Sub


		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="880" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 427px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;CGV클라우드&nbsp;세금계산서</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="적요만 수정 가능합니다."
														src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다."
														src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"></TD>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
								<!--Top Define Table End-->
								<!--Input Define Table End--></TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="880" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 880px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 880px" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" width="80" onclick="vbscript:Call gCleanField(txtTAXYEARMON,txtTAXNO)">계산서번호</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtTAXYEARMON" title="세금계산서년월" style="WIDTH: 56px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="4" name="txtTAXYEARMON">&nbsp;-
													<INPUT class="INPUT" id="txtTAXNO" title="세금계산서번호" style="WIDTH: 48px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="4" size="2" name="txtTAXNO"></TD>
												<TD class="SEARCHDATA" width="152">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><SELECT id="cmbFLAG" title="영수/청구구분" style="WIDTH: 56px" name="cmbFLAG">
																	<OPTION value="receipt" selected>청구</OPTION>
																	<OPTION value="demand">영수</OPTION>
																</SELECT></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 880px; HEIGHT: 3px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 880px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="right" border="0">
											<TR>
												<TD class="LABEL" width="90">광고주</TD>
												<TD class="DATA" width="203"><INPUT dataFld="REAL_MED_NAME" id="txtREAL_MED_NAME" title="매체사명" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="255" size="22" name="txtDEPT_NAME" class="NOINPUT">
												</TD>
												<TD class="LABEL" width="90">담당부서</TD>
												<TD class="DATA" width="203"><INPUT dataFld="DEPT_NAME" class="NOINPUT" id="txtDEPTNAME" title="담당부서" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="255" size="22" name="txtDEPTNAME">
												</TD>
												<TD class="LABEL" width="90">청구일자</TD>
												<TD class="DATA" width="203"><INPUT dataFld="DEMANDDAY" class="NOINPUT" id="txtDEMANDDAY" title="청구일" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtDEMANDDAY" accessKey="DATE"></TD>
											</TR>
											<TR>
												<TD class="LABEL">광고주사업자</TD>
												<TD class="DATA"><INPUT dataFld="REAL_MED_BISNO" class="NOINPUT" id="txtREAL_MED_BISNO" title="매체사사업자번호"
														style="WIDTH: 200px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22"
														name="txtREAL_MED_BISNO">
												</TD>
												<TD class="LABEL">전표번호</TD>
												<TD class="DATA"><INPUT class="NOINPUT" id="txtVOCHNO" title="사업자번호" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="50" size="8" name="txtVOCHNO" dataFld="VOCHNO"></TD>
												<TD class="LABEL" style="CURSOR:hand" onclick="vbscript:Call gCleanField(txtSUMM,'')">
													적요</TD>
												<TD class="DATA"><INPUT dataFld="SUMM" class="INPUT_L" id="txtSUMM" title="적요" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" name="txtSUMM">
												</TD>
											</TR>
											<TR>
												<TD class="LABEL">공급가액</TD>
												<TD class="DATA"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="공급가액" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtAMT">
												</TD>
												<TD class="LABEL">부가세액</TD>
												<TD class="DATA"><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="부가세" style="WIDTH: 200px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="100" size="22" name="txtVAT" readOnly></TD>
												<TD class="LABEL">합계금액</TD>
												<TD class="DATA"><INPUT dataFld="SUMAMT" class="NOINPUT_R" id="txtSUMAMT" title="합계" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtSUMAMT"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 880px"></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 430px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 429px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="23283">
									<PARAM NAME="_ExtentY" VALUE="11351">
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
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 880px"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
