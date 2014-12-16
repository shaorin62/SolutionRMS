<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMTRANSDTL.aspx.vb" Inherits="PD.PDCMTRANSDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>제작비 거래명세서</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
		<script language="vbscript" id="clientEventHandlersVBS">
Dim mobjPD_TRANS
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
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
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	Dim intCnt2
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		ModuleDir = "PD"
		ReportName = "PDCMTRANS.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= .txtTRANSYEARMON.value
		strTRANSNO		= .txtTRANSNO.value
		vntData = mobjPD_TRANS.Get_TRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_TRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum 
			for intCnt2=1 to 4
				strUSERID = ""
				vntDataTemp = mobjPD_TRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
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
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
	end with
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
	set mobjPD_TRANS		 = gCreateRemoteObject("cPDCO.ccPDCOTRANS")		'세금계산서조회
	

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "152px"
	'pnlTab1.style.height ="300px"
	'pnlTab1.style.left= "7px"
	

	
	mobjSCGLCtl.DoEventQueue
	
	
    gSetSheetDefaultColor
    with frmThis
		'vntInParam = window.dialogArguments
		'intNo = ubound(vntInParam)
		
		'기본값 설정
		
		'for i = 0 to intNo
		'	select case i
		'		case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
		'		case 1 : .txtTRANSNO.value = vntInParam(i)
		'	end select
		'next
		'화면의 깜박임을 방지하기 위함(Tab의 경우는 처음에 표시되는 것만 함)
		'.sprSht.style.visibility = "hidden"
		
		'**************************************************
		'***첫번째 Sheet 디자인
		'**************************************************
		
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout 디자인
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0
		
		'Binding Field 설정
	    mobjSCGLSpr.SpreadDataField .sprSht, "JOBNO|JOBNAME|INCJOBNO|INCJOBNAME|AMT|ADJAMT|BALANCE|TAXNO"
		'Header 디자인
		mobjSCGLSpr.SetHeader .sprSht,       "JOBNO|JOB명|포함청구JOBNO|포함청구JOB명|청구금액|정산액|잔액|계산서번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "11|20|11|20|10|10|10|12"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|ADJAMT|BALANCE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"JOBNO|JOBNAME|INCJOBNO|INCJOBNAME|AMT|ADJAMT|BALANCE|TAXNO"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|INCJOBNO|TAXNO",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|INCJOBNAME",-1,-1,0,2,false
		
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 8, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "JOBNO|JOBNAME|INCJOBNO|INCJOBNAME|AMT|ADJAMT|BALANCE|TAXNO"
		mobjSCGLSpr.SetText .sprSht_SUM, 1, 1, "합   계"
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUM, "JOBNO",-1,-1,2,2,false
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"1",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|ADJAMT|BALANCE", -1, -1, 0
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
		
	
	End with

	'pnlTab1.style.visibility = "visible" 
	'일단조회
	SelectRtn
	'화면 초기값 설정
	'InitPageData	
End Sub
'=================================합계쉬트 처리 시작
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub
'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
With frmThis
	mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM	
End with
end sub
'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_SUM, NewTop, NewLeft
End Sub
'시트에 금액을 합산한 값을 합계시트M에 뿌려준다.
Sub AMTSUM
	Dim lngCnt
	Dim lngAMT, lngADJAMT,lngBALANCE
	Dim lngAMTSUM,lngADJAMTSUM,lngBALANCESUM
	With frmThis
		lngAMTSUM = 0
		lngADJAMT = 0
		lngBALANCE = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			lngAMT = 0
			lngADJAMT = 0
			lngBALANCE = 0
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			lngADJAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngCnt)
			lngBALANCE = mobjSCGLSpr.GetTextBinding(.sprSht,"BALANCE", lngCnt)
			
			lngAMTSUM = lngAMTSUM + lngAMT
			lngADJAMTSUM = lngADJAMTSUM + lngADJAMT
			lngBALANCESUM = lngBALANCESUM + lngBALANCE
		Next
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, lngAMTSUM
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"ADJAMT",1, lngADJAMTSUM
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"BALANCE",1, lngBALANCESUM
	End With
End Sub

'=================================합계쉬트 처리 끝
'공급가액 금액처리
Sub txtAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub
Sub txtAMT_onblur
	
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub
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
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim intCnt
	Dim strCNT
	With frmThis
		strTRANSYEARMON		= .txtTRANSYEARMON.value
		strTRANSNO	= .txtTRANSNO.value
		IF strTRANSYEARMON = "" OR strTRANSNO = ""  THEN
			gErrorMsgBox "검색조건에 세금계산서 번호를 반드시 넣으셔야 합니다.","조회안내!"
			If strTRANSYEARMON = "" AND strTRANSNO = "" Then
			.txtTRANSYEARMON.focus
			Elseif strTRANSYEARMON = "" And strTRANSNO <> "" Then
			.txtTRANSYEARMON.focus
			Elseif strTRANSYEARMON <> "" And strTRANSNO = "" Then
			.txtTRANSNO.focus
			End If
			Exit Sub
		End If
	End With 
	
	IF not SelectRtn_HDR (strTRANSYEARMON, strTRANSNO) Then Exit Sub
	'쉬트 조회
	'Call SelectRtn_DTL 
	If not SelectRtn_DTL(strTRANSYEARMON, strTRANSNO) Then
		gErrorMsgBox "상세조회내역 조회실패","조회안내!"
		Exit Sub
	Else
		AMTSUM
	End If
	'SHEET1_SUM
	gWriteText lblStatus, "선택하신 세금계산세서 에 대하여 자료가 검색" & mePROC_DONE
End Sub
'-----------------------------------------------------------------------------------------
' 세금계산서조회HEADER
'-----------------------------------------------------------------------------------------
Function SelectRtn_HDR(ByVal strTRANSYEARMON, ByVal strTRANSNO)
	dim vntData
	'on error resume next
	'초기화
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPD_TRANS.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON,strTRANSNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 세금계산세서 번호 에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMT_onblur
			SelectRtn_HDR = True 
			'gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
		End IF
	End IF
End Function
'-----------------------------------------------------------------------------------------
' 세금계산서조회DETAIL
'-----------------------------------------------------------------------------------------
Function SelectRtn_DTL (ByVal strTRANSYEARMON, ByVal strTRANSNO)
	Dim vntData
	Dim lngCnt
	'on error resume next
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPD_TRANS.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON,strTRANSNO)
	
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		mobjSCGLSpr.SetClipbinding frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_DTL = True
	End IF
End Function

Sub EndPage()
	set mobjPD_TRANS = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	gClearAllObject frmThis
	
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="111" background="../../../images/back_p.gIF"
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
												<td class="TITLE">제작비 거래명세서</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 375px" vAlign="middle" align="right" height="28">
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
										<TABLE id="tblButton" style="WIDTH: 162px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											border="0">
											<TR>
												<TD></TD>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
												<TD><FONT face="굴림"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20"
															alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0" name="imgClose"></FONT></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
								<!--Top Define Table End-->
								<!--Input Define Table End--></TABLE>
							<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="굴림">
											<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON,txtTRANSNO)"
														width="80">명세서번호</TD>
													<TD class="SEARCHDATA"><INPUT dataFld="TRANSYEARMON" class="INPUT" id="txtTRANSYEARMON" title="세금계산서년월" style="WIDTH: 56px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtTRANSYEARMON">&nbsp;-
														<INPUT dataFld="TRANSNO" class="INPUT" id="txtTRANSNO" title="세금계산서번호" style="WIDTH: 48px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="4" size="2" name="txtTRANSNO"></TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1"
											cellPadding="0" align="left" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="90"><FONT face="굴림">광&nbsp; 고 주</FONT></TD>
												<TD class="SEARCHDATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME">
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="굴림">사&nbsp;&nbsp;업 부</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="굴림"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUTB_L" id="txtCLIENTSUBNAME" title="청구지명" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtCLIENTSUBNAME"></FONT>
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="굴림">청구일자</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="굴림"><INPUT dataFld="DEMANDDAY" class="NOINPUTB" id="txtDEMANDDAY" title="청구일" style="WIDTH: 172px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtDEMANDDAY"></FONT></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL"><FONT face="굴림">공급가액</FONT></TD>
												<TD class="SEARCHDATA"><FONT face="굴림"><INPUT dataFld="AMT" class="NOINPUTB_R" id="txtAMT" title="공급가액" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtAMT"></FONT>
												</TD>
												<TD class="SEARCHLABEL"><FONT face="굴림">부가세액</FONT></TD>
												<TD class="SEARCHDATA"></FONT></FONT><INPUT dataFld="VAT" class="NOINPUTB_R" id="txtVAT" title="부가세" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtVAT"></TD>
												<TD class="SEARCHLABEL"><FONT face="굴림">합계금액</FONT></TD>
												<TD class="SEARCHDATA"></FONT></FONT><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="합계" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtSUMAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">담당부서</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="DEPTNAME" class="NOINPUTB_L" id="txtDEPTNAME" title="사업자번호" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtDEPTNAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUMM,'')"><FONT face="굴림">비&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
														고</FONT></TD>
												<TD class="SEARCHDATA" colSpan="3"><INPUT dataFld="MEMO" class="NOINPUTB_L" id="txtMEMO" title="적요" style="WIDTH: 441px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="66" name="txtMEMO">
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 791px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 95%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31803">
									<PARAM NAME="_ExtentY" VALUE="8202">
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
								<OBJECT id="sprSht_SUM" style="WIDTH: 100%; HEIGHT: 8%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31803">
									<PARAM NAME="_ExtentY" VALUE="688">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;합 
							계 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="금액" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 794px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
