<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCOMMITLIST.aspx.vb" Inherits="MD.MDCMCOMMITLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>리포트 기초자료 생성 및 조회</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 인쇄매체
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMTRANSCONF.aspx
'기      능 : 작성된 거래명세서 의 Confirm 을 한다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/29 By Kim Tae Ho
'			 2) 
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
Dim mobjMDSRREPORTLIST
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

Sub imgQuery_Onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub


Sub EndPage()
	set mobjMDSRREPORTLIST = Nothing
	gEndPage
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()

	'서버업무객체 생성	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST") '조회

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "225px"
	'pnlTab1.style.height ="300px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor
    with frmThis
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout 디자인
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0,6
		'YEARMON|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|AMT
	    mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP"
		mobjSCGLSpr.SetHeader .sprSht,        "년월|전표구분|광고주|매체명|매체사명|사업부|브랜드|매체구분|대행사코드|취급액|부가세|소재명|광고주코드|매체코드|매체사코드|사업부코드|브랜드코드|매체구분코드|전표구분코드|부가세유무|MPP",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   7|       9|    15|    15|      15|    15|    15|       7|         0|    11|    11|    10|         0|       0|         0|         0|         0|           0|           0|         0|0"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON|VOCH_GBNNAME|MEDFLAGNAME",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|EXCLIENTCODE|MPP", true
	End with

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub SelectRtn()
	Dim vntData
	Dim i, strCols
	Dim strYEARMON
	Dim strVOCH_GBN
	Dim intCnt
	with frmThis
	'ON ERROR RESUME NEXT
		.sprSht.MaxRows = 0
		
		'월별 기초데이터 저장상태 state를 보여준다.
		SelectRtn_STATECHK
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		
		IF .rdT.checked = TRUE THEN
			strVOCH_GBN = "VOCH"
		ELSE
			strVOCH_GBN = "NOVOCH"
		END IF
		
		IF .cmbCOMMITGBN.value = 0 THEN
			vntData = mobjMDSRREPORTLIST.SelectRtn_LOW(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strVOCH_GBN)

			if not gDoErrorRtn ("SelectRtn_LOW") then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE			
   			end if
		ELSE
			vntData = mobjMDSRREPORTLIST.SelectRtn_REPORT(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strVOCH_GBN)

			if not gDoErrorRtn ("SelectRtn_REPORT") then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					
   				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			end if
		END IF
		
   		AMT_SUM
	End With
End Sub


Sub SelectRtn_STATECHK()
	Dim vntData
	Dim i, strCols
	Dim strYEAR
	Dim strVOCH_GBN
	Dim intCnt
	with frmThis
	'ON ERROR RESUME NEXT
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEAR = MID(.txtYEARMON.value,1,4)
		
		vntData1 = mobjMDSRREPORTLIST.SelectRtn_STATECHK(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEAR)
		
		If mlngRowCnt > 0 then
			for i = 1 to mlngRowCnt
				if vntData1(1,i) <> "" then
					if vntData1(1,i) = "VOCH" THEN
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = true
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
						document.getElementById("txtVOCH" & vntData1(0,i)).value = vntData1(2,i)
					elseif vntData1(1,i) = "NOVOCH" then
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = true
						document.getElementById("txtVOCH" & vntData1(0,i)).value = vntData1(2,i)
					else
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
						document.getElementById("txtVOCH" & vntData1(0,i)).value = ""
					end if
				else
					document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
					document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
					document.getElementById("txtVOCH" & vntData1(0,i)).value = ""
				end if
			next
		else
			for i = 1 to 12
				if i < 10 then
					document.getElementById("rdVOCH0" & i).checked = false
					document.getElementById("rdNOVOCH0" & i).checked = false
					document.getElementById("txtVOCH0" & i).value = ""
				else
					document.getElementById("rdVOCH" & i).checked = false
					document.getElementById("rdNOVOCH" & i).checked = false
					document.getElementById("txtVOCH" & i).value = ""
				end if
			next
		END IF
	
	End With
End Sub


'-----------------------------------------------------------------------------------------
' 화면 처리 SCRIPT
'-----------------------------------------------------------------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------
'------------------------------------------
' 삭제로직
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strYEARMON
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "삭제할 데이터가 없습니다.","처리안내!"
			Exit Sub
		end if
		
		if .cmbCOMMITGBN.value <> 1 then
			gErrorMsgBox "미생성된 자료는 삭제할 수 없습니다.","처리안내!"
			Exit Sub
		end if
		
		
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
	
		intRtn = mobjMDSRREPORTLIST.DeleteRtn(gstrConfigXml,strYEARMON)
		IF not gDoErrorRtn ("DeleteRtn") then
			gErrorMsgBox "자료가 삭제되었습니다.","삭제안내!"
			.sprSht.MaxRows = 0
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn_STATECHK
		'SelectRtn
	End with
	err.clear	
End Sub

Sub DeleteRtn_process ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strYEARMON
	with frmThis
		
		'선택된 자료를 끝에서 부터 삭제
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
	
		intRtn = mobjMDSRREPORTLIST.DeleteRtn(gstrConfigXml,strYEARMON)
		IF not gDoErrorRtn ("DeleteRtn") then
			If strDESCRIPTION <> "" Then
				gErrorMsgBox strDESCRIPTION,"삭제안내!"
				Exit Sub
			End If
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End IF
   		
	End with
	err.clear	
End Sub
'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		
		.sprSht.MaxRows = 0			
	end With
End Sub

Sub imgSave_onclick
	IF frmThis.cmbCOMMITGBN.value = "1" then
		gErrorMsgBox "미생성 상태에서만 저장가능합니다.","저장안내!"
		Exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'------------------------------------------
' 승인 저장로직
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
   	dim vntData
   	Dim vntData1
   	dIM strYEARMON
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
		vntData1 = mobjMDSRREPORTLIST.SelectRtn_EXISTREPORT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON)
		
		If mlngRowCnt > 0 then
			IF vntData1(1,1) = "VOCH" THEN
				intRtn = gYesNoMsgbox("이미 전표완료상태로 저장된 자료가 존재합니다. 다시 생성하시겠습니까?","자료저장")
				IF intRtn <> vbYes then exit Sub
				DeleteRtn_process
			elseif vntData1(1,1) = "NOVOCH" THEN
				intRtn = gYesNoMsgbox("이미 전표미완료상태로 저장된 자료가 존재합니다. 다시 생성하시겠습니까?","자료저장")
				IF intRtn <> vbYes then exit Sub
				DeleteRtn_process
			END IF
		END IF
		
		
		'On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		for i=1 to .sprSht.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP")
		
		intRtn = mobjMDSRREPORTLIST.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장성공!"
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			.cmbCOMMITGBN.value ="1"
			SelectRtn
   		end if
   	end with
End Sub

Sub AMT_SUM
	Dim lngCnt
	Dim lntTVAMT,		lntTVAMTSUM
	Dim lntRDAMT,		lntRDAMTSUM
	Dim lntDMBAMT,		lntDMBAMTSUM
	Dim lntCATVAMT,		lntCATVAMTSUM
	Dim lntINTERNETAMT, lntINTERNETAMTSUM
	Dim lntOUTDOORAMT,	lntOUTDOORAMTSUM
	Dim lntMP01AMT,		lntMP01AMTSUM
	Dim lntMP02AMT,		lntMP02AMTSUM

	With frmThis
		lntTVAMTSUM = 0
		lntRDAMTSUM = 0
		
		lntDMBAMTSUM = 0
		lntCATVAMTSUM = 0
		
		lntINTERNETAMTSUM = 0
		lntOUTDOORAMTSUM = 0
		
		lntMP01AMTSUM = 0
		lntMP02AMTSUM = 0
		
		'수수료 그리드 합계그리드 값넣기
		For lngCnt = 1 To .sprSht.MaxRows
			lntTVAMT = 0
			lntRDAMT = 0
			lntDMBAMT = 0
			lntCATVAMT = 0
			lntINTERNETAMT = 0
			lntOUTDOORAMT = 0
			lntMP01AMT = 0
			lntMP02AMT = 0
                
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "TV" THEN
				lntTVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntTVAMTSUM = lntTVAMTSUM  + lntTVAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "RD" THEN
				lntRDAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntRDAMTSUM = lntRDAMTSUM  + lntRDAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "DMB" THEN
				lntDMBAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntDMBAMTSUM = lntDMBAMTSUM  + lntDMBAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "CATV" THEN
				lntCATVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntCATVAMTSUM = lntCATVAMTSUM  + lntCATVAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "신문" THEN
				lntMP01AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntMP01AMTSUM = lntMP01AMTSUM  + lntMP01AMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "잡지" THEN
				lntMP02AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntMP02AMTSUM = lntMP02AMTSUM  + lntMP02AMT
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "인터넷" THEN
				lntINTERNETAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntINTERNETAMTSUM = lntINTERNETAMTSUM  + lntINTERNETAMT
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "옥외" THEN
				lntOUTDOORAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntOUTDOORAMTSUM = lntOUTDOORAMTSUM  + lntOUTDOORAMT
			END IF
		Next
		
		if .sprSht.MaxRows >0 Then
			.txtTV.value = lntTVAMTSUM
			.txtRD.value = lntRDAMTSUM
			
			.txtDMB.value = lntDMBAMTSUM
			.txtCATV.value = lntCATVAMTSUM
			
			.txtINTERNET.value = lntINTERNETAMTSUM
			.txtOUTDOOR.value = lntOUTDOORAMTSUM
			
			.txtMP01.value = lntMP01AMTSUM
			.txtMP02.value = lntMP02AMTSUM
			
			call gFormatNumber(.txtTV,0,true)
			call gFormatNumber(.txtRD,0,true)
			call gFormatNumber(.txtDMB,0,true)
			call gFormatNumber(.txtCATV,0,true)
			call gFormatNumber(.txtINTERNET,0,true)
			call gFormatNumber(.txtOUTDOOR,0,true)
			call gFormatNumber(.txtMP01,0,true)
			call gFormatNumber(.txtMP02,0,true)
		ELSE
			.txtTV.value = 0
			.txtRD.value = 0
			
			.txtDMB.value = 0
			.txtCATV.value = 0
			
			.txtINTERNET.value = 0
			.txtOUTDOOR.value = 0
			
			.txtMP01.value = 0
			.txtMP02.value = 0
		end if
	End With
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD style="HEIGHT: 54px">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIf" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;리포트 관리</td>
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
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 20px" vAlign="top" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
													width="90">년월&nbsp;
												</TD>
												<TD class="SEARCHDATA" width="180"><INPUT class="INPUT" id="txtYEARMON" title="년월" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="9" name="txtYEARMON">
												</TD>
												<TD class="SEARCHLABEL" width="90">생성구분</TD>
												<TD class="SEARCHDATA" width="120"><SELECT id="cmbCOMMITGBN" title="생성구분" style="WIDTH: 112px" name="cmbCOMMITGBN">
														<OPTION value="0" selected>미생성</OPTION>
														<OPTION value="1">생성</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" width="90">전표구분</TD>
												<TD class="SEARCHDATA">&nbsp;&nbsp;<INPUT id="rdT" title="확정내역조회" type="radio" CHECKED value="rdT" name="rdGBN">&nbsp;전표확정&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT id="rdF" title="미확정내역조회" type="radio" value="rdF" name="rdGBN">&nbsp;전표미확정
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20"
														alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery">
												</td>
											</TR>
										</TABLE>
										<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="굴림"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;리포트 기초자료 생성 및 조회</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																	height="20" alt="자료를 삭제합니다.." src="../../../images/imgDelete.gif" border="0" name="imgDelete"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1250px; HEIGHT: 10px">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px">
													<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">구분</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">1월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">2월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">3월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">4월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">5월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">6월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">7월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">8월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">9월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">10월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">11월</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">12월</TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">전표확정</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH01" title="1월" disabled type="checkbox" name="rdVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH02" title="2월" disabled type="checkbox" name="rdVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH03" title="3월" disabled type="checkbox" name="rdVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH04" title="4월" disabled type="checkbox" name="rdVOCH04"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH05" title="5월" disabled type="checkbox" name="rdVOCH05"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH06" title="6월" disabled type="checkbox" name="rdVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH07" title="7월" disabled type="checkbox" name="rdVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH08" title="8월" disabled type="checkbox" name="rdVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH09" title="9월" disabled type="checkbox" name="rdVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH10" title="10월" disabled type="checkbox" name="rdVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH11" title="11월" disabled type="checkbox" name="rdVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH12" title="12월" disabled type="checkbox" name="rdVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">미확정</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH01" title="1월" disabled type="checkbox" name="rdNOVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH02" title="2월" disabled type="checkbox" name="rdNOVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH03" title="3월" disabled type="checkbox" name="rdNOVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH04" title="4월" disabled type="checkbox" name="rdNOVOCH04"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH05" title="5월" disabled type="checkbox" name="rdNOVOCH05"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH06" title="6월" disabled type="checkbox" name="rdNOVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH07" title="7월" disabled type="checkbox" name="rdNOVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH08" title="8월" disabled type="checkbox" name="rdNOVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH09" title="9월" disabled type="checkbox" name="rdNOVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH10" title="10월" disabled type="checkbox" name="rdNOVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH11" title="11월" disabled type="checkbox" name="rdNOVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH12" title="12월" disabled type="checkbox" name="rdNOVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">생성일</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH01" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="7" name="txtVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH02" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH03" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH04" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH04"></TD>
															<TD class="DATA" width="80" style="TEXT-ALIGN: center"><INPUT class="NOINPUT_L" id="txtVOCH05" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH05"></TD>
															<TD class="DATA" width="80" style="TEXT-ALIGN: center"><INPUT class="NOINPUT_L" id="txtVOCH06" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH07" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH08" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH09" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH10" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH11" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH12" title="생성일" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD-->
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1038px; HEIGHT: 555px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 556px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 556px"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27437">
												<PARAM NAME="_ExtentY" VALUE="15505">
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
								<!--List End-->
								<!--BodySplit Start-->
								<!--Brench End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<TR>
									<TD>
										<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
													<TABLE class="DATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
														<TR>
															<TD class="LABEL" width="90">TV</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtTV" title="TV광고비" style="WIDTH: 152px; HEIGHT: 22px" type="text"
																	size="20" name="txtTV" readOnly></TD>
															<TD class="LABEL" width="90">RD</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtRD" title="라디오광고비" style="WIDTH: 152px; HEIGHT: 22px" type="text"
																	size="20" name="txtRD"></TD>
															<TD class="LABEL" width="90">CATV</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtCATV" title="케이블광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtCATV" readOnly></TD>
															<TD class="LABEL" width="90">지상파DMB</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtDMB" title="지상파DMB광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtDMB" readOnly></TD>
														</TR>
														<TR>
															<TD class="LABEL" width="90">신문</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtMP01" title="신문광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="20" name="txtMP01" readOnly></TD>
															<TD class="LABEL" width="90">잡지</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtMP02" title="잡지광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="21" name="txtMP02" readOnly></TD>
															<TD class="LABEL" width="90">인터넷</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtINTERNET" title="인터넷광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtINTERNET" readOnly></TD>
															<TD class="LABEL" width="90">옥외</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtOUTDOOR" title="옥외광고비" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtOUTDOOR" readOnly></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
