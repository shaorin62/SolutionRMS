<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMESTDTLSRC.aspx.vb" Inherits="PD.PDCMESTDTLSRC" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>견적 내역관리</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 견적내역관리 화면(PDCMESTDTL)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPREESTDTL.aspx
'기      능 : 가견적 내역 등록 및 확정
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/16 By Tae Ho Kim
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
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMPREESTDTL '공통코드, 클래스
Dim mstrPROCESS
Dim mstrPROCESS2 '조회상태이면 true 신규상태이면 false
Dim mstrCheck
Dim mobjMDLOGIN
Dim mobjMDCMEMP
Dim mobjPDCMGET
CONST meTAB = 9
mstrPROCESS = TRUE
mstrPROCESS2 = TRUE
mstrCheck = True

'=============================
' 이벤트 프로시져 
'=============================
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


'Sub imgDelete_onclick
'	gFlowWait meWAIT_ON
'	DeleteRtn
'	gFlowWait meWAIT_OFF
'End Sub




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


'스프레드의 행을 더블 클릭 시 발생
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


'=============================
' UI업무 프로시져 
'=============================
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub
Sub txtSUMAMT_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtSUMAMT.value,",","")
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	'서버업무객체 생성	
	dim vntInParam
	dim intNo,i
	
	set mobjPDCMPREESTDTL	= gCreateRemoteObject("cPDCO.ccPDCOPREESTDLT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "232px"
	pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
				case 1 : frmThis.txtJOBNO.value = vntInParam(i)
			end select
		next
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis

		
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT"
		mobjSCGLSpr.SetHeader .sprSht,		  "가견적번호|순번|대분류|중분류|견적항목코드|견적항목명|내역|커미션|수량|단가|금액|수수료금액"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|   0|     8|    12|          10|        18|  28|     6|  12|  13|15  |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMCODENAME|STD", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT"
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|ITEMCODESEQ|SUSUAMT|ITEMCODESEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE|ITEMCODESEQ",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME|CLASSNAME",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0
	End With
	
	'화면 초기값 설정
	InitPageData	
	
	SelectRtn
End Sub

Sub EndPage()
	set mobjPDCMPREESTDTL = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub
'-----------------------------
' 확정 및 확정취소 처리
'-----------------------------	
Sub imgSetting_onclick
	Dim intRtnConfirm
	Dim intRtn
	intRtnConfirm = gYesNoMsgbox("자료를 확정 하시겠습니까?","자료확정 확인")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_Confirm(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_Confirm") then
				gErrorMsgBox " 자료가 확정 되었습니다.","확정안내" 
			End If
			ESTCONFIRM_Search
	End with
End Sub

Sub ImgConfirmCancel_onclick
	Dim intRtnConfirm
	Dim intRtn
	intRtnConfirm = gYesNoMsgbox("자료를 확정취소 하시겠습니까?","자료확정취소 확인")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
		intRtn = mobjPDCMPREESTDTL.ProcessRtn_ConfirmCancel(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmCancel") then
			gErrorMsgBox " 자료가 확정취소 되었습니다.","확정취소안내" 
		End If
		ESTCONFIRM_Search	
	End with
End Sub
'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtPRINTDAY.value = gNowDate
		.sprSht.MaxRows = 0
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'청구확정 조회
Sub ESTCONFIRM_Search
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
				.imgSetting.disabled = true
				.ImgConfirmCancel.disabled = false
			Else
				.imgSetting.disabled = false
				.ImgConfirmCancel.disabled = true
			End if
   		end if
	end with
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	Dim strJOBCODE
	With frmThis
		strCODE = .txtPREESTNO.value
		strJOBCODE = .txtJOBNO.value
		IF strCODE = ""  THEN
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = TRUE
			IF not SelectRtn_HeadLess (strJOBCODE) Then Exit Sub
			
		Else
			IF not SelectRtn_Head (strCODE) Then Exit Sub

			'쉬트 조회
			CALL SelectRtn_Detail (strCODE)
			txtSUSUAMT_onblur
			txtCOMMITION_onblur
			txtSUMAMT_onblur
			txtNONCOMMITION_onblur
		End If
	End With
End Sub
'기존내역이 없을 경우 조회
Function SelectRtn_HeadLess (ByVal strJOBCODE)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_HeadLess = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDRLESS(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBCODE)
	
	IF not gDoErrorRtn ("SelectRtn_HeadLess") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 견적에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_HeadLess = True
		End IF
	End IF
End Function
'기존내역이 있을 경우 조회
Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 가견적에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function


'예산 테이블 조회
Function SelectRtn_Detail (ByVal strCODE)
	dim vntData
	Dim intCnt
	Dim strRows
	'on error resume next
	'초기화
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		End with
	End IF
End Function

'****************************************************************************************
'이전 검색어를 담아 놓는다.
'****************************************************************************************
Sub PreSearchFiledValue (strTBRDSTDATE,strTBRDEDDATE, strCAMPAIGN_CODE, strCAMPAIGN_NAME, strCLIENTCODE, strCLIENTNAME)
	With frmThis
		.txtTBRDSTDATE1.value = strTBRDSTDATE
		.txtTBRDEDDATE1.value = strTBRDEDDATE
		.txtCAMPAIGN_CODE1.value = strCAMPAIGN_CODE
		.txtCAMPAIGN_NAME1.value = strCAMPAIGN_NAME
		.txtCLIENTCODE1.value = strCLIENTCODE
		.txtCLIENTNAME1.value = strCLIENTNAME
	End With
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%"cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;JOB 상세내역</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0"
							width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">견적 코드</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="PREESTNO" class="NOINPUT_L" id="txtPREESTNO" title="가견적코드" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtPREESTNO"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">견적명</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="PREESTNAME" class="NOINPUT_L" id="txtPREESTNAME" title="가견적명" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtPREESTNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAGREEYEARMON,'')">견적확정일</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="AGREEYEARMON" class="NOINPUT" id="txtAGREEYEARMON" title="견적합의일" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="10" name="txtAGREEYEARMON"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">제작 건명</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="제작건명" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtJOBNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">매체부문</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="매체부문" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtJOBGUBN"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">매체분류</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="매체분류" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtCREPART"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">광고주</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="광고주" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">사업부</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="사업부" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCLIENTSUBNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">브랜드</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUBSEQNAME"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 72px" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left"  height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;내역관리</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<tr height="5">
											<td></td>
										</tr>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="1040" border="0" align="LEFT">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">제작수수료</TD>
											<TD class="DATA" width="230"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="제작수수료" style="WIDTH: 224px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="32" name="txtSUSUAMT">
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand" align="right" width="94">Commition</TD>
											<TD class="DATA" width="260"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="commition 계" style="WIDTH: 256px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand" align="right" width="80">합계</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="총합계금액" style="WIDTH: 272px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUMAMT"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">수수료율</TD>
											<TD class="DATA"><INPUT dataFld="SUSURATE" class="NOINPUT_R" id="txtSUSURATE" style="WIDTH: 200px; HEIGHT: 22px"
													accessKey=",NUM,M" dataSrc="#xmlBind" readOnly type="text" size="28" name="txtSUSURATE">&nbsp;(%)
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px">Non Commition</TD>
											<TD class="DATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="noncommition 계"
													style="WIDTH: 256px; HEIGHT: 22px" accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37"
													name="txtNONCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" width="80">견적서 출력</TD>
											<TD class="DATA"><INPUT class="NOINPUT" id="txtPRINTDAY" title="견적서발행일" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" readOnly type="text" maxLength="10" size="10" name="txtPRINTDAY">&nbsp;&nbsp;
												<IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="견적서 를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" align="absMiddle"
													border="0" name="imgPrint">&nbsp;<INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtJOBNO"><INPUT dataFld="CREDAY" id="txtCREDAY" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCREDAY"><INPUT dataFld="CLIENTSUBCODE" id="txtCLIENTSUBCODE" style="WIDTH: 16px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="hidden" size="1" name="txtCLIENTSUBCODE"><INPUT dataFld="CLIENTCODE" id="txtCLIENTCODE" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCLIENTCODE"><INPUT dataFld="SUBSEQ" id="txtSUBSEQ" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtSUBSEQ"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27464">
											<PARAM NAME="_ExtentY" VALUE="12515">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
