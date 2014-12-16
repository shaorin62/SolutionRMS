<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_PREESTCONFIRM.aspx.vb" Inherits="PD.PDCMJOBMST_PREESTCONFIRM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>본견적확인</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_PREESTCONFIRM.aspx
'기      능 : JOBMST의 두번째 탭 PDCMJOBMST_ESTDTL 의 본견적으로 확인시 처리 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/01 By KimTH
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
Dim mobjPDCMPREESTDTL
Dim mstrPREESTNO



'DIVNAME,CLASSNAME,ITEMCODENAME,ITEMCODE,IMESEQ,SAVEFLAG
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	with frmThis
	
	End with
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ImgChSave_onclick()
	Dim intRtn
	
	intRtn = gYesNoMsgbox("수정 본견적으로 반영하시겠습니까?","저장안내")
	If intRtn <> vbYes then 
		exit sub
	End If
	window.returnvalue = "TRUE"
	Window_OnUnload
	
End Sub


Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCMPREESTDTL	= gCreateRemoteObject("cPDCO.ccPDCOPREESTDLT")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정

		'mstrPREESTNO,mstrITEMCODE,mlngIMESEQ
		for i = 0 to intNo
			select case i
				case 0 : mstrPREESTNO = vntInParam(i)			'견적번호
			end select
		next
		'PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|AMT
	'**************************************************
	'***이전본견적 디자인
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 6, 0
	mobjSCGLSpr.SpreadDataField .sprSht,    "ITEMNAME|STD|QTY|PRICE|AMT|SUSUAMT"
	mobjSCGLSpr.SetHeader .sprSht,		    "견적명|규격|수량|단가|금액|수수료"
	mobjSCGLSpr.SetColWidth .sprSht, "-1",  "16    |10  |4   |11  |11  |10"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,"ITEMNAME|STD|QTY|PRICE|AMT|SUSUAMT"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMNAME|STD",-1,-1,0,2,false ' 왼쪽
	
	'**************************************************
	'***수정본견적 디자인
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht1
	mobjSCGLSpr.SpreadLayout .sprSht1, 6, 0
	mobjSCGLSpr.SpreadDataField .sprSht1,    "ITEMNAME|STD|QTY|PRICE|AMT|SUSUAMT"
	mobjSCGLSpr.SetHeader .sprSht1,		    "견적명|규격|수량|단가|금액|수수료"
	mobjSCGLSpr.SetColWidth .sprSht1, "-1",  "16    |10  |4   |11  |11  |10"
	mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
	mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"ITEMNAME|STD|QTY|PRICE|AMT|SUSUAMT"
	mobjSCGLSpr.SetCellAlign2 .sprSht1, "ITEMNAME|STD",-1,-1,0,2,false ' 왼쪽
	
	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	
	End with

	'화면 초기값 설정
	InitPageData
	SelectRtn
	
End Sub

Sub InitpageData
	with frmThis
	
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub



'================================================================
'UI
'================================================================

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
Sub txtCHSUMAMT_onfocus
	with frmThis
		.txtCHSUMAMT.value = Replace(.txtCHSUMAMT.value,",","")
	end with
End Sub
Sub txtCHSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtCHSUMAMT,0,true)
	end with
End Sub




Sub EndPage
	Set mobjPDCMPREESTDTL = Nothing
	gEndPage
End Sub


'=============================================================
'조회
'=============================================================

Sub SelectRtn

	IF not SelectRtn_Head (mstrPREESTNO) Then Exit Sub
	'쉬트 조회

	CALL SelectRtn_leftDetail (mstrPREESTNO)

	CALL SelectRtn_RightDetail (mstrPREESTNO)
	with frmThis
	txtSUMAMT_onblur
	txtCHSUMAMT_onblur
	End with
End Sub

Function SelectRtn_Head(ByVal strPREESTNO)
	Dim vntData
	'on error resume next

	'초기화
	SelectRtn_Head = false
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMPREESTDTL.SelectRtn_confirmHDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strPREESTNO)
	
	IF not gDoErrorRtn ("SelectRtn_confirmHDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 견적에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function

Function SelectRtn_leftDetail(ByVal strPREESTNO)
	Dim vntData
	
	SelectRtn_leftDetail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_confirmLeftDTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strPREESTNO)
	
	If not gDoErrorRtn ("SelectRtn_confirmLeftDTL") then
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_leftDetail = True
	End If
End Function

Function SelectRtn_RightDetail(ByVal strPREESTNO)
	Dim vntData
	
	SelectRtn_RightDetail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_confirmRightDTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strPREESTNO)
	
	If not gDoErrorRtn ("SelectRtn_confirmRightDTL") then
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
		SelectRtn_RightDetail = True
	End If
End Function






		</script>
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF" border="0" >
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="91" background="../../../images/back_p.gIF"
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
								<td class="TITLE">본견적저장확인</td>
							</tr>
						</table>
					</td>
				</tr>
				<table class="SEARCHDATA">
					<tr>
						<td class="SEARCHDATA" style="WIDTH: 911px" width="911" colSpan="7">&nbsp;CLIENT <INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 224px; HEIGHT: 20px"
								accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="24" name="txtCLIENTNAME">&nbsp;PRODUCT
							<INPUT dataFld="PRODUCT" class="NOINPUTB_L" id="txtPRODUCT" title="제작건명" style="WIDTH: 224px; HEIGHT: 20px"
								accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="26" name="txtPRODUCT">
							&nbsp;PROJECT <INPUT dataFld="PROJECT" class="NOINPUTB_L" id="txtPROJECT" title="프로젝트명" style="WIDTH: 224px; HEIGHT: 20px"
								dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtPROJECT">&nbsp;<INPUT dataFld="PREESTNO" class="INPUT" id="txtPREESTNO" title="간접비" style="WIDTH: 16px; HEIGHT: 20px"
								accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="1" name="txtPREESTNO"></td>
						<td align="right" bgColor="#ecf2f9"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
								style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="화면을 닫습니다."
								src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
					</tr>
				</table>
			</table>
			<BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">본견적 / 현재견적 확인사항&nbsp;&nbsp;&nbsp;</td>
					<TD align="right" width="600"><IMG id="ImgChSave" onmouseover="JavaScript:this.src='../../../images/ImgChSaveOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgChSave.gIF'" height="20" alt="수정 본견적으로 저장합니다."
							src="../../../images/ImgChSave.gIF" align="absMiddle" border="0" name="ImgChSave">&nbsp;
					</TD>
				</tr>
			</table>
			<table class="SEARCHDATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
				<TR>
					<!--내용-->
					<TD class="SEARCHLABEL" style="WIDTH: 212px" width="100">이전 본견적 명</TD>
					<TD class="SEARCHDATA" width="337"><INPUT dataFld="CHPREESTNAME" class="NOINPUTB_L" id="txtCHPREESTNAME" title="수정 본견적 명"
							style="WIDTH: 228px; HEIGHT: 22px" dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtCHPREESTNAME"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px" width="212">수정 본견적 명</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="PREESTNAME" class="NOINPUTB_L" id="txtPREESTNAME" title="이전 본견적 명" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtPREESTNAME"></TD>
				</TR>
				<TR>
					<!--내용-->
					<TD class="SEARCHLABEL" style="WIDTH: 212px">이전 본견적일</TD>
					<TD class="SEARCHDATA" width="337"><INPUT dataFld="CHCONFIRMFLAG" class="NOINPUTB" id="txtCHCONFIRMFLAG" title="수정 본견적일" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtCHCONFIRMFLAG"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px" width="212">수정 본견적일</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="CONFIRMFLAG" class="NOINPUTB" id="txtCONFIRMFLAG" title="이전 &#13;&#10;&#9;&#9;&#9;&#9;&#9;&#9;본견적일"
							style="WIDTH: 228px; HEIGHT: 22px" dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtCONFIRMFLAG"></TD>
				</TR>
				<TR>
					<!--내용-->
					<TD class="SEARCHLABEL" style="WIDTH: 212px" width="100">이전본견적 합계금액</TD>
					<TD class="SEARCHDATA" width="337"><INPUT dataFld="CHSUMAMT" class="NOINPUTB_R" id="txtCHSUMAMT" title="수정본견적합계금액" style="WIDTH: 228px; HEIGHT: 22px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtCHSUMAMT"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px" width="212">수정본견적 합계금액</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="이전본견적 합계금액" style="WIDTH: 228px; HEIGHT: 22px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="32" name="txtSUMAMT"></TD>
				</TR>
			</table>
			<TABLE height="390" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD width="50%" height="390">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="21140">
								<PARAM NAME="_ExtentY" VALUE="10319">
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
					<TD width="50%" height="390">
						<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht1" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="21140">
								<PARAM NAME="_ExtentY" VALUE="10319">
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
			</TABLE>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 101.05%"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 101.05%"><FONT face="굴림"></FONT></TD>
				</TR>
			</table>
		</form>
	</body>
</HTML>
