<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOPTLISTPOP.aspx.vb" Inherits="SC.SCCOPTLISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>PT_LIST 달력팝업</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'프로그램명 : SCCOPTLISTPOP.aspx
'기      능 : PT_LIST 달력팝업
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 20120503 By Oh Se Hoon
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

option explicit 
Dim mlngRowCnt, mlngColCnt		'달력데이터를 가져오기 위한 변수
Dim mlngRowCnt2, mlngColCnt2	'실제 날짜 데이터를 가져오기 위한 변수
Dim mobjSCCOPTLIST
CONST meTAB = 9

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage  '페이지 초기화'
End Sub

Sub Window_OnUnload()
	EndPage 
End Sub

Sub imgClose_onclick()
	EndPage '닫기버튼 클릭시 페이지 닫음'
End Sub

Sub imgQuery_onclick() 
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim i
	With frmThis
		mobjSCGLSpr.CellChanged .sprSht, Col, Row  '스프레드시트 변경 정보를 설정한다'
	End With
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intNo,i,vntInParam
	'서버업무객체 생성	
	set mobjSCCOPTLIST		= gCreateRemoteObject("cSCCO.ccSCCOPTLIST")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    gSetSheetDefaultColor
    
    with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		
		for i = 0 to intNo
			select case i
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : .txtSEQ.value = vntInParam(i)
			end select
		next
		
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0, 0,0  '스프레드 시트 디자인'
		mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | SUN | MON | TUE | WED | THU | FRI | SAT"
		mobjSCGLSpr.SetHeader .sprSht,		 "년월|일요일|월요일|화요일|수요일|목요일|금요일|토요일"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 0|    13|    13|    13|    13|    13|    13|    13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "45"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | SUN | MON | TUE | WED | THU | FRI | SAT", -1, -1, 200
		'mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, ""  '셀을 문자열 입력으로 설정  사용여부는 체크박스'
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | SUN | MON | TUE | WED | THU | FRI | SAT" '거래처 코드 잠금'
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON",-1,-1,2,2,False
		mobjSCGLSpr.SetCellAlign2 .sprSht, " SUN | MON | TUE | WED | THU | FRI | SAT",-1,-1,0,0,False
		
		.sprSht.style.visibility = "visible"
	
	End with
	'일단조회
	SelectRtn
End Sub

Sub EndPage()
	set mobjSCCOPTLIST = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 조회MASTER
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData,vntData2
	Dim lngCnt
	Dim strYEARMON
	Dim strYEAR
	Dim strMON
	Dim dblSEQ
	
	Dim strCLIENTNAME
	Dim strOTDATE
	Dim strPTDATE1,strPTDATE2,strPTDATE3
	
	With frmThis
	
		if .txtYEARMON.value = "" or .txtSEQ.value  = ""  then
			gErrorMsgBox "저장되지 않은 데이터이거나 년월이 비어있습니다.","조회안내!"
			exit sub
		end if
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0) : mlngRowCnt2 =clng(0): mlngColCnt2=clng(0)
		strYEARMON = "" : strYEAR = "" : strMON = "" : dblSEQ = ""
		
		strYEARMON	= .txtYEARMON.value
		strYEAR		= mid(.txtYEARMON.value ,1,4)
		strMON		= mid(.txtYEARMON.value ,6,2)
		dblSEQ		= .txtSEQ.value 
		
		strYEARMON = REPLACE(strYEARMON,"-","")

		'그리드에 달력 그리기 
		vntData = mobjSCCOPTLIST.SelectRtn_CalEndar(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEAR,strMON)
	
		vntData2 = mobjSCCOPTLIST.SelectRtn_date(gstrConfigXml, mlngRowCnt2, mlngColCnt2, strYEARMON,dblSEQ)
		
		IF not gDoErrorRtn ("SelectRtn_CalEndar") then
			IF mlngRowCnt > 0 THEN
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)

				strCLIENTNAME	= vntData2(2,1)
				strOTDATE		= vntData2(3,1)
				strPTDATE1		= vntData2(4,1)
				strPTDATE2		= vntData2(5,1)
				strPTDATE3		= vntData2(6,1)
				strOTDATE		= MID(strOTDATE,7,2)
				strPTDATE1		= MID(strPTDATE1,7,2)
				strPTDATE2		= MID(strPTDATE2,7,2)
				strPTDATE3		= MID(strPTDATE3,7,2)
				
				call sprShtdateBinding(strCLIENTNAME,strOTDATE,strPTDATE1,strPTDATE2,strPTDATE3)
				
				mobjSCGLSpr.SetCellShadow .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"SUN"), mobjSCGLSpr.CnvtDataField(.sprSht,"SUN"), -1, -1,&Hcc66ff, &H000000,False
				mobjSCGLSpr.SetCellShadow .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"SAT"), mobjSCGLSpr.CnvtDataField(.sprSht,"SAT"), -1, -1,&Hff9933, &H000000,False
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			END IF 
			
		End IF
	End With
End Sub

sub sprShtdateBinding(byval strCLIENTNAME,byval strOTDATE,byval strPTDATE1, byval strPTDATE2,byval strPTDATE3 )
	Dim i,j
	with frmThis 
		
		for i = 1 to .sprSht.MaxRows
			if strOTDATE <> "" then 
				for j = 2 to 8 
					IF strOTDATE = mobjSCGLSpr.GetTextBinding(.sprSht,j, i) THEN
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,j,i, mobjSCGLSpr.GetTextBinding(.sprSht,j, i) & "  " & strCLIENTNAME & "OT 일시"
					END IF 
				next
			end if
			
			if strPTDATE1 <> "" then 
				for j = 2 to 8 
					IF strPTDATE1 = mobjSCGLSpr.GetTextBinding(.sprSht,j, i) THEN
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,j,i, mobjSCGLSpr.GetTextBinding(.sprSht,j, i) & "  " & strCLIENTNAME & "PT 1차 일시"
					END IF 
				next
			end if
			
			if strPTDATE2 <> "" then 
				for j = 2 to 8 
					IF strPTDATE2 = mobjSCGLSpr.GetTextBinding(.sprSht,j, i) THEN
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,j,i, mobjSCGLSpr.GetTextBinding(.sprSht,j, i) & "  " & strCLIENTNAME & "PT 2차 일시"
					END IF 
				next
			end if
		
			if strPTDATE3 <> "" then 
				for j = 2 to 8 
					IF strPTDATE3 = mobjSCGLSpr.GetTextBinding(.sprSht,j, i) THEN
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,j,i, mobjSCGLSpr.GetTextBinding(.sprSht,j, i) & "  " & strCLIENTNAME & "PT 3차 일시"
					END IF 
				next
			end if
		next
	end with 
end sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="880" border="0">
				<!--Top TR Start-->
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
											<td align="left" height="4"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;&nbsp;&nbsp;&nbsp;PT리스트 월별내역</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 282px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 54px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="54" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose">
											</TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="880" border="0"> <!--TopSplit Start->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH:100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="left">
									<TABLE class="DATA" id="tblDATA" style="HEIGHT: 6px" cellSpacing="1" cellPadding="0" align="left"
										border="0">
										<TR>
											<TD class="LABEL" onclick="vbscript:Call gCleanField(txtYEARMON,'')">년월</TD>
											<TD class="DATA" width="170"></FONT><INPUT class="NOINPUTB_L" id="txtYEARMON" title="년월" style="WIDTH: 110px; HEIGHT: 22px"
													readOnly maxLength="255" align="left" size="22" name="txtYEARMON"><INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtSEQ" dataSrc="#xmlBind" class="NOINPUT_L"
													title="광고주코드" dataFld="SEQ" readOnly maxLength="6" size="3" name="txtSEQ">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--BodySplit Start-->
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 500px" vAlign="top" align="center">
									<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										DESIGNTIMEDRAGDROP="213">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="23256">
										<PARAM NAME="_ExtentY" VALUE="13229">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 880px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--Bottom Split End-->
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
