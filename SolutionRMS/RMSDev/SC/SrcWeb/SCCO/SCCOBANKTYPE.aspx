<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOBANKTYPE.aspx.vb" Inherits="SC.SCCOBANKTYPE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>BANK_TYPE 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOCUSTMPPLIST.aspx
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/01/02 By OSH
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script type="text/javascript">
		
	function Set_IframeValue(strBUSINO,intCNT) {
	var value1  = strBUSINO;
	var value2  = intCNT;
	//iframe 서버컨트롤 텍스트 박스 busino 입력
	var textbox1 = frmSapCon.document.getElementById("<%=txtSAPBUSINO.ClientID%>");
	var textbox2 = frmSapCon.document.getElementById("<%=txtCNT.ClientID%>");
	
	textbox1.value = value1;
	textbox2.value = value2;
	window.frames[0].document.forms[0].submit();
}

		</script>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
dim mobjSCCOBANKTYPE		'비지니스로직
CONST meTAB = 9


'---------------------------------------------------
' 신규 SAP 값받아오기
'---------------------------------------------------
Sub Set_CustValue (strVALUE, strBANKTYPE)
	'BANK_TYPE 관련 정보 업데이트
    Dim firstArray_bank
    Dim secondArray_bank
    Dim strSAUPNOBANK, strBVTYP, strBANKL, strBANKN, strKOINH
    Dim i, strCNT
    
	
	With frmThis
		
		If strBANKTYPE = "" Then
			gErrorMsgBox "SAP 쪽에 존재하지않는 BANKTYPE 사업자번호입니다.",""
			.txtBUSINO.focus()
			.sprSht.focus()
			mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",.sprSht.ActiveRow, ""
			Exit Sub
		Else
			strCNT = 0
			
			firstArray_bank = Split(strBANKTYPE, ":")
			
			strCNT = ubound(firstArray_bank)
			
			mobjSCGLSpr.SetMaxRows .sprSht, strCNT + 1
			
			for i = 0 to strCNT
				strSAUPNOBANK = "" :  strBVTYP = "" :  strBANKL = "" :  strBANKN = "" :  strKOINH = ""
				secondArray_bank = Split(firstArray_bank(i), "|")
				
				strSAUPNOBANK = secondArray_bank(0)
                strBVTYP = secondArray_bank(1)
                strBANKL = secondArray_bank(2)
                strBANKN = secondArray_bank(3)
                strKOINH = secondArray_bank(4)
                
                mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",i + 1, trim(strSAUPNOBANK)
                mobjSCGLSpr.SetTextBinding .sprSht,"BANK_TYPE",i + 1, trim(strBVTYP)
                mobjSCGLSpr.SetTextBinding .sprSht,"BANK_KEY",i + 1, trim(strBANKL)
                mobjSCGLSpr.SetTextBinding .sprSht,"BANK_NUM",i + 1, trim(strBANKN)
                mobjSCGLSpr.SetTextBinding .sprSht,"BANK_USER",i + 1, trim(strKOINH)
                mobjSCGLSpr.SetTextBinding .sprSht,"USE_YN",i + 1, "Y"
			
			next

			.txtBUSINO.focus()
			.sprSht.focus()
		End If

	End With
End Sub

'====================================================
' 이벤트 프로시져 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' 명령 버튼 클릭 이벤트
'---------------------------------------------------
'-----------------------------------
'조회
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' 엑셀
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'추가
'-----------------------------------
sub ImgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtBUSINO.focus
		.sprSht.focus
	End With 
End sub


'-----------------------------------
' 저장   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub


'--------------------------------------------------
' SpreadSheet 이벤트
'--------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	with frmThis
		'사업자 번호를 입력하면 banktype 을 가져와야 한다.
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BUSINO") Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow) <> "" Then
				If Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow))) = 10 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",.sprSht.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow),1,3) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow),4,2) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow),6,5)
				elseIf Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow))) = 13 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",.sprSht.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow),1,6) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow),7,7)
				else
					mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",.sprSht.ActiveRow, Trim(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow))
				End If
				Set_IframeValue TRIM(mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",Row)) , 1
			end if
		end if 
		
		mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row    '달라진 시트의 열과 행을 찾는다
	end with
End Sub

'-----------------------------------
'쉬트 더블클릭
'-----------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End if
	End With
End Sub

'--------------------------------------------------
'쉬트 키다운
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		frmThis.sprSht.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,False,frmThis.sprSht.ActiveRow,mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"BUSINO"),mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"BUSINO"),True
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtBUSINO.focus
		frmThis.sprSht.focus
	End If
End Sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
Sub InitPage()
' 페이지 화면 디자인 및 초기화 
'----------------------------------------------------------------------
	'서버업무객체 생성	
	set mobjSCCOBANKTYPE = gCreateRemoteObject("cSCCO.ccSCCOBANKTYPE")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht	
		mobjSCGLSpr.SpreadLayout .sprSht, 7, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CUSTNAME | BUSINO | BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER | USE_YN"
		mobjSCGLSpr.SetHeader .sprSht,		 " 사업자명|사업자번호|은행키|계좌번호|BANK_TYPE|계정보유자|사용유무"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "     25|        10|     8|      13|        8|        25|       4"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USE_YN"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTNAME | BUSINO | BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER ", -1, -1, 200
		'mobjSCGLSpr.ColHidden .sprSht, "", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "BUSINO | BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER" ,-1,-1,2,2,false '가운데정렬
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "CUSTNAME | BUSINO | BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER | USE_YN"
		mobjSCGLSpr.CellGroupingEach .sprSht, "CUSTNAME | BUSINO"
		.sprSht.style.visibility = "visible"
    End With
    
	'화면 초기값 설정
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOBANKTYPE = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis

	'초기 데이터 설정
	With frmThis
		.sprSht.MaxRows = 0
		.txtBUSINO.value = ""
	End With
End Sub

'------------------------------------------
' HDR 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strBUSINO
   	
	With frmThis
		strBUSINO = ""	
		'Sheet초기화
		.sprSht.MaxRows = 0  '모든열을 초기화

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strBUSINO	= REPLACE(.txtBUSINO.value,"-","")
		vntData = mobjSCCOBANKTYPE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strBUSINO)

		If not gDoErrorRtn ("SelectRtn") Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			if mlngRowCnt > 0 then
				gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
			Else
				.sprSht.MaxRows = 0
				gWriteText lblStatus, mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE
			End if 	
   		End if
   	End With
End Sub

'------------------------------------------
' HDR 데이터 저장
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
	Dim lngCol, lngRow
	Dim strDataCHK
	Dim strBUSINO
	With frmThis

		 strBUSINO = ""
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "BUSINO | BANK_TYPE",lngCol, lngRow, False) 

		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " 줄의 사업자번호/BANKTYPE 은 필수 입력사항입니다.","저장안내"
			Exit Sub		 
		 End If

		 strBUSINO = mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",.sprSht.ActiveRow)

		 '모든 데이터를 저장플래그 설정
		 mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS

		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CUSTNAME | BUSINO | BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER | USE_YN")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			Exit Sub
		End If

		intRtn = mobjSCCOBANKTYPE.ProcessRtn(gstrConfigXml,vntData)

		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"

			.txtBUSINO.value = strBUSINO
			SelectRtn
   		End If
   		
   	End With
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF"
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
											<td class="TITLE">BANK_TYPE 관리</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
												width="100">사업자번호</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtBUSINO" title="코드조회" style="WIDTH: 168px; HEIGHT: 22px" maxLength="15"
													align="left" name="txtBUSINO">
												<asp:textbox id="txtSAPBUSINO" runat="server" Visible="False" Width="8px"></asp:textbox><asp:textbox id="txtCNT" runat="server" Visible="false" Width="8px"></asp:textbox></TD>
											<TD align="right" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<tr>
								<td>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20"></TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																alt="한 행 추가" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="16378">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 600px; HEIGHT: 600px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
