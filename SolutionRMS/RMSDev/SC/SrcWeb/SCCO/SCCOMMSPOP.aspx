<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOMMSPOP.aspx.vb" Inherits="SC.SCCOMMSPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>MPP 조회</title> 
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCOMPPPOP.aspx
'기      능 : MPP 팝업
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/07 By KTY
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjSCCOGET 
Dim mobjPDCMGET
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

CONST meTAB = 9
'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
end sub

Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick
	Dim intRtn , i
	Dim vntData
	Dim vntData_info
	Dim strMODE
	Dim strCLIENTNAME 
	Dim strFromUserName, strFromUserEmail, strFromUserPhone
	Dim strToUserName, strToUserEmail, strToUserPhone
	
	with frmThis
		'모드 1은 한건 보내기 || 모드2는 모두 보내기 
		strMODE = 1 
		
		'승인권자가 업으면 
		if .txtEMPNO.value = "" then
			gErrorMsgBox "승인권자를 선택하세요.","선택안내!"
			exit Sub
		end if
		
		intRtn = gYesNoMsgbox("청구요청을 하시겠습니까?","청구요청 확인")
		IF intRtn <> vbYes then exit Sub
		
		IF .sprSht.MaxRows > 1 Then
			intRtn = gYesNoMsgbox("한건전송:예||모두전송:아니오","청구요청 선택")
			IF intRtn <> vbYes then strMODE = 2
		End if
		
		
		'보내는 사람의 정보 가져오기
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
		
		'보내는사람정보
		strFromUserName		= vntData_info(0,2)
		strFromUserEmail	= vntData_info(1,2)
		strFromUserPhone	= vntData_info(2,2)
		
		'받는사람 정보
		strToUserName		=  vntData_info(0,1)
		strToUserEmail		=  vntData_info(1,1)
		strToUserPhone		=  vntData_info(2,1)
		
		if strMODE = 1 then
			strCLIENTNAME	= mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",1)	
			call SMS_SEND(strCLIENTNAME,strFromUserName,strFromUserPhone,strToUserPhone)
		else
			for i = 0 to .sprSht.MaxRows
				strCLIENTNAME	= mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",i)
				call SMS_SEND(strCLIENTNAME,strFromUserName,strFromUserPhone,strToUserPhone)
			next
		end if
		
		msgbox ("전송이 완료되었습니다.")
		imgCancel_onclick
	End with
end sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------------------------------------------------------------------
' 사원코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'실제 데이터List 가져오기
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../../../PD/SrcWeb/PDCO/PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()				' 포커스 이동
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag 변경 알림
			gSetChangeFlag .txtEMPNAME
			
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetPDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					'.txtMEMO.focus()
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntData
	dim intCol,intRow,i,j ,intcnt
	
	'서버업무객체 생성	
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntData = window.dialogArguments
		intCol = ubound(vntData, 1)
		intRow = ubound(vntData, 2)
		
		
		'기본값 설정
		'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
    
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 5, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "CLIENTCODE|CLIENTNAME|AMT|SUMAMTVAT|USERNAME"
		mobjSCGLSpr.SetHeader .sprSht, "광고주코드|광고주|공급가액|합계금액|요청자"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 0|    16|      12|      14|     6"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "CLIENTCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "CLIENTNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "AMT"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "SUMAMTVAT"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "USERNAME"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", TRUE '
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CLIENTCODE|CLIENTNAME|AMT|SUMAMTVAT|USERNAME"
		mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
		mobjSCGLSpr.SetCellAlign2 .sprSht, "USERNAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME",-1,-1,0,2,false

		frmThis.sprSht.MAXROWS = intRow
		intcnt= 1
		
		for i = 1 to intRow
			if vntData(1,i) = "1" then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",intcnt, TRIM(vntData(2,i))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",intcnt, TRIM(vntData(3,i))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",intcnt, TRIM(vntData(4,i))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",intcnt, TRIM(vntData(5,i))
				intcnt = intcnt+1
			end if
		next
		
		'선택된 자료를 끝에서 부터	
			for i = intRow to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",i) = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End If
			Next
		
	end with		
	
end sub

Sub EndPage()
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtHIGHCUSTCODE.value,.txtCUSTNAME.value, "P")

		if not gDoErrorRtn ("GetHIGHCUSTCODE") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			if mlngRowCnt <> 0 then
   				.sprSht.focus()
   			else
   				.sprSht.MaxRows = 0
   				.txtCUSTNAME.focus()
   			end if 
   		end if
   	end with
end sub

-->
		</script>
		<script language="javascript">
		
		function SMS_SEND(strCLIENTNAME, strFromUserName , strFromUserPhone, strToUserPhone){
		
			frmSMS.location.href = "SMS.asp?CLIENTNAME="+ strCLIENTNAME + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
		</script>
	</HEAD>
	<body class="base"  bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												승인요청&nbsp;
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 20px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD width="20"><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'" height="20"
													alt="화면을 닫습니다." src="../../../images/imgCancel.gif" border="0" name="imgCancel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
										<TABLE class="SEARCHDATA" id="tblKey" style="WIDTH: 392px" cellSpacing="0" cellPadding="0" width="392"
											align="right" border="0">
											<TBODY>
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNAME,txtEMPNO)">
														승인권자</TD>
													<TD class="SEARCHDATA" colspan="2"><INPUT class="INPUT_L" id="txtEMPNAME" title="사원조회" style="WIDTH: 140px; HEIGHT: 22px"
															type="text" maxLength="255" align="left" size="20" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"  align="absMiddle"
															border="0" name="ImgEMPNO"> <INPUT class="INPUT" id="txtEMPNO" title="사번조회" style="WIDTH: 70px; HEIGHT: 22px" readOnly
															type="text" maxLength="8" align="left" size="10" name="txtEMPNO">
													</TD>
													<TD class="SEARCHDATA" width="20"><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gif'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gif'"
															height="20" alt="자료를 승인합니다." src="../../../images/ImgConfirmRequest.gif" border="0" name="imgConfirm"></TD>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT id="sprSht" style="WIDTH: 392px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="10372">
											<PARAM NAME="_ExtentY" VALUE="7250">
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
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="굴림"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
		<iframe id="frmSMS" style="DISPLAY: none; WIDTH: 10px; HEIGHT: 10px" name="frmSMS"></iframe>
		
	</body>
</HTML>