<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCDEMPMST.aspx.vb" Inherits="SC.SCCDEMPMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>사원관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/27 By KIM TAE YUB
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
'전역변수 설정
Dim mobjSCCOEMPMST
Dim mobjSCCOGET
Dim mlngRowCnt,mlngColCnt
CONST meTAB = 9
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

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	mobjSCGLSpr.ExcelExportOption = true
	mobjSCGLSpr.ExportExcelFile frmThis.sprSht
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	if frmThis.sprSht.MaxRows = 0 then
		gErrMsgBox "저장할 자료가 없습니다.","저장 안내"
		exit sub
	end if
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgpwdinit_onclick ()
	if frmThis.sprSht.MaxRows = 0 then
		gErrMsgBox "비밀번호 초기화할 자료가 없습니다.","저장 안내"
		exit sub
	end if
	
	ProcessRtn_PWDCHANGE
End Sub

Sub imgE_HR_onclick
	Dim intRtn

	with frmThis
		
		intRtn = gYesNoMsgbox("사원정보를 최신정보로 교체 하시겠습니까?","EHR-연동확인")
		
		IF intRtn <> vbYes then exit Sub
		
		gFlowWait meWAIT_ON
		Call EHR_LOAD()
		window.setTimeout "SelectRtn", 3000	
		gOkMsgBox "변경되었습니다..","변경안내!"
		gFlowWait meWAIT_OFF
	End With
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성	
	Set mobjSCCOEMPMST = gCreateRemoteObject("cSCCO.ccSCCOEMPMST")
	set mobjSCCOGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'화면 초기값 설정
	InitPageData	
End Sub

Sub Grid_Layout()
	Dim intGBN
	Dim strComboList
	
	strComboList =  "재직" & vbTab & "휴직" & vbTab & "퇴직"
	gSetSheetDefaultColor
    
    with frmThis
		
		'**************************************************
		'***Sheet 디자인
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 0
		
		mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "EMPNO | EMP_NAME | HIGHCC_NAME | CC_CODE | BTN | CC_NAME | SC_EMP_STATUS | USE_YN | PASSWORD | MANAGER | TITLENAME | BIRTH | E_MAIL | TEL | CELLPHONE"
		mobjSCGLSpr.SetHeader .sprSht,		    "사번|사원명|상위부서명|부서코드|부서명|재직구분|사용여부|비밀번호|총괄승인권한|직책|생년월일|이메일|전화번호|핸드폰번호"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "  11|    20|        15|      15|2|  40|      12|      10|       0|          10|   0|       0|     0|       0|         0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USE_YN | MANAGER"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "EMPNO | EMP_NAME | CC_CODE | SC_EMP_STATUS",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CC_NAME",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "TITLENAME | BIRTH", True
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EMPNO | EMP_NAME | CC_CODE | CC_NAME | HIGHCC_NAME | TITLENAME | BIRTH | E_MAIL | TEL | CELLPHONE"
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,7,7,,,strComboList
	End with
	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOEMPMST.GetEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value,.txtEMP_NAME.value,.cmbGUBUN.value,.txtDEPTCD.value,.txtDEPTNAME.value,.cmbUSEYN.value)
		if not gDoErrorRtn ("SelectRtn") then
			If mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
			Else
				initpageData
			End If
   			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
End Sub

'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 변경시 체크 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CC_NAME",Row)))
		vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,435)
		IF isArray(vntRet) then
			mobjSCGLSpr.SetTextBinding .sprSht,"CC_CODE",Row, vntRet(0,0)
			mobjSCGLSpr.SetTextBinding .sprSht,"CC_NAME",Row, vntRet(1,0)			
			mobjSCGLSpr.CellChanged .sprSht, Col,Row
		End IF
		.txtEMPNO.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
	End with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		If Row = 0 and Col = 7  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 7,7,,, , , , , , mstrCheck
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
'여기까지 쉬트 버튼 클릭

'Validation
Function DataValidation ()
	DataValidation = false	
	With frmThis
		'IF not gDataValidation(frmThis) then exit Function	
	End With
	DataValidation = True
End Function
'저장로직

Sub ProcessRtn()
	Dim intRtn
   	dim vntData
   	mlngRowCnt=clng(0) : mlngColCnt=clng(0)

	with frmThis
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"EMPNO|EMP_NAME|CC_CODE|CC_NAME|SC_EMP_STATUS|USE_YN|PASSWORD|MANAGER")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		'처리 업무객체 호출
		intRtn = mobjSCCOEMPMST.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			If intRtn > 0 Then
				gErrorMsgBox intRtn & " 건 이 저장되었습니다.","저장안내"
			End If
			SelectRtn
   		end if
   	end with
End Sub

Sub ProcessRtn_PWDCHANGE()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strYEARMON
	Dim strSEQ

	with frmThis
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("ProcessRtn_PWDCHANGE") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "비밀번호를 초기화할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("비밀번호를 초기화 하시겠습니까?","저장 확인")
		IF intRtn <> vbYes then exit Sub
		
		for i = intSelCnt-1 to 0 step -1
			strEMPNO = mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",vntData(i))
			
			intRtn = mobjSCCOEMPMST.ProcessRtn_PWDCHANGE(gstrConfigXml,strEMPNO)
		next
		IF not gDoErrorRtn ("ProcessRtn_PWDCHANGE") then
			gOkMsgBox "사번으로 비밀번호가 초기화 되었습니다.","변경안내"
   		End IF
		SelectRtn
	End with
End Sub

'-----------------------------------------------------------------------------------------
' 광고부서팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------	
Sub ImgCC_onclick
	Call JOBREQU_DEPTCD_POP()
End Sub

Sub JOBREQU_DEPTCD_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC 코드/명,optional(현재사용여부,사용검사일,추가조회 필드,Key Like여부)
		vntInParams = array(trim(.txtDEPTNAME.value))
		vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = vntRet(0,0)	'Code값 저장
			.txtDEPTNAME.value = vntRet(1,0)	'코드명 표시
			gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub

Sub txtDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPTNAME.value))
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,1)
					.txtDEPTNAME.value = vntData(1,1)
				Else
					Call JOBREQU_DEPTCD_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

Sub EndPage()
	set mobjSCCOEMPMST = Nothing
	set mobjSCCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.MaxRows = 0
	End with
End Sub

</script>
<script language="javascript">
function EHR_LOAD(){
	ifrm_EHR.location.href = "SCCDEHREMP.asp";		
}
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
												<TABLE cellSpacing="0" cellPadding="0" width="71" background="../../../images/back_p.gIF"
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
											<td class="TITLE">사용자 관리&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1">
								</TD>
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
									<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,'')">사번
												</TD>
												<TD class="SEARCHDATA" width="81" style="WIDTH: 81px"><INPUT class="INPUT_L" id="txtEMPNO" style="WIDTH: 80px; HEIGHT: 22px" type="text" maxLength="8"
														size="8" name="txtEMPNO" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMP_NAME,'')">사원명
												</TD>
												<TD class="SEARCHDATA" width="89" style="WIDTH: 89px"><INPUT class="INPUT_L" id="txtEMP_NAME" style="WIDTH: 88px; HEIGHT: 22px" type="text" maxLength="255"
														name="txtEMP_NAME" size="9"></TD>
												<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">부서명
												</TD>
												<TD class="SEARCHDATA" width="215" style="WIDTH: 215px"><INPUT dataFld="DEPT_NAME" class="INPUT_L" id="txtDEPTNAME" title="담당부서명" style="WIDTH: 136px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="100" size="17" name="txtDEPTNAME"> <IMG id="imgCC" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="imgCC"> <INPUT dataFld="DEPT_CD" class="INPUT_L" id="txtDEPTCD" title="담당부서코드" style="WIDTH: 55px; HEIGHT: 22px"
														accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="3" name="txtDEPTCD">
												</TD>
												<TD class="SEARCHLABEL" width="70">재직구분
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 105px"><SELECT id="cmbGUBUN" title="사용구분" style="WIDTH: 105px" name="cmbGUBUN">
														<OPTION value="A">전체</OPTION>
														<OPTION value="0" selected>재직</OPTION>
														<OPTION value="1">휴직</OPTION>
														<OPTION value="3">퇴직</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" width="70">사용구분
												</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbUSEYN" title="사용구분" style="WIDTH: 104px" name="cmbUSEYN">
														<OPTION value="A">전체</OPTION>
														<OPTION value="Y" selected>사용</OPTION>
														<OPTION value="N">미사용</OPTION>
													</SELECT>
												</TD>
												<td class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left">
																<TABLE cellSpacing="0" cellPadding="0" width="98" background="../../../images/back_p.gIF"
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
															<td class="TITLE">사용자정보 변경&nbsp;</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgE_HR" onmouseover="JavaScript:this.src='../../../images/imgE_HROn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgE_HR.gIF'" height="20" alt="신규자료를 작성합니다."
																	src="../../../images/imgE_HR.gIF" border="0" name="imgE_HR"></TD>
															<TD><IMG id="imgpwdinit" onmouseover="JavaScript:this.src='../../../images/imgpwdinitOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgpwdinit.gIF'"
																	height="20" alt="자료를 저장합니다." src="../../../images/imgpwdinit.gIF" border="0" name="imgpwdinit"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgchangeSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgchangeSave.gIF'"
																	height="20" alt="자료를 저장합니다." src="../../../images/imgchangeSave.gIF" border="0" name="imgchangeSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--내용 및 그리드-->
								<TR vAlign="top" align="left">
									<!--내용-->
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="18256">
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
												<PARAM NAME="MaxCols" VALUE="11">
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
								<tr><td><iframe id="ifrm_EHR" width="0" height="0" frameborder="0"></iframe></td></tr>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
		
	</body>
</HTML>

