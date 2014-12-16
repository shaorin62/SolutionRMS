<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECEXMAIN01.aspx.vb" Inherits="MD.MDCMELECEXMAIN01" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>시스템 공통</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/MD/공중파
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMELECEXMAIN01.aspx
'기      능 : 소재별운행현황을 DOWNLOAD 받아 PASE AND COPY 를 하여 일괄 등록한다.
'컨트롤작성 : ccMDELECEXCOM,ccMDELECEXBrowse,ccMDCMGET
'엔티티작성 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/09/04 By Kim Tae Ho
'	         2) 
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
    Option explicit
    Dim mlngRowCnt, mlngColCnt
    Dim sprSht_DataFields
    Dim vntData_DataFields	
    Dim sprSht_DisplayFields
    Dim sprSht_ColWidth
    Dim sprSht_NotNull
    Dim vntData_Nullable
    Dim sprSht_DefualtValueFields
    Dim vntData_DefaultValue
    Dim vntData_DataType
    Dim vntData_DataLength
    Dim mobjccMDELECEXCOM
    Dim mInsOKFlag 'Insert Flag 
    Dim mobjMDCMGET
    Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode '팝업사용시
    Dim mstrdeletetemp
    mstrdeletetemp = false
'=============================
' 이벤트프로시져 
'=============================
Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
End Sub

Sub InitPage()
    '서버업무객체 생성	
    Set mobjccMDELECEXCOM = gCreateRemoteObject("cMDET.ccMDETELECEXCOM")
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")

   '권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	With frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'InsOKFlag 를 false 값으로 설정한다.
		mInsOKFlag   =  false

		gSetSheetDefaultColor
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* 년월을 선택하신 후 초기화 버튼을 눌러 주시기 바랍니다.."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "123"
		
		.txtYEARMON.readOnly = False
		
		pnlTab1.style.visibility = "visible"
	end with
	
	'Call imgFind_onclick()
end Sub

Sub EndPage()
	set mobjccMDELECEXCOM = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

'=============================
' 명령버튼클릭이벤트
'=============================
Sub imgFind_onclick
    Dim vntRet, vntInParams, dblTAB_ID
    Dim vntData
    Dim intRtn
    Dim strYEARMONTEMP
    Dim i, RowNum, intRows
	with frmThis
		If .txtYEARMON.value = ""  Then
			gErrorMsgBox "년월은 필수 입니다.","처리안내"
		End If
		
		If LEN(.txtYEARMON.value) <> 6 Then
			gErrorMsgBox "년월은 6자리 입니다.","처리안내"
		End If  
	
		gFlowWait meWAIT_ON
		   makePageData
		gFlowWait meWAIT_OFF
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMONTEMP = .txtYEARMON.value
		vntData = mobjccMDELECEXCOM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, sprSht_DataFields, strYEARMONTEMP)
		IF mlngRowCnt >0 THEN 
			intRtn = gYesNoCancelMsgBox("기존오류로 인해 투입되지 않은 자료가 있습니다. 다시 보시겠습니까?" & vbCrlf & "(예:다시보기,아니요:넘어가기,취소:자료삭제)","자료삭제 확인")
			IF intRtn = vbYes then 
				mstrdeletetemp = true
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG 
				mobjSCGLSpr.SetFlag  .sprSht, meINS_FLAG
				mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
			elseif intRtn = vbNo then
			'	Insert OK Flag 를 True 로 설정한다.
	   			mInsOKFlag = true
				mstrdeletetemp = false
				'추가부분
				
				RowNum = 5001
				
				mobjSCGLSpr.SetMaxRows .sprSht, RowNum
				gOKMsgbox "데이터를 입력할 준비가 되었습니다. Excel Data를 붙여넣어 주십시요.", ""
			elseif intRtn = vbCancel then
				intRtn = mobjccMDELECEXCOM.Delete_Temp_Rtn(gstrConfigXml, .txtYEARMON.value)
			end if
		else
		'	Insert OK Flag 를 True 로 설정한다.
	   		mInsOKFlag = true
			mstrdeletetemp = false
			'추가부분
			RowNum = 5001
			
			mobjSCGLSpr.SetMaxRows .sprSht, RowNum 
			gOKMsgbox "데이터를 입력할 준비가 되었습니다. Excel Data를 붙여넣어 주십시요.", ""
		END IF
	end with
				
	mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,1
	frmThis.sprSht.focus()
End Sub

Sub imgSave_onclick()
	Dim intRtn
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
		Exit Sub
	end if
    gFlowWait(meWAIT_ON)
    if mstrdeletetemp then 
		intRtn = mobjccMDELECEXCOM.Delete_Temp_Rtn(gstrConfigXml, frmThis.txtYEARMON.value)
    end if
    
    ProcessRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgDelete_onclick
    gFlowWait(meWAIT_ON)
    DeleteRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgClose_onclick()
    Window_OnUnload()
End Sub

'=============================
'SheetEvent
'=============================
Sub sprSht_KeyDown(KeyCode, Shift)
	If KeyCode = 86 Then
		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
	end if
End Sub

Sub sprSht_KeyUp(KeyCode, shift)
	If KeyCode = 86 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,7,5001) <> "" then
			gErrorMsgbox "5000건 이상의 데이터를 한번에 올리면 오류가 발생할 수 있습니다. 다시 올려주십시오.",""
			mobjSCGLSpr.ClearText frmThis.sprSht , -1, -1, -1, -1 
			exit sub
		End If
	end if
end Sub

'==================================================
'데이터를 처리
'==================================================
Sub ProcessRtn ()
	Dim intRtn
   	Dim vntData
   	Dim intCnt
   	Dim strYEARMON
   	
	with frmThis
		'여분 Rows 삭제처리
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_CLIENTCODE",intCnt) = ""  then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			else
				if mstrdeletetemp then 
					mobjSCGLSpr.CellChanged frmThis.sprSht, 6, intCnt
				end if
				
			End If
		Next
		'변경된 데이터를 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
 	    if  not IsArray(vntData) then 
		    gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
		    exit sub
        end if
		strYEARMON = .txtYEARMON.value 
		
		intRtn = mobjccMDELECEXCOM.ProcessRtn(gstrConfigXML, vntData, sprSht_DataFields, false, strYEARMON)

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
	   	    mobjSCGLSpr.SetMaxRows frmThis.sprSht, 0 
	   	    gOKMsgbox "데이터를 성공적으로 UPLOAD 하였습니다.", "" 
	   	    
	   	    '여기서부터 코드투입여부 확정
	   	    'MsgBox "브랜드 및 소재코드 투입을 시작합니다." & vbcrlf & "작업은 몇분 정도 걸립니다. 기다리십시오."
			strYEARMON = .txtYEARMON.value 
			'vntData = mobjccMDELECEXCOM.BatchCODE(gstrConfigXml, strYEARMON)
			'if not gDoErrorRtn ("BatchCODE") then
			'	gErrorMsgBox "[브랜드,소재] 코드 가 성공적으로 투입되었습니다." & vbcrlf & "반드시 브랜드관리에서 부서코드 를 등록하시고, " & vbcrlf & "소재관리 에서 대대행사를 등록 하십시오.","저장안내!"
			'End If
	   	    'Long Type의 ByRef 변수의 초기화
	   	    '저장시에 오류체크한후에 오류체크에 걸린 데이터는 TEMP테이블에 저장되었다가
	   	    '해당 데이터를 수정할수 있도록 검색해온다.
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
	   	    vntData = mobjccMDELECEXCOM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, sprSht_DataFields, strYEARMON)

	   	    if not gDoErrorRtn ("SelectRtn") then
				if mlngRowCnt >0 then
					mstrdeletetemp = true
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					gErrorMsgBox mlngRowCnt & "건의 자료가 데이터 이상으로 저장되지 않았습니다. 수정후 재저장하십시오","저장오류"
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG 
   				end if
   			end if
   			
	   	    'Insert OK Flag 를 True 로 설정한다.
	   	    mInsOKFlag = true
   		end if

   	end with
End Sub

Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i
	On error resume next
	with frmThis
		'한 건씩 삭제할 경우
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		if gDoErrorRtn ("DeleteRtn") then exit sub
		if intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit sub
		end if
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			end if
		next
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	end with
End Sub
'======================================
'기타함수
'======================================
Sub makePageData
     Dim vntData
     
     With frmThis
        .sprSht.MaxRows = 0
        sprSht_DataFields    = "MEDGUBUN | REGIONGUBUN | CLIENTSUBCODE | STD | EXCLIENTCODE | MATTERNAME | INPUT_MEDNAME | PROGRAM | INPUT_WEEK | TYPHOUR | BRDSTTIME | BRDEDTIME | ROLLSTDATE | ROLLEDDATE | TBRDSTDATE | TBRDEDDATE | CMLAN | CNT | PRICE | AMT | BRDDIV | ADSTOCFLAG | INPUT_AREAFLAGNAME | INPUT_CLIENTCODE | INPUT_MEDCODE | INPUT_MEDFLAG | INPUT_AREAFLAG | ADLOCALFLAG | ATTR01"
        sprSht_DisplayFields = "TV/RD|서울/기타|사업부|품목|제작사|소재|방송사|편성명|요일|시급|시작시간|종료시간|운행시작일|운행종료일|소재시작일|소재종료일|초수|횟수|단가|금액|운행구분명|청약구분명|본지사명|KOBACO광고주코드|방송사코드|매체코드|본지사코드|지역|오류사항"
  
        gSetSheetDefaultColor
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 29, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
        mobjSCGLSpr.SetHeader           .sprSht, sprSht_DisplayFields
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
        mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "CMLAN | CNT | PRICE | AMT", -1, -1, 0
        
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetColWidth         .sprSht, "-1", 10
    End With
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="82" background="../../../images/back_p.gIF"
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
											<td class="TITLE" id="tblTitleName">일괄청약 관리</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" height="20" alt="Loading"
													src="../../../images/imgCho.gif" width="64" border="0" name="imgFind"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="자료를 삭제합니다." src="../../../images/imgDelete.gif" width="54" border="0"
													name="imgDelete"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gif" width="54" border="0"
													name="imgClose"></TD>
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
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 90px">년월</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="해당년도" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM,M"
													type="text" maxLength="6" size="9" name="txtYEARMON"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<tr>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27517">
											<PARAM NAME="_ExtentY" VALUE="11774">
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
							</tr>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
