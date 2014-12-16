<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPOINTADEXCEL.aspx.vb" Inherits="MD.MDCMPOINTADEXCEL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>포인트 친구 AD 엑셀 업로드</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMPOINTADEXCEL.aspx
'기      능 : 거래명세서 승인전 엑셀 파일 업로드 
'파라  메터 : 
'특이  사항 : 엑셀 파일을 업로드 하여 POINT AD 프로그램의 데이터를 관리한다.
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2012/08/01 By OH Se Hoon
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script id="clientEventHandlersVBS" language="vbscript">
		
Dim mobjMDOTPOINTADCOMMI
Dim mstrTRANSYEARMON
Dim mstrTRANSNO
Dim mCAMPAIGN_CODE

Dim mlngRowCnt, mlngColCnt
'각각의 필드를 담기위함 전역변수
Dim sprSht_DataFields
Dim sprSht_DisplayFields

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------
'팝업 닫기 버튼 
Sub imgClose_onclick()
	EndPage
End Sub

'초기화 버튼 클릭
Sub imgFind_onclick
	gFlowWait meWAIT_ON
	EXCEL_UPLOAD
	gFlowWait meWAIT_OFF
End sub

'조회버튼 
Sub imgQuery_onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'저장버튼 클릭
Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'엑셀 출력 버튼
Sub imgExcel_onclick ()
	with frmThis
		gFlowWait meWAIT_ON
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
		gFlowWait meWAIT_OFF
	end with
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
	
	set mobjMDOTPOINTADCOMMI  = gCreateRemoteObject("cMDOT.ccMDOTPOINTADCOMMI")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
    with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)

		'기본값 설정
		for i = 0 to intNo
			select case i
				case 0 : mstrTRANSYEARMON = vntInParam(i)	
				case 1 : mstrTRANSNO = vntInParam(i)
				case 2 : mCAMPAIGN_CODE = vntInParam(i)
			end select
		next
		
		gSetSheetDefaultColor
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* 엑셀 업로드를 위해 초기화 버튼을 눌러 주시기 바랍니다.."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "110"
		
		.txtTRANSYEARMON.value = mstrTRANSYEARMON
		.txtTRANSNO.value = mstrTRANSNO
		.txtCAMPAIGN_CODE.value = mCAMPAIGN_CODE
		
		if mstrTRANSYEARMON = "" or mstrTRANSNO = "" or mCAMPAIGN_CODE = "" then
			gErrorMsgBox "상세내역을 확인하는데 필요한 조건이 충분하지 않습니다. 관리자에게 문의하세요.","상세내역 오류!"
			EndPage
		end if 
	
	End with
	pnlTab1.style.visibility = "visible" 
End Sub

Sub EndPage()
	set mobjMDOTPOINTADCOMMI = Nothing
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


'-----------------------------------------------------------------------------------------
'엑셀 입력을 위한 초기화
'-----------------------------------------------------------------------------------------
Sub EXCEL_UPLOAD
	with frmThis
		
		'기본 페이지를 다시 그려서 시트에 붙여 넣거나 입력을 받도록 한다.
		makePageData
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'최초 조회시 데이터가 존재 하면 데이터를 보여주고 존재 하지 않으면 입력을 유도한다.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			intRtn = gYesNoMsgBox("이미 현재 거래명세서의 상세 내역이 존재 합니다. 보시겠습니까?" & vbCrlf & "(예:다시보기,아니요:자료삭제)","자료삭제 확인")
			IF intRtn = vbYes then
			
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG 
				mobjSCGLSpr.SetFlag  .sprSht,meINS_FLAG
				
				'조회한다.
				SelectRtn
			elseif intRtn = vbNo then
				'사용자가 원하지 않을경우 데이터를 삭제 한다.
				DeleteRtn
			end if
		ELSE
			RowNum = 500
			mobjSCGLSpr.SetMaxRows .sprSht, RowNum
			gOKMsgbox "데이터를 입력할 준비가 되었습니다. Excel Data를 붙여넣어 주십시요.[최대 500 개의 데이터입력이 가능 합니다.]", " EXCEL UPLOAD"
		end if
		
	End with
End sub


'-----------------------------------------------------------------------------------------
'엑셀 파일 조회
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	with frmThis
		
		'기본 페이지를 다시 그려서 시트에 붙여 넣거나 입력을 받도록 한다.
		makePageData
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'최초 조회시 데이터가 존재 하면 데이터를 보여주고 존재 하지 않으면 입력을 유도한다.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			'조회시 입금여부를 체크박스로 변경한다.
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "PAY_YN"
			Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,-1,1,16,True
			
			
			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
		ELSE
			.sprSht.MaxRows = 0
			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
		end if
		
	End with
End Sub


'======================================
'시트를 다시 그리기
'======================================
Sub makePageData
     
     With frmThis
        .sprSht.MaxRows = 0
        sprSht_DataFields    = "POINTNO | CHN | CLIENTNAME | GACODE | EXCLIENTNAME | ADEXCLIENTNAME | CLIENT_TYPE | ADCLIENTCODE | TITLE | TDATE | EDATE | SAND_STATUS | SAND_DATE | PAY_YN | AMT | CDATE"
        sprSht_DisplayFields = "번호|채널|광고주명|가맹점코드|대행사|업업사|광고유형|광고코드|광고명|이벤트시작일|이벤트마감일|발송상태|발송일자|입금여부|광고단가|등록일자"	
  
        gSetSheetDefaultColor
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 16, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
        mobjSCGLSpr.SetHeader           .sprSht, sprSht_DisplayFields
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
        mobjSCGLSpr.SetCellTypeFloat2	.sprSht, "AMT", -1, -1, 0
        
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetColWidth         .sprSht, "-1", 10
    End With
End Sub


'------------------------------------------------
'엑셀 업로드 저장 
'------------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim intCnt
	Dim vntData
	
	with frmThis
	
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'해당 거래명세서 내역의 상세 내역이 있는지 조회한다.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			gErrorMsgBox "저장된 데이터가 있습니다 확인하시고 다시 저장해 주십시요.!","저장취소"
			exit sub
		end if 
	
		'여분 Rows 삭제처리
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"POINTNO",intCnt) = ""  then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			END IF
		Next

		mobjSCGLSpr.SetFlag  .sprSht,meINS_FLAG
		
		'변경된 데이터를 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
	
 	    if not IsArray(vntData) then 
		    gErrorMsgBox "변경된 " & meNO_DATA,"저장취소"
		    exit sub
        end if
		
		intRtn = mobjMDOTPOINTADCOMMI.ProcessRtn_EXCEL(gstrConfigXML, vntData, sprSht_DataFields, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
	   	    mobjSCGLSpr.SetMaxRows frmThis.sprSht, 0 
	   	    gOKMsgbox "데이터를 성공적으로 UPLOAD 하였습니다.", "엑셀 업로드 안내!" 
	   	    '업로드후 조회한다.
	   	    SelectRtn
	   	 END IF

	end with
End Sub

'------------------------------------------------
'UPLOAD 했던 EXCEL 데이터를  일괄 삭제 합니다.
'------------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intRtn, i

	'On error resume next
	with frmThis
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		
		intRtn = mobjMDOTPOINTADCOMMI.DeleteRtn_EXCEL(gstrConfigXml,mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		'삭제후 초기화한다.
		InitPage
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	end with
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="880">
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIF"
							height="28">
							<TR>
								<TD style="WIDTH: 427px" height="28" width="427" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td rowSpan="2" width="14" align="left"><IMG src="../../../images/TitleIcon.gIF" width="14" height="28"></td>
											<td height="4" align="left"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;포인트 친구 AD &nbsp;엑셀 업로드</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 282px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="처리중입니다."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="880">
							<TR>
								<TD style="WIDTH: 880px" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 880px" class="KEYFRAME" vAlign="middle" align="center">
									<TABLE id="tblKey" class="DATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTAXYEARMON,txtTAXNO)"
												width="100">거래명세서번호</TD>
											<TD style="WIDTH: 124px" class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 56px; HEIGHT: 22px" id="txtTRANSYEARMON" class="NOINPUT_L"
													title="거래명세서년월" readOnly maxLength="6" size="4" name="txtTRANSYEARMON">&nbsp;-
												<INPUT accessKey="NUM" style="WIDTH: 48px; HEIGHT: 22px" id="txtTRANSNO" class="NOINPUT_L"
													title="거래명세서번호" readOnly maxLength="4" size="2" name="txtTRANSNO"></TD>
											<td class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 122px; HEIGHT: 22px" id="txtCAMPAIGN_CODE" class="NOINPUT_L"
													title="캠페인 코드" readOnly maxLength="4" size="2" name="txtCAMPAIGN_CODE">
											</td>
											<TD align = "right"><IMG style="CURSOR: hand" id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgFind"
													alt="Loading" src="../../../images/imgCho.gif" width="64" height="20">
													<IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery"
													alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgSave"
													alt="적요만 수정 가능합니다." src="../../../images/imgSave.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" border="0" name="imgClose"
													alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" height="20">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 880px; HEIGHT: 3px" class="TOPSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				<TR>
					<TD style="WIDTH: 880px" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" class="LISTFRAME" vAlign="top" align="center">
						<DIV style="POSITION: relative; WIDTH: 100%; VISIBILITY: hidden" id="pnlTab1" ms_positioning="GridLayout">
							<OBJECT style="WIDTH: 100%; HEIGHT: 550px" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="23256">
								<PARAM NAME="_ExtentY" VALUE="14552">
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
					<TD style="WIDTH: 880px" id="lblStatus" class="BOTTOMSPLIT"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
