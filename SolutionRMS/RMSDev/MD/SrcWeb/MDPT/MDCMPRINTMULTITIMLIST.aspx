<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTMULTITIMLIST.aspx.vb" Inherits="MD.MDCMPRINTMULTITIMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>광고주별 팀별 검색</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/인쇄 광고주별 팀별 신문잡지 금액 화면(MDCMPRINTMULTILIST)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : MDCMPRINTMULTILIST.aspx.aspx
'기      능 : 인쇄 광고주별 팀별 신문잡지 금액
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2010/03/17 By 황덕수
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
DIm mobjMDCOGET
Dim mobjMDSRPRINTMULTILIST
Dim mlngRowCnt,mlngColCnt
Dim mintCnt
Dim mintCnt2
Dim mvntData
Dim mvntData2
Dim mstrField
Dim mvntDataExist
Dim mintCntExist
Dim mstrFieldExist
Dim mstrClientcode
Dim mvntDataCust
Dim mvntDataCustCNT
Dim mClientsubcode


'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick
	EndPage
End Sub

Sub imgQuery_onclick
	if frmThis.txtYEAR.value = "" then
		gErrorMsgBox "년도를 입력하시오",""
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SELECTRTN
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------

'광고주팝업버튼
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub GetCLIENTSUBLIST (strCLIENTCODE)
	Dim vntData
   	Dim i, strCols
   	Dim strHTML
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strHTML = "" 
		mClientsubcode = ""
		vntData = mobjMDSRPRINTMULTILIST.GetCLIENTSUBLIST_PRINT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE, .txtYEAR.value, .cmbMED_FLAG.value)
		if not gDoErrorRtn ("GetCLIENTSUBLIST_PRINT") then
			If mlngRowCnt > 0 Then
				For i = 0 to mlngRowCnt-1
					strHTML = strHTML & "<INPUT id='"& vntData(0,i) & "' type='checkbox' name='"&  vntData(0,i) & "' checked>" & vntData(1,i) & "&nbsp;&nbsp;"
					IF i = 0 THEN
						mClientsubcode = mClientsubcode & vntData(0,i)
					ELSE
						mClientsubcode = mClientsubcode & "♥" & vntData(0,i)
					END IF
				next
			Else
				strHTML = ""
			End If
			document.getElementById("tdCLIENTSUB").innerHTML = strHTML
			
   		end if
   	end with
End Sub

sub cmbMED_FLAG_onchange
	with frmThis
		if mstrClientcode <> "" then
			Call SelectRtn	
		end if
	end with
end sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'서버업무객체 생성
	set mobjMDSRPRINTMULTILIST	= gCreateRemoteObject("cMDSC.ccMDSCPRINTMULTILIST")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 0, 0, 0, 0,5

    End With

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSRPRINTMULTILIST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		.txtYEAR.value = mid(gNowDate,1,4)
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtYEAR.focus()
	End with
End Sub

'조회
Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	Dim chkflag
   	Dim strSUBLIST
   	Dim strCLIENTSUBLIST
   	Dim intSUBRow
   	Dim strMONCNT
   	Dim strLIST
   	Dim tmon, fmon
	Dim strYEARMONLAST
	
	With frmThis
		
		SetChangeLayout
		
		.sprSht.MaxRows = 0
		strSUBLIST = ""
		chkflag = 1
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTSUBLIST=""
		
		strCLIENTSUBLIST = 	split(mClientsubcode,"♥")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
		
		strSUBLIST = replace(strSUBLIST,"없음","")
		
		vntData = mobjMDSRPRINTMULTILIST.SelectRtn_PRINTCUSTAndTIM(gstrConfigXml, mlngRowCnt, mlngColCnt, mvntDataCust, mvntDataCustCNT, .txtYEAR.value, .txtCLIENTCODE.value, strSUBLIST, .cmbMED_FLAG.value)
		
		If not gDoErrorRtn ("SelectRtn_PRINTCUSTAndTIM") then
			IF mlngRowCnt <> 0 THEN
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			else
				.sprSht.MaxRows = 0
			END IF
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		END IF
   		Layout_change
   	End With
End Sub


Sub SetChangeLayout () 
	Dim strYEAR
	Dim intAddCnt,intAddHeadCnt,intAddWith,intFieldSetting,intHide,intFloat,intAddCnt2 'For 문 Count변수
	Dim vntData
	Dim strStartHead
	Dim strClientAndMed
	Dim i
	Dim strHead
	Dim strHeadCLIENT
	Dim strAddField
	Dim strField
	Dim intLayOutCnt
	Dim chkflag
	
	mvntDataCustCNT = ""
	mvntDataCust = ""
	mstrField = ""
	gInitComParams mobjSCGLCtl,"MC"
	
	With frmThis
		
		strSUBLIST = ""
		chkflag = 1
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		strCLIENTSUBLIST=""
		
		strCLIENTSUBLIST = 	split(mClientsubcode,"♥")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
		
		strSUBLIST = replace(strSUBLIST,"없음","")
		
		mvntDataCust = mobjMDSRPRINTMULTILIST.GetPRINTCLIENTTIMCNT(gstrConfigXml,mlngRowCnt,mlngColCnt, .txtYEAR.value, .txtCLIENTCODE.value, strSUBLIST, .cmbMED_FLAG.value)
		mvntDataCustCNT = mlngRowCnt
		If not gDoErrorRtn ("GetPRINTCLIENTTIMCNT") then
			If mlngRowCnt > 0 Then 
				'필드 고정값세팅
				
				strField = "YEAR|CUST"
				
				'필드 증가값세팅 [광고주코드]
				
				strAddField = ""
				For intAddCnt = 1 To mvntDataCustCNT
					strAddField = strAddField & "|A" & intAddCnt
				Next
				
				'필드 증가값 [값]
				mstrField = strField & strAddField & "|SUMAMT"
				'헤더 고정값세팅
				
				strHead = .txtYEAR.value & "년|"
				'헤더 증가값세팅
				
				strHeadCLIENT = ""
				strStartHead = ""
				
				For intAddHeadCnt = 1 To  mvntDataCustCNT
					strHeadCLIENT   = strHeadCLIENT & "|" & TRIM(mvntDataCust(0,intAddHeadCnt))
				Next
				strStartHead = strHead & strHeadCLIENT & "|계"
				'넓이 고정값세팅
				Dim strWith
				strWith = "13|13"
				'넓이 증가값세팅
				Dim strAddWith
				Dim strEndWith
				strAddWith = ""
				strEndWith = ""
				For intAddWith = 1 To mvntDataCustCNT
					strAddWith = strAddWith & "|13"
				Next
				strEndWith = strWith & strAddWith & "|13"
				
				
				'총컬럼갯수
				intLayOutCnt = ""
				intLayOutCnt = 2 + mvntDataCustCNT + 1
				'여기까지 괜찮음
				
				Call Grid_init()
				
				gSetSheetColor mobjSCGLSpr, .sprSht
				
				'Sheet Layout 디자인
				mobjSCGLSpr.SpreadLayout .sprSht, intLayOutCnt, 0,2
				mobjSCGLSpr.SpreadDataField .sprSht, mstrField 
				mobjSCGLSpr.SetHeader .sprSht,       strStartHead ,0,1,true
				mobjSCGLSpr.AddCellSpan .sprSht, 1, SPREAD_HEADER + 0, 2    , 1      , 0 , true
				mobjSCGLSpr.SetColWidth .sprSht, "-1", strEndWith
				'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, strField, , , 50, , ,2
				mobjSCGLSpr.SetCellTypeFloat2 .sprSht, mstrField, -1, -1, 0
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEAR|CUST", , , 50, , ,0
				mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
				mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,strField
				mobjSCGLSpr.CellGroupingEach .sprSht, "YEAR"
				mobjSCGLSpr.SetCellAlign2 .sprSht, "YEAR|CUST",-1,-1,2,2,false
			ELSE
				'Sheet 기본Color 지정
				gSetSheetDefaultColor() 
				
				With frmThis
					gSetSheetColor mobjSCGLSpr, .sprSht
					
				End With
			End If
   		End if
   		
   	End With
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "MON"
		mobjSCGLSpr.SetHeader .sprSht,		 "MON"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MON", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MON"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MON",-1,-1,2,2,false
	End With
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If mobjSCGLSpr.GetTextBinding(.sprSht,"CUST",intCnt) = "계" Then
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		End If
	Next 
	End With
End Sub

		</script>
	</HEAD>
	<body class="base">
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
												<TABLE cellSpacing="0" cellPadding="0" width="190" background="../../../images/back_p.gIF"
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
											<td class="TITLE">광고비 집행내역 - 광고주별 팀별</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="창을 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End-->
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
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" width="70" title="년도을삭제합니다." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEAR,'')">년도
											</TD>
											<TD class="SEARCHDATA" width="120"><INPUT class="INPUT" id="txtYEAR" title="년을입력하세요" style="WIDTH: 120px; HEIGHT: 22px" type="text"
													maxLength="4" size="14" name="txtYEAR" accessKey="NUM">
											</TD>
											<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">광고주
											</TD>
											<TD class="SEARCHDATA" width="289"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="코드명" style="WIDTH: 180px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"  src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="코드조회" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
											</TD>
											<TD class="SEARCHLABEL" width="70">매체구분
											</TD>
											<TD class="SEARCHDATA" >
												<SELECT name="cmbMED_FLAG" id="cmbMED_FLAG" title="매체구분" style="WIDTH: 136px">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="MP01">신문</OPTION>
													<OPTION value="MP02">잡지</OPTION>
												</SELECT>
											</TD>
											<TD class="SEARCHLABEL" width="430">
											</TD>
										</TR>
										<tr>
											<TD class="SEARCHLABEL"width="70" >팀
											</TD>
											<TD id="tdCLIENTSUB" class="SEARCHDATA" colspan="6" >
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="17780">
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
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
