<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMATTERNEWPOP.aspx.vb" Inherits="MD.MDCMMATTERNEWPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>소재 조회</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/공통/공통코드 팝업
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPOP1.aspx
'기      능 : JOBNO 조회를 위한 팝업
'파라  메터 : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , 조회추가필드, 현재 사용중인 것만 조회할지 여부,
'			  코드 사용시점, 코드Like할지 여부
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
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
Dim mobjMDCMGET
Dim mobjMDCMCODETR 
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

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

sub imgConfirm_onclick ()
	if frmThis.sprSht.ActiveRow > 0 then
		sprSht_DblClick frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	else
		call Window_OnUnload()
	end if
end sub

Sub imgCancel_onclick
With frmThis
		
		window.returnvalue = "1|1"
		call Window_OnUnload()
		
End with
End Sub

Sub imgSave_onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

sub sprSht_DblClick (Col,Row)
	'선택된 로우 반환
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
		'msgbox Col & Row
		window.returnvalue = mobjSCGLSpr.GetClip (.sprSht,1,.sprSht.ActiveRow,.sprSht.MaxCols,1,1)
		call Window_OnUnload()
		end if
	End With
end sub
'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMDCMCODETR = gCreateRemoteObject("cMDCO.ccMDCOCODETR")
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	'gInitPageSetting mobjSCGLCtl,"SC"
	gInitComParams mobjSCGLCtl,"MC"
	with frmThis
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : .txtMATTERCODE.value = vntInParam(i)	
				case 1 : .txtMATTERNAME.value = vntInParam(i)
				case 2 : .txtClientcode.value  = vntInParam(i)			'조회추가필드
				case 3 : .txtCLIENTNAME.value = vntInParam(i)		'현재 사용중인 것만
				case 4 : .txtSUBSEQ.value  = vntInParam(i)		'코드 사용 시점
				case 5 : .txtSUBSEQNAME.value  = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
        With frmThis
            gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "MATTERCODE|MATTER|CUSTCODE|CUSTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SEQNO|SEQNAME|EXCLIENTCODE|EXCLIENTCODENAME|DEPTCD|DEPTNAME|ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,         "코드|소재명|광고주코드|광고주명|사업부코드|사업부명|브랜드코드|브랜드명|대대행사코드|제작대행사명|부서코드|부서명|입력건"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "5   | 15   |0         |14      |0         |14      |0         |14      | 0          |14          |0       |12    |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MATTERCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MATTER|EXCLIENTCODENAME|CUSTNAME|SEQNAME|CLIENTSUBNAME|DEPTNAME"
		mobjSCGLSpr.ColHidden .sprSht, "CUSTCODE|SEQNO|EXCLIENTCODE|CLIENTSUBCODE|DEPTCD|ATTR01",true
		mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MATTERCODE|EXCLIENTCODENAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MATTER",-1,-1,0,2,false

        End With
	end with	
	'자료조회	
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDCMCODETR = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
	Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCMCODETR.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtMATTERCODE.value,.txtMATTERNAME.value,.txtCLIENTCODE.value,.txtCLIENTNAME.value,.txtSUBSEQ.value,.txtSUBSEQNAME.value,.txtEXCLIENTCODE1.value,.txtEXCLIENTNAME1.value )

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			' mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			if mlngRowCnt < 1 Then
			.sprSht.MaxRows = 0 
				
			Else
			'sprShtToFieldBinding 1,1 '임시주석
			
			
				
				for intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",intCnt) = "1" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					Else
						If intCnt Mod 2 = 0 Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False	
						Else
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					End If
				Next
			End If
			'조회해온 추가 필드를 Hidden
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
			
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
end sub
'-----------------------------------------------------------------------------------------
' 스프레드 쉬트 클릭시 
'-----------------------------------------------------------------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	Dim intCnt, i
	
	With frmThis
		if Row > 0 and Col > 1 then		
				sprShtToFieldBinding Col,Row	
		end if 
	End With
End Sub  
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			gSetChangeFlag .txtCLIENTCODE                 ' gSetChangeFlag objectID	 Flag 변경 알림
		end if
	End with
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 브랜드코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgSUBSEQCODE_onclick
	'with frmThis
	'	If .txtCLIENTCODE.value = "" Then
	'		gErrorMsgBox "브랜드검색은 광고주코드를 먼저 조회하십시오.","검색안내!"
	'		Exit Sub
	'	End If 
	'End with
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		if isArray(vntRet) then
			if .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtSUBSEQ.value = trim(vntRet(1,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(2,0))	' 브랜드명 표시
			.txtCLIENTCODE.value = trim(vntRet(3,0))		' 광고주 표시
			.txtCLIENTNAME.value = trim(vntRet(4,0))	' 광고주명 표시
			'.txtPUB_DATE.focus()					' 포커스 이동
			gSetChangeFlag .txtSUBSEQ		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
		
			if window.event.keyCode = meEnter then
				
			
					Dim vntData
   					Dim i, strCols
					'On error resume next
					with frmThis
						'Long Type의 ByRef 변수의 초기화
						mlngRowCnt=clng(0)
						mlngColCnt=clng(0)
						vntData = mobjMDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
						if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
							If mlngRowCnt = 1 Then
								.txtSUBSEQ.value = trim(vntData(1,0))
								.txtSUBSEQNAME.value = trim(vntData(2,0))
								.txtCLIENTCODE.value = trim(vntData(3,0))		' 광고주 표시
								.txtCLIENTNAME.value = trim(vntData(4,0))	' 광고주
							Else
								Call SUBSEQCODE_POP()
							End If
   						end if
   					end with
					window.event.returnValue = false
					window.event.cancelBubble = true
				
			end if
		
End Sub
'-----------------------------------------------------------------------------------------
' 대행사 코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgEXCLIENTCODE1_onclick
	Call EXCLIENTCODE_POP1()
End Sub

'실제 데이터List 가져오기
Sub EXCLIENTCODE_POP1
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE1.value), trim(.txtEXCLIENTNAME1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtEXCLIENTCODE1.value = vntRet(0,0) and .txtEXCLIENTNAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtEXCLIENTCODE1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtEXCLIENTNAME1.value = trim(vntRet(1,0))  ' 코드명 표시
			'.txtMEDNAME.focus()					' 포커스 이동
			'gSetChangeFlag .txtEXCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtEXCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE1.value),trim(.txtEXCLIENTNAME1.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE1.value = trim(vntData(0,0))
					.txtEXCLIENTNAME1.value = trim(vntData(1,0))
					'.txtMEDNAME.focus()
					'GetBrandAndDept'광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
				Else
					Call EXCLIENTCODE_POP1()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 쉬트 바인딩 
'-----------------------------------------------------------------------------------------
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	dim vntData
	dim strCODE
	with frmThis
		strCODE =	mobjSCGLSpr.GetTextBinding(.sprSht,"MATTERCODE",Row)
		
		if strCODE ="" Then EXIT Function
		
		vntData = mobjMDCMCODETR.GetMATTER_spr(gstrConfigXml,Row,Col,strCODE)
	
		IF not gDoErrorRtn ("GetMATTER") then
			'조회한 데이터를 바인딩
			
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			.txtCLIENTCODE.focus()
			.sprSht.focus()
			'.txtMATTERNAME.focus()
		End IF
	
	END WITH
End Function

Sub processRtn
	
	Dim strMATTERNAME
	Dim strCUSTCODE
	Dim strSEQNO
	Dim strEXCLIENTCODE
	with frmThis
	strMATTERNAME = .txtMATTERNAME.value 
	strCUSTCODE = .txtCLIENTCODE.value 
	strSEQNO = .txtSUBSEQ.value 
	strEXCLIENTCODE = .txtEXCLIENTCODE1.value 
	
	
	
	End with
End Sub

Sub ProcessRtn ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strMATTERNAME

	with frmThis
	'On error resume next
  		'데이터 Validation
		if DataValidation =false then exit sub
		strMATTERNAME = .txtMATTERNAME.value 
		If strMATTERNAME = "" Then
			gErrorMsgbox "소재명은 필수 입니다.","저장안내!"
		End If
		strMasterData = gXMLGetBindingData (xmlBind)
		
		intRtn = mobjMDCMCODETR.ProcessRtn_MATTERINSERT(gstrConfigXml,strMasterData,strMATTERNAME)
		
		if not gDoErrorRtn ("ProcessRtn_MATTERINSERT") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " 자료가 신규저장" & mePROC_DONE,"저장안내" 
			SelectRtn
  		end if
 	end with
End Sub

Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols
	On error resume next
	with frmThis
		'Master 입력 데이터 Validation : 필수 입력항목 검사
   		IF not gDataValidation(frmThis) then exit Function
		
   	End with
	DataValidation = true
End Function
-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<XML id="xmlBind"></XML>
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="600" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/PopupIcon.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">소재 등록&nbsp;
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
									<TABLE id="tblButton" style="WIDTH: 168px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="168" border="0">
										<TR>
											<TD><FONT face="굴림"></FONT></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><FONT face="굴림"></FONT></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD style="WIDTH: 1px"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="자료를 조회합니다." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD width="3"><FONT face="굴림"></FONT></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCLOSEOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCLOSE.gif'"
													height="20" alt="화면을 닫습니다." src="../../../images/imgCLOSE.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="굴림"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
										<TABLE class="KEY" id="tblKey" style="WIDTH: 762px; HEIGHT: 25px" cellSpacing="0" cellPadding="0"
											width="792" align="right" border="0">
											<TBODY>
												<TR>
													<TD class="KEY" style="WIDTH: 105px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,txtCLIENTNAME)">
														광고주</TD>
													<TD class="SEARCHDATA" style="WIDTH: 88px"><INPUT class="NOINPUT" id="txtCLIENTCODE" style="WIDTH: 90px; HEIGHT: 22px" accessKey=",M"
															readOnly type="text" size="9" name="txtCLIENTCODE" dataFld="CUSTCODE" dataSrc="#xmlBind" title="광고주코드">&nbsp;</TD>
													<TD class="SEARCHDATA"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
															width="23" align="absMiddle" border="0" name="ImgCLIENTCODE">&nbsp;<INPUT class="INPUT_L" id="txtCLIENTNAME" style="WIDTH: 440px; HEIGHT: 22px" tabIndex="1"
															type="text" size="68" name="txtCLIENTNAME" dataFld="CUSTNAME" dataSrc="#xmlBind" title="광고주명"></TD>
												</TR>
												<TR>
													<TD class="KEY" style="WIDTH: 105px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQ,txtSUBSEQNAME)">
														브랜드</TD>
													<TD class="SEARCHDATA" style="WIDTH: 88px"><INPUT class="NOINPUT" id="txtSUBSEQ" style="WIDTH: 90px; HEIGHT: 24px" readOnly type="text"
															size="9" name="txtSUBSEQ" dataFld="SEQNO" accessKey=",M" dataSrc="#xmlBind" title="브랜드코드">&nbsp;</TD>
													<TD class="SEARCHDATA"><IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
															align="absMiddle" border="0" name="ImgSUBSEQCODE">&nbsp;<INPUT class="INPUT_L" id="txtSUBSEQNAME" style="WIDTH: 440px; HEIGHT: 22px" tabIndex="1"
															type="text" size="68" name="txtSUBSEQNAME" dataFld="SEQNAME" dataSrc="#xmlBind" title="브랜드 명"></TD>
												</TR>
												<TR>
													<TD class="KEY" style="WIDTH: 105px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTCODE1,txtEXCLIENTNAME1)">
														제작대행사</TD>
													<TD class="SEARCHDATA" style="WIDTH: 88px"><INPUT class="NOINPUT" id="txtEXCLIENTCODE1" style="WIDTH: 90px; HEIGHT: 22px" readOnly
															type="text" size="9" name="txtEXCLIENTCODE1" dataFld="EXCLIENTCODE" accessKey=",M" dataSrc="#xmlBind" title="제작대행사 코드">&nbsp;</TD>
													<TD class="SEARCHDATA"><IMG id="ImgEXCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
															align="absMiddle" border="0" name="ImgEXCLIENTCODE1">&nbsp;<INPUT class="INPUT_L" id="txtEXCLIENTNAME1" style="WIDTH: 440px; HEIGHT: 22px" tabIndex="1"
															type="text" size="68" name="txtEXCLIENTNAME1" dataFld="EXCLIENTCODENAME" dataSrc="#xmlBind" title="제작대행사명"></TD>
												</TR>
												<TR>
													<TD class="KEY" style="WIDTH: 105px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERCODE,txtMATTERNAME)">
														소재명</TD>
													<TD class="SEARCHDATA" style="WIDTH: 88px"><INPUT class="NOINPUT" id="txtMATTERCODE" style="WIDTH: 90px; HEIGHT: 22px" readOnly type="text"
															size="9" name="txtMATTERCODE" title="소재코드">&nbsp;</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtMATTERNAME" style="WIDTH: 467px; HEIGHT: 22px" tabIndex="1"
															type="text" size="72" name="txtMATTERNAME" accessKey=",M" title="소재명"></TD>
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
										<OBJECT id="sprSht" style="WIDTH: 762px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="20161">
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
				</FORM></TR>
		</TABLE>
	</body>
</HTML>
