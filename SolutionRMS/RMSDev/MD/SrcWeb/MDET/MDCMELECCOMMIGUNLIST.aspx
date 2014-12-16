<%@ Page CodeBehind="MDCMELECCOMMIGUNLIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECCOMMIGUNLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>공중파 수수료 거래명세 조회</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/표준샘플/스프레드쉬트
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : SpreadSheet를 이용한 조회/입력/수정/삭제/인쇄 처리 표준 샘플
'파라  메터 : 
'특이  사항 : 표준샘플을 위해 만든 것임
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
'****************************************************************************************
-->
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet 정보 --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMGET 
Dim mobjMDCMELECCOMMILIST

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
Sub imgFind_onclick
	COMMI_POP
End Sub

Sub imgQuery_onclick
	if frmThis.txtTRANSYEARMON.value = "" or frmThis.txtTRANSNO.value = "" then
		gErrorMsgBox "년월과 거래명세 번호를 입력하시오",""
		exit Sub
	end if
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub


Sub imgPrint_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSER	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
	
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjMDCMELECCOMMILIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "MD"
		ReportName = "MDCMELECCOMMI_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",1)
		
		vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		strUSER = ""
		IF not gDoErrorRtn ("Get_ELETRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt + 1
			vntDataTemp = mobjMDCMELECCOMMILIST.ProcessRtn_temp(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt,strUSER)
		End IF
		Params = strUSER
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with
	gFlowWait meWAIT_OFF
End Sub	

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' 달력
'-----------------------------------------------------------------------------------------
Sub imgPRINTDAY_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgPRINTDAY,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'광고주
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'실제 데이터List 가져오기
Sub REAL_MED_CODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMREALMEDPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtREAL_MED_CODE.value = vntRet(0,0)		             ' Code값 저장
			.txtREAL_MED_NAME1.value = vntRet(1,0)             ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtREAL_MED_NAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
		
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetREALMEDNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtREAL_MED_CODE.value,.txtREAL_MED_NAME1.value)
			if not gDoErrorRtn ("GetREALMEDNO") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = vntData(0,0)
					.txtREAL_MED_NAME1.value = vntData(1,0)
				Else
					Call REAL_MED_CODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' 거래처번호팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCOMMI_onclick
	Call COMMI_POP()
End Sub

Sub txtTRANSNO_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, .txtTRANSNO.value, .txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value, "commi", "ELEC")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)  ' Code값 저장
					.txtTRANSNO.value = vntData(1,0)  ' 코드명 표시
					.txtREAL_MED_CODE.value = vntData(2,0)  ' 코드명 표시
					.txtREAL_MED_NAME1.value = vntData(3,0)  ' 코드명 표시
					'Call SelectRtn ()
				Else
					Call COMMI_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub COMMI_POP
	dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis

	If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
	strYEARMON = .txtTRANSYEARMON.value
	End If
	'msgbox strYEARMON
		vntInParams = array(strYEARMON, .txtTRANSNO.value, .txtREAL_MED_CODE.value, .txtREAL_MED_NAME1.value, "commi", "ELEC") '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code값 저장
			.txtTRANSNO.value = vntRet(1,0)  ' 코드명 표시
			.txtREAL_MED_CODE.value = vntRet(2,0)  ' 코드명 표시
			.txtREAL_MED_NAME1.value = vntRet(3,0)  ' 코드명 표시
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' 달력
'-----------------------------------------------------------------------------------------
Sub ImgPRINTDAY_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.ImgPRINTDAY,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'발행일
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row > 0 and Col > 1 then		
			'sprShtToFieldBinding Col,Row			
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'서버업무객체 생성	
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECCOMMILIST = gCreateRemoteObject("cMDET.ccMDETELECCOMMILIST")

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	'gInitPageSetting mobjSCGLCtl,"MD"
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
    
    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "TRANSYEARMON | TRANSNO | MEDCODE | MEDNAME | CLIENTCODE |CLIENTNAME|REAL_MED_CODE | REAL_MED_NAME| DEPT_CD| DEMANDDAY| PRINTDAY| AMT|SUSURATE| SUSU| VAT| MED_NAME"
		mobjSCGLSpr.SetHeader .sprSht,		"TRANSYEARMON|TRANSNO|MEDCODE|매체명|CLIENTCODE|광고주|REAL_MED_CODE|매체사|DEPT_CD|청구일자|발행일자|대행금액|수수료율(%)|수수료|부가세|매체종류"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "        0|	    0|      0|     0|		  0|	34|	           0|     0|      0|       0|       10|  12|       10|    12|    0|     13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY| PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT| SUSU| VAT|SUSURATE", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MEDNAME|CLIENTNAME|MED_NAME", -1, -1, 20
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MED_NAME",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO | CLIENTCODE|REAL_MED_NAME|MEDCODE|REAL_MED_CODE|DEPT_CD|DEMANDDAY|MEDNAME|VAT ", true
		.sprSht.style.visibility = "visible"
    End With

	'화면 초기값 설정
	InitPageData	
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'기본값 설정
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
				case 1 : .txtTRANSNO.value = vntInParam(i)
				case 2 : mblnUseOnly = vntInParam(i)		'현재 사용중인 것만
				case 3 : mstrUseDate = vntInParam(i)		'코드 사용 시점
				case 4 : mblnLikeCode = vntInParam(i)		'조회시 코드를 Like할지 여부
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECCOMMILIST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' 화면의 초기상태 데이터 설정
'****************************************************************************************
Sub InitPageData
	'모든 데이터 클리어
	gClearAllObject frmThis
	
	'초기 데이터 설정
	with frmThis
		'.txtTRANSYEARMON.value = "200712"
		'Sheet초기화
		.sprSht.MaxRows = 0
		.txtREAL_MED_NAME1.focus
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' 데이터 조회
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE, strTRANSNO
   	Dim i, strCols
	'On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON	= .txtTRANSYEARMON.value
		strTRANSNO	= .txtTRANSNO.value
		strREAL_MED_CODE	= .txtREAL_MED_CODE.value
		
		IF not SelectRtn_HDR (strYEARMON, strTRANSNO, strREAL_MED_CODE) Then Exit Sub

		'쉬트 조회
		Call SelectRtn_DTL (strYEARMON, strTRANSNO, strREAL_MED_CODE)
		
	END WITH
	
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strREAL_MED_CODE)
	dim vntData
	on error resume next

	'초기화
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strREAL_MED_CODE)
	
	IF not gDoErrorRtn ("Get_ELECTRANS_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "선택한 거래명세번호에 대하여" & meNO_DATA, ""
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			gWriteText "", "선택한 거래명세번호에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			SelectRtn_HDR = True
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByVal strYEARMON,ByVal strTRANSNO, ByVal strREAL_MED_CODE)
	dim vntData
	on error resume next

	'초기화
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strREAL_MED_CODE)
	
	IF not gDoErrorRtn ("Get_ELECTRANS_LIST") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_DTL = True
		gWriteText "", "선택한 거래명세번호건의 상세내역에 대하여" & mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strCUSTCODE, strCUSTNAME)
	frmThis.txtTRANSYEARMON.value = strCUSTCODE
	frmThis.txtREAL_MED_CODE.value = strCUSTNAME		
End Sub

'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'단가
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

Sub txtSUMAMTVAT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMTVAT,0,true)
	end with
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base" >
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 684px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
				width="684" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;공중파&nbsp;수수료&nbsp;거래명세 조회 및 발행</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 375px" vAlign="middle" align="right" height="28">
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
									<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="203" border="0">
										<TR>
											<TD><IMG id="imgFind" onmouseover="JavaScript:this.src='../../../images/imgFindOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgFind.gIF'" height="20" alt="자료를 검색합니다."
													src="../../../images/imgFind.gIF" width="54" border="0" name="imgFind"></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다."
													src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="굴림">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)">년 
													월</TD>
												<TD class="SEARCHDATA" style="WIDTH: 176px"><INPUT class="INPUT" id="txtTRANSYEARMON" title="거래명세년월" style="WIDTH: 72px; HEIGHT: 22px"
														accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON">&nbsp;<IMG id="ImgCOMMI" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
														name="ImgCOMMI"> <INPUT class="INPUT" id="txtTRANSNO" title="거래명세번호" style="WIDTH: 72px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTRANSNO"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_CODE, txtREAL_MED_NAME1)">청구지
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 263px"><INPUT class="INPUT_L" id="txtREAL_MED_CODE" title="코드조회" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtREAL_MED_CODE">&nbsp;<IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
														name="ImgREAL_MED_CODE">&nbsp;<INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="코드명" style="WIDTH: 167px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="21" name="txtREAL_MED_NAME1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"><FONT face="굴림">발행일자</FONT></TD>
												<TD class="SEARCHDATA">
													<INPUT class="INPUT" id="txtPRINTDAY" title="담당부서명" style="WIDTH: 82px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="100" size="9" name="txtPRINTDAY">&nbsp;<IMG id="ImgPRINTDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23"
														align="absMiddle" border="0" name="ImgPRINTDAY">
												</TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 791px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="right" border="0">
										<TR>
											<TD class="LABEL" width="90"><FONT face="굴림">매체사</FONT></TD>
											<TD class="DATA" width="173"></FONT><INPUT dataFld="REAL_MED_NAME" class="NOINPUT_L" id="txtREAL_MED_NAME" title="광고주명" style="WIDTH: 176px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="24" name="txtREAL_MED_NAME">
											</TD>
											<TD class="LABEL" width="90"><FONT face="굴림">담당부서</FONT></TD>
											<TD class="DATA" width="173"><FONT face="굴림"><INPUT dataFld="DEPT_NAME" class="NOINPUT_L" id="txtDEPT_NAME" title="브랜드코드" style="WIDTH: 173px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="23" name="txtDEPT_NAME"></FONT>
											</TD>
											<TD class="LABEL" width="90"><FONT face="굴림">청구일자</FONT></TD>
											<TD class="DATA" width="173"><FONT face="굴림"><INPUT dataFld="DEMANDDAY" class="NOINPUT" id="txtDEMANDDAY" title="브랜드명" style="WIDTH: 93px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="10" name="txtDEMANDDAY"></FONT></TD>
										</TR>
										<TR>
											<TD class="LABEL"><FONT face="굴림">수수료금액</FONT></TD>
											<TD class="DATA"><FONT face="굴림"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="광고금액" style="WIDTH: 135px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="16" name="txtAMT"></FONT>
											</TD>
											<TD class="LABEL"><FONT face="굴림">부가세</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="부가세" style="WIDTH: 135px; HEIGHT: 22px"
													accessKey="NUM" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="17" name="txtVAT"></TD>
											<TD class="LABEL"><FONT face="굴림">계</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT class="NOINPUT_R" id="txtSUMAMTVAT" title="계" style="WIDTH: 154px; HEIGHT: 22px"
													type="text" maxLength="100" size="20" name="txtSUMAMTVAT" readonly dataFld="SUMAMTVAT" dataSrc="#xmlBind"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 791px; HEIGHT: 3px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD align="center">
									<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LISTFRAME" style="HEIGHT: 101px" height="101">
												<OBJECT id="sprSht" style="WIDTH: 790px; HEIGHT: 346px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="20902">
													<PARAM NAME="_ExtentY" VALUE="9155">
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
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
