<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMSUMMARYLIST.aspx.vb" Inherits="PD.PDCMSUMMARYLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_ESTDTL.aspx
'기      능 : JOBMST의 두번째 탭 - 가/본 견적서를 저장 및 수정 한다. 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/18 By KimTH
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
'=============================
' 이벤트 프로시져 
'=============================
option explicit
Const meTAB = 9
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMSUMMARY, mobjPDCMGET

'선택체크용
Dim mstrCheck
Dim mALLCHECK
' 벨리데이션에 걸렸을시에 체크 mstrValiCHECK   pub_processrtn에서 사용
Dim mstrValiCHECK
'본견적을 가져왔을때 true   아니고 exe_hdr 에 있다면  초기값인 false
Dim strACTUALFLAG
'헤더의 변경내용 여부    기본 false 변경 true
Dim mstrHEADERFLAG 
Dim mstrPROCESS

Dim mstrJOBNO 

mALLCHECK = TRUE
mstrCheck=TRUE
mstrValiCHECK = TRUE
strACTUALFLAG = FALSE
mstrPROCESS = False
mstrHEADERFLAG = false
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

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	gErrorMsgBox "양식을 제출하여 주십시오.","인쇄안내"
	gFlowWait meWAIT_OFF
End Sub	

'출력이 확정된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMCATVTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------

Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub

Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub

Sub txtDEMANDAMT_onfocus
	with frmThis
		.txtDEMANDAMT.value = Replace(.txtDEMANDAMT.value,",","")
	end with
End Sub
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub

Sub txtESTAMT_onfocus
	with frmThis
		.txtESTAMT.value = Replace(.txtESTAMT.value,",","")
	end with
End Sub
Sub txtESTAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTAMT,0,true)
	end with
End Sub

Sub txtPAYMENT_onfocus
	with frmThis
		.txtPAYMENT.value = Replace(.txtPAYMENT.value,",","")
	end with
End Sub

Sub txtPAYMENT_onblur
	with frmThis
		call gFormatNumber(.txtPAYMENT,0,true)
	end with
End Sub

Sub txtINCOM_onfocus
	with frmThis
		.txtINCOM.value = Replace(.txtINCOM.value,",","")
	end with
End Sub
Sub txtINCOM_onblur
	with frmThis
		call gFormatNumber(.txtINCOM,0,true)
	end with
End Sub

Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub

Sub txtACCAMT_onfocus
	with frmThis
		.txtACCAMT.value = Replace(.txtACCAMT.value,",","")
	end with
End Sub
Sub txtACCAMT_onblur
	with frmThis
		call gFormatNumber(.txtACCAMT,0,true)
	end with
End Sub


'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT"))   Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	Dim strColFlag
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
		strCOLUMN = ""
		strColFlag = 0
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
						End If
					Next
					
					.txtSELECTAMT.value = strSUM
				End If
				
			else
				.txtSELECTAMT.value = 0
			End If
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
	
	'서버업무객체 생성	
	set mobjPDCMSUMMARY	= gCreateRemoteObject("cPDCO.ccPDCOSUMMARY")
	set mobjPDCMGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"

	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis	
	
		'******************************************************************
		'거래명세서 생성 그리드
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_CLIENT
		mobjSCGLSpr.SpreadLayout .sprSht_CLIENT, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_CLIENT,   "CLIENTNAME|TIMNAME|DIVAMT|OUTDIVAMT"
		mobjSCGLSpr.SetHeader .sprSht_CLIENT,		   "광고주|팀|분할금액|외주비분담금액"
		mobjSCGLSpr.SetColWidth .sprSht_CLIENT, "-1", "    20| 20|     15|             14"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_CLIENT, "DIVAMT|OUTDIVAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_CLIENT, true, "CLIENTNAME|TIMNAME|DIVAMT|OUTDIVAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht_CLIENT, "CLIENTNAME|TIMNAME",-1,-1,0,2,false
	
	    .sprSht_CLIENT.style.visibility  = "visible"
		.sprSht_CLIENT.MaxRows = 0

		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,   "OUTSNAME|ITEMCLASS|ITEMNAME|OUTAMT"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		   "외주처|외주부문|외주항목|지급액"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1", "    20| 18|     18|             13"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "OUTAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, true, "OUTSNAME|ITEMCLASS|ITEMNAME|OUTAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "OUTSNAME|ITEMCLASS|ITEMNAME",-1,-1,0,2,false
	
	    .sprSht_OUT.style.visibility  = "visible"
		.sprSht_OUT.MaxRows = 0
		
		InitPageData
		'부모창의 데이터 가져오기  (전역변수에담기)
		'mstrJOBNO =  parent.document.forms("frmThis").txtJOBNO.value 
		
		mstrJOBNO = parent.document.forms("frmThis").txtJOBNO.value 
		
		SelectRtn
	End With
End Sub

Sub EndPage()
	'set mobjPDCMSUMMARY = Nothing
	'set mobjPDCMGET = Nothing
	gEndPage
End Sub

'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub



'****************************************************************************************
' 데이터 조회
'****************************************************************************************
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	with frmThis
		mstrJOBNO = parent.document.forms("frmThis").txtJOBNO.value 
		if mstrJOBNO = "" Or Len(mstrJOBNO) <> 7 Then
			gErrorMsgBox "제작번호를확인하십시오.","조회안내!"
			Exit Sub
		End if
	
		'JOBNO로 정산데이타를 가져온다. 업으면FALSE
		IF SelectRtn_Head Then 
			CALL SelectRtn_Detail ()
		else
			'본견적으로 가져올경우는 없다고 생각되어 지움...
			'call SelectRtn_Actual_Head ()
			'call SelectRtn_Actual_Detail ()	
		END IF
		
		
		if .txtJOBGUBN.value ="CF" then
			.imgDetailList.style.visibility = "visible"
		else	
			.imgDetailList.style.visibility = "hidden"
		end if
		
		txtSUSUAMT_onblur
		txtCOMMITION_onblur
		txtDEMANDAMT_onblur
		txtPAYMENT_onblur
		txtINCOM_onblur
		txtNONCOMMITION_onblur
		txtACCAMT_onblur
		txtESTAMT_onblur
		mstrHEADERFLAG = false
	End with
End Sub

Sub imgDetailListOn

End Sub

Function SelectRtn_Head
	Dim vntData
	SelectRtn_Head = false
	'on error resume next
	'초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMSUMMARY.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt <=0 then
			SelectRtn_Head = FALSE
			strACTUALFLAG = TRUE
		else
			'조회한 데이터를 바인딩
			SelectRtn_Head = True
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
		End IF
	End IF
End Function


'예산 테이블 조회
Function SelectRtn_Detail
	dim vntData_CLIENT, vntData_OUT
	Dim strRows
	Dim intCnt
	
	
	SelectRtn_Detail = false
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData_CLIENT = mobjPDCMSUMMARY.SelectRtn_DTL_CLIENT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO)

	IF not gDoErrorRtn ("SelectRtn_DTL_CLIENT") then
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_CLIENT,vntData_CLIENT,1,1,mlngColCnt,mlngRowCnt,true)

		SelectRtn_Detail = True
		
		with frmThis
			mobjSCGLSpr.SetFlag  frmThis.sprSht_CLIENT,meCLS_FLAG
			gWriteText lblStatus_CLIENT, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
		End with
	End IF
	
	
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData_OUT = mobjPDCMSUMMARY.SelectRtn_DTL_OUT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO)

	IF not gDoErrorRtn ("SelectRtn_DTL_OUT") then

		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_OUT,vntData_OUT,1,1,mlngColCnt,mlngRowCnt,true)

		SelectRtn_Detail = True
		
		with frmThis
			mobjSCGLSpr.SetFlag  frmThis.sprSht_OUT,meCLS_FLAG
			gWriteText lblStatus_OUT, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
		End with
	End IF
End Function


Function SelectRtn_Actual_Head
	Dim vntData
	'on error resume next
	
	'초기화
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'
	vntData	= mobjPDCMSUMMARY.SelectRtn_Actual_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO)
	
	IF not gDoErrorRtn ("SelectRtn_Actual_HDR") then
		IF mlngRowCnt > 0 then
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			
			'바인딩한 후에는 본견적 jobno와 preestno 로 다시 전역변수에 넣어준다.
			'mstrJOBNO	= frmThis.txtJOBNO.value
			'strPREESTNO = frmThis.txtPREESTNO.value
	
		End IF
	End IF
End Function

'예산 테이블 조회
Function SelectRtn_Actual_Detail
	Dim vntData
	Dim intCnt
	Dim strRows
	Dim intCnt2
	'on error resume next	
	'초기화
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'MSGBOX mstrJOBNO
	'eXIT Function
	vntData = mobjPDCMSUMMARY.SelectRtn_Actual_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO)

	IF not gDoErrorRtn ("SelectRtn_Actual_DTL") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정

		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt = 1 To .sprSht.MaxRows
					If .txtENDDAY.value <> "" or mobjSCGLSpr.GetTextBinding(.sprSht, "CONTRACTNO",intCnt) <> "" or mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt) <> "" then '특정값에 해당 없으면 기본색을 세팅
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,-1,-1,true
					ELSE
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False '이게 흰색
						mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,intCnt,-1,-1,true
						mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SORTSEQ|ADJDAY|CONTRACTNO"
					END IF
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",intCnt,"코드선택"
					sprSht_Change 18,intCnt
					
					End If
					
				Next
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="27" background="../../../images/back_p.gIF"
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
											<td class="TITLE">개요&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton2" style=" HEIGHT: 20px" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgDetailList" onmouseover="JavaScript:this.src='../../../images/imgDetailListOn.gif'"
													style="VISIBILITY: hidden; CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDetailList.gif'"
													height="20" alt="JOB의 추가입력사항을 조회합니다." src="../../../images/imgDetailList.gIF" border="0"
													name="imgDetailList"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top" width="100%">
						<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1"
							cellPadding="0" align="right" border="0">
							<TR>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">프로젝트명</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="PROJECTNM" class="NOINPUTB_R" id="txtPROJECTNM" title="프로젝트명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPROJECTNM"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">광고주</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_R" id="txtCLIENTNAME" title="광고주" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTNAME"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">견적금액</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="ESTAMT" class="NOINPUTB_R" id="txtESTAMT" title="견적금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtESTAMT"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">Noncommition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="수수료미지불금액"
										style="WIDTH: 152px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtNONCOMMITION"></TD>
							</TR>
							<TR>
								<TD class="SEARCHLABEL">JOB명</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="JOBNAME" class="NOINPUTB_R" id="txtJOBNAME" title="JOB명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBNAME"></TD>
								<TD class="SEARCHLABEL">팀</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="TIMNAME" class="NOINPUTB_R" id="txtTIMNAME" title="팀명" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtTIMNAME"></TD>
								<TD class="SEARCHLABEL">청구금액</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDAMT" class="NOINPUTB_R" id="txtDEMANDAMT" title="청구금액 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDAMT"></TD>
								<TD class="SEARCHLABEL">Commition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="수수료지불금액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCOMMITION"></TD>
							</TR>
							<tr>
								<TD class="SEARCHLABEL">매체부문</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="JOBGUBN" class="NOINPUTB_R" id="txtJOBGUBN" title="매체부문" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBGUBN"></TD>
								<TD class="SEARCHLABEL">브랜드</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_R" id="txtSUBSEQNAME" title="브랜드" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUBSEQNAME"></TD>
								<TD class="SEARCHLABEL">외주비</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="PAYMENT" class="NOINPUTB_R" id="txtPAYMENT" title="외주비 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPAYMENT"></TD>
								<TD class="SEARCHLABEL">수수료</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="수수료합계금액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUSUAMT"></TD>
							</tr>
							<tr>
								<TD class="SEARCHLABEL">매체분류</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUTB_R" id="txtCREPART" title="매체분류" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="6" name="txtCREPART"></TD>
								<TD class="SEARCHLABEL">청구일</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDDAY" class="NOINPUTB_R" id="txtDEMANDDAY" title="청구일" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDDAY"></TD>
								<TD class="SEARCHLABEL">진행비</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ACCAMT" class="NOINPUTB_R" id="txtACCAMT" title="비용 합계" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtACCAMT"></TD>
								<TD class="SEARCHLABEL">수수료율</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSURATE" class="NOINPUTB_R" id="txtSUSURATE" title="수수료율" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtSUSURATE">&nbsp;(%)</TD>
							</tr>
							<TR>
								<TD class="SEARCHLABEL">상태</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ENDFLAG" class="NOINPUTB_R" id="cmbENDFLAG" title="상태" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="cmbENDFLAG"></TD>
								<TD class="SEARCHLABEL">결산일</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CLOSEDAY" class="NOINPUTB_R" id="txtClOSEDAY" title="결산일" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtClOSEDAY"></TD>
								<TD class="SEARCHLABEL">내수액</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="INCOM" class="NOINPUTB_R" id="txtINCOM" title="내수액" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtINCOM"></TD>
								<TD class="SEARCHLABEL">내수율</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="RATE" class="NOINPUTB_R" id="txtRATE" title="내수율" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtRATE">&nbsp;(%)</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" id="spacebar" style="WIDTH: 100%; HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="54" background="../../../images/back_p.gIF"
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
											<td class="TITLE">1) 청구지&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="54" background="../../../images/back_p.gIF"
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
											<td class="TITLE">2) 외주비&nbsp;</td>
										</tr>
									</table>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 218px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
							<TR>
								<td style="WIDTH: 50%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_CLIENT" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15928">
											<PARAM NAME="_ExtentY" VALUE="10081">
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
								</td>
								<td style="WIDTH: 50%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_OUT" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15928">
											<PARAM NAME="_ExtentY" VALUE="10081">
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
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus_CLIENT" style="WIDTH: 1040px"></TD>
								<TD class="BOTTOMSPLIT" id="lblStatus_OUT" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</tr>
			</TABLE>
		</FORM>
	</body>
</HTML>
