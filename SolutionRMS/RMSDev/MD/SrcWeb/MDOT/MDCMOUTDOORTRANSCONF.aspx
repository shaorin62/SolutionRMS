<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORTRANSCONF.aspx.vb" Inherits="MD.MDCMOUTDOORTRANSCONF" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>거래명세서 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : 인쇄매체
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMTRANSCONF.aspx
'기      능 : 작성된 거래명세서 의 Confirm 을 한다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/29 By Kim Tae Ho
'			 2) 
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
Dim mobjMDTRANSCONF
Dim mobjMDCMGET
Dim mstrCheck
'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick()
	EndPage
End Sub

Sub imgQuery_Onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub


'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------------------------------------------------------------------
Sub InitPage()

	'서버업무객체 생성	
	set mobjMDTRANSCONF	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOORCOMMI") '조회
	set mobjMDCMGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")	  '코드

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "125px"
	'pnlTab1.style.height ="300px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
    gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***첫번째 Sheet 디자인
		'**************************************************
		
		'Sheet 칼라 지정
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout 디자인
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0,6
		'mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		'Binding Field 설정
	    mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONFIRMGBN|CONFIRMFLAG|TRANSYEARMON|TRANSNO|CLIENTCODE|CLIENTNAME|DEMANDDAY|PRINTDAY|AMT|VAT|SUMAMT|MED_FLAG_NAME|MEMO"
		'Header 디자인
		mobjSCGLSpr.SetHeader .sprSht,        "선택|승인여부|계산서여부|거래년월|번호|광고주코드|광고주|청구일|발행일|공급가액|부가세|합계금액|매체구분|비고",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|       7|         9|       8|   4|         0|    14|     8|     8|      10|     9|      11|       8|  20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY", , , ,3
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CONFIRMGBN|TRANSYEARMON|TRANSNO|CLIENTCODE|CLIENTNAME|AMT|VAT|SUMAMT|MED_FLAG_NAME|MEMO|CONFIRMFLAG"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMFLAG",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", true
	End with

	pnlTab1.style.visibility = "visible" 
	
	'화면 초기값 설정
	InitPageData	
End Sub

Sub SelectRtn()
	Dim vntData
	Dim i, strCols
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDEMANDDAYFROM
	Dim strDEMANDDAYTO
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strGUBUN
	Dim intCnt
	with frmThis
	'ON ERROR RESUME NEXT
		.sprSht.MaxRows = 0
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = .txtTRANSYEARMON.value
		strTRANSNO = .txtTRANSNO.value
		strDEMANDDAYFROM = Replace(.txtFROM.value,"-","")
		strDEMANDDAYTO = Replace(.txtTO.value,"-","")
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value
		strGUBUN = .cmbGUBUN.value
		vntData = mobjMDTRANSCONF.SelectRtn_TransList(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON,strTRANSNO,strDEMANDDAYFROM,strDEMANDDAYTO,strCLIENTCODE,strGUBUN)

		if not gDoErrorRtn ("SelectRtn_TransList") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",intCnt) = "Y" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				Else
					'체크
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
				End If			
			Next
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE			
   		end if
	End With
End Sub

Sub EndPage()
	set mobjMDTRANSCONF = Nothing
	gEndPage
End Sub
'-----------------------------------------------------------------------------------------
' 화면 처리 SCRIPT
'-----------------------------------------------------------------------------------------
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[입력용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			'.txtMEDNAME.focus()					' 포커스 이동
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag 변경 알림
     	end if
	End with
	'GetBrandAndDept '광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					'.txtMEDNAME.focus()
					'GetBrandAndDept'광고주 시퀀스와 시퀀스의 담당부서를 가져온다.
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		elseif Row = 0 and Col =1 then
		else
			strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",Row)
			strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",Row)
			
			vntInParams = array(strTRANSYEARMON, strTRANSNO) '<< 받아오는경우
			vntRet = gShowModalWindow("MDCMOUTDOORTRANSGUNLIST.aspx",vntInParams , 813,585)
			if isArray(vntRet) then
     		end if
		end if
	end with
end sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = 1  then 'AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "N"
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
				
			next
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",intCnt) = "Y" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				'Else
					'체크
				'	mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
				End If			
			Next
		end if
	end with
End Sub  
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub



'-----------------------------------------------------------------------------------------
' Field 체크
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgTO_onclick
	'CalEndar를 화면에 표시
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub
Sub txtFROM_onchange
	gSetChange
End Sub
Sub txtTO_onchange
	gSetChange
End Sub

'조회날자 이벤트
Sub txtFROM_onfocus
End Sub
Sub txtFROM_onblur
	with frmThis
	gFormatDate .txtFROM,True
	End with
End Sub
Sub txtTO_onfocus
End Sub
Sub txtTO_onblur
	with frmThis
	gFormatDate .txtTO,True
	End with
End Sub

'------------------------------------------
' 삭제로직
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	with frmThis
	strDESCRIPTION = ""
		
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"TAXFLAG",intCnt2) = "Y" AND mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt2) & " 에 대하여" &vbcrlf & "계산서가 존재하는 내역은 삭제가 되지 않습니다.","삭제안내!"
				Exit Sub
			End If
		Next
			
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
			
				intRtn = mobjMDTRANSCONF.DeleteRtn_TransList(gstrConfigXml,strTRANSYEARMON, strTRANSNO,strDESCRIPTION)
				IF not gDoErrorRtn ("DeleteRtn_TransList") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"삭제안내!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intCnt & "건이 삭제" & mePROC_DONE
   		End IF
   		
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub
'-----------------------------------------------------------------------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtTRANSYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		DateClean
		
		.sprSht.MaxRows = 0			
		.txtCLIENTNAME.focus()
	end With
End Sub

Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub

Sub imgSetting_onclick
		Call ProcessRtn_ConfirmOK()
End Sub

Sub ImgConfirmCancel_onclick
		ProcessRtn_ConfirmCancel
End Sub


Sub txtTRANSYEARMON_onblur
	With frmThis
		If .txtTRANSYEARMON.value <> "" AND Len(.txtTRANSYEARMON.value) = 6 Then DateClean
	End With
End Sub
'------------------------------------------
' 승인 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim lngCONFIRMSUM
	Dim strFLAG 
	
	strFLAG = "CONFIRM"
	
	with frmThis
   		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		lngCONFIRMSUM = 0
   		
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
				IF trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt)) = "미승인" then
					lngCONFIRMSUM = lngCONFIRMSUM + 1
				end if
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "승인처리할 데이터를 선택하십시오.","저장안내!"
			Exit Sub
		End If
		
		If lngCONFIRMSUM = 0 Then
			gErrorMsgBox "선택하신 데이터에 미승인 건이 없습니다.","저장안내!"
			Exit Sub
		End If
		
		On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG")
		
		intRtn = mobjMDTRANSCONF.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "선택하신 " & lngCHKSUM & " 건의 자료중 미승인 상태인 " & lngCONFIRMSUM & "의 자료가 승인" & mePROC_DONE,"승인안내!"
			'gWriteText "", intRtn & "건의 자료가 저장" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' 승인취소 저장로직
'------------------------------------------
Sub ProcessRtn_ConfirmCancel
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim lngCONFIRMSUM
	Dim strFLAG
	strFLAG = "CANCEL"
	with frmThis
   		'데이터 Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "조회된 건이 없으므로 저장이 불가능 합니다.","저장안내!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		lngCONFIRMSUM = 0
   		
   		For intCnt = 1 to .sprSht.MaxRows
			 IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
				IF trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt)) = "승인" then
					lngCONFIRMSUM = lngCONFIRMSUM + 1
				end if
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "승인취소할 데이터를 선택하십시오.","저장안내!"
			Exit Sub
		End If
		
		If lngCONFIRMSUM = 0 Then
			gErrorMsgBox "선택하신 데이터에 승인 건이 없습니다.","저장안내!"
			Exit Sub
		End If
		
		On error resume next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG")
		
		intRtn = mobjMDTRANSCONF.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
	
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "선택하신 " & lngCHKSUM & " 건의 자료중 승인 상태인 " & lngCONFIRMSUM & "의 자료가 승인취소" & mePROC_DONE,"승인안내!"
			SelectRtn
   		end if
   		
   	end with
End Sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD style="HEIGHT: 54px">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIf" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">
													&nbsp;거래명세서 관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 20px" vAlign="top" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON,txtTRANSNO)"
													width="90">거래명세서
												</TD>
												<TD class="SEARCHDATA" width="200"><INPUT dataFld="TRANSYEARMON" class="INPUT" id="txtTRANSYEARMON" title="세금계산서년월" style="WIDTH: 56px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtTRANSYEARMON">&nbsp;-&nbsp;<INPUT dataFld="TRANSNO" class="INPUT" id="txtTRANSNO" title="세금계산서번호" style="WIDTH: 48px; HEIGHT: 22px"
														accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="4" size="2" name="txtTRANSNO">&nbsp;<SELECT dataFld="MED_FLAG" id="cmbGUBUN" title="매체구분" style="WIDTH: 72px" name="cmbGUBUN">
														<OPTION value="X" selected>전체</OPTION>
														<OPTION value="0">미승인</OPTION>
														<OPTION value="1">승인</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM,txtTO)" align="right"
													width="80">청구일자
												</TD>
												<TD class="SEARCHDATA" width="220"><INPUT class="INPUT" id="txtFROM" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
														type="text" maxLength="8" size="6" name="txtFROM"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="청구일자" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
														type="text" maxLength="8" size="6" name="txtTO"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgTo">
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,txtCLIENTNAME)"
													width="80">광고주</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUTB_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 168px; HEIGHT: 21px"
														type="text" maxLength="255" size="22" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 72px" type="text" maxLength="6"
														name="txtCLIENTCODE">
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="WIDTH: 56px; CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" width="56" align="absMiddle" border="0" name="imgQuery">
												</td>
											</TR>
										</TABLE>
										<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="굴림"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;거래명세서 조회 및 승인</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
																	height="20" alt="자료를승인처리합니다." src="../../../images/imgAgree.gIF" border="0" name="imgSetting"></TD>
															<td><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCancelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCancel.gIF'"
																	height="20" alt="승인처리를 취소합니다." src="../../../images/imgAgreeCancel.gif" border="0"
																	name="ImgConfirmCancel"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1038px; HEIGHT: 700px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 700px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 700px"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27437">
												<PARAM NAME="_ExtentY" VALUE="18521">
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
								<!--List End-->
								<!--BodySplit Start-->
								<!--Brench End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
