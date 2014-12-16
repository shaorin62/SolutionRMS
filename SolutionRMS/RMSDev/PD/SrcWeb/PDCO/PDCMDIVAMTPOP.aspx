<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDIVAMTPOP.aspx.vb" Inherits="PD.PDCMDIVAMTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>분할금액 관리</title> 
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
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjPDCDGet
Dim mobjPDCMDIVAMT
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Const meTab = 9
'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgClose_onclick()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtJOBNO_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub



Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------
' Spread Sheet Event
'-----------------------------	
'onblour 이벤트
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub
Sub txtDIVAMT_onblur
	with frmThis
		call gFormatNumber(.txtDIVAMT,0,true)
	end with
End Sub
Sub sprSht_change(ByVal Col,ByVal Row)
	
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName,strCodeName2
   	Dim strQTY,strPRICE,strAMT 
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		IF  Col = 11 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)
			vntData = mobjPDCDGet.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)
			
			if not gDoErrorRtn ("GetCUSTNO_HIGHCUSTCODE") then
			
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(5,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(6,0)			
					'mobjSCGLSpr.CellChanged .sprSht, frmThis.sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					.txtDIVAMT.focus
					.sprSht.focus 
					mobjSCGLSpr.ActiveCell .sprSht, Col+4,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht, 10, Row
					.txtDIVAMT.focus
					.sprSht.focus 
				End If
   			end if
   		ElseIF  Col = 14 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)
			
			vntData = mobjPDCDGet.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,0)			
					'mobjSCGLSpr.CellChanged .sprSht, frmThis.sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					.txtDIVAMT.focus
					.sprSht.focus 
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht, 13, Row
					.txtDIVAMT.focus
					.sprSht.focus 
				End If
   			end if
   		ElseIF  Col = 8 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)
			vntData = mobjPDCDGet.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)

			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntData(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntData(3,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntData(4,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(7,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(8,0)		
					'mobjSCGLSpr.CellChanged .sprSht, frmThis.sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					.txtDIVAMT.focus
					.sprSht.focus 
					mobjSCGLSpr.ActiveCell .sprSht, Col+7,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht, 7, Row
					.txtDIVAMT.focus
					.sprSht.focus 
				End If
   			end if
		end if
   	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
	SUM_AMT
End Sub	
sub sprSht_DblClick (Col,Row)
	'선택된 로우 반환
	'window.returnvalue = mobjSCGLSpr.GetClip (frmThis.sprSht,1,frmThis.sprSht.ActiveRow,frmThis.sprSht.MaxCols,1,1)
	'call Window_OnUnload()
end sub
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub
sub imgDelRow_onclick ()
	With frmThis
		call sprSht_Keydown(meDEL_ROW, 0)
	End With 
end sub

Sub sprSht_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 13 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub

Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, .txtJOBNO.value 		
		mobjSCGLSpr.SetTextBinding .sprSht,"CREDAY",.sprSht.ActiveRow, .txtCREDAY.value  
	End with
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""
		IF Col = 10 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(6,0)				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+5,Row
			End IF
		elseIF Col = 13 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		elseIF Col = 7 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN0") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(8,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+8,Row
			End IF
		
		end if
		.txtDIVAMT.focus
		.sprSht.focus 

	End with
	
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
dim vntRet, vntInParams
	with frmThis
		IF Col = 10 Then			
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN1") then exit Sub
			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(6,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+4,Row
			End IF
		elseIF Col = 13 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		elseIF Col = 7 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN2") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",Row, vntRet(8,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtDIVAMT.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+7,Row
			End IF
		
		end if
		.txtDIVAMT.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
		.sprSht.Focus
	end with
End Sub
'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()
	Dim intNo,i,vntInParam
	
	set mobjPDCMDIVAMT = gCreateRemoteObject("cPDCO.ccPDCODIVAMT")
	set mobjPDCDGet = gCreateRemoteObject("cPDCO.ccPDCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	with frmThis
		.txtPREESTNO.style.visibility = "hidden"
		.txtYEARMON.style.visibility = "hidden"
		.txtCREDAY.style.visibility = "hidden"
		'★★★★★★★★★★★★★★★★★★IN 파라메터 및 조회를 위한 추가 파라메터 부터
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
	
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		''PREESTNO,YEARMON,JOBNO,CREDAY,DIVAMT
		for i = 0 to intNo
			select case i
				case 0 : .txtPREESTNO.value = vntInParam(i)	
				case 1 : .txtYEARMON.value = vntInParam(i)
				case 2 : .txtJOBNO.value = vntInParam(i)
				case 3 : .txtCREDAY.value = vntInParam(i)
				case 4 : .txtDIVAMT.value = vntInParam(i)
			end select
		next
		'★★★★★★★★★★★★★★★★★★IN 파라메터 및 조회를 위한 추가 파라메터 까지
		'msgbox .txtJOBYEARMON.value
		'SpreadSheet 디자인
		gSetSheetDefaultColor()
		txtDIVAMT_onblur
	End with
        With frmThis
			'메인쉬트
            gSetSheetColor mobjSCGLSpr, .sprSht 
			mobjSCGLSpr.SpreadLayout .sprSht, 17, 0
			mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.AddCellSpan  .sprSht,10, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.AddCellSpan  .sprSht,13, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.SpreadDataField .sprSht, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT"
			mobjSCGLSpr.SetHeader .sprSht,         "견적번호|순번|제작번호|년월|견적합의일|브랜드|브랜드명|사업부|사업부명|광고주|광고주명|분할금액|제작건명|청구금액"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "0       |0   |0       |0   |10        |6   |2|12      |6     |2|12    |6     |2|15    |10      |22      |0"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN0"
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN2"
			'PREESTNO|SEQ|JOBNO|YEARMON|CREDAY
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
			mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|SEQ|JOBNO|YEARMON|ADJAMT", true
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|JOBNAME|SUBSEQ|SUBSEQNAME", -1, -1, 255
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT", -1, -1, 0
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQ",-1,-1,1,2,false
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTCODE",-1,-1,2,2,false
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTNAME",-1,-1,0,2,false
			'mobjSCGLSpr.ColHidden .sprSht, "SEQ", true
			'mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			'Sum 쉬트
			gSetSheetColor mobjSCGLSpr, .sprShtSum
			mobjSCGLSpr.SpreadLayout .sprShtSum, 17, 1, 0,0,1,1,1,false,true,true,1
			mobjSCGLSpr.SpreadDataField .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT"
			mobjSCGLSpr.AddCellSpan  .sprShtSum, 2, 1, 2, 1
			mobjSCGLSpr.SetText .sprShtSum, 2, 1, "합 계"
			mobjSCGLSpr.SetScrollBar .sprShtSum, 0
			mobjSCGLSpr.SetBackColor .sprShtSum,"1|2",rgb(205,219,215),false
			mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "DIVAMT", -1, -1, 0
			mobjSCGLSpr.ColHidden .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON", true
			mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
			mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "15"
			.sprSht.focus
        End With
        
        SelectRtn
        SUM_AMT
end sub

Sub EndPage()
	set mobjPDCMDIVAMT = Nothing
	set mobjPDCDGet = Nothing
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

		vntData = mobjPDCMDIVAMT.SelectRtn_DIV(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)

		if not gDoErrorRtn ("SelectRtn_DIV") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			If mlngRowCnt < 1 Then
			frmThis.sprSht.MaxRows = 20 '최초로우개수 세팅할부분
			'mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
			End If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'Call SUM_AMT ()
   		end if
   	end with
end sub
Sub DeleteRtn_DTL
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	dim strJOBNO,strCUST,strSEQ
	Dim lngSUMAMT,lngSUMAMT2
	Dim strPREESTNO
	Dim dblSEQ
	Dim strRow
	'On error resume next
	
	with frmThis
		'한 건씩 삭제할 경우
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)

		if gDoErrorRtn ("DeleteRtn_Dtl") then exit sub

		if intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		if intRtn <> vbYes then exit sub
		
		strJOBNO = ""
		strCUST = ""
		strSEQ = 0
		lngSUMAMT = 0
		lngSUMAMT2 = 0
		'합계가 맞는지 여부검사
		'현재저장되어 있는 금액
		
		
		'선택된 자료를 끝에서 부터 삭제
		for i = intSelCnt-1 to 0 step -1
			strJOBNO = Trim(.txtJOBNO.value) 
			strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",vntData(i))	
			dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))	
			'Insert Transaction이 아닐 경우 삭제 업무객체 호출
			if cstr(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))) <> "" AND cstr(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))) <> "1" then
				If cstr(mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR02",vntData(i))) <> "" Then
					gErrorMsgBox "거래명세서 작성내역은 삭제될수 없습니다.","삭제오류"
					Exit Sub
				End If
				intRtn = mobjPDCMDIVAMT.DeleteRtn(gstrConfigXml,strJOBNO,strPREESTNO,dblSEQ)
			Elseif cstr(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))) = "1" Then
				gErrorMsgBox "최초생성 견적내역은 삭제될수 없습니다.","삭제오류"
				Exit Sub
			Else
				
			end if
			
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				'합계재계산
				gWriteText "", "자료가 삭제" & mePROC_DONE
   			end if
		next
		'ProcessRtn
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		
		SelectRtn
		
		
	end with
End Sub

'기본그리드의 헤더WIDTH가 변할시에 합계 그리드도 함께변한다.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
	End with
end sub
'스크롤이동시 합계 그리도도 함께 움직인다.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub
Sub SUM_AMT()
	Dim lngCnt
	Dim strSUMDEMANDAMT
	Dim strDIVAMT
	strSUMDEMANDAMT = 0
	With frmThis
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		
		mobjSCGLSpr.SetTextBinding .sprShtSum,"DIVAMT",1, strSUMDEMANDAMT
	End With
End Sub
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strJOBNO,strDEMANDAMT,strJOBYEARMON
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt,intCnt2
	
	with frmThis
   		'데이터 Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		'회의결과 달라도 저장될수 있음.. 분담금액이 청구금액보다 크다면 에러,,
		'만약 작다면 바로저장 청구금액이 예산에서 삭제 또는 삭감 되면 기존 분담 PD_GROUP_DIVAMT 의 내역 삭제 
		If CDBL(.txtDIVAMT.value) < strSUMDEMANDAMT Then
   			msgbox "분할금액의 합은 청구금액을 넘을수 없습니다."
   			Exit Sub
   		End IF
		
		'제작건명 처음의 로우와 일치 시키기
		For intCnt2 = 2 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt2) = "" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"JOBNAME",intCnt2, mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",1)  
			end if
		Next
		'쉬트의 변경된 데이터만 가져온다.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|DIVAMT|JOBNAME")
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "디테일 데이터를 입력 하십시오"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		intRtn = mobjPDCMDIVAMT.ProcessRtn(gstrConfigXml,vntData,.txtJOBNO.value )
	
		if not gDoErrorRtn ("ProcessRtn") then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "건의 자료가 저장" & mePROC_DONE,"저장안내!"
			SelectRtn
   		end if
   		
   	end with
End Sub
'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master 입력 데이터 Validation : 필수 입력항목 검사
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  		If mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",1) = "" Then
  			gErrorMsgBox "첫번째 행의 제작건명은 반드시 입력하셔야 합니다.","입력오류"
  			Exit Function
  		End if
  		for intCnt = 1 to .sprSht.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " 번째 행의 광고주코드를 확인하십시오","입력오류"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " 번째 행의 사업부코드를 확인하십시오","입력오류"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = 0 Then 
					gErrorMsgBox intCnt & " 번째 행의 분할금액을 확인하십시오","입력오류"
					Exit Function
			 End if
		next
		'AND mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBCODE",intCnt) = "" AND (mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",intCnt) = 0) 
   	End with
	DataValidation = true
End Function

-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%"border="0">
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
											<td class="TITLE" id="objTitle">분할금액 관리
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 225px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 108px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD style="WIDTH: 126px"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF"
													width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="자료를 닫습니다." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="HEIGHT: 100%" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="굴림">
										<TABLE class="KEY" id="tblKey" cellSpacing="0" cellPadding="0" width="1040" align="LEFT"
											border="0">
											<TBODY>
												<TR>
													<TD style="WIDTH: 109px" align="right">제작번호&nbsp;
													</TD>
													<td style="WIDTH: 313px"><INPUT class="NOINPUT" id="txtJOBNO" style="WIDTH: 144px; HEIGHT: 22px" readOnly type="text"
															size="18" name="txtJOBNO">
													</td>
													<TD style="WIDTH: 114px" align="right">
													견적확정금액&nbsp;
													<td><INPUT class="NOINPUT" id="txtDIVAMT" style="WIDTH: 200px; HEIGHT: 22px" tabIndex="1" readOnly
															type="text" size="28" name="txtDIVAMT">
													</td>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD style="HEIGHT: 26px" vAlign="bottom" align="right" width="100%"><INPUT class="NOINPUT" id="txtPREESTNO" style="WIDTH: 82px; HEIGHT: 22px" tabIndex="1"
										type="text" size="8" name="txtPREESTNO"><INPUT class="NOINPUT" id="txtYEARMON" style="WIDTH: 80px; HEIGHT: 22px" tabIndex="1" type="text"
										size="8" name="txtYEARMON"><INPUT class="NOINPUT" id="txtCREDAY" style="WIDTH: 98px; HEIGHT: 22px" tabIndex="1" type="text"
										size="11" name="txtCREDAY"><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="한 행 추가" src="../../../images/imgAddRow.gif"
										width="54" align="absMiddle" border="0" name="imgAddRow"><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'" alt="한 행 삭제" src="../../../images/imgDelRow.gif" width="54"
										align="absMiddle" border="0" name="imgDelRow">
								</TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="굴림">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 90%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="23125">
											<PARAM NAME="_ExtentY" VALUE="6641">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
										<OBJECT id="sprShtSum" style="WIDTH: 100%; HEIGHT: 5%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											 VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="23125">
											<PARAM NAME="_ExtentY" VALUE="609">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
