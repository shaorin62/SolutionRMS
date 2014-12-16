<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDEMANDPOP.aspx.vb" Inherits="PD.PDCMDEMANDPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>청구요청 미리보기</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMDEMANDPOP.aspx
'기      능 : 청구요청 화면의 청구요청미리보기 버튼 클릭시 미리보기 화면으로 제공되며, 청구요청을 하여 PD_DIVAMT 에 투입또는 업데이트 한다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/06 By KimTH
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
		
Dim mlngRowCnt,mlngColCnt
Dim mobjPDCODEMAND
Dim mobjPDCMGET
Dim mobjSCCOGET
Dim mstrYEARMON1,mstrYEARMON2, mstrUSENO
Dim mstrCheck	
Dim mstrGBN
Dim mlngTempRowCnt,mlngTempColCnt
Dim mstrITEMCODESEQ


Dim mvntData

mstrCheck = True	

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	
	with frmThis
		window.returnvalue = "SAVETRUE"
	End with
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
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
Sub imgRowDel_onclick()

End Sub

Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCODEMAND = gCreateRemoteObject("cPDCO.ccPDCODEMANDLIST")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정

		'mstrPREESTNO,mstrITEMCODE,mlngIMESEQ
		for i = 0 to intNo
			select case i
				case 0 : mstrYEARMON1 = vntInParam(i)			'해당월
				case 1 : mstrYEARMON2 = vntInParam(i)			'해당월
				case 2 : mstrUSENO = vntInParam(i)				'해당사용자
			end select
		next
		'PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|AMT
	'**************************************************
	'***Sum Sheet 디자인
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 29, 0
	mobjSCGLSpr.SpreadDataField .sprSht,    "YEARMON|PREESTNO|JOBNAME|JOBNO|SEQ|CREDAY|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|DEMANDFLAGNAME|MEMO|TAXCODE|TAXCODENAME|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON"
	mobjSCGLSpr.SetHeader .sprSht,		    "요청월|견적번호|제작건명|JOBNO|SUBNO.|견적일|광고주코드|광고주|팀코드|팀명|브랜드코드|브랜드|견적금액|청구금액|잔액|청구기준|청구기준|비고내역|청구방법|청구방법|담당자|완료구분|승인구분|SORT|그룹핑|상세시퀀스|승인권자|차월이력|승인요청월"
	mobjSCGLSpr.SetColWidth .sprSht, "-1",  "     7|      10|15      |7    |6     |9     |0         |13    |0     |13  |0         |13    |11      |11      |11  |10      |10      |10      |10      |10      |6     |10      |10      |10  |10    |10        |10      |10      |10"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	'mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT|CHARGE", -1, -1, 0
	'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|", -1, -1, 10
	'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,"YEARMON|PREESTNO|JOBNAME|JOBNO|SEQ|CREDAY|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV|DEMANDFLAGNAME|TAXCODENAME|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "PREESTNO|JOBNAME|CLIENTNAME|TIMNAME|SUBSEQNAME|MEMO|DEMANDFLAGNAME|TAXCODENAME",-1,-1,0,2,false ' 왼쪽
	mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON|JOBNO|SEQ|CREDAY|CLIENTCODE|TIMCODE|SUBSEQ|DEMANDFLAG|TAXCODE|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV",-1,-1,2,2,false '가운데
	'mobjSCGLSpr.ColHidden .sprSht, "ATTR", true 
	'CHK|PREESTNO|SEQ|ITEMCODESEQ|ITEMCODE|AMT
	'mobjSCGLSpr.ColHidden .sprSht, "PREESTNO", true

	pnlTab1.style.visibility = "visible" 
	.txtYEARMON1.value = mstrYEARMON1
	.txtYEARMON2.value = mstrYEARMON2
	.txtUSENO.value = mstrUSENO
	
	SelectRtn
	.txtEMPNAME.focus()
	End with
	 
End Sub

Sub InitpageData
	with frmThis
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub imgRowAdd_onclick ()
call sprSht_Keydown(meINS_ROW, 0)
End Sub
'================================================================
'UI
'================================================================
Sub txtDIVAMT_onfocus
	with frmThis
		.txtDIVAMT.value = Replace(.txtDIVAMT.value,",","")
	end with
End Sub
Sub txtDIVAMT_onblur
	with frmThis
		call gFormatNumber(.txtDIVAMT,0,true)
	end with
End Sub

Sub txtADJAMT_onfocus
	with frmThis
		.txtADJAMT.value = Replace(.txtADJAMT.value,",","")
	end with
End Sub
Sub txtADJAMT_onblur
	with frmThis
		call gFormatNumber(.txtADJAMT,0,true)
	end with
End Sub

Sub txtCHARGE_onfocus
	with frmThis
		.txtCHARGE.value = Replace(.txtCHARGE.value,",","")
	end with
End Sub
Sub txtCHARGE_onblur
	with frmThis
		call gFormatNumber(.txtCHARGE,0,true)
	end with
End Sub



'================================================================
'SpreadSheet Event
'================================================================
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
		'필드 To 바인딩 존재시 기입
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
				strCOLUMN = "DIVAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
				strCOLUMN = "CHARGE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
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
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	Dim lngChCnt,IntChAMT,IntChAMTSUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtDIVAMT.value = 0
		else
			.txtDIVAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtDIVAMT,0,True)
		End If
		
		IntEXEAMTSUM = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtADJAMT.value = 0
		else
			.txtADJAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtADJAMT,0,True)
		End If
		
		IntChAMTSUM = 0
		For lngChCnt = 1 To .sprSht.MaxRows
			IntChAMT = 0	
			IntChAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHARGE", lngChCnt)
			IntChAMTSUM = IntChAMTSUM + IntChAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtCHARGE.value = 0
		else
			.txtCHARGE.value = IntChAMTSUM
			Call gFormatNumber(frmThis.txtCHARGE,0,True)
		End If
	End With
End Sub



Sub EndPage
	Set mobjPDCODEMAND = Nothing
	Set mobjPDCMGET = Nothing
	Set mobjSCCOGET = Nothing
	
	gEndPage
End Sub
'=============================================================
'Sheet Event
'=============================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub


Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	
End Sub

'=============================================================
'조회
'=============================================================

Sub SelectRtn
	Dim vntData
   	Dim i, strCols
    Dim strCHK
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCODEMAND.SelectRtn_PreView(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value,.txtYEARMON2.value,.txtUSENO.value)
		
		if not gDoErrorRtn ("SelectRtn_PreView") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht.MaxRows 
					'JOB별 컬러 통일
					If mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",intCnt) <> "전월이월분" Then
						If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKDIV",intCnt) Mod 2 = "0" Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					End If
				Next
   		
   			Else
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			
   		end if
   	window.setTimeout "AMT_SUM",1	
	.txtSELECTAMT.value = 0
   	end with
   	
   	
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
		
		vntRet = gShowModalWindow("../../../PD/SrcWeb/PDCO/PDCMEMPPOP_MANAGER.aspx",vntInParams , 413,435)
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
			vntData = mobjPDCMGET.GetPDEMP_MANAGER(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
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

'-----------------------------------------------------------------------------------------
' 승인요청
'-----------------------------------------------------------------------------------------
Sub processRtn
	Dim vntData
	Dim intRtn
	Dim strSAVEGBN
	Dim intCnt,intCnt2,intCnt3,intMsgCnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg
	'SMS 정보
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		
		strMasterData = gXMLGetBindingData (xmlBind)
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "청구요청건이 없습니다.","청구요청안내"
			Exit Sub
		End If
		
		'쉬트의 변경된 데이터만 가져온다.
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "승인권자를 선택 하십시오.","청구요청안내"
			Exit Sub
		End If
		
		
		'승인권자 를 그리드에 탑재
		intMsgCnt = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"MANAGER",intCnt2,Trim(.txtEMPNO.value)
			'그리드의 제작건명 을 가져온다
			If intCnt2 = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt2)
			End If
			intMsgCnt = intMsgCnt +1
		Next
	
	
		If intMsgCnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] 승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 승인요청이있습니다"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] 외" & intMsgCnt-1 & "건의승인요청이있습니다"
			Else
				strMstMsg = "[ " & strMsg & "] 외" & intMsgCnt-1 & "건의승인요청이있습니다"
			End If
		End If
		
		if DataValidation =false then exit sub 	

		intSaveRtn = gYesNoMsgbox("해당데이터를 청구요청 하시겠습니까?","청구요청 확인")
		IF intSaveRtn <> vbYes then 
			exit Sub
		Else
		
			'전체행을 가져와야 한다.
			For intCnt = 1 To .sprSht.MaxRows
				mobjSCGLSpr.CellChanged .sprSht, 1, intCnt	
			Next
			
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|JOBNO|SEQ|PREESTNO|JOBNAME|CLIENTCODE|TIMCODE|SUBSEQ|CREDAY|DIVAMT|ADJAMT|CHARGE|ENDFLAG|DEMANDFLAG|CONFIRMFLAG|MEMO|USENO|TAXCODE|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON")
			
			intRtn = mobjPDCODEMAND.ProcessRtn_Demand(gstrConfigXml,vntData, .txtYEARMON1.value,.txtYEARMON2.value,.txtUSENO.value)
			If not gDoErrorRtn ("ProcessRtn_Demand") Then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gOkMsgBox "저장되었습니다.","저장안내!"
				
				'승인을 수락하였으므로 SMS 발송
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
			
				
				strAMT = .txtADJAMT.value 
				call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
				
				
				Window_OnUnload
			End If
		End If
		
	End with
End Sub
'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	
   	Dim intCnt
	'On error resume next
	with frmThis
		
   		for intCnt = 1 to .sprSht.MaxRows
   			'Sheet 필수 입력사항
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht,"MANAGER",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 행의 승인권자 입력에 문제가 있습니다" & vbcrlf & "운영팀 에게 문의 하십시오.","청구요청안내"
				Exit Function
			End if
			
		next
   	
   	End with
   	
	DataValidation = true
End Function


		</script>
		<script language="javascript">
		//SMS 발송
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
		</script>
		
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis"><br>
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table style="WIDTH: 100%; HEIGHT: 24px" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="138" background="../../../images/back_p.gIF"
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
								<td class="TITLE">청구요청내역 미리보기</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<TD>
						<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
							<TR>
								<td class="SEARCHDATA" style="WIDTH: 911px" width="911" colSpan="7">&nbsp;청구요청월 <INPUT class="NOINPUTB" id="txtYEARMON1" title="청구요청월" style="WIDTH: 96px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="10" size="10" name="txtYEARMON1">&nbsp;~&nbsp;<INPUT class="NOINPUTB" id="txtYEARMON2" title="청구요청월" style="WIDTH: 96px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="10" size="10" name="txtYEARMON2">
									담당자&nbsp; <INPUT class="NOINPUTB_R" id="txtUSENO" title="간접비" style="WIDTH: 112px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="15" size="13" name="txtUSENO">&nbsp;님 
									께서 청구 요청하실 내역입니다.</td>
								<td align="right" ><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="화면을 닫습니다."
										src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
							</TR>
						</TABLE>
					</TD>
				</tr>
			</table>
			<BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">합 계 : <INPUT class="NOINPUTB_R" id="txtDIVAMT" title="견적금액합계" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtDIVAMT"> <INPUT class="NOINPUTB_R" id="txtADJAMT" title="청구금액합계" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtADJAMT">&nbsp;<INPUT class="NOINPUTB_R" id="txtCHARGE" title="잔액합계" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtCHARGE">&nbsp;<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<td style="FONT-WEIGHT: bold; FONT-SIZE: 12px" align="right" width="600"><span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">승인자:</span>
						&nbsp;<INPUT class="NOINPUTB_L" id="txtEMPNAME" title="승인권자" style="WIDTH: 96px; HEIGHT: 20px"
							type="text" maxLength="100" size="10" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
							name="ImgEMPNO" title="승인권자선택"> <INPUT class="NOINPUTB" id="txtEMPNO" title="승인권자사번" style="WIDTH: 58px; HEIGHT: 20px"
							type="text" maxLength="100" size="4" name="txtEMPNO">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgDivDemandOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDivDemand.gIF'" height="20" alt="청구요청을 합니다.." src="../../../images/imgDivDemand.gif"
							align="absMiddle" border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"
							width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</td>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--내용-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="30506">
								<PARAM NAME="_ExtentY" VALUE="12435">
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
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
				</TR>
			</table>
		</form>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 0px;HEIGHT: 0px" name="frmSMS"></iframe> <!--DISPLAY: none; -->
	</body>
</HTML>
