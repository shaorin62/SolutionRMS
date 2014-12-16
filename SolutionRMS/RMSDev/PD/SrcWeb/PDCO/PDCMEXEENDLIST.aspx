<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMEXEENDLIST.aspx.vb" Inherits="PD.PDCMEXEENDLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>결산 관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : SFAR/TR/차입금 등록 화면(TRLNREGMGMT0)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SheetSample.aspx
'기      능 : 차입금에 대한 MAIN 정보를 조회/입력/수정/삭제 처리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMEXEENDLIST, mobjPDCMGET
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
Dim mstrCHKROW
mstrCHKROW = false
Const meTab = 9
mALLCHECK = TRUE
mstrCheck=True
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
Sub ImgExeConfirm_onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub ImgExeConfirmCancel_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Cancel
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================
'****************************************************************************************
' 페이지 화면 디자인 및 초기화 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	
	'서버업무객체 생성	
	set mobjPDCMEXEENDLIST	= gCreateRemoteObject("cPDCO.ccPDCOEXEENDLIST")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
   
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	'탭 위치 설정 및 초기화
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "130px"
	'pnlTab1.style.left= "7px"
	
	'pnlTab2.style.position = "absolute"
	'pnlTab2.style.top = "593px"
	'pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'JOB 리스트
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|PROJECTNO|JOBNO|JOBNAME|DIVAMT|AMT|DIVFLAG|ENDDAY|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|ENDFLAG|JOBGUBN|CREPART|CREGUBN|REQDAY|COMMITION|CLIENTCODE|PREESTNO|DEMANDYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		   "선택|프로젝트번호|JOBNO|JOB명|청구예정금액|청구금액|청구상태|결산일|광고주|사업부|사업부명|브랜드|브랜드명|상태|매체부분|매체분류|신규|작성일|수수료율|광고주코드|확정견적코드|청구일자"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|0           |7    |   19|12          |10      |10      |10    |13    |6     |12      |6     |13      |   0|12      |12      |6   |10    |0       |0         |0			  | 10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|ENDDAY|DEMANDYEARMON", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "DIVAMT|AMT|DIVFLAG|PROJECTNO|JOBNO|JOBNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|JOBGUBN|CREPART|CREGUBN|REQDAY|ENDFLAG|CLIENTNAME|PREESTNO|DEMANDYEARMON"
		
		mobjSCGLSpr.ColHidden .sprSht, "PROJECTNO|COMMITION|CLIENTCODE|PREESTNO|ENDFLAG", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTSUBNAME|SUBSEQNAME|CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTSUBCODE|SUBSEQ|JOBGUBN|CREPART|CREGUBN|JOBNO|ENDFLAG|DIVFLAG|DEMANDYEARMON",-1,-1,2,2,false
		
		
	    
	    '******************************************************************
		'세부정산내역 리스트
		'******************************************************************
	   gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 20, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht1, 11, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1,   "JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|BTN|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ|PURCHASENO"
		mobjSCGLSpr.SetHeader .sprSht1,		   "제작번호|견적번호|순번|외주항목순번|외주항목코드|대분류|견적항목|수량|단가|금액|외주처코드|외주처|지급액|내역|전표번호|정산일|삽입구분|번호|정산번호"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "       0|      0|   4|           0|           0|    10|14      |7   |9   |11  |       9|2|16    |11    |12  |0       |9     |0       |0   |9"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "SORTSEQ|QTY|PRICE|AMT|ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "ITEMCODESEQ|ITEMCLASS|ITEMNAME", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "OUTSNAME|STD", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "ADJDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht1, true, "SORTSEQ|QTY|PRICE|AMT|ADJDAY|OUTSCODE|BTN|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ"
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "OUTSCODE|PURCHASENO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht1, "JOBNO|PREESTNO|ITEMCODESEQ|ITEMCODE|ADDFLAG|SEQ|VOCHNO|BTN", true
		
	    
	    .ImgExeConfirmCancel.disabled = true
		.ImgExeConfirm.disabled =  true		
    End With    
	'pnlTab1.style.visibility = "visible"
	'pnlTab2.style.visibility = "visible"
	'화면 초기값 설정
	InitPageData	
	
	
End Sub
Sub EndPage()
	set mobjPDCMEXEENDLIST = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

Sub initpageData

	with frmThis
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub
'-----------------------------------------------------------------------------------------
' 날자컨트롤 및 달력 / Onchange Event
'-----------------------------------------------------------------------------------------
Sub imgCalEndarFROM_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub


Sub cmbGUBN_onchange
	SelectRTN
	with frmThis
		if .cmbGUBN.value = "T" Then
		.ImgExeConfirm.disabled = true
		.ImgExeConfirmCancel.disabled = false
		Elseif  .cmbGUBN.value  = "F" Then
			.ImgExeConfirmCancel.disabled = true
			.ImgExeConfirm.disabled =  false
		Elseif  .cmbGUBN.value  = "" Then
			.ImgExeConfirmCancel.disabled = true
			.ImgExeConfirm.disabled =  true
		End If
	End with
End Sub
'-----------------------------------------------------------------------------------------
' Project 팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'ProjectNO 조회팝업
Sub ImgPROJECTNO1_onclick
	Call PONO_POP()
End Sub
'실제 데이터List 가져오기
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code값 저장
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' 코드명 표시
			'.txtCLIENTNAME1.focus()					' 포커스 이동
     	end if
	End with
	gSetChange
End Sub
'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' 광고주코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' 코드명 표시
		
				
     		'GetBrandDefaultFind	
     			
			
			.txtPROJECTNM1.focus()					' 포커스 이동
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
			
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
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
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim intCnt
	On error resume next
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		
		
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		'세금계산서 완료조회
		vntData = mobjPDCMEXEENDLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtPROJECTNO1.value),Trim(.txtPROJECTNM1.value),Trim(.txtCLIENTCODE.value),Trim(.txtCLIENTNAME.value),Trim(.cmbGUBN.value))
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
				.sprSht1.MaxRows = 0
			ELSE
				For intCnt = 1 To .sprSht.MaxRows
					If  .cmbGUBN.value  = "" Then
						'스태틱
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,8,8,true
					Else
						'체크
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,False
						
						If .cmbGUBN.value  = "T" Then
							mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,8,8,true
						Else
							If mobjSCGLSpr.GetTextBinding(.sprSht,"DIVFLAG",intCnt) = "청구미완료" Then
								mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
								mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
								mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,8,8,true
							End If
						End If
					End If	
   				Next
   				Call sprSht_Click(2,1)
			End If
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
End Sub

'------------------------------------------
' 정산내역 조회
'------------------------------------------
Sub SelectRtn_DTL (ByVal strJOBNO)
	Dim vntData1
	
	Dim intCnt
	'on error resume next
	with frmThis
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData1 = mobjPDCMEXEENDLIST.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
	If not gDoErrorRtn ("SelectRtn_DTL") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht1.MaxRows = 0	
			Else
				mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
				gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
			End If
	End If	
	End with
End SUB
'****************************************************************************************
' 쉬트 클릭 이벤트
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	Dim intcnt
	Dim strJOBNO
	
	
	with frmThis
	strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",Row)
		If Row = 0 and Col = 1  then 
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
				If  .cmbGUBN.value = "" Then
					'스태틱
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				Elseif .cmbGUBN.value = "F" Then
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DIVFLAG",intCnt) = "청구미완료" Then
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					End If
				End If			
			Next
		Else
			if Col = 1 Or Col = 6 Then
				
			Else
				SelectRtn_DTL(strJOBNO)
			End If
		end if
	end with
End Sub
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
'-----------------------------------------------------------------------------------------
' 확정 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strYEARMON
	Dim intCnt2
	Dim lngCHK
	with frmThis
	'On error resume next
		IF .cmbGUBN.value <> "F" THEN
			gErrorMsgBox "결산 은 미완료 조회시 가능합니다.","결산안내"
			Exit Sub
		End if
		lngCHK = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1"  Then
				lngCHK = lngCHK + 1
			End If
		Next
		If lngCHK = 0  Then 
		gErrorMsgBox "선택된건이 없습니다.","결산안내"
		Exit Sub
		End If
		
  		'데이터 Validation
		if DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|ENDDAY|JOBNO")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"결산안내"
			exit sub
		End If
	
		
		
		intRtn = mobjPDCMEXEENDLIST.ProcessRtn(gstrConfigXml,vntData)
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " 결산 처리가" & mePROC_DONE,"결산안내" 
			SelectRtn
  		end if
 	end with
End Sub
'-----------------------------------------------------------------------------------------
' 확정 취소 Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn_Cancel ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strYEARMON
	Dim intCnt2
	Dim lngCHK
	with frmThis
	'On error resume next
		IF .cmbGUBN.value <> "T" THEN
			gErrorMsgBox "결산취소는 완료 조회시 가능합니다.","결산취소안내"
			Exit Sub
		End if
		lngCHK = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = "1"  Then
				lngCHK = lngCHK + 1
			End If
		Next
		If lngCHK = 0  Then 
		gErrorMsgBox "선택된건이 없습니다.","결산취소안내"
		Exit Sub
		End If
  		'데이터 Validation
		if DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|ENDDAY|JOBNO")
		if  not IsArray(vntData) then 
			gErrorMsgBox "변경된 " & meNO_DATA,"마감취소안내"
			exit sub
		End If
	
		'처리 업무객체 호출
		
		
		intRtn = mobjPDCMEXEENDLIST.ProcessRtn_Cancel(gstrConfigXml,vntData)
		

		if not gDoErrorRtn ("ProcessRtn_Cancle") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " 결산취소가" & mePROC_DONE,"결산취소안내" 
			SelectRtn
  		end if
 	end with
End Sub
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
  	
		'Master 입력 데이터 Validation : 필수 입력항목 검사 TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" And mobjSCGLSpr.GetTextBinding(.sprSht,"ENDDAY",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 완료일자를 확인하십시오.","입력오류"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" height="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;정산관리</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 50px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="50">의뢰일자</TD>
									<TD class="SEARCHDATA" style="WIDTH: 200px"><INPUT class="INPUT" id="txtFROM" title="기간검색(FROM)" style="WIDTH: 72px; HEIGHT: 22px"
											accessKey="DATE" type="text" maxLength="10" size="5" name="txtFROM"><IMG id="imgCalEndarFROM" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
											border="0" name="imgCalEndarFROM">~<INPUT class="INPUT" id="txtTO" title="기간검색(TO)" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="8" name="txtTO"><IMG id="imgCalEndarTO" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgCalEndarTO"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 50px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
										width="50">광고주</TD>
									<TD class="SEARCHDATA" style="WIDTH: 233px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 136px; HEIGHT: 22px"
											type="text" maxLength="100" size="17" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" size="5" name="txtCLIENTCODE">
									</TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPROJECTNM1, txtPROJECTNO1)"
										width="50">프로젝트</TD>
									<TD class="SEARCHDATA" style="WIDTH: 225px"><INPUT class="INPUT_L" id="txtPROJECTNM1" title="프로젝트명 조회" style="WIDTH: 136px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="17" name="txtPROJECTNM1"><IMG id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgPROJECTNO1"><INPUT class="INPUT" id="txtPROJECTNO1" title="프로젝트명 조회" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="7" align="left" size="3" name="txtPROJECTNO1"></TD>
									<TD class="SEARCHLABEL" width="50">결산구분</TD>
									<TD class="SEARCHDATA"><SELECT id="cmbGUBN" style="WIDTH: 88px" name="cmbGUBN">
											<OPTION value="" selected>전체</OPTION>
											<OPTION value="F">미결산</OPTION>
											<OPTION value="T">결산</OPTION>
										</SELECT></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;JOB 리스트</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="ImgExeConfirm" onmouseover="JavaScript:this.src='../../../images/ImgExeSetOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgExeSet.gIF'"
														height="20" alt="결산처리를 합니다." src="../../../images/ImgExeSet.gIF" border="0" name="ImgExeConfirm"></TD>
												<TD><IMG id="ImgExeConfirmCanCel" onmouseover="JavaScript:this.src='../../../images/ImgExeSetCancelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgExeSetCancel.gif'"
														height="20" alt="결산을 취소합니다." src="../../../images/ImgExeSetCancel.gIF" border="0"
														name="ImgExeConfirmCanCel"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
							<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="10663">
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
						<TD>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="굴림"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="굴림"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;세부 정산내역</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgExcel1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" border="0" name="imgExcel1"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody1" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
							<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="6535">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;합 
							계 :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="금액" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="굴림"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
