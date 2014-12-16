<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVMEDIUMSEARCH.aspx.vb" Inherits="MD.MDCMCATVMEDIUMSEARCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>매체관리</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : PROJECT 등록 화면(PDCMPONO)
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPONO.aspx
'기      능 : 프로젝트 등록 및 관리
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/10/27 By Tae Ho Kim
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
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMMEDIUMLIST
Dim mobjMDCOGET
Dim mcomecalender, mcomecalender2
Dim mstrHIDDEN
CONST meTAB = 9
mcomecalender = FALSE
mcomecalender2 = FALSE
mstrHIDDEN = 0

'=============================
' 이벤트 프로시져 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------

'입력 필드 숨기기
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblBody").style.display = "inline"
		Else
			document.getElementById("tblBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
		SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'출력 인쇄버튼 클릭시 이벤트
Sub imgPrint99999999_onclick ()
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,j,k
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
	'체크가 된 데이터가 있는지 없는지 체크한다.
	intCount = 0
	for i=1 to frmThis.sprSht.MaxRows
		IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
			intCount = 1
		end if
	next
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if intCount = 0 then
		gErrorMsgBox "선택된 데이터가 없습니다. 인쇄할 데이터를 체크하시오",""
		Exit Sub
	end if
	
	gFlowWait meWAIT_ON
	with frmThis
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "PD"
		ReportName = "PDCMTRANS.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjPD_TRANS.Get_TRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_TRANS_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					
					datacnt = strcntsum
					strUSERID = ""
					vntDataTemp = mobjPD_TRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
					
				End IF
			END IF
		next
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		window.setTimeout "printSetTimeout", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub

'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndar,"txtFOM_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgCalEndarREQ_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		mcomecalender2 = true
		gShowPopupCalEndar frmThis.txtTO,frmThis.imgCalEndar,"txtTO_onchange()"
		mcomecalender2 = false
		gSetChange
	end with
End Sub

'-----------------------------------------------------------------------------------------
' 매체명코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'이미지버튼 클릭시
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'실제 데이터List 가져오기
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array("", "",trim(.txtMEDCODE.value), trim(.txtMEDNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtMEDCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtMEDNAME.value = trim(vntRet(1,0))       ' 코드명 표시
			'.txtREAL_MED_CODE.value = trim(vntRet(3,0))       ' 코드명 표시
			'.txtREAL_MED_NAME.value = trim(vntRet(4,0))       ' 코드명 표시
			
		End If
	End With
	gSetChange
End Sub

'한건을 찾을경우 엔터 이벤트로써 해당값을 뿌려줌
Sub txtMEDNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetMEDCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"","", trim(.txtMEDCODE.value),trim(.txtMEDNAME.value))
			
			If not gDoErrorRtn ("GetMEDCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,1))	    ' Code값 저장
					.txtMEDNAME.value = trim(vntData(1,1))       ' 코드명 표시
					'.txtREAL_MED_CODE.value = trim(vntData(3,1))
					'.txtREAL_MED_NAME.value = trim(vntData(4,1))
				Else
					Call MEDCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 소재명팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgMATTERCODE_onclick
	Call MATTERCODE_POP()
End Sub

Sub MATTERCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTNAME.value),"" , trim(.txtSUBSEQNAME.value),"", _
							trim(.txtMATTERNAME.value), "" , "B") '<< 받아오는경우
		
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
		
		If isArray(vntRet) Then
			If .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtMATTERCODE.value = trim(vntRet(0,0))	' 소재코드 표시
			.txtMATTERNAME.value = trim(vntRet(1,0))	' 소재명 표시
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주코드 표시
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
			'.txtTIMCODE.value = trim(vntRet(4,0))		' 팀코드 표시
			'.txtTIMNAME.value = trim(vntRet(5,0))		' 팀명 표시
			.txtSUBSEQ.value = trim(vntRet(6,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(7,0))	' 브랜드명 표시
			'.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' 제작사코드 표시
			'.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' 제작사코드 표시
			'.txtDEPT_CD.value = trim(vntRet(10,0))		' 부서코드 표시
			'.txtDEPT_NAME.value = trim(vntRet(11,0))	' 부서명 표시
			
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME.value),"", trim(.txtSUBSEQNAME.value), "" , _
											trim(.txtMATTERNAME.value), "" , "B")
											
			If not gDoErrorRtn ("GetMATTER") Then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntRet(0,1))	' 소재코드 표시
					.txtMATTERNAME.value = trim(vntRet(1,1))	' 소재명 표시
					.txtCLIENTCODE.value = trim(vntRet(2,1))	' 광고주코드 표시
					.txtCLIENTNAME.value = trim(vntRet(3,1))	' 광고주명 표시
					.txtTIMCODE.value	 = trim(vntRet(4,1))	' 팀코드 표시
					.txtTIMNAME.value	 = trim(vntRet(5,1))	' 팀명 표시
					.txtSUBSEQ.value	 = trim(vntRet(6,1))	' 브랜드 표시
					.txtSUBSEQNAME.value = trim(vntRet(7,1))	' 브랜드명 표시
					'.txtEXCLIENTCODE.value = trim(vntRet(8,1))	' 제작사코드 표시
					'.txtDEPT_CD.value	 = trim(vntRet(10,1))	' 부서코드 표시
					'.txtDEPT_NAME.value	 = trim(vntRet(11,1))	' 부서명 표시
				
				Else
					Call MATTERCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' 브랜드코드팝업 버튼[조회용]
'-----------------------------------------------------------------------------------------
'광고주 시퀀스가져오기
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< 받아오는경우
		
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))		' 브랜드 표시
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	' 브랜드명 표시
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' 광고주 표시
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' 광고주명 표시
			'.txtTIMCODE.value = trim(vntRet(4,0))	' 광고주명 표시
			'.txtTIMNAME.value = trim(vntRet(5,0))	' 광고주명 표시
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
												trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	' 광고주 표시
					.txtCLIENTNAME.value = trim(vntData(3,1))	' 광고주
					'.txtTIMCODE.value = trim(vntData(4,1))	' 팀모드
					'.txtTIMNAME.value = trim(vntData(5,1))	' 팀명
				Else
					Call SUBSEQCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
'CIC/사업부 팝업  버튼[조회용]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

Sub CLIENTSUBCODE_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCLIENTSUBPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtCLIENTSUBCODE.value = trim(vntRet(0,0))	'Code값 저장
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))	'코드명 표시
			.txtCLIENTCODE.value = trim(vntRet(3,0))	'Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(4,0))	'코드명 표시
			
			gSetChangeFlag .txtCLIENTCODE
		End If
	end With
End Sub

Sub txtCLIENTSUBNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type의 ByRef 변수의 초기화
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
			If not gDoErrorRtn ("GetCLIENTSUBCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))	'Code값 저장
					.txtCLIENTNAME.value = trim(vntData(1,0))	'코드명 표시
					.txtCLIENTSUBCODE.value = trim(vntData(3,0))	'Code값 저장
					.txtCLIENTSUBNAME.value = trim(vntData(4,0))	'코드명 표시
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' 변경된 데이터가 없다면 exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code값 저장
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' 코드명 표시
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
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 입력필드 키다운 이벤트
'****************************************************************************************
Sub cmbVAT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtFROM.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' 입력필드 체인지 이벤트
'****************************************************************************************


Sub txtFROM_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	Dim strOLDYEARMON
	strdate = ""
	strFROM =""
	strFROM2 = ""
	With frmThis
		strdate=.txtFROM.value
		'달력팝업후 오는 데이터는 2000-01-01이런식으로 들어오고 직접입력은 20000101이런식으로 들어오므로
		If mcomecalender Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strFROM2 = strdate
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		
		
		IF .txtFROM.value <> "" THEN
			.txtTO.value = strFROM
			DateClean strFROM
		END IF
	
	End With
	gSetChange
End Sub

'****************************************************************************************
'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
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

'****************************************************************************************
' Amt 선택한 값들의 합계를 텍스트박스에 뿌려준다
'****************************************************************************************
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")  Then
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


'-----------------------------------------------------------------------------------------
' 천단위 나눔점 표시 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------
'금액
Sub txtFROMAMOUNT_onblur
	with frmThis
		'COMMI_RATE_Cal
		call gFormatNumber(.txtFROMAMOUNT,0,true)
	end with
End Sub

'금액
Sub txtTOAMOUNT_onblur
	with frmThis
		'COMMI_RATE_Cal
		call gFormatNumber(.txtTOAMOUNT,0,true)
	end with
End Sub

'-----------------------------------------------------------------------------------------
' 천단위 나눔점 없애기 ( 단가, 금액, 수수료)
'-----------------------------------------------------------------------------------------

Sub txtFROMAMOUNT_onfocus
	with frmThis
		.txtFROMAMOUNT.value = Replace(.txtFROMAMOUNT.value,",","")
	end with
End Sub

Sub txtTOAMOUNT_onfocus
	with frmThis
		.txtTOAMOUNT.value = Replace(.txtTOAMOUNT.value,",","")
	end with
End Sub



'****************************************************************************************
' SpreadSheet 이벤트
'****************************************************************************************
Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
Sub cmbMED_onchange
	with frmThis
		If .cmbMED.selectedIndex = 0 Then
		End If
		'인쇄
		If .cmbMED.selectedIndex = 1 Then
		End IF
	End With
end Sub



'****************************************************************************************
'****************************************************************************************
'=============================
' UI업무 프로시져 
'=============================
'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	'서버업무객체 생성	
	set mobjMDCMMEDIUMLIST = gCreateRemoteObject("cMDCO.ccMDCOMEDIUMLIST")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis
	   	gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 20, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, " CHK|MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|MEDCODE|MEDNAME|MATTERCODE|MATTERNAME|SUBSEQ|SUBSEQNAME|TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|AMT|COMMI_RATE|VAT"
		mobjSCGLSpr.SetHeader .sprSht,        "선택|매체구분|청약일|청약구분|청구구분|청구일|매체코드|매체명|소재코드|소재명|브래드코드|브랜드|팀코드|CIC/팀|광고주코드|광고주|집행금액|수수료율|VAT"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|       8|     8|       8|        8|     8|       0|    15|       0|    15|         0|    15|     0|    15|         0|    14|      14|      8|  6|" 
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|MATTERNAME|SUBSEQNAME|TIMNAME|CLIENTNAME",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|VAT",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|MEDCODE|MEDNAME|MATTERCODE|MATTERNAME|SUBSEQ|SUBSEQNAME|TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|AMT|COMMI_RATE|VAT"
		mobjSCGLSpr.ColHidden .sprSht, "MEDCODE | MATTERCODE | SUBSEQ | TIMCODE | CLIENTCODE", True
		.sprSht.style.visibility = "visible"
	End With
	'화면 초기값 설정
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	gEndPage
End Sub


'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'초기 데이터 설정
	with frmThis
		.sprSht.MaxRows = 0
		.txtFROM.value = gNowDate
		.txtTO.value  = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)	
		
		'날짜셋팅 - 시작달의 마지막일
		DateClean .txtTO.value
		
		.txtFROM.focus
	End with
End Sub

'날짜 조회조건 생성
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE

	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
	With frmThis
		.txtTO.value = date2
	End With
End Sub

'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strFROM
   	Dim strTO
   	Dim strVOCH_TYPE
   	Dim strVOCH_TYPE2
   	Dim strMEDCODE
   	Dim strMATTERCODE
   	Dim strSUBSEQ
   	Dim	strCLIENTSUBCODE
   	Dim strCLIENTCODE
   	Dim strFROMAMOUNT
   	Dim strTOAMOUNT
   	Dim strCOMMI_RATE
   	Dim strVAT
   	
	with frmThis
		'Sheet초기화
		.sprSht.MaxRows = 0
		
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM				= .txtFROM.value
		strTO				= .txtTO.value
		strVOCH_TYPE		= .cmbVOCH_TYPE.value
		strVOCH_TYPE2		= .cmbVOCH_TYPE2.value
		strMEDCODE			= .txtMEDCODE.value
		strMATTERCODE		= .txtMATTERCODE.value
		strSUBSEQ			= .txtSUBSEQ.value
		strCLIENTSUBCODE	= .txtCLIENTSUBCODE.value
		strCLIENTCODE		= .txtCLIENTCODE.value
		strFROMAMOUNT		= .txtFROMAMOUNT.value
		strTOAMOUNT			= .txtTOAMOUNT.value
		strCOMMI_RATE		= .cmbCOMMI_RATE.value
		strVAT				= .cmbVAT.value

		vntData = mobjMDCMMEDIUMLIST.SelectRtn_CATV(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strVOCH_TYPE,strVOCH_TYPE2,strMEDCODE,strMATTERCODE,strSUBSEQ,strCLIENTSUBCODE,strCLIENTCODE,strFROMAMOUNT,strTOAMOUNT,strCOMMI_RATE,strVAT)
		
		If not gDoErrorRtn ("SelectRtn") then
			'조회한 데이터를 바인딩
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'초기 상태로 설정
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht.MaxRows = 0	
			End If
			gWriteText lblstatus, "선택한 자료에 대해서 " & mlngRowCnt & " 건의 자료가 검색" & mePROC_DONE			
		
		End If		
	END WITH
	'조회완료메세지
	gWriteText "", "자료가 검색" & mePROC_DONE
	AMT_SUM
End Sub



'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
				</TR>
				<!--TopSplit End-->
				<!--Input Start-->
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="430" colSpan="2" height="28">
									<table cellSpacing="0" cellPadding="0" width="800" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;청약관리-세부내역조회 <span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()">
													(숨기기)</span>
											</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" colSpan="2" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 600px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<td><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF" align="right" border="0"
													name="imgQuery"></td>
											<td><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="자료를 인쇄합니다." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></td>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE style="WIDTH: 100%; HEIGHT: 90%" cellSpacing="0" cellPadding="0" align="left" border="0">
							<TR>
								<TD style="HEIGHT: 4px"><FONT face="굴림"></FONT></TD>
							</TR>
							<TR>
								<TD id="tblBody" style="WIDTH: 280px; HEIGHT: 91%" vAlign="top">
									<table class="DATA" id="tblKey2" style="WIDTH: 272px; HEIGHT: 302px" cellSpacing="1" cellPadding="0"
										width="272" align="left" border="0">
										<tr>
											<td class="TITLE" width="272" colSpan="2">전체합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="WIDTH: 202px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT"></td>
										</tr>
										<tr>
											<td class="TITLE" colSpan="2">선택합계 : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 202px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
										<tr>
											<TD class="GROUP" colSpan="2">조회조건</TD>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">매체구분</TD>
											<td class="DATA" width="184"><SELECT id="cmbMED" title="매체구분" style="WIDTH: 111px" name="cmbMED">
													<OPTION value="" selected>CATV</OPTION>
												</SELECT></td>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">기간</TD>
											<td class="DATA"><INPUT class="INPUT" id="txtFROM" title="의뢰일 검색(FROM)" style="WIDTH: 76px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarFROM1">~<INPUT class="INPUT" id="txtTO" title="의뢰일 검색(TO)" style="WIDTH: 76px; HEIGHT: 22px" accessKey="DATE"
													type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">청약구분</TD>
											<td class="DATA"><SELECT id="cmbVOCH_TYPE" title="청약구분" style="WIDTH: 111px" name="cmbVOCH_TYPE">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="0">위수탁</OPTION>
													<OPTION value="1">협찬</OPTION>
													<OPTION value="2">일반</OPTION>
												</SELECT></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">청구구분</TD>
											<td class="DATA"><SELECT id="cmbVOCH_TYPE2" title="청구구분" style="WIDTH: 111px" name="cmbVOCH_TYPE2">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="0">위수탁</OPTION>
													<OPTION value="2">일반</OPTION>
												</SELECT></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME, txtMEDCODE)"
												width="88">매체명</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtMEDNAME" title="매체명" style="WIDTH: 125px; HEIGHT: 22px" type="text"
													maxLength="100" size="12" name="txtMEDNAME"><IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
													width="22" align="absMiddle" border="0" name="ImgMEDCODE"><INPUT class="INPUT_L" id="txtMEDCODE" title="매체명코드" style="WIDTH: 59px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="4" name="txtMEDCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtMATTERNAME, txtMATTERCODE)"
												width="88">소재명</TD>
											<td class="DATA"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="소재명" style="WIDTH: 125px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="500" size="30" name="txtMATTERNAME"><IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="22" align="absMiddle" border="0"
													name="ImgMATTERCODE"><INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="소재코드" style="WIDTH: 59px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="10" size="4" name="txtMATTERCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME, txtSUBSEQ)"
												width="88">브랜드</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="브랜드명" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" size="12" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgSUBSEQCODE"><INPUT class="INPUT_L" id="txtSUBSEQ" title="시퀀스코드" style="WIDTH: 59px; HEIGHT: 22px" type="text"
													maxLength="9" name="txtSUBSEQ"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)"
												width="83">CIC/팀</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="팀사업부명" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" size="26" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT_L" id="txtCLIENTSUBCODE" title="사업부코드" style="WIDTH: 59px; HEIGHT: 22px"
													type="text" maxLength="9" name="txtCLIENTSUBCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="83">광고주</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="광고주명" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="16" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="광고주코드" style="WIDTH: 59px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROMAMOUNT, txtTOAMOUNT)"
												width="83">집행금액</TD>
											<td class="DATA"><INPUT class="INPUT_R" id="txtFROMAMOUNT" title="금액" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="13" size="20" name="txtFROMAMOUNT">~<INPUT class="INPUT_R" id="txtTOAMOUNT" title="금액" style="WIDTH: 99px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="13" size="9" name="txtTOAMOUNT"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px" width="83">수수료율</TD>
											<td class="DATA"><SELECT id="cmbCOMMI_RATE" title="수수료율" style="WIDTH: 80px" name="cmbCOMMI_RATE">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="1">1</OPTION>
													<OPTION value="2">2</OPTION>
													<OPTION value="3">3</OPTION>
													<OPTION value="4">4</OPTION>
													<OPTION value="5">5</OPTION>
													<OPTION value="6">6</OPTION>
													<OPTION value="7">7</OPTION>
													<OPTION value="8">8</OPTION>
													<OPTION value="9">9</OPTION>
													<OPTION value="10">10</OPTION>
													<OPTION value="15">15</OPTION>
													<OPTION value="20">20</OPTION>
													<OPTION value="25">25</OPTION>
													<OPTION value="30">30</OPTION>
													<OPTION value="35">35</OPTION>
													<OPTION value="40">40</OPTION>
													<OPTION value="35">45</OPTION>
													<OPTION value="40">50</OPTION>
													<OPTION value="100">50초과</OPTION>
												</SELECT>
												(%)</td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px" width="83">VAT</TD>
											<td class="DATA"><SELECT id="cmbVAT" title="VAT" style="WIDTH: 111px" name="cmbVAT">
													<OPTION value="" selected>전체</OPTION>
													<OPTION value="1">과세</OPTION>
													<OPTION value="01">면세</OPTION>
													<OPTION value="02">영세</OPTION>
												</SELECT></td>
										</tr>
									</table>
								</TD>
								<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" ms_positioning="GridLayout">
										<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 100%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32438">
											<PARAM NAME="_ExtentY" VALUE="20929">
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
											<PARAM NAME="ReDraw" VALUE="-1">
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
								<TD><!--좌측여백! 지우지말것 지울경우 아래TD에 COLSPAN 추가--></TD>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
