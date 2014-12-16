<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_ESTDTL.aspx.vb" Inherits="PD.PDCMJOBMST_ESTDTL" %>
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
'HISTORY    :1) 2011/05/04 By kty
'****************************************************************************************
-->
		<meta content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script id="clientEventHandlersVBS" language="vbscript">

option explicit

Dim mobjPDCOPREESTDTL, mobjPDCOGET
Dim mlngTempRowCnt
Dim mlngTempColCnt
Dim mlngRowCnt, mlngColCnt		
Dim mstrFIRSTPRODUCTIONCHECK
Dim mstrSELECT		'견적조회시 해당 JOBNO 와 보여지는 JOBNO 의 차이를  확인  다르면 "F" 같으면 "T"
Dim mstrCheck

mstrFIRSTPRODUCTIONCHECK = "N"
CONST meTAB = 9
mstrCheck = TRUE

'=============================
' 이벤트 프로시져 
'=============================
Sub window_onload
	window.setTimeout "call Initpage()",1000 
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'기본포멧 적용
Sub ImgBasicFormat_onclick
	Dim vntData
	Dim vntRet, vntInParams
	Dim intCnt
	Dim strPREESTNO
	
	with frmThis
		If .sprSht.MaxRows <> 0 Or .txtPREESTGBN.value = "본견적" Then
			gErrorMsgBox "본견적 및 상세내역 존재시 견적유형을 설정할수 없습니다." & vbcrlf & "상세내역 삭제 및 본견적을 가견적으로 전환하여 처리하십시오.","처리안내"
			Exit Sub
		End If
		
		vntInParams = array("","")
		vntRet = gShowModalWindow("PDCMESTTYPEPOP.aspx",vntInParams , 590,490)
		if isArray(vntRet) then
			'그리드 일단 초기화
			.sprSht.MaxRows = 0
			If .txtPREESTNO.value = "" Then
				strPREESTNO = "9999999999"
			Else
				strPREESTNO = .txtPREESTNO.value 
			End If
			.txtSUMAMT.value = trim(vntRet(3,0))
			.txtAMT.value = trim(vntRet(4,0))
			.txtSUSURATE.value  = trim(vntRet(5,0))
			.txtSUSUAMT.value = trim(vntRet(6,0))
			.txtCOMMITION.value = trim(vntRet(7,0))
			.txtNONCOMMITION.value = trim(vntRet(8,0))
			
			txtSUSUAMT_onblur
			txtSUMAMT_onblur
			txtAMT_onblur
			
			txtESTSUSUAMT_onblur
			txtESTSUMAMT_onblur
			txtESTAMT_onblur
			
			txtCOMMITION_onblur
			
			
			txtNONCOMMITION_onblur
			
			
		    '하단 Sheet 적용
		    mlngRowCnt=clng(0): mlngColCnt=clng(0)
		    vntData = mobjPDCOPREESTDTL.SelectRtn_ProcEST(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(vntRet(0,0)),strPREESTNO)
			
			IF not gDoErrorRtn ("SelectRtn_ProcEST") then
				'조회한 데이터를 바인딩
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
				'여기서 부터 Detail 버튼 설정
				
	
				For intCnt =1 To .sprSht.MaxRows 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",intCnt) = "Y"  Then
						'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",intCnt,intCnt,false
						'버튼형태로 변경
						If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
							mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",intCnt,intCnt,,false
						Else
							mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",intCnt,intCnt,,false
						End If
						
					Else
						'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
						'mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DETAIL_BTN|QTY|PRICE|AMT",Row,Row,false	
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",intCnt,intCnt,false
						'일반형태로 변경
						'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "DETAIL_BTN",Row,Row,255,,,,,False
						mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",intCnt,intCnt,0,,,,,,,,False
					End If
					sprSht_Change 1,intCnt
				Next
			End If
		end if
	End with
End Sub	


Sub imgBonSave_onclick
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 본견적으로 저장하기위해서는 일단 저장을 먼저 해야 합니다.","본견적저장안내!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	ExeProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'엑셀버튼
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	with frmThis
		'msgbox .txtPREESTNO.value & "--" & .txtPREESTNOVIEW.value
		
		'  "" 인경우는 다른견적 복사시 //  <>"" 인경우는 해당job의 견적 수정시
		if .txtPREESTNO.value <> "" then 
			if frmThis.txtENDFLAG.value = "T" Then
				gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 저장이 불가능 합니다.","저장안내!"
				Exit Sub
			End If
		end if 
	end with
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDel_onclick
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 변경하기 위해서는 일단 저장을 먼저 해야 합니다.","삭제안내!"
		exit sub
	else
		if frmThis.txtENDFLAG.value = "T" Then
			gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 행삭제가 불가능 합니다.","삭제안내!"
			Exit Sub
		End If
	end if 
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgCFInput_onclick
	Dim vntInParams
	Dim vntRet
	Dim strPREESTNO
	Dim i, SaveCHK  '저장 체크
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "저장이 되지 않은 데이터가 있습니다. 저장을 완료하신 후 진행하세요.! ","CF 내역 입력 안내.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 변경하기 위해서는 일단 저장을 먼저 해야 합니다. ","CF 내역 입력 안내.!"
		exit sub
	end if 
	
	with frmThis
	
		If .txtPREESTNO.value = "" Then
			strPREESTNO = "9999999999"
		Else
			strPREESTNO = .txtPREESTNO.value 
		End If
			
		vntInParams = array(strPREESTNO)
		vntRet = gShowModalWindow("PDCMJOBMST_CFINPUT.aspx",vntInParams , 1149,400)
	End With
End Sub

Sub imgTableUP_onclick
	Dim strRow
	Dim intCnt
	Dim i
	
	
	with frmThis
		
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				intCnt = intCnt + 1
			End if 
		next
		
		if intCnt > 1 then
			gErrormsgbox "데이터 이동시 한 데이터만 선택하셔야 합니다.","이동안내!"
			exit sub
		end if
			
			
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "선택된 데이터가 없습니다.","이동안내!"
			Exit Sub
		end if 
		if strRow = 1 then exit sub
		
		'자기자신을 넘긴다.
		sprSht_UpCopy strRow
	
	end with
End Sub

Sub imgTableDown_onclick
	Dim strRow
	Dim intCnt
	Dim i	
	
	with frmThis
		
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				intCnt = intCnt + 1
			End if 
		next
		
		if intCnt > 1 then
			gErrormsgbox "데이터 이동시 한 데이터만 선택하셔야 합니다.","이동안내!"
			exit sub
		end if	
	
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "선택된 데이터가 없습니다.","이동안내!"
			Exit Sub
		end if 
		
		if strRow = (.sprSht.MaxRows) then exit sub	
		sprSht_DownCopy strRow
	end with
End Sub

Sub sprSht_UpCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	Dim i
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row셋팅	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'msgbox strRow
		'돌면서 자신과 printseq만큼 위에꺼 복사
		for i=1 to .sprSht_copy.MaxRows
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow- strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow-strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow -strPRINT_SEQ)

			mobjSCGLSpr.CellChanged frmThis.sprSht_copy, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)			
			strPRINT_SEQ = strPRINT_SEQ -1

		next

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
			End if
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)	
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


Sub sprSht_DownCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	Dim i
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row셋팅	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'돌면서 자신과 printseq만큼 위에꺼 복사
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow+ strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow+strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow +strPRINT_SEQ)
			
			mobjSCGLSpr.CellChanged frmThis.sprSht_copy, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 		
		next
		.sprSht_copy.MaxRows = 0
	end with
End Sub


'견적서 출력
Sub imgPrintEstBasic_onclick
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,SaveCHK		'인쇄전 저장 체크
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim intRtn
	Dim strUSERID
	
	SaveCHK = 0
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "저장이 되지 않은 데이터가 있습니다. 저장을 완료하신 후 출력하세요.! ","출력 데이터 안내.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 출력하기 위해서는 일단 저장을 먼저 해야 합니다. ","출력 데이터 안내.!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp삭제 끝
		
		ModuleDir = "PD"
		
		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow)
		
		if cdate(.txtAGREEYEARMON.value) < cdate("2013-01-31") then
			ReportName = "ESTIMATE_ONE.rpt"
		else
			ReportName = "ESTIMATE_ONE_P.rpt"
		end if 
		
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		strUSERID = ""
		vntData = mobjPDCOPREESTDTL.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, 1, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		
		window.setTimeout "call printSetTimeout('" & strPREESTNO & "')", 10000
		
	end with
	gFlowWait meWAIT_OFF
End Sub

'견적서 출력
Sub imgPrintEst_onclick
	Dim ModuleDir 	    '사용할 모듈명
	Dim ReportName      '리포트 이름
	Dim Params		    '파라메터(VARCHAR2)
	Dim Opt             '미리보기 "A" : 미리보기, "B" : 출력
	Dim i,SaveCHK		'인쇄전 저장 체크
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim intRtn
	Dim strUSERID
	
	'체크된 데이터가 없다면 메시지를 뿌린후 Sub를 나간다
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "인쇄할 데이터가 없습니다.",""
		Exit Sub
	end if
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "저장이 되지 않은 데이터가 있습니다. 저장을 완료하신 후 출력하세요.! ","출력 데이터 안내.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 출력하기 위해서는 일단 저장을 먼저 해야 합니다. ","출력 데이터 안내.!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'인쇄버튼을 클릭하기 전에 md_trans_temp테이블에 내용을 삭제한다
		'인쇄후에 temp테이블을 삭제하게 되면 크리스탈 리포트뷰어에 파라메터 값이 넘어가기전에
		'데이터가 삭제되므로 파라메터가 넘어가지 않는다. by kty
		'md_trans_temp삭제 시작
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)

		'md_trans_temp삭제 끝
		
		ModuleDir = "PD"
		
		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow)
		
		IF .cmbESTTYPE.value = 1 THEN
			IF .txtPREESTGBN.value = "가견적" THEN
				if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
					ReportName = "ESTIMATE.rpt"
				else
					ReportName = "ESTIMATE_P.rpt"
				end if
			ELSE
				if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
					ReportName = "ESTIMATE_ESTBACK.rpt"
				else
					ReportName = "ESTIMATE_ESTBACK_P.rpt"
				end if
				
			END IF 
		ELSEIF .cmbESTTYPE.value = 2 THEN
			IF .txtPREESTGBN.value = "가견적" THEN
				gErrorMsgBox "본견적만 출력 가능합니다.","인쇄안내!"
				EXIT SUB
			END IF 
			
			if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
				ReportName = "ESTIMATE_ACTUAL.rpt"
			else
				ReportName = "ESTIMATE_ACTUAL_P.rpt"
			end if
			
		ELSE
			IF .txtPREESTGBN.value = "가견적" THEN
				gErrorMsgBox "본견적만 출력 가능합니다.","인쇄안내!"
				EXIT SUB
			END IF
			
			if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
				ReportName = "ACTUAL.rpt"
			else
				ReportName = "ACTUAL_P.rpt"
			end if
			
		END IF

		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		strUSERID = ""
		vntData = mobjPDCOPREESTDTL.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, 1, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10초후에 printSetTimeout 펑션을 호출하여 temp테이블을 삭제한다.
		'출력화면이 뜨는 속도보다 삭제하는 속도가 빨라서 밑에서 바로 삭제가 안되기때문에 시간을 임의로 줌..
		
		window.setTimeout "call printSetTimeout('" & strPREESTNO & "')", 10000
		
	end with
	gFlowWait meWAIT_OFF
End Sub


'출력이 완료된후 md_trans_temp(다중출력을 위한 임시테이블)을 지운다
Sub printSetTimeout(strPREESTNO)
	Dim intRtn, intRtn2
	with frmThis
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)
		'intRtn2 = mobjPDCOPREESTDTL.DeleteRtnUpdate_ATTR07(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

Sub imgCalEndarAGREE_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtAGREEYEARMON,frmThis.imgCalEndarAGREE,"txtAGREEYEARMON_onchange()"
		gSetChange
	end with
End Sub

Sub imgimgCalEndarCREDAY_onclick
	WITH frmThis
		'CalEndar를 화면에 표시
		gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgimgCalEndarCREDAY,"txtPRINTDAY_onchange()"
		gSetChange
	end with
End Sub

'행추가시 참고사항=================================================================================================
'본견적인경우는 txtEndflag 즉 청구정산이 있는 경우 추가불가
'가견적 및 신규입력시는 바로 추가 가능 << 가견적 저장후는 해당 JOB 의 청구진행을 조회하여 endflag 를 다시 가져오니,
'본견적으로 저장 한다면 frmThis.txtENDFLAG.value = "T" 문장을 제대로 수행 할것이다,
'==================================================================================================================

Sub imgRowAdd_onclick ()
	'행추가는 언제나 가능 가견적이므로 본견적으로 확정시에만 판단하자... 가견적은 언제든지 만들어질수 있음을 명시 하자!!
	'본견적이라면
	if mstrSELECT = "F" then
		gErrorMsgBox "다른 JOBNO 의 내역을 변경하기 위해서는 일단 저장을 먼저 해야 합니다.","추가안내!"
		exit sub
	end if 
	
	If frmThis.txtPREESTGBN.value = "본견적" And frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 행추가가 불가능 합니다.","처리안내!"
		Exit Sub
	End IF
	
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DETAIL_BTN")  Then
			If frmThis.txtPREESTGBN.value = "본견적" Then
				If frmThis.txtENDFLAG.value <> "T" Then
					intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
					DefaultValue
				end if
			Else
				intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
				DefaultValue
			End If
		End If
	Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select
	End If
End Sub



Sub DefaultValue
	Dim i
	Dim imeseqCHECK
	Dim highCnt, cnt1, cnt2
	Dim highvalue
	
	imeseqCHECK = true
	highCnt = 0	
	cnt2 = 0
	
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",.sprSht.ActiveRow, "N"
		If .sprSht.MaxRows = 1 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow,1
			mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow,1
		Else
			'IMESEQ 의 경우 나중에 ITEMCODESEQ 와 값이 같아져야한다. 테이블값의 최대값으로 생성한다.
			for i = 1  to .sprSht.MaxRows
				cnt1 = 0
				if mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i) <> "" then 
					cnt1 = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i))
				end if
				if cnt2 < cnt1 then 
					highcnt = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i) + 1)
					cnt2 = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i))
				end if 
			next
				highvalue = cstr(highcnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow, highvalue
			
			mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"PRINT_SEQ",.sprSht.MaxRows -1)+1
		End If
	End With
End Sub


'-----------------------------------------
'시트 에서 키를 떼었을때 선택 금액 합산. 
'-----------------------------------------
Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCOLUMN
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	'키 움직일때 바인딩

	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") or  _
																		   mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) or _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT")) Then
				
				
				
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
END SUB

'-----------------------------------
'시트에서 마우스를 떼었늘때 이벤트
'-----------------------------------
Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
	
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") or _
																			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT")  Then
																			
				
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
				
				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					End If
				Next
				
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if	
		
	End With
End Sub

'-----------------------------
' UI
'-----------------------------	
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
		
		If .txtSUSURATE.value = 0 Then
			.txtSUSUAMT.readOnly = false
		Else
			.txtSUSUAMT.readOnly = true
		End If
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtSUSURATE_onblur
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		
	End with
End Sub
Sub CleanFieldRate
	with frmThis
		.txtSUSURATE.value = 0
	End With
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
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
		
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
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
Sub txtAMT_onfocus
	with frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
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

Sub txtESTSUMAMT_onfocus
	with frmThis
		.txtESTSUMAMT.value = Replace(.txtESTSUMAMT.value,",","")
	end with
End Sub
Sub txtESTSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTSUMAMT,0,true)
	end with
End Sub

Sub txtESTSUSUAMT_onfocus
	with frmThis
		.txtESTSUSUAMT.value = Replace(.txtESTSUSUAMT.value,",","")
	end with
End Sub
Sub txtESTSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTSUSUAMT,0,true)
	end with
End Sub

Sub txtESTSUSURATE_onblur
	with frmThis
		If .txtESTSUSURATE.value = "" Then
			.txtESTSUSURATE.value = 0
		End If
		
	End with
End Sub




Sub txtAGREEYEARMON_onchange
	gSetChange
End Sub


Sub txtPRINTDAY_onchange
	gSetChange
End Sub

Sub txtSUSURATE_onchange
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM

		gSetChangeFlag .txtSUSURATE  
	End with
End Sub
Sub txtESTSUSURATE_onchange
	with frmThis
		If .txtESTSUSURATE.value = "" Then
			.txtESTSUSURATE.value = 0
		End If
		ESTSUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtESTSUSURATE  
	End with
End Sub


Sub txtSUSUAMT_onchange
	with frmThis
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSUAMT  
	End with
End Sub

Sub txtESTSUSUAMT_onchange
	with frmThis
		ESTSUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSUAMT  
	End with
End Sub





'-----------------------------------
' SpreadSheet 이벤트
'-----------------------------------
'스프레드의 모든 항목의 값이 변경 될때 발생 하는 이벤트 입니다.
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, strCols
	Dim strCode, strCodeName
	Dim strQTY, strPRICE, strAMT
	Dim lngPrice
	Dim lngVALUE
	Dim lngVALUE1
	Dim lngVALUE2

	with frmThis
				'Long Type의 ByRef 변수의 초기화
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				strCode = ""
				strCodeName = ""
				IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
					strCode = ""
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",.sprSht.ActiveRow)
					vntData = mobjPDCOGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",strCodeName)
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntData(4,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntData(7,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntData(8,1)
						
						'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
						If vntData(7,1) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
						
							'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
							mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
							'수량,단가,금액은 0처리
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
							'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
							If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
							Else
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
							End If
						Else
							'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
							'일반형태로 변경
							mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
						End If	
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						
						'SUSUAMT_CHANGEVALUE2
						IF .txtPREESTGBN.value ="가견적" THEN
							Call ESTSUSUAMT_CHANGEVALUE2
						ELSE	
							Call SUSUAMT_CHANGEVALUE2
						END IF
						
						BUDGET_AMT_SUM
						
						
						If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
							mstrFIRSTPRODUCTIONCHECK = "Y"
						
							call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
						end if 
					Else
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End If
					.txtSUSURATE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
					.sprSht.Focus	
					mobjSCGLSpr.ActiveCell .sprSht, Col, Row
				'수량로직	
				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strPRICE <> "" And strAMT = "" Then
   						lngVALUE = strQTY * strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE
   					ElseIf strPRICE = "" And strAMT <> "" Then
   						lngVALUE1 = gRound(strAMT/strQTY,0)
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngVALUE1
   					ElseIf strPRICE <> "" And strAMT <> "" Then
   						lngVALUE2 = strQTY * strPRICE
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE2
   					End IF
   					
   					IF .txtPREESTGBN.value ="가견적" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
					
   					BUDGET_AMT_SUM
   				'단가 로직
   				ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
   					strQTY		= mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",.sprSht.ActiveRow)
					strPRICE   = mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",.sprSht.ActiveRow)
					strAMT = strQTY * strPRICE
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT	
					
					'Call SUSUAMT_CHANGEVALUE(Row)
					IF .txtPREESTGBN.value ="가견적" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
					BUDGET_AMT_SUM
				'금액로직	
   				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strAMT = 0 Then
   						mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, strAMT
   					Else 
   						If strQTY <> 0  Then
   							lngPrice = gRound(strAMT/strQTY,0)
   							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngPrice
   						End IF
   					End IF
   					'Call SUSUAMT_CHANGEVALUE(Row)
   					IF .txtPREESTGBN.value ="가견적" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
   					BUDGET_AMT_SUM
   				Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMIFLAG") Then
   					'Call SUSUAMT_CHANGEVALUE2
   					IF .txtPREESTGBN.value ="가견적" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
   					BUDGET_AMT_SUM
				END IF
	end with
	'변경 플래그 설정
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub SUSUAMT_CHANGEVALUE(ByVal Row)
	Dim strAMT,strCOMMIFLAG
	with frmThis
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		if strCOMMIFLAG = "1" Then
			'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), Row
	End with
End SUb

Sub ESTSUSUAMT_CHANGEVALUE(ByVal Row)
	Dim strAMT,strCOMMIFLAG
	with frmThis
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		if strCOMMIFLAG = "1" Then
			'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * .txtESTSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), Row
	End with
End SUb


Sub SUSUAMT_CHANGEVALUE2
Dim intCnt
Dim strAMT,strCOMMIFLAG
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		if strCOMMIFLAG = "1" Then
		'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), intCnt
	Next
	
	End with
End Sub

Sub ESTSUSUAMT_CHANGEVALUE2
	Dim intCnt
	Dim strAMT,strCOMMIFLAG
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		if strCOMMIFLAG = "1" Then
		'수수료율 설정시 .txtSUSURATE 가 Null 일 경우 오류
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * .txtESTSUSURATE.value /100),0)
		
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), intCnt
	Next
	
	End with
End Sub


'시트에 금액을 합산한 값을 합계시트에 뿌려준다.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	
	With frmThis
	
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		
		If .sprSht.MaxRows > 0 Then
			.txtSUMAMT_TOTAL.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT_TOTAL,0,True)
			
		else
			.txtSUMAMT_TOTAL.value = 0
		End If
		
	End With
End Sub


Sub BUDGET_AMT_SUM
	'총합계 변수
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM, intAMTSUB
	Dim lngSUSU
	'수수료 계산 변수
	Dim intCnt,intSUSU,intSUSUSUM 
	'commition 계산 변수
	Dim intCnt1,intCOM,intCOMSUM 
	'noncommition 계산변수
	Dim intCnt2,intNON,intNONSUM 
	
	Dim intESTSUSU , intESTSUSUSUM
	Dim IntESTAMT,IntESTAMTSUM, intESTAMTSUB
	
	with frmThis
	
		IntAMTSUM = 0
		IntPRICESUM = 0
		intSUSU = 0
		intSUSUSUM = 0
		intCOM = 0
		intCOMSUM = 0
		intNON = 0
		intNONSUM = 0
		intAMTSUB = 0
		
		intESTSUSU = 0
		intESTSUSUSUM = 0
		IntESTAMT = 0
		IntESTAMTSUM = 0
		intESTAMTSUB = 0
	
		IF .txtPREESTGBN.value = "가견적" THEN
			
			For intCnt = 1 To .sprSht.MaxRows
				intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
				intSUSUSUM = intSUSUSUM + intSUSU
			Next
			
			'If .txtESTSUSURATE.value <> 0 Then
				.txtESTSUSUAMT.value = intSUSUSUM
			'End If
			
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			intAMTSUB = IntAMTSUM
		
			If .txtESTSUSURATE.value <> 0 Then
				IntAMTSUM = IntAMTSUM + intSUSUSUM
			Else
				IntAMTSUM = IntAMTSUM + replace(.txtESTSUSUAMT.value,",","")
			End If
			
			.txtESTAMT.value = intAMTSUB
			.txtESTSUMAMT.value = IntAMTSUM
		ELSE
			
			For intCnt = 1 To .sprSht.MaxRows
				intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
				intSUSUSUM = intSUSUSUM + intSUSU
			Next
			
			'If .txtSUSURATE.value <> 0 Then
				.txtSUSUAMT.value = intSUSUSUM
			'End If
			
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			intAMTSUB = IntAMTSUM
		
			If .txtSUSURATE.value <> 0 Then
				IntAMTSUM = IntAMTSUM + intSUSUSUM
			Else
				IntAMTSUM = IntAMTSUM + replace(.txtSUSUAMT.value,",","")
			End If
			
			.txtAMT.value = intAMTSUB
			.txtSUMAMT.value = IntAMTSUM
		END IF 
		
		
		For intCnt1 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG", intCnt1) = "1" Then
				
				intCOM = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intCOMSUM = intCOMSUM + intCOM
			Else
				intNON = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intNONSUM = intNONSUM + intNON
			end if
		Next
		.txtCOMMITION.value = intCOMSUM
		.txtNONCOMMITION.value = intNONSUM
		
		txtSUSUAMT_onblur
		txtSUMAMT_onblur
		txtAMT_onblur
		
		txtESTSUSUAMT_onblur
		txtESTSUMAMT_onblur
		txtESTAMT_onblur
		
		txtCOMMITION_onblur
		txtNONCOMMITION_onblur
		
	End With
End Sub

'=================================
'스프레드의 행을 클릭 시 발생
'=================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	end With
End Sub

'=================================
'스프레드의 행을 더블 클릭 시 발생
'=================================
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'=========================================================
'스프레드의 항목이 변할시 어떠한 함수를 태우고자 할때 사용
'=========================================================
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding( sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)	
				
				'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
				If vntRet(7,0) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
				
					'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
					'수량,단가,금액은 0처리
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
					'일반형태로 변경
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If	
				'mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
	
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				IF .txtPREESTGBN.value ="가견적" THEN
					ESTSUSUAMT_CHANGEVALUE2
				ELSE	
					SUSUAMT_CHANGEVALUE2
				END IF 
				
				BUDGET_AMT_SUM
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
					mstrFIRSTPRODUCTIONCHECK = "Y"
					call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
				end if 
			End IF
			.txtSUSURATE.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row +1
		end if
	End With
End Sub

'=================================================
'스프레드 내 버튼을 클릭 하였을때 발생 하는 이벤트
'=================================================
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strPREESTNO
	Dim strGBN
	Dim strITEMCODE
	Dim dblAMT, dblESTAMT
	Dim vntData, vntData_Temp
	Dim strRTN , strRTN2
	Dim intCnt,intRtn
	Dim strChk
	Dim intTempChkCnt
	Dim dblSUSUAMT  , dblESTSUSUAMT
	Dim intColorCnt
	Dim strITEMCODESEQ
	Dim returnPOP		'상세내역 팝업 리턴 값 받는 변수
	Dim indirecPOP		'간접비 팝업 리턴 값 변수
	Dim processData		'간접비 임시 저장 변수
	Dim intProductionCnt'간접비 로우수 [2개 이상을 막는다.]
	
	Dim strOldITEMCODE
	with frmThis
	    '외주항목코드 선택 버튼 부분
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			strOldITEMCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row)
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)
				'외주항목상세구분이 "Y" 이고, JOB 구분이 전파일때만 상세견적 버튼이 출현
				
				If vntRet(7,0) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
					'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
					'수량,단가,금액은 0처리
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'버튼형태로 변경(간접비와 부문 입력으로 분기처리!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
					'일반형태로 변경
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				IF .txtPREESTGBN.value ="가견적" THEN
					ESTSUSUAMT_CHANGEVALUE2
				ELSE	
					SUSUAMT_CHANGEVALUE2
				END IF
				BUDGET_AMT_SUM
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
					mstrFIRSTPRODUCTIONCHECK = "Y"
					call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
				end if 
			End IF
			.txtSUSURATE.focus()	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row+1
			
		'CF견적의 부문을 선택시 버튼처리....	
		'상세견적 저장 및 조회 팝업 호출
'------------------------------------------------------------------------------------------------------------			
'***********************1.   간접비 부문 처리와 상세내역 처리로 나눠짐.
'------------------------------------------------------------------------------------------------------------
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DETAIL_BTN") Then
			'프로덕션 제작 수수료 부분
			If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then 
			
				'다른 JOBNO 의 견적을 그대로 복사 하고싶은 경우 일단 저장을 유도한다.
				if mstrSELECT = "F" then
					gErrorMsgBox "다른 JOBNO 의 내역을 변경하기 위해서는 일단 저장을 먼저 해야 합니다.","간접비 변경 안내!"
					exit sub
				end if 
				
				intProductionCnt = 0
				'강제로 쉬트의 값을 가져온다.
				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) <> "242001" Then
						If mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",intCnt) = "T" Then
							mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",intCnt, "F"
						Else
							mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",intCnt, "T"
						End If
						mobjSCGLSpr.CellChanged frmThis.sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"INDIRECFLAG"), intCnt
					ELSE
						'간접비 로우체크
						intProductionCnt = intProductionCnt + 1
					End If
				Next
				
				if intProductionCnt > 1 then
					gErrorMsgBox "이미 다른 간접비 데이터가 있습니다. 간접비는 한가지만 입력 가능합니다." & vbcrlf & "데이터를 확인하세요.!" ,"간접비 안내"
					mobjSCGLSpr.DeleteRow .sprSht,Row
					exit sub
				end if 
				'indirectflag =  모든 행을 가져오기위해,,, For 문의 Cellchange 를 태운다.
				'PD_PRODUCTIONCOMMI_INPUT 테이블 처리사항============================================================
				'견적과 간접비 팝업의 경우 팝업에서 일어나는 모든 처리는 INPUT테이블에서 해결된다.
				'저장 삭제 수정이 모두 INPUT 에서 이루어지며 메인 화면의 저장이 실제 모든 변경의 저장 시점이 된다.
				'따라서 견적을 입력하다 조회를 하거나 다른 작업을 할경우 데이터가 사라지며 저장 하지 않은 데이터는 영향을 미치지 않는다.
				'====================================================================================================
				strPREESTNO = .txtPREESTNO.value 
				
				strChk = 0
				For intTempChkCnt=1 To .sprSht.MaxRows
					'프로덕션 간접비 를 제외한 다른 견적이 실제 저장되어 있지 않으면 [이말은 즉 견적번호가 없다면 간접비를 입력할수 없다는 뜻..]
					If  (mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",intTempChkCnt) = "" and _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intTempChkCnt) <> "242001") or _
													 (mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intTempChkCnt) <> "242001" and _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",intTempChkCnt) = "Y" ) Then 
						strChk = strChk +1
					End If
				Next
				
				If strChk <> 0 Then
					if mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) = "" THEN
						gErrorMsgBox "견적이 저장되지않은 노란색 행을 확인하거나 견적을 저장 하십시오." & vbcrlf & "모든행이 저장되어야 간접비를 입력하실수 있습니다.","처리안내"
						mobjSCGLSpr.DeleteRow .sprSht,Row
					END IF
					'저장이 되어있는 데이터의경우 메시지 박스를 띄우는데 행은 삭제 하지 않음
					gErrorMsgBox "견적이 저장되지않은 노란색 행을 확인하거나 견적을 저장 하십시오." & vbcrlf & "모든행이 저장되어야 간접비를 입력하실수 있습니다.","처리안내"
					EXIT SUB
				Else
				
					'시트의 변경된 데이터를 가져온다.
					vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | ITEMCODE | AMT | PRODUCTIONCOMMISSION | PRINT_SEQ")
				
					if IsArray(vntData) then
						processData = vntData
					else 
						gErrorMsgBox "처리할 데이터가 없습니다." & vbcrlf & "간접비생성은 간접비 외 견적데이터가 필요합니다.","처리안내!"
						Exit Sub
					end if
					
					dblAMT = 0
					dblAMT = mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",Row)
					mstrFIRSTPRODUCTIONCHECK = "Y"
					
					vntInParams = array(strPREESTNO,Trim(.txtENDFLAG.value), mstrFIRSTPRODUCTIONCHECK,.txtPREESTGBN.value, processData)
					
					vntRet = gShowModalWindow("PDCMJOBMST_INDIRECTCOST.aspx",vntInParams , 1149,650)
					
					indirecPOP = vntRet(0)
					
					'팝업창에서 간접비가 없다면 메인 페이지에 영향을 주지 않는다.
					if indirecPOP <> "False" then
						'해당 팝업의 결과값 데이터를 추출하여 투입이전값과 다르다면 견적상세 내역을 저장 할수 있도록 유도한다.
						mlngTempRowCnt=clng(0): mlngTempColCnt=clng(0)
						vntData = mobjPDCOPREESTDTL.SelectRtn_CommiHDRSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row))			

						If mlngTempRowCnt > 0 Then
							strRTN = Cstr(vntData(0,1)) '가견적
							strRTN2 = Cstr(vntData(1,1)) '가견적
						Else 
							strRTN = "0"
							strRTN2 = "0"
						End If

						IF .txtPREESTGBN.value = "가견적" THEN
							If strRTN <> Cstr(indirecPOP) Then
								'값을 받아오는 것이 아니라, PD_SUBITEM_DTL 또는 PD_SUBITEM_INPUT 에서 가져온다.
								mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
								If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
									dblSUSUAMT = indirecPOP * .txtESTSUSURATE.value * 0.01
									mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
								End If
								
								BUDGET_AMT_SUM
								
								'값이 변경되었음을 표기
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							End IF
						ELSE
							If strRTN2	<> Cstr(indirecPOP) Then
								'값을 받아오는 것이 아니라, PD_SUBITEM_DTL 또는 PD_SUBITEM_INPUT 에서 가져온다.
								mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
								If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
									dblSUSUAMT = indirecPOP * .txtSUSURATE.value * 0.01
									mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
								End If
								BUDGET_AMT_SUM
								
								'값이 변경되었음을 표기
								If vntRet(1) = "T" then
									mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
								else 
									mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "N"
								end if  
							End IF
						END IF 
					
						mobjSCGLSpr.CellChanged .sprSht, Col,Row
						
						For intColorCnt = 1 To .sprSht.MaxRows
							If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
							Else
								If intColorCnt Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
								End If
							End if
						Next
					end if
				End If	
'------------------------------------------------------------------------------------------------------------			
'***********************2.   TV-CF 상세 내역 처리 
'------------------------------------------------------------------------------------------------------------
			Else
				strRTN = 0
				strRTN2 = 0
				dblAMT = 0
				returnPOP = 0
				
				
				'다른 JOBNO 의 견적을 그대로 복사 하고싶은 경우 일단 저장을 유도한다.
				
				if mstrSELECT = "F" then
					gErrorMsgBox "다른 JOBNO 의 내역을 변경하기 위해서는 일단 저장을 먼저 해야 합니다.","상세 견적 안내!"
					exit sub
				end if 
				
				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
						If mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",intCnt) = "Y" Then
							gErrorMsgBox "간접비가 수정 되어있습니다. 저장하신후 상세 견적을 입력하세요.","상세 견적 안내!"
							exit sub
						end if
					End If
				Next
				
				If .txtPREESTNO.value = "" Then
					strPREESTNO = "9999999999"
					
				Else
					strPREESTNO = .txtPREESTNO.value 
				End If	
				
				If .txtPREESTGBN.value = "가견적" Then
					strGBN="F"
				ElseIf .txtPREESTGBN.value = "본견적" Then
					strGBN = "T"
				End If
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) = "" Then
				'ITEMCODESEQ 가 없다면[저장되지 않고 상세내역부터 입력한 값이라면] IMESEQ 를 가져온다.[추후에 ITEMCODESEQ 와 IMESEQ 는 값이 같아진다.!]
					strITEMCODESEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row)
				Else
					strITEMCODESEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row)
				End If
				
				dblAMT = mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",Row)
				vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",Row), _
									strPREESTNO,strGBN,strITEMCODESEQ,Trim(.txtENDFLAG.value), _
									.txtJOBNO.value)
				
				vntRet = gShowModalWindow("PDCMJOBMST_SUBITEM.aspx",vntInParams , 1149,650)
			
				'팝업에서 수정되거나 저장된값
				returnPOP = vntRet(0)
				
				'팝업에서 저장이 이루어져 input에 값이 있다면 값을 아니면 false 를 반환	
				if returnPOP <> "False" then		
					
					mlngTempRowCnt=clng(0): mlngTempColCnt=clng(0)
					vntData = mobjPDCOPREESTDTL.SelectRtn_DtlSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt, _
																.txtJOBNO.value, _
																mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row), _
																mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row))
					
					If mlngTempRowCnt > 0 Then
						strRTN = Cstr(vntData(0,1)) 
						strRTN2 = Cstr(vntData(1,1)) 
					Else 
						strRTN = "0"
						strRTN2 = "0"
					End If
		
					IF .txtPREESTGBN.value = "가견적" THEN
						If strRTN <> Cstr(returnPOP) Then
							'값을 받아오는 것이 아니라, PD_SUBITEM_DTL 또는 PD_SUBITEM_INPUT 에서 가져온다.
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtESTSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							'값이 변경되었음을 표기
							mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
						'값이 같다면 같은 값을 뿌려줘도 상관 없다 다만 팝업에서 변경했다가 다시 원래대로 돌릴경우 
						else 
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtESTSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'금액이 같더라고 상세 내역이 다를수 있기때문에 팝업에서 이벤트가 있었다면 저장을 유도한다.
							If vntRet(1) = "T" then
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							end if  
						End IF
					ELSE
						If strRTN2	<> Cstr(returnPOP) Then
							'값을 받아오는 것이 아니라, PD_SUBITEM_DTL 또는 PD_SUBITEM_INPUT 에서 가져온다.
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'값이 변경되었음을 표기
							mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
						
						else 
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'금액이 같더라고 상세 내역이 다를수 있기때문에 팝업에서 이벤트가 있었다면 저장을 유도한다.							
							If vntRet(1) = "T" then
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							end if 
						End IF
					END IF 
				else
					'팝업의 값이 false 이며[input 에 값이 없으며] [CHANGEFLAG = "T" 저장이나 삭제를 했다면.](전체 삭제일경우임)
					IF vntRet(1) = "T" then
						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
						mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
						If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
							mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, 0
						End If
						
						BUDGET_AMT_SUM
						
						'값이 변경되었음을 표기
						mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"					
					end if 
				End if 
					
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next
				
				mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SAVEFLAG"),Row
			End If
		end if
	End with
End Sub

'-----------------------------
' 페이지 화면 디자인 및 초기화 
'-----------------------------	
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	Dim strMSG
	Dim strPREESTNO
	
	'서버업무객체 생성	
	set mobjPDCOPREESTDTL = gCreateRemoteObject("cPDCO.ccPDCOPREESTDTL")
	set mobjPDCOGET		  = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet 기본Color 지정
	gSetSheetDefaultColor()
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 24, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | PRINT_SEQ | PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | DETAIL_BTN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION"
		mobjSCGLSpr.SetHeader .sprSht,		  "선택|인쇄|가견적번호|순번|대분류|중분류|견적항목코드|견적항목명|견적명|내역|커미션|수량|단가|금액|수수료금액|저장구분|상세견적|상세견적여부|가짜순번|상세저장여부|상세부분여부|다이렉트플레그|간접비"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|   4|        10|   4|     8|    12|        8 |2|        15|12    |  18|     6|  6|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     "
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRINT_SEQ | QTY | PRICE | AMT | SUSUAMT | PRODUCTIONCOMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY", -1, -1, 1
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | ITEMCODENAME | FAKENAME | STD | GBN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PRINT_SEQ | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | PREESTNO | DETAIL_BTN | IMESEQ | SAVEFLAG | DETAILYNFLAG"
		mobjSCGLSpr.ColHidden .sprSht, "GBN | SUBDETAIL | INDIRECFLAG | PRODUCTIONCOMMISSION | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION | PREESTNO", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE | ITEMCODESEQ",-1,-1,2,2,false


		'견적 순서를 조정을 위한 그리드
		gSetSheetColor mobjSCGLSpr, .sprSht_copy
		mobjSCGLSpr.SpreadLayout .sprSht_copy, 24, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_copy, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_copy, "CHK | PRINT_SEQ | PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | DETAIL_BTN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION"
		mobjSCGLSpr.SetHeader .sprSht_copy,		  "선택|인쇄|가견적번호|순번|대분류|중분류|견적항목코드|견적항목명|견적명|내역|커미션|수량|단가|금액|수수료금액|저장구분|상세견적|상세견적여부|가짜순번|상세저장여부|상세부분여부|다이렉트플레그|간접비"
		mobjSCGLSpr.SetColWidth .sprSht_copy, "-1","   4|   4|        10|   4|     8|    12|        8 |2|        15|12    |  18|     6|  6|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     "
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_copy,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_copy, "CHK | COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "PRINT_SEQ | QTY | PRICE | AMT | SUSUAMT | PRODUCTIONCOMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "QTY", -1, -1, 1
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_copy, "PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | ITEMCODENAME | FAKENAME | STD | GBN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht_copy, true, "ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | PREESTNO | DETAIL_BTN | IMESEQ | SAVEFLAG | DETAILYNFLAG"
		mobjSCGLSpr.ColHidden .sprSht_copy, "GBN | SUBDETAIL | INDIRECFLAG | PRODUCTIONCOMMISSION | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION | PREESTNO", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "ITEMCODE | ITEMCODESEQ",-1,-1,2,2,false	    

		.sprSht.style.visibility  = "visible"
		.sprSht_copy.style.visibility  = "visible"

		InitPageData
		
		'부모창의 데이터 가져오기
		.txtJOBNO.value =  parent. document.forms("frmThis").txtJOBNO.value 
		
		
		strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value
		.txtPREESTNO.value = strPREESTNO

		'견적이 없는 경우 strPREESTNO 부분에 .txtPREESTNO.value를 그냥 조건으로 넣으면 될거 같지만
		'필드에 입력되는거보다 아래 조건을 더빨리 타는 것 같다.
		'그래서 부모창의 PREESTNO를 변수로 받아서 변수로 조건을 넣음...KTY 20110504
		'견적이 없다면
		if strPREESTNO = "" Then
			document.getElementById("strMsgBox").innerHTML = "- 견적내역이 없습니다."
			window.setTimeout "call NewDataSet()",1000 
		Else
			'견적이 있다면
			SelectRtn

			'견적이 있으나, 본견적이 없는 경우와 본견적이 있는 경우 청구와 정산내역을 조회함.
			if .txtSETCONFIRMFLAG.value = "F" Then
				strMSG = "- 본견적이 없습니다."
			Else
				if .txtENDFLAG.value = "T" AND .txtENDFLAGEXE.value = "T" Then
					strMSG = "- 청구,정산 내역이 있습니다"
				Elseif .txtENDFLAG.value = "T" AND .txtENDFLAGEXE.value = "F" Then
					strMSG = "- 청구 내역이 있습니다"
				Elseif .txtENDFLAG.value = "F" AND .txtENDFLAGEXE.value = "T" Then
					strMSG = "- 정산 내역이 있습니다"
				Elseif .txtENDFLAG.value = "F" AND .txtENDFLAGEXE.value = "F" Then
					strMSG = "- 청구,정산 내역이 없습니다."
				End If
			End IF
			document.getElementById("strMsgBox").innerHTML = strMSG
		End If
	End With
End Sub

Sub EndPage()
	set mobjPDCOPREESTDTL = Nothing
	set mobjPDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------
' 화면의 초기상태 데이터 설정
'-----------------------------	
Sub InitPageData
	'모든 데이터 클리어
	'gClearAllObject frmThis
	with frmThis

		.sprSht.MaxRows = 0
		.sprSht_copy.MaxRows = 0
		.txtPRINTDAY.value = gNowDate
	End with
	'새로운 XML 바인딩을 생성
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'주의사항 현재 본견적이 청구상태 라면 본견적으로 저장 될수 없음. - 본견적으로저장 버튼에 임의 조회를 하여 해당 JOBNO 가 청구되어 있는지를 확인 해야 함....
Sub NewDataSet
	with frmThis
		.txtAGREEYEARMON.value = gNowDate
		.txtPREESTGBN.value ="가견적"
		.txtPREESTNO.value = ""
		.txtMEMO.value = ""
		.txtPREESTNAME.value = ""
		.txtAMT.value = 0
		.txtCOMMITION.value = 0
		.txtNONCOMMITION.value = 0
		.txtSUSUAMT.value = 0
		.txtSUMAMT.value = 0
		.txtSUSURATE.value = 10
		.txtESTAMT.value = 0
		.txtESTSUMAMT.value = 0
		.txtESTSUSUAMT.value = 0
		.txtESTSUSURATE.value = 10
		'기존의 광고주,팀,브랜드를 복사한다.
		.txtCLIENTCODE.value = parent.document.forms("frmThis").txtCLIENTCODE.value
		.txtTIMCODE.value = parent.document.forms("frmThis").txtTIMCODE.value
		.txtSUBSEQ.value  = parent.document.forms("frmThis").txtSUBSEQ.value

		.txtJOBNAME.value  = parent.document.forms("frmThis").txtPRIJOBNAME.value
		.txtJOBNO.value  = parent.document.forms("frmThis").txtJOBNO.value 
		.txtPREESTNAME.value = parent.document.forms("frmThis").txtPRIJOBNAME.value
		txtSUSUAMT_onfocus
	End With
End Sub

'------------------------------------------
' 데이터 조회
'------------------------------------------
Sub SelectRtn ()
	window.setTimeout "call SelectRtn_Real()", 1000
End Sub

Sub SelectRtn_Real()
	Dim strCODE
	Dim strCHK
	With frmThis
		
		strCODE = parent.document.forms("frmThis").txtPREESTNO.value
		
		IF not SelectRtn_Head (strCODE) Then Exit Sub
	
		'쉬트 조회
		CALL SelectRtn_Detail (strCODE)
		
		txtSUSUAMT_onblur
		txtSUMAMT_onblur
		txtAMT_onblur
		
		txtESTSUSUAMT_onblur
		txtESTSUMAMT_onblur
		txtESTAMT_onblur
		
		txtCOMMITION_onblur
		txtNONCOMMITION_onblur
	
		'거래명세서 청구건에 따라 견적명을 수정저장 할수 있다.
		If .txtENDFLAG.value = "T" AND .txtPREESTGBN.value = "본견적" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
			.txtAGREEYEARMON.className = "NOINPUT"
			.txtAGREEYEARMON.readOnly = true
			.imgCalEndarAGREE.disabled = true
	
			
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = true
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = true
			
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			
		ElseIF  .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "본견적" Then
			
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtSUSUAMT.className = "INPUT_R"
			.txtSUSUAMT.readOnly = FALSE
			.txtSUSURATE.className = "INPUT_R"
			.txtSUSURATE.readOnly = FALSE
			
			
		ELSEIF .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "가견적" Then
			
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = TRUE
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = TRUE
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtESTSUSUAMT.className = "INPUT_R"
			.txtESTSUSUAMT.readOnly =  FALSE
			.txtESTSUSURATE.className = "INPUT_R"
			.txtESTSUSURATE.readOnly = FALSE
			
		End If
			
		.txtSELECTAMT.value = 0
		.txtSUMAMT_TOTAL.value = 0
		AMT_SUM
	End with
	
	'견적리스트에서 원래값이 아닌 다른값을 불러올경우[복사할려고] F 가 들어있다.
	mstrSELECT = parent. document.forms("frmThis").txtSELECT.value 
	
	
	
	If mstrSELECT = "F" Then
		Est_Copy
	End If
	
	CFINPUT_VISIBLE
		
End SUb

'견적 리스트에서 다른견적을 가져올경우.....
Sub Est_Copy
	Dim intCnt
	Dim intRtn
	Dim intSaveRtn
	Dim strOldCode
	Dim strJOBNO
	
	with frmThis		
		'현재 남아있는 템프값 무조건 삭제
		intRtn = mobjPDCOPREESTDTL.ProcessRtn_TempDelete(gstrConfigXml)
				
		if gDoErrorRtn ("ProcessRtn_TempDelete") then
			gErrorMsgBox "최초견적작성시 상세외주항목 Temp값을 지우는데 실패하였습니다." & vbcrlf & "관리자에게 문의 하십시오.","처리안내!"
		End If
	
		'텍스트 필드의 활성화
		.txtPREESTNAME.className = "INPUT_L"
		.txtPREESTNAME.readOnly = false
		.txtSUSURATE.className = "INPUT_R"
		.txtSUSURATE.readOnly = false
		.txtSUSUAMT.className = "INPUT_R"
		.txtSUSUAMT.readOnly = false
			
		'내역복사를 하기위하여 준비한다.	
		.txtAGREEYEARMON.value = gNowDate
		.txtPREESTGBN.value ="가견적"
		.txtPREESTNO.value = ""
		.txtMEMO.value = ""
		.txtPREESTNAME.value = ""
		
		'기존의 광고주,팀,브랜드를 복사한다.
		.txtCLIENTCODE.value = parent.document.forms("frmThis").txtCLIENTCODE.value
		.txtTIMCODE.value = parent.document.forms("frmThis").txtTIMCODE.value
		.txtSUBSEQ.value  = parent.document.forms("frmThis").txtSUBSEQ.value
		
		.txtJOBNAME.value  = parent.document.forms("frmThis").txtPRIJOBNAME.value
		.txtJOBNO.value  = parent.document.forms("frmThis").txtJOBNO.value 
		
		'상세항목 내역 pd_subitem_input 에 복제
		strOldCode = parent.document.forms("frmThis").txtPREESTNO.value
		
		'복사하려는 값도 조회나 저장시 현재의 값에 저장 을위해 jobno 를 박는다.
		strJOBNO = parent.document.forms("frmThis").txtJOBNO.value
		
		intSaveRtn = mobjPDCOPREESTDTL.ProcessRtn_TempInsert(gstrConfigXml,strOldCode,strJOBNO)
		
		'상세견적에 
		For intCnt = 1 To .sprSht.MaxRows 
			mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",intCnt, ""
			
			'mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",intCnt, ""
			
			'mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",intCnt, ""
			'If mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",intCnt) = "Y" Then
			'	mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",intCnt, "Y"
			'End IF
			'sprSht_Change 1,intCnt
		Next
		'금액 재계산
		IF .txtPREESTGBN.value ="가견적" THEN
			Call ESTSUSUAMT_CHANGEVALUE2
		ELSE	
			Call SUSUAMT_CHANGEVALUE2
			
		END IF
		BUDGET_AMT_SUM	
		
	End with
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'==============================================
'---삭제나 저장후에 해당 데이터를 재조회 함----
'==============================================
Sub SelectRtn_ProcessRtn (strCODE)
	Dim strCHK
	
	IF not SelectRtn_Head (strCODE) Then Exit Sub
	CALL SelectRtn_Detail (strCODE)
	
	txtSUSUAMT_onblur
	txtSUMAMT_onblur
	txtAMT_onblur
	
	txtESTSUSUAMT_onblur
	txtESTSUMAMT_onblur
	txtESTAMT_onblur
	
	txtCOMMITION_onblur
	txtNONCOMMITION_onblur
	
	with frmThis
		If .txtAGREEYEARMON.value = "" Then
		'	.txtAGREEYEARMON.value = gNowDate
		End if
		
		If .txtENDFLAG.value = "T" AND .txtPREESTGBN.value = "본견적" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
			.txtAGREEYEARMON.className = "NOINPUT"
			.txtAGREEYEARMON.readOnly = true
			.imgCalEndarAGREE.disabled = true
	
			
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = true
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = true
			
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			
		ElseIF  .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "본견적" Then
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtSUSUAMT.className = "INPUT_R"
			.txtSUSUAMT.readOnly = FALSE
			.txtSUSURATE.className = "INPUT_R"
			.txtSUSURATE.readOnly = FALSE
			
			
		ELSEIF .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "가견적" Then
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = TRUE
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = TRUE
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			.txtESTSUSUAMT.className = "INPUT_R"
			.txtESTSUSUAMT.readOnly =  FALSE
			.txtESTSUSURATE.className = "INPUT_R"
			.txtESTSUSURATE.readOnly = FALSE
		End If
		
		
	End with
	CFINPUT_VISIBLE
End Sub

Sub CFINPUT_VISIBLE
	with frmThis
		
		If parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" AND .txtPREESTNO.value <> "" Then
			.imgCFInput.style.visibility = "visible"
		Else
			.imgCFInput.style.visibility = "hidden"
		End If
	End with
End Sub

Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next
	'초기화
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCOPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			'gErrorMsgBox "선택한 가견적에 대하여" & meNO_DATA, "" '요청사항이나, 여러 탭동시 조회시 MSG 창은 사용자에게 혼돈을 줘,, 무리하게 뺐음...By TH
			frmThis.sprSht.MaxRows = 0 
			NewDataSet			
			exit Function
		else
			'조회한 데이터를 바인딩
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function

'예산 테이블 조회
Function SelectRtn_Detail (ByVal strCODE)
	Dim vntData
	Dim intCnt
	Dim intCnt_SubDetail
	Dim strRows
	Dim intColorCnt
	'on error resume next
	'초기화
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCOPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'조회한 데이터를 바인딩
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'초기 상태로 설정
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
				
				Detail_Yn
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
			window.setTimeout "AMT_SUM",100	
			
		End with
	End IF
End Function

Sub Detail_Yn()
	Dim intCnt
	with frmThis
		For intCnt =1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",intCnt) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
				'단가,수량,금액 lock / 버튼은 입력상태 - QTY|PRICE|AMT
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",intCnt,intCnt,false
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",intCnt,intCnt,false
				'버튼형태로 변경
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"간접비입력","DETAIL_BTN",intCnt,intCnt,,false
				Else
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"상세견적","DETAIL_BTN",intCnt,intCnt,,false
				End If
			Else
				'단가,수량,금액 입력받을수 있도록 변경 / 버튼은 lock
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",intCnt,intCnt,false
				'일반형태로 변경
				mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",intCnt,intCnt,0,,,,,,,,False
		End If
		Next
	End With
End Sub

'------------------------------------------
' 데이터 저장
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim strPREESTNO
	Dim intCnt
	Dim vntData
	Dim strAGREEYEARMON
	Dim strJOBNO
	Dim strCHKCONFIRM
	Dim strNEWPREESTNO
	Dim strPRODUCTIONCHK
	Dim i
	
	if DataValidation =false then exit sub
	strMasterData = gXMLGetBindingData (xmlBind)
	
	with frmThis
	
			
		'모든 데이터를 저장한다. print_seq 가 제대로 저장 되지 않는경우가 발생하여 정렬이 움직이는 경우를 막는다 _20120323_ SH
		for i = 1 to .sprSht.maxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		next
		
		'저장의 공통사항 처리
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | IMESEQ | SAVEFLAG | PRINT_SEQ")		
		
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
			Else
				gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				exit sub
			End If
		End If
		
		'헤더 Insert Update 분기처리	
		If .txtPREESTNO.value <> "" Then   
		
			'가견적일경우 Validation 없음
			If .txtPREESTGBN.value  = "가견적" Then
				strCHKCONFIRM = "F"	
				
			Elseif .txtPREESTGBN.value  = "본견적" Then
			'본견적일경우 저장 Validation 필요 - 거래명세서가 작성되어 청구가 된상태라도 변경 될수 있는 여지가 있음.......
				strCHKCONFIRM = "T"
				if .txtENDFLAG.value = "T" Then
					gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 저장이 불가능 합니다.","저장안내!"
					Exit Sub
				End If
			End If
			strPREESTNO = .txtPREESTNO.value 
			
			'간접비가 있는지 없는지 채크
			If PRODUCTIONCHK = True Then
				strPRODUCTIONCHK = "T"
			Else
				strPRODUCTIONCHK = "F"
			End IF
			
			'이곳만 간접비 체크를 하면 된다.
			
			intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,strCHKCONFIRM,"UU",strPRODUCTIONCHK,"PROCESS")
				
			if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
				strJOBNO = .txtJOBNO.value
				SelectRtn_ProcessRtn(strPREESTNO)
				'1번탭 재조회
				parent.jobMst_Tab1Search
				parent.jobMst_Tab5Search
			End If
		Else
		'견적번호가 없다 즉. 견적 리스트에서 다른 JOB 의 견적을 복사 한값이다.
			'일단 여기 견적번호가 있는 이미 저장되어있는 다른 JOB 을 불러왔을 경우...(과거 견적 번호로 데이터를 조회해 올수 있다.)
			if parent.document.forms("frmThis").txtPREESTNO.value <> "" Then 
			 
				'선택한 JOB 과 다르므로 Copy Rule 을 따른다.	
				strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value
				strNEWPREESTNO = ""
				strJOBNO = .txtJOBNO.value
				
				intRtn = mobjPDCOPREESTDTL.ProcessRtn_HDRLESS_COPY(gstrConfigXml, strMasterData, strPREESTNO,strNEWPREESTNO, strJOBNO, strAGREEYEARMON)
				
				if not gDoErrorRtn ("ProcessRtn_HDRLESS_COPY") then
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
					gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
					strJOBNO = .txtJOBNO.value
					SelectRtn_ProcessRtn(strNEWPREESTNO)
					'1번탭 재조회
					
					parent.jobMst_Tab1Search_EstCopy
					parent.jobMst_Tab5Search
				End If	
			
			else
				'선택한 JOB 과 다르므로 Copy Rule 을 따른다.	
				strPREESTNO = ""
				
				intRtn = mobjPDCOPREESTDTL.ProcessRtn_HDRLESS(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON)
				
				if not gDoErrorRtn ("ProcessRtn_HDRLESS") then
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
					gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
					strJOBNO = .txtJOBNO.value
					SelectRtn_ProcessRtn(strPREESTNO)
					'1번탭 재조회
					
					parent.jobMst_Tab1Search_EstCopy
					parent.jobMst_Tab5Search
				End If	
			end if
		End If
		mstrSELECT = "T"
		mstrFIRSTPRODUCTIONCHECK = "N"
	End with
End Sub

'------------------------------------------
' 청구견적을 본견적(실행견적) 으로 저장
'------------------------------------------
Sub ExeProcessRtn
	Dim intCnt
	Dim intRtn
	Dim strPREESTNO
	Dim intRtnSave
	Dim vntData, vntInParams, vntRet
	Dim intRtnChk
	Dim strPRODUCTIONCHK
	Dim strJOBNO, strAGREEYEARMON
	Dim strMasterData
	
	With frmThis
		if DataValidation =false then exit sub
		strMasterData = gXMLGetBindingData (xmlBind)
		
		
   		If PRODUCTIONCHK = True Then
			strPRODUCTIONCHK = "T"
		Else
			strPRODUCTIONCHK = "F"
		End IF

		strPREESTNO = .txtPREESTNO.value 
		strJOBNO = .txtJOBNO.value 
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'견적 삭제와는 다르게 jobno로만 하는 이유는 
		'해당잡의 여러 견적(preestno)이 있지만 실질적으로 본견적인것만 청구진행이 가능하기때문에 
		'해당잡의 많은 견적중에서  confrimflag가 3이상인 건이 있다면 가견적을 본견적으로 바꿀수 없다.
		vntData = mobjPDCOPREESTDTL.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strJOBNO)
		
		If mlngRowCnt > 0  Then
			gOkMsgBox "[본견적]이 승인상태 또는 청구가 진행되었습니다. 삭제할수 없습니다","삭제안내!"
			Exit Sub
		end if 
		
		
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | IMESEQ | SAVEFLAG | PRINT_SEQ")		
	
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
			Else
				'gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
				intRtnChk = gYesNoMsgbox("변경된자료가 없습니다. 본견적으로 반영하시겠습니까?","저장안내")
				If intRtnChk <> vbYes then 
					exit sub
				End If
			End If
		End If
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		intCnt = mobjPDCOPREESTDTL.SelectRtn_ExeCount(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
		IF not gDoErrorRtn ("SelectRtn_ExeCount") then
			If intCnt = "" then
				
				'본견적이 없는경우 (Type1. 가견적을 본견적으로 즉시 저장 됨, Type2. 가견적 내역을 작성하는 도중 본견적으로 저장)
				If .txtPREESTNO.value <> "" Then
				'Type1.
					intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
					if not gDoErrorRtn ("ProcessRtn_Confirm") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
						
						SelectRtn_ProcessRtn(strPREESTNO)
						'1번탭 재조회
						parent.jobMst_Tab1Search
						parent.jobMst_Tab5Search
					End If
				Else
				'본견적이 있는 경우
				'Type2.
				'이경우 간접비가 있는지 체크 해야 함
					strPREESTNO = ""
					intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","I",strPRODUCTIONCHK,"EXPROCESS")
					if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
						strJOBNO = .txtJOBNO.value
						SelectRtn_ProcessRtn(strPREESTNO)
						'1번탭 재조회
						parent.jobMst_Tab1Search
						parent.jobMst_Tab5Search
					End If
				End If
			else
				
				'본견적이 있는경우 -
				if .txtENDFLAG.value = "T" Then
					gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 본견적 변경이 불가능 합니다.","저장안내!"
					Exit Sub
				End If
				
				intRtnSave = gYesNoCancelMsgBox("이미 본견적이 있습니다. 이전 본견적으로 확인하시겠습니까?" & vbcrlf & vbcrlf & "[예:확인후저장 / 아니오:본견적으로 덮어쓰기]","본견적확인저장 및 취소")
				If intRtnSave = 6 then
					'이전본견적 확인 팝업 오픈후,,,본견적이 있는데도 가견적 없이 즉시 본견적을 누를수 있음.
					'아래 intRtnSave = 7 인경우와 동일한 프로세스를 타지만,,,, 팝업이 호출되어 비교할수 있어야 함. 
					'플레그에에 따라 수정본견적으로 저장 하면아래프로세스와 동일하도록 처리.,,,취소를 눌렀다면 본견적으로 저장 하지 아니한 상태값 그대로 반영
					'intCnt 는 이전 본견적번호 임.
				
					if .txtPREESTNO.value = "" then 
						gErrorMsgBox " 견적이 저장되어 있지 않습니다. " & vbcrlf & " 견적을 저장한 후 원래의 견적과 비교해야 합니다.","저장안내" 
						exit sub
					else 
						vntInParams = array(Trim(.txtPREESTNO.value))
					end if
					
					vntRet = gShowModalWindow("PDCMJOBMST_PREESTCONFIRM.aspx",vntInParams, 1149,650)
					
					If vntRet = "TRUE" Then
							'바로 본견적으로 갱신
						If .txtPREESTNO.value <> "" Then
							intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
							if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
								mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
								gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
								strJOBNO = .txtJOBNO.value
								SelectRtn_ProcessRtn(strPREESTNO)
								'1번탭 재조회
								parent.jobMst_Tab1Search
								parent.jobMst_Tab5Search
							End If
						Else
							
						End If
					
					End If
					
				Elseif intRtnSave = 7 then
					
					'바로 본견적으로 갱신
					If .txtPREESTNO.value <> "" Then
					'Type1. 견적이 현재 존재하며 이존재하는 견적을 본견적으로 수정 하는것,,,,
						intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
						if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
							mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
							gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
							strJOBNO = .txtJOBNO.value
							SelectRtn_ProcessRtn(strPREESTNO)
							'1번탭 재조회
							parent.jobMst_Tab1Search
							parent.jobMst_Tab5Search
						End If
					Else
					'Type2. 견적이 현재 존재 하지 아니하며, 현재 입력카피 된값을 본견적으로 즉시 저장 하는것
						strPREESTNO = ""
						intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","I",strPRODUCTIONCHK,"EXPROCESS")
						
						if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
							mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
							gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 
							strJOBNO = .txtJOBNO.value
							
							msgbox "저장된견적번호는: " & strPREESTNO
							SelectRtn_ProcessRtn(strPREESTNO)
							'1번탭 재조회
							parent.jobMst_Tab1Search
							parent.jobMst_Tab5Search
						End If
					End If
				End If
			End IF
		End IF
		mstrSELECT = "T"
	End With	
End Sub

'===============================
'간접비 투입여부를 알아보는 함수
'===============================
Function PRODUCTIONCHK
	Dim intCnt_prod
	Dim intCnt_ProdChk
	intCnt_prodChk = 0
	with frmThis
		For intCnt_prod = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt_prod) = "242001" Then
				intCnt_prodChk = intCnt_prodChk + 1
			End If
		Next
		
		If intCnt_prodChk = 0  Then
			PRODUCTIONCHK = False	
		Else
			PRODUCTIONCHK = True
		End If
	End with
End Function

'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt, intcnt2
   	Dim intRtnChk
   	
	'On error resume next
	with frmThis
  		'Field 필수 입력 항목 검사
  		If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "견적명을 입력하십시오.","저장안내"
			Exit Function
		End If
		If .txtAGREEYEARMON.value = "" Then
			gErrorMsgBox "견적일을 입력하십시오.","저장안내"
			Exit Function
		End If
		
		'Sheet 필수 입력 항목 검사 
		If .sprSht.MaxRows = 0 Then
				gErrorMsgBox "저장할 상세 내역이 존재 하지 않습니다.","저장안내"
				Exit Function
		End IF
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		intcnt2 = 0
   		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNAME",intCnt) = "" _
				Or mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNAME",intCnt) = "" _
				Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt) = "" Or _
				mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODENAME",intCnt) = "" Then 
				
				gErrorMsgBox intCnt & " 번째 행의 외주항목 내용 을 확인하십시오","입력오류"
				Exit Function
			End if
		next
   	End with
	DataValidation = true
End Function

'================================================
'--------------------자료삭제--------------------
'================================================
Sub DeleteRtn ()
	Dim intRtn ,intRtn2
	Dim i, lngCnt
	Dim strPREESTNO
	Dim strJOBNO
	Dim strITEMCODESEQ
	Dim strITEMCODE
	Dim strSUBDETAIL
	Dim intCnt_Prod
	Dim strPRODUCTSUSUCHK
	Dim strCHKCONFIRM
	Dim intCount
	Dim intCount2
	
	with frmThis
		
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "조회된 데이터가 없습니다.","자료 삭제 안내"
			Exit Sub
		end if
		
		'체크된 데이터 확인 
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				intCount = intCount + 1
			end if
		next
		
		if intCount = 0 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, "삭제안내"
			Exit Sub
		end if

		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		
	
		If .txtPREESTGBN.value  = "가견적" Then
			strCHKCONFIRM = "F"	
			
		Elseif .txtPREESTGBN.value  = "본견적" Then'txtJOBNO
		'본견적일경우 저장 Validation 필요 - 거래명세서가 작성되어 청구가 된상태라도 변경 될수 있는 여지가 있음.......
			strCHKCONFIRM = "T"
		End If
		
		strJOBNO = .txtJOBNO.value
		
		lngCnt =0
		intRtn2 = 0
		intCnt_Prod = 0
				
		for i = .sprSht.MaxRows to 1 step -1
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN	
				If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",i) <> ""  Then
					strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)
					strITEMCODESEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",i))
					strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
					strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
					For intCount2 = 1 to .sprSht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCount2) = "242001" Then
							intCnt_Prod = intCnt_Prod +1
						End If
					Next
					
					If intCnt_Prod <> 0 Then
						strPRODUCTSUSUCHK = "T"
					Else 
						strPRODUCTSUSUCHK = "F"
					End If
					
					intRtn2 = mobjPDCOPREESTDTL.DeleteRtn(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL,strPRODUCTSUSUCHK,strJOBNO,strCHKCONFIRM)
					
					IF not gDoErrorRtn ("DeleteRtn") then
						lngCnt = lngCnt +1
						mobjSCGLSpr.DeleteRow .sprSht, i
					end if
				else
					strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)
					
					If strPREESTNO = "" Then 
						strPREESTNO = "9999999999"
						
						strITEMCODESEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i)
						
						strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
						strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
						intRtn2 = mobjPDCOPREESTDTL.DeleteRtn_TempDel(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL)
						IF not gDoErrorRtn ("DeleteRtn_TempDel") then
							lngCnt = lngCnt +1
							mobjSCGLSpr.DeleteRow .sprSht, i					
						end if
					else 
					
						strITEMCODESEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i)
						strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
						strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
						
						intRtn2 = mobjPDCOPREESTDTL.DeleteRtn_TempDel(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL)
						IF not gDoErrorRtn ("DeleteRtn_TempDel") then
							lngCnt = lngCnt +1
							mobjSCGLSpr.DeleteRow .sprSht, i					
						end if
						
					end if 
				end if	
			end if
		next
		'헤더재계산
		IF .txtPREESTGBN.value ="가견적" THEN
			Call ESTSUSUAMT_CHANGEVALUE2
		ELSE	
			Call SUSUAMT_CHANGEVALUE2
		END IF
		BUDGET_AMT_SUM
		'저장되어있는 값이 있으면 DB 에 헤더재계산 값을 저장 
		If intRtn2 = 0 Then
   		Else
			DelProc
		End If
		'1건이라도 삭제건이 있다면 메세지 출력
		If lngCnt <> 0 Then
			gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
		End If
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
	End with
	err.clear
End Sub

'삭제후에 헤더 다시 계산
Sub DelProc
	Dim intHDR
	Dim strMasterData
	Dim strPREESTNO
	Dim strCHKCONFIRM
	Dim strAGREEYEARMON
	Dim strJOBNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		If .txtPREESTGBN.value  = "가견적" Then
			strCHKCONFIRM = "F"	
			
		Elseif .txtPREESTGBN.value  = "본견적" Then
		'본견적일경우 저장 Validation 필요 - 거래명세서가 작성되어 청구가 된상태라도 변경 될수 있는 여지가 있음.......
			strCHKCONFIRM = "T"
		End If
		strPREESTNO = .txtPREESTNO.value
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		'intHDR = mobjPDCOPREESTDTL.ProcessRtn_DelProc(gstrConfigXml,strMasterData,strPREESTNO,strAGREEYEARMON,strCHKCONFIRM,"U")
		
		if not gDoErrorRtn ("ProcessRtn_DelProc") then
			strJOBNO = .txtJOBNO.value
			SelectRtn_ProcessRtn(strPREESTNO)
			'1번탭 재조회
			parent.jobMst_Tab1Search
		End If
	End with
End Sub

Function CleanField (ByVal objField, ByVal objField1)
	if isobject(objField) then objField.value = ""
	if isobject(objField1) then objField1.value = ""
end Function


		</script>
	</HEAD>
	<body style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px" class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE border="0" cellSpacing="1" cellPadding="0" width="100%" align="left" height="98%">
				<TR>
					<TD>
						<TABLE id="tblTitle1" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD0" height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<TR>
											<td class="TITLE">청구견적
											</td>
										</TR>
									</table>
								</TD>
								<TD style="WIDTH: 100%" height="20" vAlign="middle" align="right">
									<!--Common Button Start--></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<TABLE id="tblDATA" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
							align="left">
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80">가/본</TD>
								<TD class="SEARCHDATA" width="150"><INPUT style="WIDTH: 148px; HEIGHT: 22px" id="txtPREESTGBN" dataSrc="#xmlBind" class="NOINPUTB_L"
										title="가/본 견적구분" dataFld="PREESTGBN" readOnly maxLength="10" size="32" name="txtPREESTGBN"></TD>
								<TD style="WIDTH: 92px; CURSOR: hand" class="SEARCHLABEL" width="92">대표JOB명</TD>
								<TD class="SEARCHDATA" width="260"><INPUT accessKey=",M" style="WIDTH: 256px; HEIGHT: 22px" id="txtJOBNAME" dataSrc="#xmlBind"
										class="NOINPUTB_L" title="JOB명" dataFld="JOBNAME" readOnly maxLength="255" size="37" name="txtJOBNAME"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL"><span id="strMsg_Amt">가견적금액</span></TD>
								<TD class="SEARCHDATA" width="105"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUMAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="수수료 + 금액" dataFld="ESTSUMAMT" readOnly maxLength="20" size="32" name="txtESTSUMAMT"></SPAN></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL">실행견적금액</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtSUMAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="수수료 + 금액" dataFld="SUMAMT" readOnly maxLength="20" size="32" name="txtSUMAMT">
								</TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80">견적코드</TD>
								<TD class="SEARCHDATA" width="150"><INPUT style="WIDTH: 148px; HEIGHT: 22px" id="txtPREESTNO" dataSrc="#xmlBind" class="NOINPUTB_L"
										title="견적코드" dataFld="PREESTNO" readOnly maxLength="10" size="32" name="txtPREESTNO">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call CleanField('','')"
									width="94">견적명</TD>
								<TD class="SEARCHDATA" width="260"><INPUT accessKey=",M" style="WIDTH: 256px; HEIGHT: 22px" id="txtPREESTNAME" dataSrc="#xmlBind"
										class="INPUT_L" title="견적명" dataFld="PREESTNAME" maxLength="255" size="37" name="txtPREESTNAME"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80" align="right">(가)금액</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="금액합계" dataFld="ESTAMT" readOnly maxLength="20" size="32" name="txtESTAMT"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80" align="right">(실)금액</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="금액합계" dataFld="AMT" readOnly maxLength="20" size="32" name="txtAMT"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')">견적일</TD>
								<TD class="SEARCHDATA"><INPUT accessKey="DATE,M" style="WIDTH: 72px; HEIGHT: 22px" id="txtAGREEYEARMON" dataSrc="#xmlBind"
										class="INPUT" title="견적합의일" dataFld="AGREEYEARMON" maxLength="10" size="6" name="txtAGREEYEARMON">
									<IMG style="CURSOR: hand" id="imgCalEndarAGREE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarAGREE"
										align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')">비고</TD>
								<TD class="SEARCHDATA"><TEXTAREA style="WIDTH: 256px; HEIGHT: 22px" id="txtMEMO" dataSrc="#xmlBind" dataFld="MEMO"
										wrap="hard" cols="10" name="txtMEMO"></TEXTAREA></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')"
									width="80">(가)수수료</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUSUAMT" dataSrc="#xmlBind"
										class="INPUT_R" title="수수료금액합계" dataFld="ESTSUSUAMT" maxLength="20" size="32" name="txtESTSUSUAMT"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')"
									width="80">(실)수수료</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtSUSUAMT" dataSrc="#xmlBind"
										class="INPUT_R" title="수수료금액합계" dataFld="SUSUAMT" maxLength="20" size="32" name="txtSUSUAMT"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtENDFLAG" dataSrc="#xmlBind" dataFld="ENDFLAG"
										size="1" type="hidden" name="txtENDFLAG"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtENDFLAGEXE" dataSrc="#xmlBind" dataFld="ENDFLAGEXE"
										size="1" type="hidden" name="txtENDFLAGEXE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtSETCONFIRMFLAG" dataSrc="#xmlBind" dataFld="SETCONFIRMFLAG"
										size="1" type="hidden" name="txtSETCONFIRMFLAG"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">Commission</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 148px; HEIGHT: 22px" id="txtCOMMITION" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="수수료대상금액" dataFld="COMMITION" readOnly maxLength="20" size="32" name="COMMITION">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">NonCommission</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 256px; HEIGHT: 22px" id="txtNONCOMMITION" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="수수료제외금액" dataFld="NONCOMMITION" readOnly maxLength="20" size="37" name="txtNONCOMMITION"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField(txtMEMO, '')"
									width="80">(가)수수료율</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUSURATE" dataSrc="#xmlBind" class="INPUT_R"
										title="가/본 견적구분" dataFld="ESTSUSURATE" maxLength="20" size="37" name="txtESTSUSURATE"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField(txtMEMO, '')"
									width="80">(실)수수료율</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 100px; HEIGHT: 22px" id="txtSUSURATE" dataSrc="#xmlBind" class="INPUT_R"
										title="가/본 견적구분" dataFld="SUSURATE" maxLength="20" size="37" name="txtSUSURATE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtJOBNO" dataSrc="#xmlBind" dataFld="JOBNO"
										size="1" type="hidden" name="txtJOBNO"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtCREDAY" dataSrc="#xmlBind" dataFld="CREDAY"
										size="1" type="hidden" name="txtCREDAY"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtTIMCODE" dataSrc="#xmlBind" dataFld="TIMCODE"
										size="1" type="hidden" name="txtTIMCODE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtCLIENTCODE" dataSrc="#xmlBind" dataFld="CLIENTCODE"
										size="1" type="hidden" name="txtCLIENTCODE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtSUBSEQ" dataSrc="#xmlBind" dataFld="SUBSEQ"
										size="1" type="hidden" name="txtSUBSEQ"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">부가항목</TD>
								<TD class="SEARCHDATA" colSpan="7"><INPUT accessKey="DATE" style="WIDTH: 72px; HEIGHT: 22px" id="txtPRINTDAY" class="INPUT"
										title="견적서출력일" maxLength="10" size="6" name="txtPRINTDAY"> <IMG style="CURSOR: hand" id="imgimgCalEndarCREDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgimgCalEndarCREDAY" align="absMiddle" src="../../../images/btnCalEndar.gIF"
										height="15">&nbsp;<SELECT style="WIDTH: 132px" id="cmbESTTYPE" title="견적유형선택" name="cmbESTTYPE">
										<OPTION selected value="1">ESTIMATE</OPTION>
										<OPTION value="2">ESTIMATE/ACTUAL</OPTION>
										<OPTION value="3">ACTUAL</OPTION>
									</SELECT><IMG style="CURSOR: hand" id="imgPrintEst" onmouseover="JavaScript:this.src='../../../images/imgPrintEstOn.gIF'"
										title="견적일 및 견적유형을 선택하시어 선택적 견적서를 출력합니다" onmouseout="JavaScript:this.src='../../../images/imgPrintEst.gif'"
										border="0" name="imgPrintEst" alt="견적서출력(상세)." align="absMiddle" src="../../../images/imgPrintEst.gIF"
										width="100" height="20">&nbsp;<IMG style="CURSOR: hand" id="imgPrintEstBasic" onmouseover="JavaScript:this.src='../../../images/imgPrintEstBasicOn.gIF'"
										title="견적일 및 견적유형을 선택하시어 선택적 기본견적서를 출력합니다" onmouseout="JavaScript:this.src='../../../images/imgPrintEstBasic.gif'"
										border="0" name="imgPrintEstBasic" alt="견적서출력(기본)." align="absMiddle" src="../../../images/imgPrintEstBasic.gIF" width="120"
										height="20">&nbsp;<IMG style="CURSOR: hand" id="imgCFInput" onmouseover="JavaScript:this.src='../../../images/imgCFInputOn.gIF'"
										title="CF 외주내역 항목을 기입하여 견적서에 적용합니다." onmouseout="JavaScript:this.src='../../../images/imgCFInput.gif'"
										border="0" name="imgCFInput" alt="" align="absMiddle" src="../../../images/imgCFInput.gIF" height="20"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 25px" id="spacebar" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" height="20" width="120" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="200">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">세부내역&nbsp;<span id="strMsgBox"></span>
											</td>
										</tr>
									</table>
								</TD>
								<td class="TITLE">금액합계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT_TOTAL" title="합계금액" style="WIDTH: 120px; HEIGHT: 22px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT_TOTAL">&nbsp; 
									선택합계 : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
								<TD height="20" vAlign="middle" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
										<TR>
											<td width="62" align="left"><input accessKey="NUM," style="VISIBILITY: hidden; WIDTH: 5px" id="txtPRINT_SEQ" value="1"
													maxLength="2" name="txtPRINT_SEQ"><IMG style="CURSOR: hand" id="imgTableUp" border="0" name="imgTableUp" alt="자료를 올립니다."
													align="absMiddle" src="../../../images/imgTableUp.gif"> <IMG style="CURSOR: hand" id="imgTableDown" border="0" name="imgTableDown" alt="자료를 내립니다."
													align="absMiddle" src="../../../images/imgTableDown.gif"></td>
											<TD><IMG style="CURSOR: hand" id="ImgBasicFormat" onmouseover="JavaScript:this.src='../../../images/ImgBasicFormatOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/ImgBasicFormat.gIF'" border="0"
													name="ImgBasicFormat" alt="견적타입별 기본값을 투입합니다" src="../../../images/ImgBasicFormat.gIF"
													height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" border="0" name="imgRowAdd"
													alt="자료입력을 위해 행을추가합니다." src="../../../images/imgRowAdd.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" border="0" name="imgRowDel"
													alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgSave"
													alt="자료를 저장합니다." src="../../../images/imgSave.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgBonSave" onmouseover="JavaScript:this.src='../../../images/imgBonSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgBonSave.gIF'" border="0" name="imgBonSave"
													alt="본견적으로저장" src="../../../images/imgBonSave.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 99%" vAlign="top" align="center">
						<DIV style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" id="pnlTab1"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="9737">
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
				</tr>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 1%" vAlign="top" align="center">
						<DIV style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" id="pnlTab2"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_copy" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="503">
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
				</tr>
				<TR>
					<TD style="WIDTH: 1040px" id="lblStatus" class="BOTTOMSPLIT"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
