<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_SUBITEM.aspx.vb" Inherits="PD.PDCMJOBMST_SUBITEM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>상세견적관리</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST_SUBITEM.aspx
'기      능 : JOBMST의 두번째 탭 PDCMJOBMST_ESTDTL 의 상세견적 버튼을 클릭하였을때 처리 
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/19 By KimTH
'****************************************************************************************
-->
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
Dim mobjPDCOPREESTSUB		'cc 개체 생성을위한 변수
Dim mstrITEMCODE			'외주 항목 코드
Dim mstrITEMCODESEQ			'외주 항목 코드 순번
Dim mstrSAVEFLAG			'상세내역이 저장되어있는지 가져오는 변수
Dim mstrPREESTNO			'견적 번호
Dim mstrCheck	
Dim mstrGBN					'본견적 / 가견적
Dim mlngTempRowCnt,mlngTempColCnt
Dim mstrSAVEGBN
Dim mstrJOBNO				'JOBNO
Dim mstrCHANGEFALG			'변경확인 플래그(삭제를 할경우 전체 삭제일 경우의 예외를 처리한다.)  [ T/F  (T 일반 이벤트시 / F 삭제 이벤트가 발생할경우)]


mstrCheck = True	
mstrCHANGEFALG = "F"

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	Dim lngPRICEAMT
	Dim vntData
	Dim INPUTAMT
	'가견적일 경우 견적 금액을 리턴 본견적일 경우 실행 견적을 리턴한다.
	with frmThis
		
		'저장이 되어있는지 아닌지 결정 [저장이 되었다는건 INPUT 에 금액이 있다는 뜻!]
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'닫기 버튼을 누를경우 window_onunload 를 닫기 버튼에 의해서 1번타고 정말 윈도우가 닫히면서 한번(두번) 타는거같음. 그래서 개체 생성 삽입
		'set mobjPDCOPREESTSUB = gCreateRemoteObject("cPDCO.ccPDCOPREESTSUB")
		
		vntData = mobjPDCOPREESTSUB.SelectRtn_INPUTAMT(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrJOBNO, mstrITEMCODE, mstrITEMCODESEQ)
		
		'Set mobjPDCOPREESTSUB = Nothing

		'가견적일경우
		If mstrGBN = "F" Then
			if vntData(0,1) = "" then
				INPUTAMT = split("False;" & mstrCHANGEFALG ,";")
			ELSE 
				INPUTAMT = split(vntData(0,1) & ";" & mstrCHANGEFALG ,";")
			END IF		
		'본견적일경우
		Elseif mstrGBN = "T" Then
		
			if vntData(1,1) = "" then
				INPUTAMT = split("False;" & mstrCHANGEFALG ,";")
			ELSE 
				INPUTAMT = split(vntData(1,1) & ";" & mstrCHANGEFALG ,";")
			END IF
		End If
		window.returnvalue = INPUTAMT
		
	End with
	Set mobjPDCOPREESTSUB = Nothing
End Sub

Sub imgClose_onclick ()
	EndPage
	'Window_OnUnload
End Sub

'-----------------------------------
' 명령 버튼 클릭 이벤트
'-----------------------------------

Sub imgRowAdd_onclick ()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 행추가가 불가능 합니다.","행추가안내!"
		Exit Sub
	End If
	call sprSht_Keydown(meINS_ROW, 0)
End Sub


Sub imgSave_onclick ()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 저장이 불가능 합니다.","저장안내!"
		Exit Sub
	End If
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
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 삭제 불가능 합니다.","행삭제안내!"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgAllType_onclick()
	Dim intCnt
	Dim strSEQ
	Dim strSEQString
	Dim strLen
	Dim vntData
	Dim dblRow
	
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "청구요청 및 거래명세서 진행중이므로 전체타입항목 보기가 불가능 합니다.","처리안내!"
		Exit Sub
	End If
	
	With frmThis
		
		For intCnt = 1 To .sprSht.MaxRows
			strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBITEMCODESEQ",intCnt)
			strSEQString = strSEQString & strSEQ & ","
		Next 
		
		dblRow = .sprSht.MaxRows + 1
		if strSEQString <> "" then
			strSEQString = MID(strSEQString,1, Len(strSEQString) -1)
		else 
			strSEQString = "999999"
		end if 
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		vntData = mobjPDCOPREESTSUB.SelectRtn_Fill(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,strSEQString)
		if not gDoErrorRtn ("SelectRtn_Fill") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, dblRow, mlngColCnt, mlngRowCnt, True
				gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			end If
   		end if
   					
	End With
End Sub

Sub ImgMoveData_onclick
	Dim intCnt
	with frmThis
		If mstrGBN = "F" Then 
			gErrorMsgBox "가견적 상태는 실행견적을 입력할수 없습니다.","입력안내!"
			EXIT SUB
		END IF 
		
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"QTY", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"TERM", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt,"1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
	End with
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
	Dim strCheckRow
	
	with frmThis
	
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'PRINT_SEQ 가 1부터 순서대로가 아닐수도있음. 전체를 돌면서 제일 작은 값을 찾아낸다.
		strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",1)
		for i=1 to .sprSht.MaxRows-1
			for j=i+1 to .sprSht.MaxRows
				IF strCheckRow > mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j) then
					strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j)
				end if
			Next
		Next	
		strCheckRow= strCheckRow -1		
		
		
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
		
		for i=1 to .sprSht_copy.MaxRows
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow- strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow-strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMCODESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"TERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"TERM",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"MEMO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEPRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEPRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEQTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEQTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXETERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXETERM",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEAMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEMEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEMEMO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"NEWFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"NEWFLAG",strRow -strPRINT_SEQ)
			
			'기본으로 Y를 박는다
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


Sub sprSht_DownCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		
		'PRINT_SEQ 가 1부터 순서대로가 아닐수도있음. 전체를 돌면서 제일 작은 값을 찾아낸다.
		strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",1)
		for i=1 to .sprSht.MaxRows-1
			for j=i+1 to .sprSht.MaxRows
				IF strCheckRow > mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j) then
					strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j)
				end if
			Next
		Next	
		strCheckRow= strCheckRow -1	
		
		
		'row셋팅	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow+ strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow+strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMCODESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"TERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"TERM",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"MEMO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEPRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEPRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEQTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEQTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXETERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXETERM",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEAMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEMEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEMEMO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"NEWFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"NEWFLAG",strRow +strPRINT_SEQ)
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 		
			
			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


'=============================================================
'Sheet Event
'=============================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if
	end with	
End Sub


Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim lngPRICE
	Dim lngQTY
	Dim lngTERM
	Dim lngAMT
	Dim lngCalCul
	
	Dim lngEXEPRICE
	Dim lngEXEQTY
	Dim lngEXETERM
	Dim lngEXEAMT
	Dim lngEXECalCul
	
	Dim lngOldAMT
	Dim lngOldPRICE
	Dim lngOldExeAmt
	Dim lngOldExePRICE
	
	With frmThis	
			
			lngOldAMT		= mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			lngOldPRICE		= mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
			lngOldExeAmt	= mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
			lngOldExePRICE	= mobjSCGLSpr.GetTextBinding(.sprSht,"EXEPRICE",Row)
				
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Or  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") Or _
																	Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TERM") Then
   				lngPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   				lngQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   				lngTERM = mobjSCGLSpr.GetTextBinding(.sprSht,"TERM",Row)
		   		
   				lngCalCul = lngPRICE * lngQTY * lngTERM
   				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngCalCul
		   		
   				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   				.txtDIVNAME.focus
				.sprSht.focus
		   		
   					'상세내역금액을 0으로 수정할경우 변경되야함..._20110830_SH
   		   		IF lngOldAMT <> mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) THEN
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				ELSE
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   				END IF 
	   	
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
   				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	   			
   				If lngAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"TERM",Row, "1"
   				End If
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Or  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEQTY") Or _
   																		Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXETERM") Then
	   			
   				lngEXEPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEPRICE",Row)
   				lngEXEQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEQTY",Row)
   				lngEXETERM = mobjSCGLSpr.GetTextBinding(.sprSht,"EXETERM",Row)
		   		
   				lngEXECalCul = lngEXEPRICE * lngEXEQTY * lngEXETERM
   				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",Row, lngEXECalCul
		   		
   				lngEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
   				.txtDIVNAME.focus
				.sprSht.focus


   				IF lngOldEXEAMT <> mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) THEN
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				ELSE
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   				END IF 
	   				   			
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
   				lngEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
	   			
   				If lngEXEAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",Row, "1"
   				End If
   			End If
   		AMT_SUM
   	End with 
   	'실제 Sprsht 변경에 대한 플레그 처리
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
   	
End Sub



Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"NEWFLAG")  Then
			
				intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
				DefaultValue
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
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, mstrPREESTNO
		
		mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"NEWFLAG",.sprSht.MaxRows -1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"QTY",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"TERM",.sprSht.ActiveRow, 1
		mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",.sprSht.ActiveRow, 1
		
		mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"PRINT_SEQ",.sprSht.MaxRows -1)+1	
	End With
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
		'sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow	
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
				strCOLUMN = "EXEPRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
				strCOLUMN = "EXEAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
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

Sub InitPage()
	'서버업무객체 생성	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCOPREESTSUB = gCreateRemoteObject("cPDCO.ccPDCOPREESTSUB")

	gInitComParams mobjSCGLCtl,"MC"
	'탭 위치 설정 및 초기화
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정

		for i = 0 to intNo
			select case i
				case 0 : .txtDIVNAME.value = vntInParam(i)			'대분류명
				case 1 : .txtCLASSNAME.value = vntInParam(i)		'중분류명
				case 2 : .txtITEMCODENAME.value = vntInParam(i)		'외주항목명
				case 3 : mstrITEMCODE = vntInParam(i)				'외주항목코드
				case 4 : mstrSAVEFLAG = vntInParam(i)				'조회시 최초입력 조회인지, 저장된내역 조회인지 구분
				case 5 : mstrPREESTNO = vntInParam(i)				'견적번호를 가져온다.
				case 6 : mstrGBN = vntInParam(i)					'T:본견적시 처리, F:가견적시처리
				case 7 : mstrITEMCODESEQ = vntInParam(i)			'외주항목코드 순번 저장 안하고 넘어올시 가짜 순번따라옴[IMESEQ]
				case 8 : mstrSAVEGBN = vntInParam(i)				'청구 확인용 플래그
				case 9 : mstrJOBNO = vntInParam(i)					'부모창의 JOBNO
			end select
		next
			
	'**************************************************
	'***Sum Sheet 디자인
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 19, 0
	mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG "
	mobjSCGLSpr.SetHeader .sprSht,		    "선택|견적번호|이동|순번|코드|상세견적항목|단가|수량|기간|금액|비고|실행단가|수량|기간|실행금액|실행비고|저장구분|신규투입여부|이동여부"
	mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|      10|   4|   4|   4|          21|  12|   6|   4|  12|  12|      12|   6|   4|      12|      12|      10|          10|       6"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRINT_SEQ|SUBITEMCODESEQ|PRICE|AMT|EXEPRICE|EXEAMT", -1, -1, 0
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|TERM|EXEQTY|EXETERM", -1, -1, 1
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SAVEFLAG|SEQ"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "SUBITEMNAME|MEMO|EXEMEMO",-1,-1,0,2,false ' 왼쪽
	mobjSCGLSpr.SetCellAlign2 .sprSht, "PRINT_SEQ | SEQ | SUBITEMCODESEQ",-1,-1,2,2,false ' 왼쪽
	
	
	gSetSheetColor mobjSCGLSpr, .sprSht_copy
	mobjSCGLSpr.SpreadLayout .sprSht_copy, 19, 0
	mobjSCGLSpr.SpreadDataField .sprSht_copy,    "CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG "
	mobjSCGLSpr.SetHeader .sprSht_copy,		    "선택|견적번호|이동|순번|코드|상세견적항목|단가|수량|기간|금액|비고|실행단가|수량|기간|실행금액|실행비고|저장구분|신규투입여부|이동여부"
	mobjSCGLSpr.SetColWidth .sprSht_copy, "-1",  "   4|      10|   4|   4|   4|          21|  12|   6|   4|  12|  12|      12|   6|   4|      12|      12|      10|          10|     6"
	mobjSCGLSpr.SetRowHeight .sprSht_copy, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht_copy, "-1", "13"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_copy, "CHK "
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "PRINT_SEQ|SUBITEMCODESEQ|PRICE|AMT|EXEPRICE|EXEAMT", -1, -1, 0
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "QTY|TERM|EXEQTY|EXETERM", -1, -1, 1
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht_copy, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht_copy,true,"PRINT_SEQ|SAVEFLAG|SEQ|NEWFLAG|SUBITEMCODESEQ"
	mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "SUBITEMNAME|MEMO|EXEMEMO",-1,-1,0,2,false ' 왼쪽
	mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "PRINT_SEQ | SEQ | SUBITEMCODESEQ",-1,-1,2,2,false ' 왼쪽
	mobjSCGLSpr.ColHidden .sprSht, "SAVEFLAG | NEWFLAG | MOVEFLAG", true
	
	'본견적이라면 가견적내역 수정불가
	If mstrGBN = "F" Then
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO"

	'가견적이라면 본견적내역 수정불가
	Else
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PRICE | QTY | TERM | AMT | MEMO"
	
	End If

	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	End with
	
	'화면 초기값 설정
	SelectRtn
	
End Sub

Sub EndPage
	gEndPage
End Sub

'=============================================================
'조회
'=============================================================
Sub SelectRtn
	Dim vntData
   	Dim vntData_Temp
   	Dim vntData_Dtl
    Dim intCnt
    Dim intRtn
	
    
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		mlngTempRowCnt=clng(0)
		mlngTempColCnt=clng(0)
			
			'진행중인 상세 내역을 먼저 조회한다.(아직 INPUT에 저장되어있는 값)
			intCnt = mobjPDCOPREESTSUB.SelectRtn_TempCount(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt, mstrJOBNO, mstrITEMCODE, mstrITEMCODESEQ,mstrPREESTNO)
			IF not gDoErrorRtn ("SelectRtn_ExeCount") then
				If mlngTempRowCnt > 0 Then
					'상세내역이 있다면 상세내역을 보여준다.
					vntData_Temp = mobjPDCOPREESTSUB.SelectRtn_Temp(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ,mstrPREESTNO)
					if not gDoErrorRtn ("SelectRtn_Temp") then
						if mlngRowCnt > 0 Then
							mobjSCGLSpr.SetClipbinding .sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True
							gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
						Else	
							.sprSht.MaxRows = 0
							gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   						end If
   					end if
				ELSE
					'상세 내역이 없으면 저장되어있는 데이터 가져오기
					intRtn = mobjPDCOPREESTSUB.SelectRtn_DtlCount(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ,mstrPREESTNO)
					IF not gDoErrorRtn ("SelectRtn_DtlCount") then
						If mlngTempRowCnt > 0 Then
							vntData_Dtl = mobjPDCOPREESTSUB.SelectRtn_Dtl(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,mstrITEMCODESEQ,mstrJOBNO)
							if not gDoErrorRtn ("SelectRtn_Dtl") then
								if mlngRowCnt > 0 Then
									mobjSCGLSpr.SetClipbinding .sprSht, vntData_Dtl, 1, 1, mlngColCnt, mlngRowCnt, True
									gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   								Else	
   									.sprSht.MaxRows = 0
   									gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   								end If
   							end if
						Else
							'상세내역이 없다면 전체 내역을 보여준다.
							vntData = mobjPDCOPREESTSUB.SelectRtn_Empty(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,mstrITEMCODESEQ)
							if not gDoErrorRtn ("SelectRtn_Empty") then
								if mlngRowCnt > 0 Then
									mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
									gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   								Else	
   									.sprSht.MaxRows = 0
   									gWriteText lblStatus, 0 & "건의 자료가 검색" & mePROC_DONE
   								end If
   							end if
						End If
					END IF
				END IF 
			END IF 
		
	window.setTimeout "AMT_SUM",1	
	.txtSELECTAMT.value = 0
	mstrCHANGEFALG = "F"
   	end with
End Sub


Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	
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
		
		IntEXEAMT = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtEXESUMAMT.value = 0
		else
			.txtEXESUMAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtEXESUMAMT,0,True)
		End If
	End With
End Sub
'------------------------------------------------------------------
'---------------------------------저장----------------------------- 
'------------------------------------------------------------------
Sub processRtn
	Dim vntData
	Dim intRtn
	Dim intCnt
	Dim ProcessCHK
	Dim i
	Dim strNEWFALG

	with frmThis
		if .sprSht.MaxRows = 0 Then 
			gErrorMsgBox "저장할 데이터가 없습니다.","저장안내"
			EXIT Sub
		end if
		
		'첫번째 NEWFLAG 를 가져온다 (순서는 상관없음)
		strNEWFALG = mobjSCGLSpr.GetTextBinding(.sprSht,"NEWFLAG",1)
		
		if strNEWFALG = "Y" then
			For i = 1 To .sprSht.MaxRows 
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				intCnt = intCnt + 1
				end if 
			Next
			
			If intCnt = 0 Then
				gErrorMsgBox "선택하신(변경된) 자료가 없습니다 저장 하실 자료를 선택하세요.","저장취소"
				EXIT Sub
			End If
		end if 
		
		'자료를 저장하기전에 한번더 확인한다.[삭제후 저장하기 때문에 삭제될 자료와 저장될 자료의 구분을 명확이 해야 한다.]
		ProcessCHK = gYesNoMsgbox("자료를 저장하시겠습니까?","자료저장 확인")
		IF ProcessCHK <> vbYes then exit Sub
		
				'insert 플래그 변경 
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
   		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | PREESTNO | PRINT_SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | SEQ | NEWFLAG | MOVEFLAG")
		
		
		intRtn = mobjPDCOPREESTSUB.ProcessRtn(gstrConfigXml,vntData, mstrJOBNO,mstrPREESTNO, mstrITEMCODE, mstrITEMCODESEQ, mstrGBN,strNEWFALG)
	
		If not gDoErrorRtn ("ProcessRtn") Then
			'모든 플래그 클리어
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "저장되었습니다.","저장안내!"
			SelectRtn
			.sprSht.focus()
			mstrCHANGEFALG = "T"
		End If
	End with
End Sub

'----------------------------------------
'--------------------삭제----------------
'----------------------------------------
Sub DeleteRtn
	Dim vntData
	Dim intCnt, intRtn, i
	'삭제Key 설정
	Dim dblSEQ
	
	with frmThis
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		'선택된 자료를 끝에서 부터 삭제
		
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				mobjSCGLSpr.DeleteRow .sprSht,i
				intCnt = intCnt + 1
			end if 
		next
		
				'insert 플래그 변경 
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
   		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG ")
		
		'해당 INPUT의 모든 데이터 삭제후 저장 
		intRtn = mobjPDCOPREESTSUB.DeleteRtn_Temp(gstrConfigXml,vntData,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ, mstrGBN)
		
		If not gDoErrorRtn ("DeleteRtn_Temp") Then
			gOkMsgBox  intCnt & "건이 삭제" & mePROC_DONE, "삭제 안내!"
   		End IF
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		
		if intRtn = 0 then 
			'삭제 확인플래그(전체 삭제를위한 플래그)
			mstrCHANGEFALG = "T"
			'전체 삭제 했을경우 비우는것이 아니라 모든 타입을 선택할수 있도록 한다.
			call ImgAllType_onclick
		end if 
		.sprSht.focus()	
	End with
	err.clear	
End Sub

		</script>
	</HEAD>
	<BODY class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table style="WIDTH: 100%; HEIGHT: 24px" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="76" background="../../../images/back_p.gIF"
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
								<td class="TITLE">상세견적내역&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="SEARCHDATA" width="100%">
				<tr>
					<td class="searchDATA" colSpan="7">&nbsp;대분류 <INPUT class="NOINPUTB_L" id="txtDIVNAME" title="대분류명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="32" name="txtDIVNAME">&nbsp;&nbsp;&nbsp;중분류
						<INPUT class="NOINPUTB_L" id="txtCLASSNAME" title="중분류명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="29" name="txtCLASSNAME"> &nbsp;&nbsp;&nbsp;&nbsp;견적항목
						<INPUT class="NOINPUTB_L" id="txtITEMCODENAME" title="견적항목명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="30" name="txtITEMCODENAME">&nbsp;</td>
					<td align="right" width="54"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="화면을 닫습니다."
							src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
				</tr>
			</table>
			</TABLE> 
			<!--Top Define Table Start-->
			<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
				border="0">
				<TR>
					<TD align="left" width="100%" height="10"></TD>
				</TR>
			</TABLE>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">합 계 : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="합계금액" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtEXESUMAMT" title="합계금액" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtEXESUMAMT">&nbsp; <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="선택금액" style="WIDTH: 120px; HEIGHT: 22px"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<TD align="right" width="700">
						<input id="txtPRINT_SEQ" style="VISIBILITY: hidden; WIDTH: 5px" type="text" maxLength="2"
							value="1" name="txtPRINT_SEQ" accessKey="NUM,"><IMG id="imgTableUp" style="CURSOR: hand" alt="자료를 올립니다." src="../../../images/imgTableUp.gif"
							align="absMiddle" border="0" name="imgTableUp"> <IMG id="imgTableDown" style="CURSOR: hand" alt="자료를 내립니다." src="../../../images/imgTableDown.gif"
							align="absMiddle" border="0" name="imgTableDown"> <IMG id="ImgAllType" onmouseover="JavaScript:this.src='../../../images/ImgAllTypeOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAllType.gIF'" height="20" alt="모든상세항목을보여줍니다." src="../../../images/ImgAllType.gIF"
							align="absMiddle" border="0" name="ImgAllType">&nbsp;<IMG id="ImgMoveData" onmouseover="JavaScript:this.src='../../../images/ImgMoveDataOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMoveData.gIF'" height="20" alt="가견적내역 을 실행견적으로 복제합니다." src="../../../images/ImgMoveData.gIF"
							align="absMiddle" border="0" name="ImgMoveData">&nbsp;<IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" height="20" alt="자료입력을 위해 행을추가합니다." src="../../../images/imgRowAdd.gIF"
							align="absMiddle" border="0" name="imgRowAdd">&nbsp;<IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF"
							align="absMiddle" border="0" name="imgRowDel">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
							onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다." src="../../../images/imgSave.gIF" align="absMiddle"
							border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다."
							src="../../../images/imgExcel.gIF" width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</TD>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--내용-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="26009">
								<PARAM NAME="_ExtentY" VALUE="11933">
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
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 0%" vAlign="top" align="center">
						<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_copy" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="26009">
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
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
				</TR>
			</table>
		</form>
	</BODY>
</HTML>
