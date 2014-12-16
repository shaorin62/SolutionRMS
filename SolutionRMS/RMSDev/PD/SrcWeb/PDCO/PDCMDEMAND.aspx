<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDEMAND.aspx.vb" Inherits="PD.PDCMDEMAND" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>청구요청</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMDEMAND.aspx
'기      능 : SpreadSheet를 이용한 청구요청/JOB분할/조회 의 기능을 가진다.
'파라  메터 : 
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/10 By KimTH
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
Dim mlngRowCnt, mlngColCnt			'데이터의 로우및 컬럼 반환
Dim mobjPDCODEMAND					'청구요청 의 Control Class
Dim mobjPDCOGET						'제작공통 Control Class
Dim mobjSCCOGET						'전체공통 Control Class
Dim mstrCheck						'전체 선택 및 해제 구분자
Dim mstrSelect						'조회구분 (저장후 이력조회 Or 최초 입력대상 조회)
Dim mlngRowChk						'하단그리드 저장발리데이션 사용
Dim mstrDEPTCD						'로그인사용자부서
Dim mstrMANAGER						'로그인사용자의 슈퍼관리 여부

Dim mlngTaxRowCnt
Dim mlngTaxColCnt
Const meTab = 9
mstrCheck = True					'전체선택은 최초 해제	
mstrSelect = false					'조회구분 Default Value: 입력대상 조회

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgDivDemand_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' 명령버튼
'=========================================================================================
Sub imgQuery_onclick
	with frmThis
		
	End with
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'인쇄 - 해당사항 없음
Sub imgPrint_onclick ()
	
End Sub	



Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel2_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht1
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDelUp_onclick
	gFlowWait meWAIT_ON
	DeleteRtnProc
	gFlowWait meWAIT_OFF

End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' SpreadSheet 이벤트 
'=========================================================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intRtn
	Dim dblChk
	Dim dblChkSum
	Dim vntData
	Dim intRtnChk
	'mlngRowChk
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		Else
			dblChk = 0
			dblChkSum = 0
			For intCnt = 1 To .sprSht1.MaxRows
				If mobjSCGLSpr.GetTextBinding( .sprSht1,"SEQ",intCnt) = "" Then	
					dblChk = 1
					dblChkSum = dblChkSum + dblChk
				End IF	
			Next
			'모두 저장이 되었다면
			If dblChkSum = 0 Then
				SelectRtn_Detail Col,Row
			'하나라도 저장이 안된것이 있다면
			Else
				'If mlngRowChk = .sprSht.ActiveRow Then
				'Else
					intRtnChk = gYesNoMsgbox("상세내역이 저장완료 되지 않았습니다." & vbcrlf & "이전 작업을 처리하지 않고, 새로운청구요청 데이터를 작업 하시겠습니까 ?","처리안내")
					If intRtnChk = vbYes then 
						'새로운데이터 바인딩
						SelectRtn_Detail Col,Row
					Else
						mobjSCGLSpr.ActiveCell .sprSht, 2, mlngRowChk
								
					End If
				'End If
				
			End If
		end if
	end with	
End Sub

Sub SelectRtn_Detail(ByVal Col, ByVal Row)
	Dim intRtn
	Dim strJOBNO,strPREESTNO
	Dim vntData
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		strPREESTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row)	
		strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row)

		vntData = mobjPDCODEMAND.SelectRtn_MST(gstrConfigXml,mlngRowCnt,mlngColCnt,strPREESTNO,strJOBNO)
		if not gDoErrorRtn ("SelectRtn_MST") then
			If mlngRowCnt < 1 Then
				frmThis.sprSht1.MaxRows = 0
			Else
				mobjSCGLSpr.SetClipBinding .sprSht1, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
			End If
			
   			gWriteText lblStatus2, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht1
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
		end if
	End with
	Field_SettingDTL
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'변경 플래그 설정
	Dim intAMT
	Dim intBALANCE
	Dim intADJAMT
	Dim intCalCul
	Dim strComboList
	Dim vntData_TaxCode

	strComboList =  "차액이월" & vbTab & "마감"
	With frmThis	
		If mstrSelect = false Then
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")   Then
   				intAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",Row)
   				intBALANCE = mobjSCGLSpr.GetTextBinding(.sprSht,"OLDCHARGE",Row)
   				intADJAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT",Row)
   				intCalCul = intBALANCE - intADJAMT
   				If intADJAMT > intBALANCE Then
   					gErrorMsgBox "정산금액은 잔액 보다 클수 없습니다.","입력안내!"
   					mobjSCGLSpr.SetTextBinding .sprSht,"ADJAMT",Row,0
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHARGE",Row,mobjSCGLSpr.GetTextBinding(.sprSht,"OLDCHARGE",Row)
   					mobjSCGLSpr.ActiveCell .sprSht,Col,Row
	   				
   					.txtYEARMON1.focus
					.sprSht.focus 
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,Row ,Row , "", , , , , False
					Exit Sub
   				Else
   					If intADJAMT = intBALANCE Then
   						mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDFLAG",Row, "DI01"
   						mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, "마감"
   					End If
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHARGE",Row, intCalCul
   					.txtYEARMON1.focus
					.sprSht.focus 
   				End If
   				If intADJAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDFLAG",Row, ""
   					'mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, ""
   				End If
   				
			'가장중요한 청구기준이 변경될때 - 값변경은 물론이고, 모든데이터를 가져와 아래그리드에 복제 하여야 한다.
			Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEMANDFLAG")  Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",Row) = "DI02" Then
					'mobjSCGLSpr.SetCellsLock2 .sprSht,false, "MEMO"
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"MEMO",Row,Row,false
					'mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, ""
					mobjSCGLSpr.SetCellTypeComboBox .sprSht,15,15,Row,Row,strComboList ,,80
				Else
					mobjSCGLSpr.SetTextBinding .sprSht,"ADJAMT",Row, mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT",Row)
					mobjSCGLSpr.SetTextBinding .sprSht,"CHARGE",Row, 0
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"MEMO",Row,Row,false
					mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEMO",Row,Row,255,,,,,False
					'mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, ""
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",Row) = "DI01" Then
						mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, "마감"
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"MEMO",Row,Row,false
					End If
				End If
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",Row) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",Row) = "DI04" Then
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"TAXCODE",Row,Row,false
					mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXCODE",Row,Row,255,,,,,False
					mobjSCGLSpr.SetTextBinding .sprSht,"TAXCODE",Row, ""
					
				Else
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"TAXCODE",Row,Row,false
					vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
					
					mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",Row,Row,vntData_TaxCode,,80,,true
					'mobjSCGLSpr.SetCellTypeComboBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"TAXCODE"),mobjSCGLSpr.CnvtDataField(.sprSht,"TAXCODE"),Row,Row,strComboList ,,80
				End If
				
				
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,"-1"
					
				'실질적인 로우복사는 체크가 되어있을때,,, 금액이 들어오면 자동으로 CellChange 이벤트가 일어나고, 그때 로우가 복제 된다.
							
				mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row-1
   			End If
   		
   		End If
   	End with 
   	'실제 Sprsht 변경에 대한 플레그 처리
   	
   	
   
		
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


Sub sprSht1_Change(ByVal Col, ByVal Row)
	'광고주
	Dim strDeptCodeName
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	DIm strTIMCODE
	Dim strTIMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	Dim strComboList
	Dim intAMT
	Dim intADJAMT
	Dim intCalCul
	
	
	strComboList =  "차액이월" & vbTab & "마감"
	with frmThis
		if  Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"CLIENTNAME")  Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A")
			
			If mlngRowCnt = 1 Then	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(1,1)
				mobjSCGLSpr.CellChanged .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"CLIENTCODE"),frmThis.sprSht1.ActiveRow
			Else
				mobjSCGLSpr_ClickProc "sprSht1", Col, .sprSht1.ActiveRow
			End If
			.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus	
			If Row <> .sprSht1.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
			End IF
'팀
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"TIMNAME")  Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"TIMNAME",.sprSht1.ActiveRow)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			
			vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE,strCLIENTNAME,"",strCodeName)
			
			
	
			If mlngRowCnt = 1 Then	
				mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",Row, vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(4,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(5,1)
				mobjSCGLSpr.CellChanged .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"TIMCODE"),frmThis.sprSht1.ActiveRow
			Else
				mobjSCGLSpr_ClickProc "sprSht1", Col, .sprSht1.ActiveRow
			End If
			.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus	
			If Row <> .sprSht1.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
			End IF
		'브랜드
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"SUBSEQNAME")  Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			
			vntData = mobjSCCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,strCLIENTCODE,strCLIENTNAME)
			'msgbox "브랜드명찾아오는갯수" & mlngRowCnt
			If mlngRowCnt = 1 Then	
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(2,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(3,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",Row, vntData(4,1)
				mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",Row, vntData(5,1)
				
				mobjSCGLSpr.CellChanged .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"SUBSEQ"),frmThis.sprSht1.ActiveRow
			Else
				mobjSCGLSpr_ClickProc "sprSht1", Col, .sprSht1.ActiveRow
			End If
			.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
			.sprSht1.Focus	
			If Row <> .sprSht1.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
			End IF
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"DEMANDFLAG")  Then
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"DEMANDFLAG",Row) = "DI02" Then
				'mobjSCGLSpr.SetCellsLock2 .sprSht,false, "MEMO"
				mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"MEMO",Row,Row,false
				mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",Row, ""
				mobjSCGLSpr.SetCellTypeComboBox .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),Row,Row,strComboList ,,80
			Else
				
				mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"MEMO",Row,Row,false
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "MEMO",Row,Row,255,,,,,False
				mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",Row, ""
				If mobjSCGLSpr.GetTextBinding(.sprSht1,"DEMANDFLAG",Row) = "DI01" Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",Row, "마감"
				Else
					mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"MEMO",Row,Row,false
				End If
				mobjSCGLSpr.ActiveCell .sprSht1, Col+1, Row-1
				
			End If
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"DIVAMT")  Then
			mobjSCGLSpr.SetTextBinding .sprSht1,"ADJAMT",Row, mobjSCGLSpr.GetTextBinding( .sprSht1,"DIVAMT",.sprSht1.ActiveRow)
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",Row)
   			intADJAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"ADJAMT",Row)
   			intCalCul = intAMT - intADJAMT
   			mobjSCGLSpr.SetTextBinding .sprSht1,"CHARGE",Row, intCalCul
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"ADJAMT")  Then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",Row)
   			
   			intADJAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"ADJAMT",Row)
   			intCalCul = intAMT - intADJAMT
   			
   			If intADJAMT > intAMT Then
   				gErrorMsgBox "정산금액은 대상금액 보다 클수 없습니다.","입력안내!"
   				mobjSCGLSpr.SetTextBinding .sprSht1,"CHARGE",Row,0
   				mobjSCGLSpr.SetTextBinding .sprSht1,"ADJAMT",Row,intAMT
   				mobjSCGLSpr.ActiveCell .sprSht1,Col,Row
	   			
   				.txtYEARMON1.focus
				.sprSht1.focus 
				'mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,Row ,Row , "", , , , , False
				Exit Sub
   			Else
   				'If intADJAMT = intBALANCE Then
   				'	mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDFLAG",Row, "DI01"
   				'	mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",Row, "마감"
   				'End If
   				mobjSCGLSpr.SetTextBinding .sprSht1,"CHARGE",Row, intCalCul
   				.txtYEARMON1.focus
				.sprSht1.focus 
   			End If
   			If intCalCul <> 0 Then
   				
   				mobjSCGLSpr.SetTextBinding .sprSht1,"DEMANDFLAG",Row, "DI02"
   				mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",Row, "차액이월"
   				mobjSCGLSpr.SetCellTypeComboBox .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),Row,Row,strComboList ,,80
   			Else
   				mobjSCGLSpr.SetTextBinding .sprSht1,"DEMANDFLAG",Row, "DI01"
   				mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",Row, "마감"
   				mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "MEMO",Row,Row,255,,,,,False
   				
   			End If
		End If
		
		
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col, Row
End Sub

'상세내역 쉬트 버튼클릭
Sub sprSht1_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	Dim strCLIENTCODE , strCLIENTNAME,strTIMCODE,strTIMNAME
	Dim strSUBSEQ , strSUBSEQNM
	Dim strCPDEPTCD , strCPDEPTNAME
	Dim strCPEMPNO , strCPEMPNAME
	
	with frmThis

		'광고주
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
			if isArray(vntRet) then
				if strCLIENTCODE = vntRet(0,0) and strCLIENTNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow, trim(vntRet(1,0))
					
				mobjSCGLSpr.CellChanged .sprSht1, .sprSht1.ActiveCol,.sprSht1.ActiveRow
			end if
			.txtYEARMON1.focus()
			.sprSht1.focus()	
			gSetChange
     	'팀
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") then exit Sub
			strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"TIMCODE",.sprSht1.ActiveRow)
			strTIMNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"TIMNAME",.sprSht1.ActiveRow)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME),"", trim(strTIMNAME) ) '<< 받아오는경우
			
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if strTIMCODE = vntRet(0,0) and strTIMNAME = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit
				
				if .sprSht1.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",.sprSht1.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",.sprSht1.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow,  trim(vntRet(5,0))
					
					
					mobjSCGLSpr.CellChanged .sprSht1, .sprSht1.ActiveCol,.sprSht1.ActiveRow
				end if
				.txtYEARMON1.focus()
				.sprSht1.focus()					' 포커스 이동
				gSetChange 
     		end if
     	'브랜드
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN0") Then
		
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN0") then exit Sub
			
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
			strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQ",.sprSht1.ActiveRow)
			strSUBSEQNM = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow)
		
			
			vntInParams = array("", trim(strSUBSEQNM),"", trim(strCLIENTNAME)) '<< 받아오는경우
	
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 413,425)
			if isArray(vntRet) then
				if strSUBSEQ = vntRet(0,0) and strSUBSEQNM = vntRet(1,0) then exit Sub ' 변경된 데이터가 없다면 exit

				if .sprSht1.ActiveRow >0 Then
							mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",.sprSht1.ActiveRow, trim(vntRet(0,0))
							mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow, trim(vntRet(1,0))
							mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, trim(vntRet(2,0))
							mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow, trim(vntRet(3,0))
							mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",.sprSht1.ActiveRow, trim(vntRet(4,0))
							mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",.sprSht1.ActiveRow, trim(vntRet(5,0))
					
							mobjSCGLSpr.CellChanged .sprSht1, .sprSht1.ActiveCol,.sprSht1.ActiveRow
				end if
				.txtYEARMON1.focus()
				.sprSht1.focus()					' 포커스 이동
				gSetChange	
     		end if
     	END IF	
	End with
End Sub 
'txtYEARMON1


Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)

	Dim vntRet, vntInParams
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	DIm strTIMCODE
	Dim strTIMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	
	
	With frmThis
		'PROJECT 그리드
		If sprSht = "sprSht1" Then
			
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"CLIENTNAME") Then
			
				vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME",Row))
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				End IF
				
				.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht1.Focus	
				If Row <> .sprSht1.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
				End If
			'팀
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"TIMNAME") Then
				strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"TIMCODE",.sprSht1.ActiveRow)
				strTIMNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"TIMNAME",.sprSht1.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
					
				vntInParams = array("", trim(strCLIENTNAME),"", trim(strTIMNAME) )  '<< 받아오는경우
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",.sprSht1.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",.sprSht1.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow,  trim(vntRet(5,0))
					mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				End IF
				
				.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht1.Focus	
				If Row <> .sprSht1.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
				End If
			
			'브랜드
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht1,"SUBSEQNAME") Then
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow)
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQ",.sprSht1.ActiveRow)
				strSUBSEQNAME = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow)
					
				vntInParams = array("", trim(strSUBSEQNAME),"", trim(strCLIENTNAME))  '<< 받아오는경우
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",.sprSht1.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow, trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow, trim(vntRet(2,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow, trim(vntRet(3,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",.sprSht1.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",.sprSht1.ActiveRow, trim(vntRet(5,0))
					mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				End IF
				
				.txtYEARMON1.focus	'팝업창에 갔다 오면서 잃어버린 포커스를 다시 시트로 옮겨준다
				.sprSht1.Focus	
				If Row <> .sprSht1.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2, Row
				End If
			End If
		
		
		
		End If	
	
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
	
	'키가 움질일때 바인딩
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprSht_Click frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
End Sub
'=========================================================================================
' UI업무 프로시져 
'=========================================================================================

' 페이지 화면 디자인 및 초기화 
Sub InitPage()
	'서버업무객체 생성	
	set mobjPDCODEMAND	= gCreateRemoteObject("cPDCO.ccPDCODEMAND")
	set mobjPDCOGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet 기본Color 지정
    gSetSheetDefaultColor() 
    
    With frmThis
		'=========================================================================================
		'청구요청 SHEET
		'=========================================================================================
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 26, 0
		mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|YEARMON|PREESTNO|JOBNAME|JOBNO|SUBNO|CLIENTCODE|CLIENTNAME|DIVAMT|ADJAMT|CHARGE|OLDCHARGE|DEMANDFLAG|PRIORITYJOB|MEMO|TAXCODE|ATTR01|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|CREDAY|USENO|SAVEFLAG|CONFIRMFLAG|DATAYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		  "선택|요청월|PREESTNO|JOB명|JOBno.|SUBno.|광고주코드|광고주|견적금액|청구금액|차액|이전차액|청구기준|대표JOB|비고|청구방법|동작구분|팀코드|팀명|브랜드|브랜드코드|견적합의일|작성자|저장구분|승인구분|데이터년월"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|7      |0       |25   |9     |6     |0         |21    |12      |12      |12  |0       |10      |25     |10  |10      |10      |0     |0   |0     |0         |10        |10    |10      |10      |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT|CHARGE|OLDCHARGE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "JOBNAME|JOBNO|SUBNO|CLIENTCODE|CLIENTNAME|PRIORITYJOB|DIVAMT|CHARGE|ATTR01|USENO|SAVEFLAG|CONFIRMFLAG"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "DEMANDFLAG|TAXCODE|MEMO", -1, -1, 255
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|JOBNO|SUBNO|YEARMON|USENO|SAVEFLAG",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|PRIORITYJOB",-1,-1,0,2,false '왼쪽
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|PREESTNO|OLDCHARGE|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DATAYEARMON|CONFIRMFLAG", true
		.sprSht.style.visibility = "visible"
		'=========================================================================================
		'상세내역[미리보기] SHEET
		'=========================================================================================
		gSetSheetColor mobjSCGLSpr, .sprSht1 
		mobjSCGLSpr.SpreadLayout .sprSht1, 26, 0
		mobjSCGLSpr.AddCellSpan  .sprSht1, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,10, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1, "PREESTNO|SEQ|SUBNO|JOBNO|YEARMON|CREDAY|CLIENTCODE|BTN2|CLIENTNAME|TIMCODE|BTN|TIMNAME|SUBSEQ|BTN0|SUBSEQNAME|JOBNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|SAVEFLAG|USENO|CONFIRMFLAG|DATAYEARMON"
		mobjSCGLSpr.SetHeader .sprSht1,         "견적번호|순번|부번|제작번호|청구년월|견적합의일|광고주|광고주명|광고팀|광고팀명|브랜드|브랜드명|JOB명|분할대상금액|청구금액|잔액|청구기준|비고|청구방법|저장구분|작성자|승인구분|데이터년월"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "0       |6   |6   |0       |7       |10        |8   |2|14      |6   |2|18      |6     |2|18    |28   |12          |10      |10  |10      |10  |10      |10      |10    |10      |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN0"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN2"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "CREDAY", -1, -1, 10
		mobjSCGLSpr.ColHidden .sprSht1, "PREESTNO|JOBNO|DATAYEARMON", true
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|JOBNAME|SUBSEQ|SUBSEQNAME|YEARMON", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "DIVAMT|ADJAMT|CHARGE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht1,true, "SAVEFLAG|SUBNO|SEQ|USENO|CHARGE|CONFIRMFLAG"
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "SEQ|SAVEFLAG|USENO|CONFIRMFLAG|SUBNO",-1,-1,2,2,false '가운데
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "MEMO",-1,-1,0,2,false '왼쪽
		.sprSht1.style.visibility = "visible"
		
    End With
	If mstrSelect = false Then
		'그리드 콤보세팅
		Get_COMBO_UPVALUE
		Get_COMBO_VALUE
	End If
	
	'화면 초기값 설정
	InitPageData	
	SelectRtn
End Sub

Sub EndPage()
	set mobjPDCODEMAND = Nothing
	set mobjPDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


' 화면의 초기상태 데이터 설정

Sub InitPageData
	'모든 데이터 클리어
	Dim vntData
	
	gClearAllObject frmThis
	'초기 데이터 설정
	with frmThis
		.sprSht.maxrows = 0
		.txtYEARMON1.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2) '추후 이것으로 대처 임시로 테스트값 연결 하였음
		.txtYEARMON2.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2) '추후 이것으로 대처 임시로 테스트값 연결 하였음
		'.txtYEARMON1.value = "200910"

	vntData = mobjPDCODEMAND.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
		if mlngRowCnt > 0 Then
		mstrDEPTCD = vntData(0,1)
		mstrMANAGER = vntData(1,1)
		end if
   	end if	
		
	End with
	'새로운 XML 바인딩을 생성
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

' 그리드콤보
'자동콤보 변경
Sub Get_COMBO_PVALUE (ByVal blnRow)		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",blnRow,blnRow,vntData_TaxCode,,80,,true
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	
'상단그리드 콤보
Sub Get_COMBO_UPVALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DEMANDFLAG",,,vntData_Demand,,80	
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",,,vntData_TaxCode,,80						
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		
'하단그리드콤보
Sub Get_COMBO_VALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet초기화
		.sprSht.MaxRows = 0

		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "DEMANDFLAG",,,vntData_Demand,,80		
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "TAXCODE",,,vntData_TaxCode,,80

			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

'****************************************************************************************
' 데이터 처리 
'****************************************************************************************
'행추가
Sub imgRowAdd_onclick ()
	with frmThis
		If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI04" Then 
			call sprSht1_Keydown(meINS_ROW, 0)
			mlngRowChk = .sprSht.ActiveRow
		Else 
			gErrorMsgBox "상단선택된 데이터 청구기준이 분할내역 대상이 아닙니다." & vbcrlf & "청구기준을 확인하십시오.","행추가처리안내"
		End If
	End with
End Sub

	'intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht2, cint(meINS_ROW), 0, -1, 1)
Sub sprSht1_Keydown(KeyCode, Shift)

	Dim intRtn
	
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	'if KeyCode = meCR  Or KeyCode = meTab Then
	'	if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht1,"SAVEFLAG")  Then ' 예제 frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DETAIL_BTN")
	'		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(13), cint(Shift), -1, 1)
	'		DefaultValue
	'	End If
	'Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select

	'End If
End Sub

'신규디펄드 값을 생성
Sub DefaultValue
	Dim intCnt
	Dim dblAMT
	Dim dblSumAMT
	Dim dblSetAmt
	with frmThis
		
		
			If mobjSCGLSpr.getTextBinding(.sprSht,"SUBNO",.sprSht.ActiveRow) = "1" And .sprSht1.ActiveRow = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBNO",.sprSht1.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"SUBNO",.sprSht.ActiveRow) 
			End If
			dblSumAMT = 0
			dblSetAmt = 0
			For intCnt = 1 To .sprSht1.MaxRows
				dblAMT = 0
				dblAMT = mobjSCGLSpr.getTextBinding(.sprSht1,"DIVAMT",intCnt) 
				dblSumAMT = dblSumAMT + dblAMT
			Next
			dblSetAmt = mobjSCGLSpr.getTextBinding(.sprSht,"DIVAMT",.sprSht.ActiveRow) - dblSumAMT
			
			If dblSetAmt <= 0 Then
				gErrorMsgBox "청구분배금액은 청구확정 금액을 넘을수 없습니다.","처리안내"
				mobjSCGLSpr.DeleteRow .sprSht1,.sprSht1.ActiveRow
				Exit Sub
			End If
			mobjSCGLSpr.SetTextBinding .sprSht1,"DIVAMT",.sprSht1.MaxRows,dblSetAmt
			mobjSCGLSpr.SetTextBinding .sprSht1,"ADJAMT",.sprSht1.MaxRows,dblSetAmt
			mobjSCGLSpr.SetTextBinding .sprSht1,"CHARGE",.sprSht1.MaxRows,0
		'If mobjSCGLSpr.getTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI04" Then
			mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow) 
			mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"CLIENTNAME",.sprSht.ActiveRow) 
			mobjSCGLSpr.SetTextBinding .sprSht1,"TIMCODE",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"TIMCODE",.sprSht.ActiveRow) 
			mobjSCGLSpr.SetTextBinding .sprSht1,"TIMNAME",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"TIMNAME",.sprSht.ActiveRow) 
			mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"SUBSEQ",.sprSht.ActiveRow) 
			mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"SUBSEQNAME",.sprSht.ActiveRow) 
		'End If
		mobjSCGLSpr.SetTextBinding .sprSht1,"SAVEFLAG",.sprSht1.ActiveRow,"N"
		mobjSCGLSpr.SetTextBinding .sprSht1,"USENO",.sprSht1.ActiveRow,gstrEmpNo
		mobjSCGLSpr.SetTextBinding .sprSht1,"TAXCODE",.sprSht1.ActiveRow,"TA01"
		mobjSCGLSpr.SetTextBinding .sprSht1,"DEMANDFLAG",.sprSht1.ActiveRow,"DI01"
		mobjSCGLSpr.SetTextBinding .sprSht1,"MEMO",.sprSht1.ActiveRow,"마감"
		mobjSCGLSpr.SetTextBinding .sprSht1,"YEARMON",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"YEARMON",.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding .sprSht1,"PREESTNO",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding .sprSht1,"CREDAY",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"CREDAY",.sprSht.ActiveRow)
		mobjSCGLSpr.SetTextBinding .sprSht1,"CONFIRMFLAG",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"CONFIRMFLAG",.sprSht.ActiveRow) 
		mobjSCGLSpr.SetTextBinding .sprSht1,"DATAYEARMON",.sprSht1.ActiveRow,mobjSCGLSpr.getTextBinding(.sprSht,"DATAYEARMON",.sprSht.ActiveRow)  
		
	End With
End Sub
'조회
Sub SelectRtn
	Dim vntData
   	Dim i, strCols
    Dim strCHK
    
    
    Dim intCnt
    Dim vntData_TaxCode
    
	'On error resume next
	with frmThis
		'Long Type의 ByRef 변수의 초기화
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		mlngTaxRowCnt=clng(0)
		mlngTaxColCnt=clng(0)
		
		vntData = mobjPDCODEMAND.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value,.txtYEARMON2.value,mstrDEPTCD,mstrMANAGER)
		
		if not gDoErrorRtn ("SelectRtn") then
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngTaxRowCnt,mlngTaxColCnt,"PD_TAXCODE")	
		
			if mlngRowCnt > 0 Then
				
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				'mobjSCGLSpr.ColHidden .sprSht,strCols,true
				
   				Call sprSht_Click(2,.sprSht.ActiveRow)
   				'Fleld_Setting	
   				For intCnt = 1 To .sprSht.MaxRows
   					If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI04" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"TAXCODE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"TAXCODE",intCnt,intCnt,false
					End If
   				Next
   			Else
   			.sprSht.MaxRows = 0
   			.sprSht1.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "건의 자료가 검색" & mePROC_DONE
   		end if
   	end with
	'Fleld_Setting		
End Sub
'조회후 컬럼 데이터 Type 일치화
Sub Fleld_Setting
	Dim intCnt,intCnt2
	Dim vntData_TaxCode
	Dim strComboList
	
	strComboList =  "차액이월" & vbTab & "마감"
	with frmThis
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")	
		For intCnt = 1 To .sprSht.MaxRows
			'세금계산서 콤보 및 텍스트 형태 처리
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI04" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"TAXCODE",intCnt,intCnt,false
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXCODE",intCnt,intCnt,255,,,,,False
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"TAXCODE",intCnt,intCnt,false
				mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",intCnt,intCnt,vntData_TaxCode,,80,,true
			End If
			
			'비고 콤보 및 텍스트 형태 처리
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI02" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"MEMO",intCnt,intCnt,false
				mobjSCGLSpr.SetCellTypeComboBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO"),mobjSCGLSpr.CnvtDataField(.sprSht,"MEMO"),intCnt,intCnt,strComboList ,,80
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"MEMO",intCnt,intCnt,false
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEMO",intCnt,intCnt,255,,,,,False
			End If
		Next
	End with
End Sub
Sub Field_SettingDTL
	Dim intCnt,intCnt2
	Dim vntData_TaxCode
	Dim strComboList
	
	strComboList =  "차액이월" & vbTab & "마감"
	with frmThis
		For intCnt2 = 1 To .sprSht1.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"DEMANDFLAG",intCnt2) = "DI02" Then
				
				mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"MEMO",intCnt2,intCnt2,false
				
				mobjSCGLSpr.SetCellTypeComboBox .sprSht1,mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),mobjSCGLSpr.CnvtDataField(.sprSht1,"MEMO"),intCnt2,intCnt2,strComboList ,,80
				
			Else
				
				mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"MEMO",intCnt2,intCnt2,false
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "MEMO",intCnt2,intCnt2,255,,,,,False
				
			End If
		Next
	End with
End Sub

' 저장
Sub ProcessRtn ()
	Dim vntData
	Dim intRtn
	Dim vntData_Hdr
	Dim strPREESTNO
	with frmThis
		if DataValidation =false then exit sub 	
		mobjSCGLSpr.SetTextBinding .sprSht,"CHK",.sprSht.ActiveRow, "-1"
		mobjSCGLSpr.CellChanged .sprSht, 1, .sprSht.ActiveRow
		
		vntData_Hdr = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|PREESTNO|JOBNO|SUBNO|CLIENTCODE|TIMCODE|SUBSEQ|CREDAY|DIVAMT|ADJAMT|CHARGE|OLDCHARGE|PRIORITYJOB|DEMANDFLAG|MEMO|TAXCODE|USENO|SAVEFLAG|CONFIRMFLAG|DATAYEARMON")
		strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"PREESTNO|SEQ|SUBNO|JOBNO|YEARMON|CREDAY|SUBSEQ|SUBSEQNAME|TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|JOBNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|SAVEFLAG|USENO|CONFIRMFLAG|DATAYEARMON")		
		
		
		if  not IsArray(vntData)  then 
			gErrorMsgBox "변경된 " & meNO_DATA,"저장안내"
			exit sub
		End If
		
		'동시에 상단,하단을 저장하되 상단은 하위내역이 저장될때 같이 저장된다.
		intRtn = mobjPDCODEMAND.ProcessRtn(gstrConfigXml,vntData,vntData_Hdr,strPREESTNO)	
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			gErrorMsgBox " 자료가 저장" & mePROC_DONE,"저장안내" 	
			'다같이 조회하면 결국 이미 처리한 상단만 저장할 내역이 사라지기 때문에 조회하지 않는다.
			'SelectRtn
			SelectRtn_Detail .sprSht.activeCol,.sprSht.activeRow
			SelectRtn
		End If
	End with
End Sub


Sub ProcessRtn_HDR

	Dim vntData
	Dim intRtn
	Dim vntData_Hdr
	Dim vntInParams
	Dim vntRet
	Dim intCnt
	Dim dblChk
	
	with frmThis
	
		if DataValidationHDR = false then exit sub 	
		
		dblChk = 0
		
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				dblChk = dblChk +1
			End If		
		Next
		
		If dblChk = 0 Then
			gErrorMsgBox "청구요청할 데이터를 선택 하십시오.","처리안내"
			Exit Sub
		End If
		 
		vntData_Hdr = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|PREESTNO|JOBNO|SUBNO|CLIENTCODE|TIMCODE|SUBSEQ|CREDAY|DIVAMT|ADJAMT|CHARGE|OLDCHARGE|PRIORITYJOB|DEMANDFLAG|MEMO|TAXCODE|USENO|SAVEFLAG|CONFIRMFLAG|DATAYEARMON")
			if  not IsArray(vntData_Hdr)  then
			Else 
				intRtn = mobjPDCODEMAND.ProcessRtn_HDR(gstrConfigXml,vntData_Hdr)	
			End If
		
		'동시에 상단,하단을 저장하되 상단은 하위내역이 저장될때 같이 저장된다.
		
		
		'보내어지는 키는 년월,USENO,
		
		if not gDoErrorRtn ("ProcessRtn_HDR") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
		Else
			gErrorMsgBox "청구요청 저장에 실패 하였습니다! 관리자에게 문의 하십시오.","처리안내"
			Exit Sub
		End If
		'년월과 현재 로그인 사용자를 보낸다. - 사용자가 해당월에 작성한 내역만 보이게 된다.
		'vntInParams = array(Trim(.txtYEARMON1.value),gstrEmpNo)
		vntInParams = array(Trim(.txtYEARMON1.value),Trim(.txtYEARMON2.value),gstrEmpNo)
		vntRet = gShowModalWindow("PDCMDEMANDPOP.aspx",vntInParams , 1149,650)
		If vntRet = "SAVETRUE" Then
			SelectRtn
		End If
	End With
End Sub
'------------------------------------------
' 데이터 처리를 위한 데이타 검증
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim dblSumAmt
   	Dim dblAMT
	'On error resume next
	with frmThis
  	
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		dblSumAmt = 0
		
   		for intCnt = 1 to .sprSht1.MaxRows
   			'Sheet 필수 입력사항
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTCODE",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"YEARMON",intCnt) = "" Then 
				gErrorMsgBox intCnt & " 번째 행의 내용 기입여부 를 확인하십시오","저장오류"
				Exit Function
			End if
			dblAMT = 0
			dblAMT = mobjSCGLSpr.getTextBinding(.sprSht1,"DIVAMT",intCnt) 
			dblSumAmt = dblSumAmt + dblAMT
			'금액 오류사항
		next
   		If mobjSCGLSpr.getTextBinding(.sprSht,"DIVAMT",.sprSht.ActiveRow) < dblSumAmt Then
   			gErrorMsgBox "분할대상금액의 합은 견적금액 을 초과할수 없습니다","저장오류"
   			Exit Function
   		End If
   	End with
   	
	DataValidation = true
End Function


Function DataValidationHDR ()
	DataValidationHDR = false
	
	Dim intCnt
	with frmThis
	
	For intCnt = 1 To .sprSht.MaxRows
		If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI04" Then
			
			Else
				If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "" Then
					gErrorMsgBox intCnt & " 번째 행의 청구기준을 기입하십시오","저장오류"
					Exit Function
				End If
				if mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI01" Or  mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",intCnt) = "DI02" Then
					If mobjSCGLSpr.GetTextBinding(.sprSht,"TAXCODE",intCnt) = "" Then
						gErrorMsgBox intCnt & " 번째 행의 세금계산서를 기입하십시오","저장오류"
						Exit Function
						
					End If	
					If mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",intCnt) = "차액이월" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"CHARGE",intCnt)  = 0 Then
						gErrorMsgBox intCnt & " 번째 행의 차액이월 금액을 기입하십시오","저장오류"
						Exit Function
					End If
				End If
			End If
		End If
	Next
	End With
	
	DataValidationHDR = true
	
End Function
'삭제
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2,lngCnt
	dim strYEARMON
	Dim dblSEQ
	Dim strPREESTNO
	Dim strJOBNO
	Dim strITEMCODE
	Dim strDETAILYNFLAG
	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "삭제할 자료" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		'PREESTNO,ITEMCODESEQ
		'선택된 자료를 끝에서 부터 삭제
		lngCnt =0
		intRtn2 = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i)) <> ""  Then
		
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i))
				strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNO",vntData(i))
				dblSEQ = CDBL(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i)))
				
				intRtn2 = mobjPDCODEMAND.DeleteRtn(gstrConfigXml,strPREESTNO, dblSEQ, strJOBNO)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				gWriteText lblStatus2, "자료가 삭제되었습니다."
   			End IF
		next
		'헤더재계산
	
		If lngCnt <> 0 Then
			gOkMsgBox "자료가 삭제되었습니다.","삭제안내!"
		End If
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht1
	End with
	err.clear
End Sub

Sub DeleteRtnProc
	Dim vntData
	Dim intCnt, intRtn, i
	'삭제Key 설정
	Dim strYEARMON
	Dim strPREESTNO
	Dim strJOBNO
	Dim dblSUBNO
	Dim strDEMANDFLAG
	Dim strDelChk
	
	Dim strDESCRIPTION
	with frmThis

	strDESCRIPTION = ""
		
		intRtn = gYesNoMsgbox("자료를 삭제하시겠습니까?","자료삭제 확인")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		'선택된 자료를 끝에서 부터 삭제
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)
				strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",i)
				dblSUBNO = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SUBNO",i))	
				strDEMANDFLAG =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",i)
				
				intRtn = mobjPDCODEMAND.DeleteRtnProc(gstrConfigXml,strYEARMON,strPREESTNO,strJOBNO,dblSUBNO,strDEMANDFLAG)
			
				IF not gDoErrorRtn ("DeleteRtnProc") then
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		IF not gDoErrorRtn ("DeleteRtnProc") then
			gWriteText lblStatus, intCnt & "건이 삭제" & mePROC_DONE
   		End IF
		'선택 블럭을 해제
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub
-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0">
				<TR valign="top">
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF"
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
											<td class="TITLE">청구요청</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 326px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="처리중입니다." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="굴림"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="세금계산서조회 조건을 삭제합니다" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
												width="80">등록월</TD>
											<TD class="SEARCHDATA" ><INPUT class="INPUT" id="txtYEARMON1" title="등록월" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON1)" size="9" name="txtYEARMON1"> ~ 
													<INPUT class="INPUT" id="txtYEARMON2" title="등록월" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON2)" size="9" name="txtYEARMON2"></TD>
											<td class="SEARCHDATA" width="312" align="right"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="자료를 검색합니다." src="../../../images/imgQuery.gIF"
													align="absMiddle" border="0" name="imgQuery"> <IMG id="imgRowDelUp" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF"
													align="absMiddle" border="0" name="imgRowDelUp"> <IMG id="imgDivDemand" onmouseover="JavaScript:this.src='../../../images/imgDivPreDemandOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDivPreDemand.gIF'" height="20" alt="청구요청자료를 미리보기합니다." src="../../../images/imgDivPreDemand.gIF"
													align="absMiddle" border="0" name="imgDivDemand"> <IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF"
													align="absMiddle" border="0" name="imgExcel">
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR valign="top">
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="굴림"></FONT></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
						<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="42439">
							<PARAM NAME="_ExtentY" VALUE="7594">
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
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="300" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="255" background="../../../images/back_p.gIF"
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
											<td class="TITLE">청구요청 상세내역[승인요청자료 미리보기]&nbsp;<span id="strMsgBox"></span></td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'"
													height="20" alt="자료입력을 위해 행을추가합니다." src="../../../images/imgRowAdd.gIF" border="0"
													name="imgRowAdd"></TD>
											<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'"
													height="20" alt="선택한 행을삭제합니다." src="../../../images/imgRowDel.gIF" border="0" name="imgRowDel"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="자료를 저장합니다."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD><IMG id="imgExcel2" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="자료를 엑셀로 받습니다." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel2"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 218px; HEIGHT: 4px" colSpan="2"><FONT face="굴림"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
						<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="42439">
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
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus2"><FONT face="굴림"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TBODY></TABLE></form>
	</body>
</HTML>

