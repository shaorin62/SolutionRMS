<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPREEXEVOCH.aspx.vb" Inherits="PD.PDCMPREEXEVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ޱ� ��� ó��</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/12/19 By KTY
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLEs.CSS" type="text/css" rel="STYLESHEET">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">

'�������� ����
Dim mlngRowCnt,mlngColCnt
Dim mobjPDCOPREEXEVOCH
Dim mobjPDCOGET
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim mstrPROCESS
Dim mstrSTAY

mstrSTAY = True

mstrGUBUN = "B"
mstrPROCESS = ""
mstrCheck = True

CONST meTAB = 9
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'�������� ��ư �����
Sub Set_delete(byVal strmode)
	With frmThis
		If .rdT.checked = True Then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		Else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		End If
	End With
End Sub

'-----------------------------------
'��ư Ŭ�� �̺�Ʈ
'-----------------------------------
'��ȸ��ư
Sub imgQuery_onclick
	If frmThis.txtYEARMON.value = "" Then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'������ư Ŭ��
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = True
		mobjSCGLSpr.ExcelExportOption = True
 
		mobjSCGLSpr.ExportExcelFile .sprSht_OUT
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'��ǥ���� Ŭ��
Sub ImgvochCre_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� Ŭ��
Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� ���� Ŭ��
Sub imgVochDelco_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'������ǥ����Ŭ��
Sub ImgErrVochDel_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'--�����ư Ŭ��
Sub ImgSUMMApp_onclick()
	Dim intRtn
	
	with frmThis
		if .cmbSETTING.value = "" then
			gErrorMsgBox "�����Ͻ� �÷� ���� �����ϴ�. ","�������"
			exit sub
		end if 
		
		if mstrGUBUN = "P"  then 
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_SUSU)
		elseif mstrGUBUN = "B"  then  
			intRtn = gYesNoMsgbox("üũ�Ͻ� �������� ������ ����˴ϴ� �����Ͻðڽ��ϱ�? ","ó���ȳ�!")
			IF intRtn <> vbYes then exit Sub
			
			settingRowChange (.sprSht_OUT)
		end if
	End With
End Sub

sub settingRowChange(sprsht)
	Dim strSETTINGDATA
	Dim intCnt 
	Dim i ,j

	with frmThis
		intCnt = 0
		
		for j = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) <> "DTL" Then
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",j) = "1" Then
					intCnt = intCnt + 1
				End if 
			END IF
		next
		
		if right(sprsht.ID,3) <> "DTL" Then
			if intCnt = 0 Then
				gErrorMsgBox "üũ�� �����Ͱ� �����ϴ�. �����Ͻ� �����͸� üũ�ϼ���. ","�������"
				EXIT SUB
			End if
		End if
		
		strSETTINGDATA = ""
		strSETTINGDATA = .txtSUMM.value
		
		for i = 1 to sprsht.MaxRows
			if right(sprsht.ID,3) = "DTL" Then
				mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
			ELSE 
				If mobjSCGLSpr.GetTextBinding(sprsht,"CHK",i) = "1" Then
					mobjSCGLSpr.SetTextBinding sprsht,.cmbSETTING.value,i, strSETTINGDATA
				End if
			End if 
		next 
	End with
end sub
'-----------------------------------------------------------------------------------------
' �������˾�(��ȸ)
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)	
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then Exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
   		Dim strGBN
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),"A")
			
			If Not gDoErrorRtn ("txtCLIENTNAME_onkeydown") Then
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

'�Ϸ�üũ
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'����üũ
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'���θ��üũ
Sub rdPRO_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'�����θ��üũ
Sub rdNONPRO_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUMM_onchange
	Dim blnByteCHk
	Dim intRtn
	blnByteCHk =  checkBytes(frmThis.txtSUMM.value)
	
	If blnByteCHk  > 23 Then
		intRtn = gYesNoMsgbox("������ ũ��� 23Byte �� ������ �����ϴ�. �ʱ�ȭ �Ͻðڽ��ϱ�?","ó���ȳ�!")
		If intRtn <> vbYes Then Exit Sub
		frmThis.txtSUMM.value = ""
	End If
End Sub

function checkBytes(expression)
	Dim VLength
	Dim temp
	Dim EscTemp
	Dim i
	VLength=0
	
	temp = expression
	If temp <> "" Then
		for i=1 to len(temp) 
			If mid(temp,i,1) <> escape(mid(temp,i,1))  Then
				EscTemp=escape(mid(temp,i,1))
				If (len(EscTemp)>=6) Then
					VLength = VLength +2
				Else
				VLength = VLength +1
				End If
			Else
				VLength = VLength +1
			End If
		Next
	End If

	checkBytes = VLength
end function

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
'-----------------------------------
' SpreadSheet ü����
'-----------------------------------
Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
	With frmThis
		If	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"paycode") Then
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "G" Then
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "409903"
			ElseIf mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "T" Then
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "410999"
			ElseIf mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "C" Then
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "410999"
			Else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "410904"
			End If 
		End If 
		
		If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"FROMDATE",Row,Row,False
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"TODATE",Row,Row,False
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,False
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,False
		End If
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

'-----------------------------------
' SpreadSheet Ŭ��
'-----------------------------------
Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		If Col = 1 and Row = 0 Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck

			If mstrCheck = True Then  
				for intCnt = 1 To .sprSht_OUT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
				Next    
				mstrCheck = False
			ElseIf mstrCheck = False Then 
				mstrCheck = True
			End If
		End If 
	End With
End Sub 

'-----------------------------------
' SpreadSheet ���� Ŭ��
'-----------------------------------
sub sprSht_OUT_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUT, ""
		End If
	End With
end sub

Sub sprSht_OUT_Keyup(KeyCode, Shift)
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
	
	With frmThis
		If .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"AMT") Or .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"VAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"AMT") Or .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"VAT") Then
				strCOLUMN = "AMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_OUT,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_OUT,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"AMT")) Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"VAT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_OUT,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		Else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_OUT.MaxRows >0 Then
			If .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"AMT")  Or .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"VAT") Then
				If .sprSht_OUT.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_OUT,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_OUT,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						Exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_OUT,strCol,vntData_row(j))
						End If
					Next
					
					.txtSELECTAMT.value = strSUM
				End If
				
			Else
				.txtSELECTAMT.value = 0
			End If
		Else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjPDCOPREEXEVOCH	= gCreateRemoteObject("cPDCO.ccPDCOPREEXEVOCH")
	Set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	Set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    
    Dim strComboPREPAYMENT
	Dim strSemuComboListB, strSemuComboListA
	
	gSetSheetDefaultColor
	
    With frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strSemuComboListB =  "B5" & vbTab & "BR" & vbTab & "BH"
		strSemuComboListA =  "  " & vbTab & "A0" & vbTab & "AI" & vbTab & "A8" & vbTab & "AZ"
		
		'**************************************************
		'���� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 32, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PAYCODE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������|����|��������|�뺯����|������|���޹��|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|���ⱸ��| AMTGBN|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1",  "   4|       8|        10|    15|    15|  20|       5|         8|  10|    10|       7| 5|       8|    10|   0|       7|       7|     8|       0|         0|             0|             0|      20|      7|      7|       9|       0|        10|    0|      10|      0|       10"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,False '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,False '����
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | JOBBASE | TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN | GFLAG | JOBBASE | DUEDATE | AMTGBN | PAYCODE | PREPAYMENT | FROMDATE | TODATE", True 
		
	End With
	pnlTab_gen.style.visibility = "visible" 
    
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
	
'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	With frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht_OUT.MaxRows = 0
		.txtYEARMON.focus	
		
		Get_COMBO_VALUE
		
		'ó���� ���� ���� ����
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
	End With
End Sub

Sub EndPage()
	set mobjPDCOPREEXEVOCH = Nothing
	set mobjSCCOGET = Nothing
	Set mobjPDCOGET = Nothing
	
	gEndPage	
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht_OUT.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		vntData = mobjPDCOPREEXEVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If Not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUT, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub SelectRtn (strVOCH_TYPE)	
	With frmThis
		.sprSht_OUT.MaxRows = 0
		
		CALL SelectRtn_OUT()
		
		mstrSTAY = True
   	End With
End Sub

Sub SelectRtn_OUT ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
    Dim strPROGBN
	
	With frmThis
		.sprSht_OUT.MaxRows = 0
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		If .rdT.checked Then
			strGBN = .rdT.value
		ElseIf .rdF.checked Then
			strGBN = .rdF.value
		ElseIf .rdE.checked Then
			strGBN = .rdE.value
		End If 
		
		If .rdPRO.checked Then
			strPROGBN = .rdPRO.value
		Else
			strPROGBN = .rdNONPRO.value
		End If

		vntData = mobjPDCOPREEXEVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
												   strCLIENTCODE, strCLIENTNAME, _
												   strGBN, strPROGBN)

		If Not gDoErrorRtn ("SelectRtn_OUT") Then
			If mlngRowCnt > 0 Then
				mstrGUBUN = "B"
				
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked Then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"DEMANDDAY",intCnt,intCnt,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"DUEDATE",intCnt,intCnt,False
					ElseIf .rdF.checked or .rdE.checked Then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"DEMANDDAY",intCnt,intCnt,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"DUEDATE",intCnt,intCnt,False
					End If
					
				
					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"FROMDATE",intCnt,intCnt,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,False,"TODATE",intCnt,intCnt,False
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,False
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,False
					End If	
				Next
				
				Call AMT_SUM (.sprSht_OUT)
			Else
				.txtSELECTAMT.value = 0
   			End If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		End If
   	End With
End Sub

Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		Else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

Function DataValidation_OUT ()
	DataValidation_OUT = False	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End If
		Next
	End With
	DataValidation_OUT = True
End Function

'�������
Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn
	With frmThis
		If mstrPROCESS = "Create" Then
			If Not .rdF.checked Then
				gErrorMsgBox "�̿Ϸ���ȸ�� �����մϴ�.","�����׻���"
				Exit Sub
			End If 
		End If 
		
		If mstrPROCESS = "Delete" Then
			If Not .rdT.checked Then
				gErrorMsgBox "�Ϸ���ȸ�� �����մϴ�.","�����׻���"
				Exit Sub
			End If 
		End If 
		
		If mstrSTAY Then 
			mstrSTAY = False
			If strVOCH_TYPE = "B" Then
				If DataValidation_OUT =False Then Exit Sub
				CALL ProcessRtn_OUT()
			End If
		Else
			gErrorMsgBox "��ǥó�� �������Դϴ�.","��ǥó�� �ȳ�"
		End If
   	End With
End Sub

'�������
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim strCUSTOMERCODE
	Dim intColFlag, bsdiv, intMaxCnt
	
	With frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PAYCODE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN | TRANSRANK")
		'ó�� ������ü ȣ��
		If  Not IsArray(vntData_ProcesssRtn) Then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			Exit Sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		
		If .rdPRO.checked Then
			IF_GUBUN = "RMS_0014"
		Else
			IF_GUBUN = "RMS_0014"
		End If
		
		'�ִ밪
		intColFlag = 0
		For intMaxCnt = 1 To .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TRANSRANK",intMaxCnt))
				If intColFlag < bsdiv Then
					intColFlag = bsdiv
				End If
			End If
		Next
		
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		
		If mstrPROCESS = "Create" Then
			For intCnt = 1 To .sprSht_OUT.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" Then		
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "N"

					If strIF_CNT = "1" Then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									"" + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									"" + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					Else
						strHSEQ = strHSEQ + 1
						strISEQ = 1
						
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									"" + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									"" + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					End If
				End If 
			Next
		ElseIf mstrPROCESS = "Delete" Then
			For intCnt = 1 To .sprSht_OUT.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" Then		
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "Z"
		
					If strIF_CNT = "1" Then

						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									"" + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									"" + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					Else
						strHSEQ = strHSEQ + 1
						
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									"" + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									"" + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt),"<","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					End If
				End If 
			Next
		End If 
		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)

   	End With
End Sub

'---------------------------------------------------
' ��ǥ���� �� ��ǥ��ȣ �޾ƿ��� �� ���� RMS������Ʈ
'---------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	gFlowWait meWAIT_ON
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
		If mstrPROCESS ="Create" Then
			strRETURNLIST = replace(strRETURNLIST,"'"," ")
			If mstrGUBUN = "B" Then
				intRtn = mobjPDCOPREEXEVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN)
			End If 
			

			If Not gDoErrorRtn ("ProcessRtn") Then
				'��� �÷��� Ŭ����
				If mstrGUBUN = "B" Then
					mobjSCGLSpr.SetFlag  .sprSht_OUT, meCLS_FLAG
				End If
				
				If intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				Else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			End If
   		ElseIf mstrPROCESS ="Delete" Then
   			intRtn = mobjPDCOPREEXEVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN)
   			
   			If Not gDoErrorRtn ("VOCHDELL") Then
				'��� �÷��� Ŭ����
				If mstrGUBUN = "B" Then
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				End If
				
				If intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			End If
   		End If 
   		If mstrGUBUN = "B" Then
			.sprSht_OUT.focus()
		End If
	End With
	gFlowWait meWAIT_OFF
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
   	
	With frmThis
   		If Not .rdE.checked Then
			gErrorMsgBox "������ȸ�� �����մϴ�.","�����׻���"
			Exit Sub
		End If 
		
		If mstrGUBUN = "B" Then
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN")
		End If
		
		'ó�� ������ü ȣ��
		If  Not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			Exit Sub
		End If
		
		intRtn = mobjPDCOPREEXEVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		If Not gDoErrorRtn ("DeleteRtn") Then
			'��� �÷��� Ŭ����
			If mstrGUBUN = "B" Then
				mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
			End If
			
			If intRtn > 0 Then
			gErrorMsgBox "���� ��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
			End If
			
			SelectRtn(mstrGUBUN)
   		End If
   	End With
End Sub

'-----------------------------------------
'��ǥ ���� ����
'-----------------------------------------
Sub DeleteRtn (strGUBUN)
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strVOCHNO
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
	
		If mstrGUBUN = "B"  Then
			If .sprSht_OUT.MaxRows = 0 Then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUT.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					lngchkCnt = lngchkCnt + 1
				End If
			Next
			If lngchkCnt = 0 Then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				Exit Sub
			End If
		End If
		
		intRtn = gYesNoMsgbox("���������� SAP���� ���ε� ��ǥ�� SAP���� ����Ͽ� RMS�ʿ��� ������ �� ������ RMS�� ��ǥ�� ������ �����Ҷ� ����մϴ�. " & vbCrlf & "  " & vbCrlf & " ��ǥ�� ������ �����Ͻðڽ��ϱ�?","�������� Ȯ��")
		If intRtn <> vbYes Then Exit Sub
		
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		If mstrGUBUN = "B"  Then
			for i = .sprSht_OUT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
					
					intRtn = mobjPDCOPREEXEVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN)
					
					If Not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		End If
		
		If Not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
			SelectRtn (strGUBUN)
	End With
	err.clear	
End Sub


		</script>
		<script language="javascript">
		//##########################################################################################################################################
		//******************************************��1) frmSapCon ���� ������ �� �̿��Ͽ� Submit �ϴ� �Լ�
		//##########################################################################################################################################

		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {
		
			//���
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			
			//dtl
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			
			window.frames[0].document.forms[0].submit();
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���ޱ� ���ó��</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%" height="93%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
												width="70">&nbsp;�����
											</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="8" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="75">&nbsp;������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 142px; HEIGHT: 22px"
													maxLength="100" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<td class="SEARCHDATA" align="right" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</td>
										</TR>
										<TR>
											<TD class="SEARCHLABEL">����
											</TD>
											<TD class="SEARCHDATA">
												<INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;�Ϸ�&nbsp;
												<INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" value="rdF" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')"
													CHECKED>&nbsp;�̿Ϸ�&nbsp; <INPUT id="rdE" title="������ǥ ������ȸ" type="radio" value="rdE" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;����&nbsp;
											</TD>
											<TD class="SEARCHLABEL">����
											</TD>
											<TD class="SEARCHDATA">
												<INPUT id="rdPRO" title="���θ��" type="radio" value="rdPRO" name="rdPROGBN">&nbsp;���θ��&nbsp;
												<INPUT id="rdNONPRO" title="�����θ��" type="radio" CHECKED value="rdNONPRO" name="rdPROGBN">&nbsp;�����θ��&nbsp;
											</TD>
											<TD class="SEARCHLABEL">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 15px"></TD>
							</TR>
							<TR>
								<TD vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="HEIGHT: 26px" align="left" width="100%">
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
													<TR>
														<td><IMG id="ImgvochCre" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/ImgvochCre.gIF" border="0" name="ImgvochCre"></td>
														<td><IMG id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/imgVochDel.gIF" border="0" name="imgVochDel"></td>
														<td><IMG id="ImgErrVochDel" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																height="20" alt="������ǥ �� �����մϴ�." src="../../../images/ImgErrVochDel.gIF" border="0"
																name="ImgErrVochDel"></td>
														<td><IMG id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																height="20" alt="��ǥ�� ������ �����մϴ�." src="../../../images/imgVochDelco.gIF" border="0"
																name="imgVochDelco" title="SAP���� ���������Ͽ� RMS���� ������ �� ������ RMS��ǥ�� ������ �����Ѵ�."></td>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHDATA" width="90" onclick="vbscript:Call gCleanField(txtSUMM,'')">
												<select id="cmbSETTING" style="WIDTH: 90px">
													<OPTION value="" selected>����</OPTION>
													<OPTION value="POSTINGDATE">��ǥ����</OPTION>
													<OPTION value="SUMM">����</OPTION>
													<OPTION value="BA">�������</OPTION>
													<OPTION value="COSTCENTER">�ڽ�Ʈ����</OPTION>
													<OPTION value="SEMU">�����ڵ�</OPTION>
													<OPTION value="DEMANDDAY">���ޱ���</OPTION>
													<OPTION value="DUEDATE">�Աݱ���</OPTION>
													<OPTION value="ACCOUNT">��������</OPTION>
													<OPTION value="DEBTOR">����</OPTION>
													<OPTION value="PREPAYMENT">�����ݱ���</OPTION>
													<OPTION value="SUMMTEXT">����TEXT</OPTION>
												</select></TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUMM" title="��������" style="WIDTH: 402px; HEIGHT: 21px" size="61"
													name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													title="���並 �ϰ� �����մϴ�" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
													height="20" alt="���並 �ϰ� �����մϴ�" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0"
													name="ImgSUMMApp">
											</TD>
											<TD align="right"><INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
													accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"><INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab_gen" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_OUT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31882">
											<PARAM NAME="_ExtentY" VALUE="13070">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
				</TR>
			</TABLE>
			<P>
				<!--Input Define Table End--> </TD></TR> 
				<!--Top TR End--> </TABLE> 
				<!--Main End--></P>
		</FORM>
		</TR></TABLE><iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 100%; HEIGHT: 300px" name="frmSapCon"
			src="../../../PD/WebService/VOCHWEBSERVICE.aspx"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
