<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMFEEVOCH.aspx.vb" Inherits="MD.MDCMFEEVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>FEE ��ǥ����</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/24 By Ȳ����
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
Dim mobjMDCMFEEVVOCH
Dim mobjMDCOGET
Dim mobjMDCOVOCH
Dim mlngRowCnt,mlngColCnt
Dim mstrCheck
Dim mstrGUBUN
Dim vntData_ProcesssRtn
Dim mstrSTAY
Dim mstrPROCESS

mstrPROCESS = ""
mstrSTAY = TRUE
'FEE ����
mstrGUBUN = "F"
mstrCheck=True

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
		IF .rdT.checked = TRUE then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
'��ȸ��ư Ŭ�� �̺�Ʈ
Sub imgFEE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn_FEELIST (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� ��ư Ŭ�� �̺�Ʈ
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'������ǥ ���� ��ư
Sub imgDelete_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'������ǥ ���� ��ư
Sub imgVochDelco_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn_FEELIST(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn_FEELIST(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'����üũ
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn_FEELIST(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'�����ư Ŭ�� �̺�Ʈ
Sub btnApp_onclick ()
	Dim intCnt
	Dim intRtn
	Dim strDEMANDDAY
	With frmThis
		IF .sprSht.MaxRows = 0 then 
			gErrorMsgBox "������ �����Ͱ� �����ϴ�..","ó���ȳ�!"
			Exit Sub
		End if 
		
		if .txtDEMANDDAY.value = "" then
			gErrorMsgBox "������ û������ �Է��Ͻÿ�.","ó���ȳ�!"
			Exit Sub
		end if
		
		strDEMANDDAY = .txtDEMANDDAY.value
		intRtn = gYesNoMsgbox("���õ� ��ǥ�� û������ ���� �Ͻðڽ��ϱ�?","���� Ȯ��")
		IF intRtn <> vbYes then exit Sub

			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
					mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
				End If
			Next
	End With
End Sub

'���� �ؽ�Ʈ �ڽ��� ����Ʈ�� ����Ͽ� �����Ѵ�.
function checkBytes(expression)
	dim VLength
	dim temp
	dim EscTemp
	dIM i
	VLength=0
	
	temp = expression
	if temp <> "" then
		for i=1 to len(temp) 
			if mid(temp,i,1) <> escape(mid(temp,i,1))  then
				EscTemp=escape(mid(temp,i,1))
				if (len(EscTemp)>=6) then
					VLength = VLength +2
				else
				VLength = VLength +1
				end if
			else
				VLength = VLength +1
			end if
		Next
	end if
	checkBytes = VLength
end function

'****************************************************************************************
'  �޷�
'****************************************************************************************
Sub imgDEMANDDAY_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgDEMANDDAY,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
Sub txtYEARMON_onkeydown
	'or window.event.keyCode = meTAB ���϶��� �ƴ� �����϶��� ��ȸ
	If window.event.keyCode = meEnter Then
		DateClean
		SelectRtn_FEELIST(mstrGUBUN)
		frmThis.sprSht.focus()
	End If
End Sub

'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_FEELIST_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 AND Col > 1 then
			'mstrGrid = TRUE
			Call SelectRtn (Col, Row)
			'mstrGrid = false
		end if
	end with
End Sub

'���������Ʈ Ŭ�� �̺�Ʈ
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
	if Row > 0 and Col > 1 then		
	elseif Col = 1 and Row = 0  then
	
		mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
		if mstrCheck = True then 
			for intCnt = 1 To .sprSht.MaxRows
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
				'End If
			Next
			mstrCheck = False
		elseif mstrCheck = False then 
			mstrCheck = True
		end if
	end if 
	End With
End Sub 

'��Ʈ ����Ŭ�� �̺�Ʈ
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'���� ������ Ŭ���� ��ǥ ������ ��ȸ
Sub sprSht_FEELIST_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn frmThis.sprSht_FEELIST.ActiveCol,frmThis.sprSht_FEELIST.ActiveRow
	End If
End Sub


Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	with frmThis
	End with
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
End Sub

'-----------------------------
'�ݾ� �ڵ� ���
'-----------------------------
Sub sprSht_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt .sprSht
	End With
End Sub


SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"VAT") Then
		
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"VAT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

'----------------------------
'��Ʈ ���콺 ��
'----------------------------
Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht
	end with
End Sub
'-----------------------------------
'��Ʈ���� ���콺�� �ݾ��ջ� �̺�Ʈ
'-----------------------------------
sub MouseUp_SumAmt(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0 : intRowCnt = 0
			
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"AMT") or SprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(SprSht,"VAT") then
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)

				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	Dim strUSEYN
	Dim vntData
	Dim strCC
	Dim blnByteCHk
	Dim intRtn
	strUSEYN = ""
	strCC = ""
	
	with frmThis
		If mobjSCGLSpr.GetTextBinding(.sprSht,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,"TODATE",Row,Row,false
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMM")  Then
			blnByteCHk =  checkBytes(mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",Row))
			If blnByteCHk  > 23 Then
				intRtn = gYesNoMsgbox("������ ũ��� 23Byte �� ������ �����ϴ�. �ʱ�ȭ �Ͻðڽ��ϱ�?","ó���ȳ�!")
				IF intRtn <> vbYes then exit Sub
				mobjSCGLSpr.SetTextBinding .sprSht,"SUMM",Row,""
			End If
		END IF
		 
		 '���̳� ���� ��ȯ
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PAYCODE") then
			strCODE = mobjSCGLSpr.GetTextBinding( frmThis.sprSht,"CUSTOMERCODE",Row)
			Call Get_SUBCOMBO_VALUE(strCODE, Row, .sprSht)
		end if
		
	End With
End Sub


'�ݾ� �ջ� ����
Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDCMFEEVVOCH = gCreateRemoteObject("cMDCO.ccMDCOFEEVOCH")
	Set mobjMDCOGET		 = gCreateRemoteObject("cMDCO.ccMDCOGET")
	Set mobjMDCOVOCH = gCreateRemoteObject("cMDCO.ccMDCOVOCH")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub Grid_Layout()
	Dim intGBN
	Dim strComboList
	Dim strBMORDER
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***��ǥ ����Ʈ
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_FEELIST
		mobjSCGLSpr.SpreadLayout .sprSht_FEELIST, 4, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_FEELIST,   "YEARMON | TAXNO | CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.SetHeader .sprSht_FEELIST,		   "���|����|�ڵ�|�ŷ�ó"
		mobjSCGLSpr.SetColWidth .sprSht_FEELIST, "-1", "   0|   0|   0|    18"
		mobjSCGLSpr.SetRowHeight .sprSht_FEELIST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_FEELIST, "0", "15"
		mobjSCGLSpr.SetCellsLock2 .sprSht_FEELIST, true, "YEARMON | TAXNO | CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht_FEELIST, "CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht_FEELIST, "YEARMON | TAXNO | CLIENTCODE", true
		
		'**************************************************
		'***��ǥ �� ����Ʈ
		'**************************************************	
		strComboList =  "Y" & vbTab & " "
		
		strBMORDER = "AD0110" & vbTab & "AD0120" & vbTab & "AD0130" & vbTab & "AD0140" & vbTab & "AD0150" & vbTab & "AD0160" & vbTab & "AD0190" _
					& vbTab & "AD0210" & vbTab & "AD0220" & vbTab & "AD0290" & vbTab & "AD0310" & vbTab & "AD0320" & vbTab & "AD0390" & vbTab & "AD0410" _ 
					& vbTab & "AD0420" & vbTab & "AD0430" & vbTab & "AD0440" & vbTab & "AD0450" & vbTab & "AD0510" & vbTab & "AD0610" & vbTab & ""
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 35, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE |  BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK"
		mobjSCGLSpr.SetHeader .sprSht,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|�Աݱ���|���VENDOR|��ü���|����|��������|����|BMORDER|������|���޹��|BANKTYPE|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|MEDFLAG|AMTGBN|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|       8|        10|    15|  17|       5|         8|  10|    10|       6| 5|       8|       8|        10|      15|   0|       7|   7|      7|     8|      20|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      0|     0|        0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | SEMU | BP | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PAYCODE | BANKTYPE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht,"PREPAYMENT"),-1,-1,strComboList,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"BMORDER"),mobjSCGLSpr.CnvtDataField(.sprSht,"BMORDER"),-1,-1,strBMORDER,,80
		mobjSCGLSpr.SetCellAlign2 .sprSht, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE | VENDOR",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CUSTOMERCODE | CUSTNAME | REAL_MED_NAME | SUMM | AMT | BP | VENDOR | GBN |  TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG| TRANSRANK"
		mobjSCGLSpr.ColHidden .sprSht, "GBN  | GFLAG | MEDFLAG | ERRCODE | AMTGBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"TAXNO | VOCHNO | ERRCODE | ERRMSG"
		
		.sprSht_FEELIST.style.visibility  = "visible"
		.sprSht.style.visibility  = "visible"
		pnlFLAG.style.visibility = "visible" 
		
		.sprSht_FEELIST.MaxRows = 0
		.sprSht.MaxRows = 0
	End with
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		DateClean
		'Sheet�ʱ�ȭ
		.sprSht_FEELIST.MaxRows = 0
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
		Get_COMBO_VALUE		
		Set_delete ""
	End with
End Sub

Sub EndPage()
	set mobjMDCMFEEVVOCH = Nothing
	Set mobjMDCOGET = Nothing
	Set mobjMDCOVOCH = Nothing
	gEndPage	
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		vntData = mobjMDCMFEEVVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then		
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	

'-----------------------------------------------------------------------------------------
' �׸��� ���� �޺� ����
'-----------------------------------------------------------------------------------------
Sub Get_SUBCOMBO_VALUE(strCODE, row, sprsht)
	Dim vntData
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCODE = replace(strCODE,"-","")

       	vntData = mobjMDCMFEEVVOCH.Get_SUBCOMBO_VALUE(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)
		If not gDoErrorRtn ("Get_SUBCOMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 sprsht, "BANKTYPE",Row,Row,vntData,,160 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE

	with frmThis
		strDATE = MID(.txtYEARMON.value,1,4) & "-" & MID(.txtYEARMON.value,5,2)
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
		.txtDEMANDDAY.value = date2
	End With
End Sub

Sub SelectRtn_FEELIST (strVOCH_TYPE)
   	Dim vntData
   	Dim i, strCols
    Dim strYEARMON
    Dim strGBN
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt= clng(0) : mlngColCnt=clng(0)
		strYEARMON = .txtYEARMON.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF

		vntData = mobjMDCMFEEVVOCH.SelectRtn_FEELIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strVOCH_TYPE,strGBN)

		if not gDoErrorRtn ("SelectRtn_FEELIST") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_FEELIST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			Else
   				.sprSht_FEELIST.MaxRows = 0
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus_FEELIST, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			Call SelectRtn (1, 1)
   		end if
   	end with
End Sub

Sub SelectRtn (Col, Row)
   	Dim vntData
   	Dim i, strCols
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME
    Dim strGBN
    Dim strRANKCLIENT
    Dim lngRANK
	'On error resume next
	
	lngRANK = 1
	
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		IF .sprSht_FEELIST.MaxRows = 0 THEN EXIT SUB

		strVOCH_TYPE ="F"
		strYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_FEELIST,"YEARMON",.sprSht_FEELIST.ActiveRow)
		strTAXNO = mobjSCGLSpr.GetTextBinding( .sprSht_FEELIST,"TAXNO",.sprSht_FEELIST.ActiveRow)
		strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_FEELIST,"CLIENTCODE",.sprSht_FEELIST.ActiveRow)
		strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_FEELIST,"CLIENTNAME",.sprSht_FEELIST.ActiveRow)
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF
		vntData = mobjMDCMFEEVVOCH.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,strCLIENTCODE, strCLIENTNAME, strVOCH_TYPE,strGBN)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For i = 1 To .sprSht.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DEMANDDAY",i,i,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DUEDATE",i,i,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DEMANDDAY",i,i,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DUEDATE",i,i,false
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"PREPAYMENT",i) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"FROMDATE",i,i,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"TODATE",i,i,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,True,"FROMDATE",i,i,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,True,"TODATE",i,i,false
					End If
				Next
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				AMT_SUM .sprSht
   			Else
   				.sprSht.MaxRows = 0
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   			mstrSTAY = true
   		end if
   	end with
End Sub

Sub ProcessRtn(strVOCH_TYPE)
Dim intRtn
	with frmThis
		IF mstrPROCESS = "Create" THEN
			IF NOT .rdF.checked THEN
				gErrorMsgBox "�̿Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if

		IF mstrPROCESS = "Delete" THEN
			IF NOT .rdT.checked THEN
				gErrorMsgBox "�Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if 

		IF mstrSTAY THEN 
			mstrSTAY = FALSE
			IF strVOCH_TYPE = "F" THEN
				if DataValidation_OUT =false then exit sub
				CALL ProcessRtn_OUT()
			END IF
		ELSE
			gErrorMsgBox "��ǥó�� �������Դϴ�.","��ǥó�� �ȳ�"
		END IF
   	end with
END SUB

Function DataValidation_OUT ()
	DataValidation_OUT = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	With frmThis
		For intCnt =1  To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_OUT = True
End Function

'��ǥ���� ����
Sub ProcessRtn_OUT()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim intColFlag, bsdiv, intMaxCnt
	
	'--��ǥ ä���� ���� ���� ����
	Dim strGROUPSEQ : strGROUPSEQ = TRUE
	Dim vntData
	Dim strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strVOCHNORMS, strGROUP, strTYPE
	
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE | VENDOR | REAL_MED_NAME | GBN | ACCOUNT | DEBTOR | BMORDER | DOCUMENTDATE | PAYCODE |  BANKTYPE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | MEDFLAG | AMTGBN | TRANSRANK")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
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
		IF_GUBUN = "RMS_0012"
		
		'�ִ밪
		intColFlag = 0
		For intMaxCnt = 1 To .sprsht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",intMaxCnt) = 1 Then
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",intMaxCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			End IF
		Next
		
		Dim lngAMT, lngSUMAMT, lngVAT, lngSUMVAT
		Dim strBA, strCOSTCENTER
		Dim i, j, intCnt2
		
		IF .rdDIV.checked THEN
			if mstrPROCESS = "Create" then
				For intCnt = 1 To .sprsht.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprsht,"chk",intCnt) = "1" then		
					
						'ä���� �����Ѵ�.
						'--------------------------------------------------------------------------------------

						strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

						strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprSht,"POSTINGDATE",intCnt),"-","")
						strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAG",intCnt)
						strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt)
						strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt)'
						strTYPE				= "4"

						if strGROUPSEQ = true then
							strGROUP = TRUE
						else 
							strGROUP = FALSE
						END IF 

						If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
							gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
							Exit Sub
						END IF 

						strGROUPSEQ = FALSE
						
						'���� ������ RMS ä�� ��������
						vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
						
						strVOCHNORMS =  vntData(0,1)

						'---------------------------------------------------------------------------------------
						
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "O"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1
							strISEQ = 1
							
							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",intCnt),"-","") + "|" + _
										strVOCHNORMS + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",intCnt)
						end if
					end if 
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To .sprsht.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",intCnt) = "1" then		
						strIF_CNT = strIF_CNT + 1
				
						strRMS_DOC_TYPE = "Z"
						if strIF_CNT = "1" then

							strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",intCnt)
						else
							strHSEQ = strHSEQ + 1

							strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
										cstr(strISEQ) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BA",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",intCnt) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",intCnt)) + "|" + _
										cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",intCnt)) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",intCnt) + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"BP",intCnt) + "|" + _ 
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",intCnt) + "|" + _
										strRMS_DOC_TYPE + "|" + _ 
										mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",intCnt) + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",intCnt),"-","") + "|" + _
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",intCnt) + "|" + _  
										replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",intCnt),"-","") + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"VOCHNO",intCnt) + "|" + _
										mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",intCnt) + "|" + _  
										mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",intCnt)
						end if
					end if 
				Next
			end if 
		ELSE
			if mstrPROCESS = "Create" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "M" 
	                
					For i = 1 To .sprsht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprsht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprsht,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",i)) = intCnt Then
								'û���հ�,�ΰ����հ�,û������ ����� ������ ����
								If intCnt2 = intCnt Then
								Else
								
									'ä���� �����Ѵ�.(�ջ���ǥ�� ä�� ����)
									'--------------------------------------------------------------------------------------
									strPOSTINGDATE = "" :  strMEDFLAG = "" : strRMSTAXYEARMON = "" :  strRMSTAXNO = "" : strVOCHNORMS = "" : strTYPE = ""

									strPOSTINGDATE		= replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",intCnt),"-","")
									strMEDFLAG			= mobjSCGLSpr.GetTextBinding(.sprsht,"MEDFLAG",intCnt)
									strRMSTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",intCnt)
									strRMSTAXNO			= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",intCnt)'
									strTYPE				= "4"

									if strGROUPSEQ = true then
										strGROUP = TRUE
									else 
										strGROUP = FALSE
									END IF 

									If not InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strRMSTAXYEARMON, strRMSTAXNO, strGROUP, strTYPE) Then 
										gErrorMsgBox "��ǥ ��ȣ�� ����� �������� �ʾҽ��ϴ�. �����ڿ��� �����ϼ��� ","��ǥ ���� ���"
										Exit Sub
									END IF 

									strGROUPSEQ = FALSE
									
									'���� ������ RMS ä�� ��������
									vntData = mobjMDCOVOCH.SelectRtnVOCHNORMS(gstrConfigXml,mlngRowCnt,mlngColCnt,strPOSTINGDATE,strMEDFLAG,strRMSTAXYEARMON,strRMSTAXNO)
									
									strVOCHNORMS =  vntData(0,1)
									'---------------------------------------------------------------------------------------
									
									strIF_CNT = strIF_CNT + 1

									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprsht,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprsht,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",i)
									strRMS_DOC_TYPE = "M"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",i)
									strVOCHNO		= strVOCHNORMS
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",i)
									
									if strIF_CNT = "1" then
										strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									else
										strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
													cstr(strISEQ) + "|" + _
													replace(strPOSTINGDATE,"-","") + "|" + _
													strVENDOR + "|" + _
													strSUMM + "|" + _
													strBA + "|" + _
													strCOSTCENTER + "|" + _
													cstr(strAMT) + "|" + _
													cstr(strVAT) + "|" + _
													strSEMU + "|" + _ 
													strBP + "|" + _ 
													replace(strDEMANDDAY,"-","") + "|" + _
													strCUSTOMERCODE + "|" + _
													strTAXYEARMON + "|" + _
													strTAXNO + "|" + _
													strGFLAG + "|" + _
													strRMS_DOC_TYPE + "|" + _ 
													strACCOUNT + "|" + _
													strDEBTOR + "|" + _
													replace(strDOCUMENTDATE,"-","") + "|" + _
													strPREPAYMENT + "|" + _
													replace(strFROMDATE,"-","") + "|" + _
													replace(strTODATE,"-","") + "|" + _
													strSUMMTEXT + "|" + _
													strAMTGBN + "|" + _
													strPAYCODE + "|" + _  
													replace(strDUEDATE,"-","") + "|" + _
													strVOCHNO + "|" + _
													strBANKTYPE + "|" + _
													strBMORDER
									end if
												
									For j = 1 To .sprsht.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprsht,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",j),"-","") + "|" + _
															strVOCHNORMS + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",j) + "|" + _  
															mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",j)
											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
									intCnt2 = intCnt
								End If
								'���Ѿ�����Ʈ.
							End If
						End If
					Next
				Next
			elseif mstrPROCESS = "Delete" then
				For intCnt = 1 To intColFlag
					intCnt2 = 0
					lngAMT = 0
					lngSUMAMT = 0
					lngVAT = 0
					lngSUMVAT = 0
					strRMS_DOC_TYPE = "Z" 
	                
					For i = 1 To .sprsht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",i) = 1 Then
							'û���հ�
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",i)) = intCnt Then
								lngAMT = CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",i))
								lngSUMAMT = lngSUMAMT + lngAMT
								lngVAT = CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"VAT",i))
								lngSUMVAT = lngSUMVAT + lngVAT
							End If
						End If
					Next

					For i = 1 To .sprsht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprsht,"chk",i) = 1 Then
							If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",i)) = intCnt Then
								'û���հ�,�ΰ����հ�,û������ ����� ������ ����
								If intCnt2 = intCnt Then
								Else
									strIF_CNT = strIF_CNT + 1
									
									strPOSTINGDATE	= mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",i)
									strVENDOR		= mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",i)
									strSUMM			= mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",i)
									strBA			= mobjSCGLSpr.GetTextBinding(.sprsht,"BA",i)
									strCOSTCENTER	= mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",i)
									strAMT			= lngSUMAMT
									strVAT			= lngSUMVAT
									strSEMU			= mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",i)
									strBP			= mobjSCGLSpr.GetTextBinding(.sprsht,"BP",i)
									strDEMANDDAY	= mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",i)
									strCUSTOMERCODE = mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",i)
									strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",i)
									strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",i)
									strGFLAG		= mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",i)
									strRMS_DOC_TYPE = "Z"
									strACCOUNT		= ""
									strDEBTOR		= mobjSCGLSpr.GetTextBinding(.sprsht,"DEBTOR",i)
									strDOCUMENTDATE = mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",i)
									strPREPAYMENT	= mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",i)
									strFROMDATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",i)
									strTODATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",i)
									strSUMMTEXT		= mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",i)
									strAMTGBN		= mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",i)
									strPAYCODE		= mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",i)
									strDUEDATE		= mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",i)
									strVOCHNO		= mobjSCGLSpr.GetTextBinding(.sprsht,"VOCHNO",i)
									strBANKTYPE		= mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",i)
									strBMORDER		= mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",i)
									
									strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
												cstr(strISEQ) + "|" + _
												replace(strPOSTINGDATE,"-","") + "|" + _
												strVENDOR + "|" + _
												strSUMM + "|" + _
												strBA + "|" + _
												strCOSTCENTER + "|" + _
												cstr(strAMT) + "|" + _
												cstr(strVAT) + "|" + _
												strSEMU + "|" + _ 
												strBP + "|" + _ 
												replace(strDEMANDDAY,"-","") + "|" + _
												strCUSTOMERCODE + "|" + _
												strTAXYEARMON + "|" + _
												strTAXNO + "|" + _
												strGFLAG + "|" + _
												strRMS_DOC_TYPE + "|" + _ 
												strACCOUNT + "|" + _
												strDEBTOR + "|" + _
												replace(strDOCUMENTDATE,"-","") + "|" + _
												strPREPAYMENT + "|" + _
												replace(strFROMDATE,"-","") + "|" + _
												replace(strTODATE,"-","") + "|" + _
												strSUMMTEXT + "|" + _
												strAMTGBN + "|" + _
												strPAYCODE + "|" + _  
												replace(strDUEDATE,"-","") + "|" + _
												strVOCHNO + "|" + _
												strBANKTYPE + "|" + _
												strBMORDER
												
									For j = 1 To .sprsht.MaxRows
										If mobjSCGLSpr.GetTextBinding(.sprsht,"CHK",j) = 1 Then

											If CDbl(mobjSCGLSpr.GetTextBinding(.sprsht,"TRANSRANK",j)) = intCnt Then	
												strIF_CNT = strIF_CNT + 1
												
												strISEQ = strISEQ+1
												
												strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
															cstr(strISEQ) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"POSTINGDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"VENDOR",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SUMM",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"BA",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"COSTCENTER",j) + "|" + _
															cstr(mobjSCGLSpr.GetTextBinding(.sprsht,"AMT",j)) + "|" + _
															cstr(0) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SEMU",j) + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprsht,"BP",j) + "|" + _ 
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DEMANDDAY",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"CUSTOMERCODE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"TAXYEARMON",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"TAXNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"GFLAG",j) + "|" + _
															strRMS_DOC_TYPE + "|" + _ 
															mobjSCGLSpr.GetTextBinding(.sprsht,"ACCOUNT",j) + "|" + _
															"" + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DOCUMENTDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"PREPAYMENT",j) + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"FROMDATE",j),"-","") + "|" + _
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"TODATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"SUMMTEXT",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"AMTGBN",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"PAYCODE",j) + "|" + _  
															replace(mobjSCGLSpr.GetTextBinding(.sprsht,"DUEDATE",j),"-","") + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"VOCHNO",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"BANKTYPE",j) + "|" + _
															mobjSCGLSpr.GetTextBinding(.sprsht,"BMORDER",j)
															
											end if
										End If
									Next
									strHSEQ = strHSEQ + 1
									strISEQ = 1
								End If
								'���Ѿ�����Ʈ.
								intCnt2 = intCnt
							End If
						End If
					Next
				Next
			end if 
		END IF

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub


'****************************************************************************************
' ä�� ����ó��
'****************************************************************************************
Function InsertRtn_VOCHNO (strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
	InsertRtn_VOCHNO = false
   	Dim strVOCHNO
	With frmThis
		
		'ä���� ����& �����Ѵ� (������ �ߺ��� ���� SAP �ʿ��� ������ �� ��쿡�� ���� ��ȣ�� �����Ǵ� ���� ���´�.).
		intRtn = mobjMDCOVOCH.InsertRtn_VOCHNO(gstrConfigXml,strPOSTINGDATE, strMEDFLAG, strTAXYEARMON, strTAXNO, strGROUP, strTYPE)
		If not gDoErrorRtn ("InsertRtn_VOCHNO") Then
		
			If intRtn = 0 Then
				Exit Function
			End If		
   		End If
   	end With
   	InsertRtn_VOCHNO = true
End Function


'---------------------------------------------------
' ��ǥ���� �� ��ǥ��ȣ �޾ƿ��� �� ���� RMS������Ʈ
'---------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
		if mstrPROCESS ="Create" then
			IF mstrGUBUN = "F" THEN
				intRtn = mobjMDCMFEEVVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN, "FEEALL")
			END IF
			
			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "F" THEN
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				END IF
				
				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "�����Դϴ�..","����ȳ�"
				End If
				
				SelectRtn_FEELIST (mstrGUBUN)
				
   			end if
   		elseif mstrPROCESS ="Delete" then
   			IF mstrGUBUN = "F" THEN
				intRtn = mobjMDCMFEEVVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN, "FEEALL" )
			END IF
   			
   			if not gDoErrorRtn ("VOCHDELL") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "F" THEN
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				END IF

				gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				SelectRtn_FEELIST (mstrGUBUN)
   			end if
   		end if 
   		IF mstrGUBUN = "T" THEN
			.sprSht.focus()
		END IF
	End With
End Sub

'���� ��ǥ ���� ����
sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   	
		IF NOT .rdE.checked THEN
			gErrorMsgBox "������ȸ�� �����մϴ�.","�����׻���"
			exit sub
		end if 
		
		IF mstrGUBUN = "F" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | TAXYEARMON | TAXNO | ERRCODE | GBN | MEDFLAG")
		END IF

		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If

		intRtn = mobjMDCMFEEVVOCH.DeleteRtn(gstrConfigXml,vntData)

		if not gDoErrorRtn ("DeleteRtn") then
			'��� �÷��� Ŭ����
			IF mstrGUBUN = "F" THEN
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			END IF
			if intRtn > 0 Then
				gErrorMsgBox "���� ��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
			End If
			
			SelectRtn_FEELIST (mstrGUBUN)
   		end if
   	end with
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

		If mstrGUBUN = "F"  then  
			If .sprSht.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
		end if 
	
		intRtn = gYesNoMsgbox("���������� SAP���� ���ε� ��ǥ�� SAP���� ����Ͽ� RMS�ʿ��� ������ �� ������ RMS�� ��ǥ�� ������ �����Ҷ� ����մϴ�. " & vbCrlf & "  " & vbCrlf & " ��ǥ�� ������ �����Ͻðڽ��ϱ�?","�������� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		'���õ� �ڷḦ ������ ���� ����
		If mstrGUBUN = "F"  then  
			for i = .sprSht.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",i)
					
					intRtn = mobjMDCMFEEVVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN, "FEEALL" )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		END IF

		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
			SelectRtn_FEELIST(mstrGUBUN)
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
		<div id="contentWapperDiv"></div>
		<div id="popupDiv"></div>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28">
							<TR>
								<TD id="TD1" height="20" width="400" align="left" runat="server">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="76" background="../../../images/back_p.gIF">
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
											<td class="TITLE">FEE ��ǥ����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 101; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE style="WIDTH: 100%" id="tblBody" border="0" cellSpacing="0" cellPadding="0" height="93%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" class="KEYFRAME" vAlign="top" colSpan="2" align="center">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
												width="60">&nbsp;���</TD>
											<TD class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 88px; HEIGHT: 22px" id="txtYEARMON" class="INPUT"
													maxLength="6" size="9" name="txtYEARMON"></TD>
											<td class="SEARCHDATA" width="150">
												<TABLE border="0" cellSpacing="0" cellPadding="2" align="right">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgFEE" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgFEE"
																alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 30px" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2" width="50">
													<TR>
														<td><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'" border="0" name="imgSave"
																alt="�ڷḦ �����մϴ�." src="../../../images/ImgvochCre.gIF" height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'" border="0" name="imgVochDel"
																alt="��ǥ�� �����մϴ�." src="../../../images/imgVochDel.gIF" height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'" border="0" name="imgDelete"
																alt="������ǥ �� �����մϴ�." src="../../../images/ImgErrVochDel.gIF" height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																title="SAP���� ���������Ͽ� RMS���� ������ �� ������ RMS��ǥ�� ������ �����Ѵ�." onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																border="0" name="imgVochDelco" alt="��ǥ�� ������ �����մϴ�." src="../../../images/imgVochDelco.gIF"
																height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
																alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></td>
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
								<TD style="WIDTH: 100%" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE id="tblKey1" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<td style="WIDTH: 290px" class="DATA">�հ� : <INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 20px" id="txtSUMAMT" class="NOINPUTB_R"
													title="�հ�ݾ�" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT style="WIDTH: 120px; HEIGHT: 20px" id="txtSELECTAMT" class="NOINPUTB_R" title="���ñݾ�"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
											<TD style="WIDTH: 67px" class="LABEL">����������</TD>
											<TD style="WIDTH: 200px" class="DATA"><INPUT accessKey="date" style="WIDTH: 120px; HEIGHT: 22px" id="txtDEMANDDAY" class="INPUT"
													title="��������" maxLength="10" size="14" name="txtDEMANDDAY">&nbsp;<IMG style="CURSOR: hand" id="imgDEMANDDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgDEMANDDAY" align="absMiddle" src="../../../images/btnCalEndar.gIF"
													height="16">&nbsp;<IMG style="CURSOR: hand" id="btnApp" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" border="0" name="btnApp" alt="�������� �����մϴ�."
													align="absMiddle" src="../../../images/imgApp.gIF" width="54" height="20">
											</TD>
											<TD class="SEARCHDATA"><INPUT id="rdT" title="�Ϸ᳻����ȸ" onclick="vbscript:Call Set_delete('')" value="rdT" type="radio"
													name="rdGBN">&nbsp;�Ϸ�&nbsp; <INPUT id="rdF" title="�̿Ϸ� ������ȸ" onclick="vbscript:Call Set_delete('')" value="rdF" CHECKED
													type="radio" name="rdGBN">&nbsp;�̿Ϸ�&nbsp; <INPUT id="rdE" title="������ǥ ������ȸ" onclick="vbscript:Call Set_delete('')" value="rdE" type="radio"
													name="rdGBN">&nbsp;����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
													<DIV id="pnlFLAG" align="center" style="VISIBILITY: hidden; WIDTH: 250px; POSITION: absolute; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdDIV" title="����" type="radio"  value="rdDIV" name="rdDIVGUBUN">&nbsp;����&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdSUM" title="�ջ�" type="radio" value="rdSUM" CHECKED name="rdDIVGUBUN">&nbsp;�ջ�</DIV>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="TOPSPLIT"></TD>
							</TR>
							<!--���� �� �׸���-->
							<tr>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<TABLE border="0" cellSpacing="1" cellPadding="0" width="100%" align="left" height="98%">
										<TR>
											<td style="WIDTH: 200px; HEIGHT: 100%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 200px; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_FEELIST" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="200"
														height="100%" >
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="5291">
														<PARAM NAME="_ExtentY" VALUE="12567">
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
											<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab2"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="100%" height="100%"
														>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="26431">
														<PARAM NAME="_ExtentY" VALUE="12567">
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
											<TD style="WIDTH: 1040px" id="lblStatus_FEELIST" class="BOTTOMSPLIT"></TD>
											<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
										</TR>
										<tr>
											<td colSpan="2"><asp:textbox id="txtVOCHRETURN" runat="server" Width="8px" Height="0" Visible="False"></asp:textbox>
												<!--style="DISPLAY: none"--></td>
										</tr>
									</TABLE>
								</TD>
							</tr>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
		<iframe style="WIDTH: 100%; DISPLAY: none; HEIGHT: 300px" id="frmSapCon" src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx"
			name="frmSapCon"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
