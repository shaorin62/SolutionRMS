<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOMEDLIST.aspx.vb" Inherits="SC.SCCOMEDLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ��� ���(��ü)</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : �ŷ�ó���� (��ü��) 
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : �ŷ�ó ���� MAIN ������ ��ȸ/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/07 By KTY
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script type="text/javascript">
		
function Set_IframeValue(strBUSINO,intCNT) {
	var value1  = strBUSINO;
	var value2  = intCNT;
	//iframe ������Ʈ�� �ؽ�Ʈ �ڽ� busino �Է�
	var textbox1 = frmSapCon.document.getElementById("<%=txtSAPBUSINO.ClientID%>");
	var textbox2 = frmSapCon.document.getElementById("<%=txtCNT.ClientID%>");
	
	textbox1.value = value1;
	textbox2.value = value2;
	window.frames[0].document.forms[0].submit();
}

		</script>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOCUSTLIST '�����ڵ�, Ŭ����
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'---------------------------------------------------
' �ű� SAP ���޾ƿ���
'---------------------------------------------------
Sub Set_CustValue (strVALUE, strBANKTYPE)
	Dim strCUSTINFO
	Dim strCUSTNAME
	Dim strCOMPANYNAME
	Dim strADDRESS1
	Dim strADDRESS2
	Dim strZIPCODE
	Dim strCUSTOWNER
	Dim strBUSISTAT
	Dim strBUSITYPE
	Dim strACCUSTCODE
	Dim strTEL
	Dim arraylist
	
	With frmThis
		If MID(strVALUE,InStr(1,strVALUE,"|"),len(strVALUE)) = "||||||||||||||" Then
			gErrorMsgBox "SAP �ʿ� ���������ʴ� �ŷ�ó�� ����ڹ�ȣ�Դϴ�.",""
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, ""
			Exit Sub
		Else
			strCUSTINFO = split(strVALUE,"|")

			strCUSTNAME = "" : strCOMPANYNAME = "" : strADDRESS1 = "" : strADDRESS2 = "" : strZIPCODE = "" 
			strCUSTOWNER = "" : strBUSISTAT = "" : strBUSITYPE = "" : strACCUSTCODE = "" : strTEL = ""

			strCUSTNAME		= strCUSTINFO(1)
			strCOMPANYNAME	= strCUSTINFO(2)
			strADDRESS1		= strCUSTINFO(3)
			strADDRESS2		= strCUSTINFO(4)
			strZIPCODE		= strCUSTINFO(5)
			strTEL			= strCUSTINFO(6)
			strCUSTOWNER	= strCUSTINFO(7)
			strBUSISTAT		= strCUSTINFO(8)
			strBUSITYPE		= strCUSTINFO(9)
			strACCUSTCODE	= strCUSTINFO(11)
			
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"COMPANYNAME",	.sprSht_CUST.ActiveRow, trim(strCOMPANYNAME)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"CUSTNAME",		.sprSht_CUST.ActiveRow, trim(strCUSTNAME)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"CUSTOWNER",	.sprSht_CUST.ActiveRow, trim(strCUSTOWNER)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSISTAT",		.sprSht_CUST.ActiveRow, trim(strBUSISTAT)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSITYPE",		.sprSht_CUST.ActiveRow, trim(strBUSITYPE)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"ZIPCODE",		.sprSht_CUST.ActiveRow, trim(strZIPCODE)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"ADDRESS1",		.sprSht_CUST.ActiveRow, trim(strADDRESS1)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"ADDRESS2",		.sprSht_CUST.ActiveRow, trim(strADDRESS2)
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"TEL",			.sprSht_CUST.ActiveRow, trim(strTEL)
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
		End If

	End With
End Sub
'====================================================
' �̺�Ʈ ���ν��� 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'---------------------------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'---------------------------------------------------
'-----------------------------------
'��ȸ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'�߰�
'-----------------------------------
sub ImgAddRow_onclick ()
	With frmThis
		call sprSht_CUST_Keydown(meINS_ROW, 0)
		.txtBUSINO.focus
		.sprSht_CUST.focus
	End With 
end sub

sub ImgAddRowDTR_onclick ()
	With frmThis
		If .sprSht_CUST.MaxRows = 0 Then
			gErrorMsgBox "����� û���� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			exit Sub
		End If
		
		If mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"HIGHCUSTCODE",frmThis.sprSht_CUST.ActiveRow) = "" Then
			gErrorMsgBox "����� û���� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			exit Sub
		End If
		call sprSht_DTL_Keydown(meINS_ROW, 0)
		.txtBUSINO.focus
		.sprSht_DTL.focus
	End With 
end sub

'-----------------------------------
' ����   
'-----------------------------------
Sub imgSave_onclick ()
	IF frmThis.sprSht_CUST.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn_CUSTHDR
	gFlowWait meWAIT_OFF
End Sub

Sub imgSaveDTL_onclick ()
	IF frmThis.sprSht_DTL.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn_MEDDTL
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_CUST
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'����
'-----------------------------------
Sub imgDelete_onclick ()
	Dim i
	If frmThis.sprSht_CUST.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_DTL_onclick ()
	Dim i
	If frmThis.sprSht_DTL.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn_DTL
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' �ݱ�
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'--------------------------------------------------
' SpreadSheet �̺�Ʈ
'--------------------------------------------------
Sub sprSht_CUST_Change(ByVal Col, ByVal Row)
	Dim i
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_CUST,"BUSINO") Then
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow) <> "" Then
				If Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))) = 10 Then
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),1,3) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),4,2) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),6,5)
				elseIf Len(Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))) = 13 Then
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),1,6) & "-" & MID(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),7,7)
				else
					mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow, Trim(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow))
				End If
				If Busino_Check =False Then Exit Sub
			End If
			
			Set_IframeValue TRIM(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",Row)) , 1
		End If
		mobjSCGLSpr.CellChanged .sprSht_CUST, Col, Row
	End With
End Sub

'��ܱ׸��� ����� ��ȣ �Է½� ����ڹ�ȣ �ߺ� üũ
Function Busino_Check ()
	Busino_Check = false
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCCOCUSTLIST.Busino_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												trim(Replace(mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),"-","")), "B")
		
		If mlngRowCnt > 0 Then
			gErrorMsgBox "�����ŷ�ó�� �ߺ��� ����ڹ�ȣ�� �ֽ��ϴ�.",""
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow,""
			'��Ŀ���� ��Ʈ�� �̵���Ų��.
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
			Exit Function
   		End if
   	End With
   	Busino_Check = True
End Function

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		IF  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMPANYNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row))
			
			IF strCode = "" AND strCodeName <> "" THEN			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName,"B")

				if not gDoErrorRtn ("GetHIGHCUSTCODE") then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntData(1,1)
						
						.txtBUSINO.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_DTL_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMPANYNAME"), Row
						.txtBUSINO.focus
						.sprSht_DTL.focus 
					End If
   				end if
   			END IF
		end if
		
		IF  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MPPNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MPP",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MPPNAME",Row))
			
			IF strCode = "" AND strCodeName <> "" THEN			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "P")

				if not gDoErrorRtn ("GetHIGHCUSTCODE") then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPP",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPPNAME",Row, vntData(1,1)
						
						.txtBUSINO.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_DTL_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MPPNAME"), Row
						.txtBUSINO.focus
						.sprSht_DTL.focus 
					End If
   				end if
   			END IF
		end if
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub

Sub mobjSCGLSpr_DTL_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMPANYNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)))
			
			vntRet = gShowModalWindow("SCCOREAL_MEDPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End IF
		end if
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"MPPNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MPPNAME",Row)))
			
			vntRet = gShowModalWindow("SCCOMPPPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPP",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPPNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End IF
		end if
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtBUSINO.focus
		.sprSht_DTL.Focus
	end with
End Sub

'-----------------------------------
'��Ʈ Ŭ��
'-----------------------------------
Sub sprSht_CUST_Click(ByVal Col, ByVal Row)
	with frmThis		
		If Row > 0 and Col > 1 Then
			SelectRtn_DTLBinding Col, Row, .txtMEDNAME.value
		END IF
	End with
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
sub sprSht_CUST_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CUST, ""
		end if
	end with
end sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		end if
	end with
end sub
'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_CUST_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meINS_ROW Then
		frmThis.sprSht_DTL.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_CUST, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht_CUST,false,frmThis.sprSht_CUST.ActiveRow,mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),true
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"CUSTTYPE",frmThis.sprSht_CUST.ActiveRow, "��迭"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"USE_FLAG",frmThis.sprSht_CUST.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_CUST, 1,frmThis.sprSht_CUST.MaxRows
	End if
End Sub

Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"HIGHCUSTCODE",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"HIGHCUSTCODE",frmThis.sprSht_CUST.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"COMPANYNAME",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"COMPANYNAME",frmThis.sprSht_CUST.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"USE_FLAG",frmThis.sprSht_DTL.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_DTL, 1,frmThis.sprSht_DTL.MaxRows
	End if
End Sub

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_DTL_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNHIGH") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)))
			vntRet = gShowModalWindow("SCCOREAL_MEDPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNMPP") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MPP",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"MPPNAME",Row)))
			vntRet = gShowModalWindow("SCCOMPPPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPP",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"MPPNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
		End If	
		.txtBUSINO.focus
		.sprSht_DTL.Focus
		mobjSCGLSpr.ActiveCell .sprSht_DTL, Col, Row
	End With
End Sub

Sub sprSht_CUST_Keyup(KeyCode, Shift)
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
		SelectRtn_DTLBinding frmThis.sprSht_CUST.ActiveCol,frmThis.sprSht_CUST.ActiveRow, frmThis.txtMEDNAME.value
		
	End If
End Sub

Sub cmbMEDDIV_onChange
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtREAL_MED_NAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtMEDNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtBUSINO_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' ������ ȭ�� ������ �� �ʱ�ȭ 
'----------------------------------------------------------------------
	'����������ü ����	
	set mobjSCCOCUSTLIST = gCreateRemoteObject("cSCCO.ccSCCOCUSTLIST")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
	'���� �ŷ�ó �׸���(��ü��)
	gSetSheetColor mobjSCGLSpr, .sprSht_CUST	
	mobjSCGLSpr.SpreadLayout .sprSht_CUST, 17, 0, 0, 0,0
	mobjSCGLSpr.SpreadDataField .sprSht_CUST, "CHK | BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO | UUSER"
	mobjSCGLSpr.SetHeader .sprSht_CUST,		  "����|����ڹ�ȣ|��ü���|�ŷ�ó��|�ڵ�|��ǥ��|���|�迭|����|����|�����ȣ|�ּ�1|�ּ�2|��ȭ��ȣ|�ѽ�|���|�Է���"
	mobjSCGLSpr.SetColWidth .sprSht_CUST, "-1", " 4|        13|      25|      20|   7|    10|   5|   7|  10|  10|       0|   15|   15|       0|   0|   0|     6"
	mobjSCGLSpr.SetRowHeight .sprSht_CUST, "-1", "13"
	mobjSCGLSpr.SetRowHeight .sprSht_CUST, "0", "15"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CUST, "CHK | USE_FLAG"
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | CUSTTYPE | BUSISTAT | BUSITYPE |ADDRESS1 |ADDRESS2", -1, -1, 200
	mobjSCGLSpr.SetCellsLock2 .sprSht_CUST, true, "BUSINO | HIGHCUSTCODE | UUSER"
	mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_CUST, "CUSTTYPE", -1, -1, "�迭" & vbTab & "��迭" , 10, 60, FALSE, FALSE
	mobjSCGLSpr.colhidden .sprSht_CUST, "ZIPCODE|TEL|FAX|MEMO",true
	mobjSCGLSpr.SetCellAlign2 .sprSht_CUST, "BUSINO | HIGHCUSTCODE | ZIPCODE | CUSTTYPE" ,-1,-1,2,2,false
	
	
	'���� �ŷ�ó �׸���(��ü)
	gSetSheetColor mobjSCGLSpr, .sprSht_DTL
	mobjSCGLSpr.SpreadLayout .sprSht_DTL, 20, 0, 0, 0,0
	mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 4, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 17, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | CUSTNAME | CUSTCODE | HIGHCUSTCODE | BTNHIGH | COMPANYNAME | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_GEN | MED_PAP | MED_MAG | MED_NET | MED_OUT | MED_ETC | MPP | BTNMPP | MPPNAME | USE_FLAG"
	mobjSCGLSpr.SetHeader .sprSht_DTL,		 "����|��ü/ä�θ�|�ŷ�ó�ڵ�|û�����ڵ�|û����|TV|RD|DMB|CATV|����|�Ź�|����|���ͳ�|����|CGV|MPP�ڵ�|MPP|���"
	mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 4|         20|         8|         8|2|  20| 4| 4|  4|   4|   4|   4|   4|     4|   4|  4|      8|2|20|  5"
	mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
	mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
	mobjSCGLSpr.SetCellTYpeButton2 .sprSht_DTL,"..", "BTNHIGH | BTNMPP"
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "CUSTNAME | CUSTCODE | HIGHCUSTCODE | COMPANYNAME | MPP | MPPNAME", -1, -1, 200
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_GEN | MED_PAP | MED_MAG | MED_NET | MED_OUT | MED_ETC | USE_FLAG "
	
	
	mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CUSTCODE"
	mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CUSTCODE | HIGHCUSTCODE | MPP",-1,-1,2,2,False
		
	
	.sprSht_CUST.style.visibility = "visible"
	.sprSht_DTL.style.visibility = "visible"
	
	
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOCUSTLIST = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis

	'�ʱ� ������ ����
	With frmThis
		.sprSht_CUST.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
	End With
End Sub

'------------------------------------------
' HDR ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strREAL_MED_NAME, strMEDNAME, strBUSINO
   	Dim intCnt
   	Dim strMEDDIV
   	
	With frmThis

		'Sheet�ʱ�ȭ
		.sprSht_CUST.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'���� �ʱ�ȭ
		strREAL_MED_NAME = "" : strMEDNAME = "" :  strBUSINO = "" :  strMEDDIV = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strMEDDIV			= .cmbMEDDIV.value
		strREAL_MED_NAME	= .txtREAL_MED_NAME.value 
		strMEDNAME			= .txtMEDNAME.value
		strBUSINO			= .txtBUSINO.value
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_MEDHDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strMEDDIV, strREAL_MED_NAME, strMEDNAME, strBUSINO)

		If not gDoErrorRtn ("SelectRtn_MEDHDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_CUST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			For intCnt = 1 To .sprSht_CUST.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"USE_FLAG",intCnt) = "0" Then
					mobjSCGLSpr.SetCellShadow .sprSht_CUST, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
				End If
			Next
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   			
   			Call SelectRtn_DTLBinding(1,1, strMEDNAME)
   		End if
   	End With
End Sub

'------------------------------------------
' DTL ������ ��ȸ
'------------------------------------------
Sub SelectRtn_DTLBinding(ByVal Col, ByVal Row, ByVal strMEDNAME)
	Dim strHIGHCUSTCODEHRD
	Dim vntData
	Dim i, strCols
	Dim strRows
	Dim intCnt, intCnt2
	
	With frmThis
		'sprSht_DTL�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		
		If mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row) <> "" Then
			strHIGHCUSTCODEHRD = ""
		
			strHIGHCUSTCODEHRD = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row)
				
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			intCnt2 = 1
			
			vntData = mobjSCCOCUSTLIST.SelectRtn_MEDDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODEHRD, strMEDNAME)

			If not gDoErrorRtn ("SelectRtn_MEDDTL") Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			End if	
		
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht_DTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"USE_FLAG",intCnt) = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
					End If
				Next
			End IF
	   		
   			gWriteText lblStatusDTR, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   		End If
	End With
End Sub

'------------------------------------------
' HDR ������ ����
'------------------------------------------
Sub ProcessRtn_CUSTHDR ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	With frmThis
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_CUST, "BUSINO|COMPANYNAME|CUSTNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ����ڹ�ȣ/��ȣ��/�ŷ�ó���� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_CUST,"CHK | BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_CUSTHDR(gstrConfigXml,vntData, "B")
	
		If not gDoErrorRtn ("ProcessRtn_CUSTHDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_CUST,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht_CUST.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht_CUST, 1, strRow
			Call SelectRtn_DTLBinding(1,strRow, .txtMEDNAME.value)
   		End If
   		
   	End With
End Sub

'------------------------------------------
' DTL ������ ����
'------------------------------------------
Sub ProcessRtn_MEDDTL ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strRow
   	Dim lngCol, lngRow
   	Dim strDataCHK
	Dim lngCnt,intCnt,intCnt2
	
	With frmThis
   		'������ Validation
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_DTL, "CUSTNAME|HIGHCUSTCODE|COMPANYNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ��ü��/û�����ڵ�/û�������� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | CUSTNAME | CUSTCODE | HIGHCUSTCODE | BTNHIGH | COMPANYNAME | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_GEN | MED_PAP | MED_MAG | MED_NET | MED_OUT | MED_ETC | MPP | BTNMPP | MPPNAME | USE_FLAG")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_MEDDTL(gstrConfigXml,vntData, "B")
	
		If not gDoErrorRtn ("ProcessRtn_MEDDTL") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht_CUST.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht_CUST, 1, strRow
			Call SelectRtn_DTLBinding(1,strRow, .txtMEDNAME.value)
   		End If
   	End With
End Sub

'------------------------------------------
'������ ���� ���
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strHIGHCUSTCODE
	Dim strHIGHCUSTCODE2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht_CUST.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",i)
				If strHIGHCUSTCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					mlngRowCnt=clng(0)
					mlngColCnt=clng(0)
					vntData = mobjSCCOCUSTLIST.SelectRtn_MEDCountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODE) 
					If mlngRowCnt > 0 Then
						gErrorMsgBox i & "���� �ڵ�� ������ü�� �����մϴ�. ������ü�� ���� �Ǵ� û������ �����ϼž� ������ �� �ֽ��ϴ�.","�����ȳ�!"
						Exit Sub
					End If
				
					mlngRowCnt=clng(0)
					mlngColCnt=clng(0)
					vntData = mobjSCCOCUSTLIST.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODE, "R") 
					If mlngRowCnt > 0 Then
						strMSG = ""
						For intCnt = 0 To mlngRowCnt-1
							If vntData(0,intCnt) = "B" Then
								strMSG = strMSG & " �μ�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A2" Then
								strMSG = strMSG & " ���̺�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A" Then
								strMSG = strMSG & " ������: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "O" Then
								strMSG = strMSG & " ���ͳ�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "D" Then
								strMSG = strMSG & " ����: " & vntData(1,intCnt) & "��" 
							End If
						Next
						gErrorMsgBox i & "���� �ڵ�� " & strMSG & " �� û�൥���ͷ� ����Ǿ��ֽ��ϴ�.","�����ȳ�!"
						Exit Sub
					End If
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		For i = .sprSht_CUST.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE2 = mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"HIGHCUSTCODE",i)
			
				If strHIGHCUSTCODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					intRtn = mobjSCCOCUSTLIST.DeleteRtn_REAL(gstrConfigXml, strHIGHCUSTCODE2, "R")
					
					IF not gDoErrorRtn ("DeleteRtn_REAL") Then
						mobjSCGLSpr.DeleteRow .sprSht_CUST,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn") Then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		SelectRtn
	End With
	err.clear
End Sub


'------------------------------------------
'������ ���� ������
'------------------------------------------
Sub DeleteRtn_DTL()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strCUSTCODE
	Dim strCUSTCODE2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht_DTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 Then
				strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CUSTCODE",i)
				If strCUSTCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_DTL,i
				Else
					vntData = mobjSCCOCUSTLIST.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strCUSTCODE, "B") 
					If mlngRowCnt > 0 Then
						strMSG = ""
						For intCnt = 0 To mlngRowCnt-1
							If vntData(0,intCnt) = "B" Then
								strMSG = strMSG & " �μ�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A2" Then
								strMSG = strMSG & " ���̺�: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "A" Then
								strMSG = strMSG & " ������: " & vntData(1,intCnt) & "��" 
							ElseIf vntData(0,intCnt) = "O" Then
								strMSG = strMSG & " ���ͳ�: " & vntData(1,intCnt) & "��" 
							End If
						Next
						gErrorMsgBox i & "���� �ڵ�� " & strMSG & " �� û�൥���ͷ� ����Ǿ��ֽ��ϴ�.","�����ȳ�!"
						Exit Sub
					End If
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		For i = .sprSht_DTL.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 Then
				strCUSTCODE2 = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CUSTCODE",i)
			
				If strCUSTCODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_DTL,i
				Else
					intRtn = mobjSCCOCUSTLIST.DeleteRtn_REAL(gstrConfigXml, strCUSTCODE2, "B")
					
					IF not gDoErrorRtn ("DeleteRtn_REAL") Then
						mobjSCGLSpr.DeleteRow .sprSht_DTL,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn") Then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		SelectRtn_DTLBinding .sprSht_CUST.ActiveCol, .sprSht_CUST.ActiveRow, .txtMEDNAME.value
	End With
	err.clear
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
												<td class="TITLE">��ü�� ����&nbsp;</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
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
									<TD align="left" width="100%" height="1">
									</TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" width="60">��ü����</TD>
												<TD class="SEARCHDATA" width="120"><SELECT id="cmbMEDDIV" style="WIDTH: 111px" name="cmbMEDDIV">
														<OPTION value="" selected>��ü</OPTION>
														<OPTION value="MED_TV">������TV</OPTION>
														<OPTION value="MED_RD">������RD</OPTION>
														<OPTION value="MED_DMB">������DMB</OPTION>
														<OPTION value="MED_CATV">CABLE-TV</OPTION>
														<OPTION value="MED_PAP">�μ�</OPTION>
														<OPTION value="MED_NET">���ͳ�</OPTION>
														<OPTION value="MED_OUT">����</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME,'')"
													width="60">��ü���</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 198px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="26" name="txtREAL_MED_NAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME, '')"
													width="60">��ü��</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtMEDNAME" title="��ü��" style="WIDTH: 198px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="7" name="txtMEDNAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
													width="70">����ڹ�ȣ</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtBUSINO" title="����ڹ�ȣ" style="WIDTH: 168px; HEIGHT: 22px"
														type="text" maxLength="15" align="left" name="txtBUSINO" size="22">
													<asp:textbox id="txtSAPBUSINO" runat="server" Width="8px" Visible="False"></asp:textbox>
													<asp:textbox id="txtCNT" runat="server" Visible="false" Width="8px"></asp:textbox></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_CUST" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27517">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgAddRowDTR" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRowDTR"></TD>
															<TD><IMG id="imgSaveDTL" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" border="0" name="imgSaveDTL"></TD>
															<TD><IMG id="imgDelete_DTL" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete_DTL"></TD>
															<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27517">
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
									<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
								</TR>
								<TR>
									<TD></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 600px; HEIGHT: 500px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
