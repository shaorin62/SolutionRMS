<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOCUSTLIST.aspx.vb" Inherits="SC.SCCOCUSTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ��� ���(������)</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOCUSTLIST.aspx
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/08 By KTY
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" VIEWASTEXT>
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
	'EndPage
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
End sub

sub ImgAddRowDTR_onclick ()
	With frmThis
		If .sprSht_CUST.MaxRows = 0 Then
			gErrorMsgBox "����� û���� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		
		If mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"HIGHCUSTCODE",frmThis.sprSht_CUST.ActiveRow) = "" Then
			gErrorMsgBox "����� û���� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		call sprSht_DTL_Keydown(meINS_ROW, 0)
		.txtBUSINO.focus
		.sprSht_CUST.focus
	End With 
End sub

'-----------------------------------
' ����   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht_CUST.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn_CUSTHDR
	gFlowWait meWAIT_OFF
End Sub

Sub imgSaveDTL_onclick ()
	If frmThis.sprSht_DTL.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn_CUSTDTL
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_CUST
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	End With
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
												trim(Replace(mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow),"-","")), "A")
		
		If mlngRowCnt > 0 Then
			gErrorMsgBox "�����ŷ�ó�� �ߺ��� ����ڹ�ȣ�� �ֽ��ϴ�.",""
			mobjSCGLSpr.SetTextBinding .sprSht_CUST,"BUSINO",.sprSht_CUST.ActiveRow,""
			'��Ŀ���� ��Ʈ�� �̵���Ų��.
			.txtBUSINO.focus()
			.sprSht_CUST.focus()
			Exit Function
   		End If
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
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"GBNFLAG") Then
			If mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"GBNFLAG",Row) = "��" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,false,Row,2,4,true
			ElseIf mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"GBNFLAG",Row) = "CIC/�����" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,Row,2,4,true
			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTSUBNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTSUBCODE",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTSUBNAME",Row))
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													   TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row)), _
													   TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)), _
													   strCode, strCodeName)

				If not gDoErrorRtn ("GetCLIENTSUBCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntData(4,1)	
						
						.txtBUSINO.focus
						.sprSht_DTL.focus
					Else
						mobjSCGLSpr_DTL_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTSUBNAME"), Row
						.txtBUSINO.focus
						.sprSht_DTL.focus 
					End If
   				End If
   			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMPANYNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row))
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
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
   				End If
   			End If
		End If
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub

Sub mobjSCGLSpr_DTL_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"CLIENTSUBNAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)),  _
								"", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTSUBNAME",Row)))
			
			vntRet = gShowModalWindow("SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMPANYNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)))
			
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_DTL, Col+2,Row
			End If
		End If
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtBUSINO.focus
		.sprSht_DTL.Focus
	End With
End Sub

'-----------------------------------
'��Ʈ Ŭ��
'-----------------------------------
Sub sprSht_CUST_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_DTLBinding Col, Row
		End If
	End With
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
sub sprSht_CUST_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CUST, ""
		End If
	End With
End sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End sub
'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_CUST_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		frmThis.sprSht_DTL.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_CUST, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht_CUST,False,frmThis.sprSht_CUST.ActiveRow,mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),mobjSCGLSpr.CnvtDataField(frmThis.sprSht_CUST,"BUSINO"),True
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"CUSTTYPE",frmThis.sprSht_CUST.ActiveRow, "��迭"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"USE_FLAG",frmThis.sprSht_CUST.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_CUST, 1,frmThis.sprSht_CUST.MaxRows
		frmThis.txtBUSINO.focus
		frmThis.sprSht_CUST.focus
	End If
End Sub

Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"GBNFLAG",frmThis.sprSht_DTL.ActiveRow, "��"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"HIGHCUSTCODE",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"HIGHCUSTCODE",frmThis.sprSht_CUST.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"COMPANYNAME",frmThis.sprSht_DTL.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht_CUST,"COMPANYNAME",frmThis.sprSht_CUST.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"USE_FLAG",frmThis.sprSht_DTL.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_DTL, 1,frmThis.sprSht_DTL.MaxRows
		
		frmThis.txtBUSINO.focus
		frmThis.sprSht_DTL.focus
	End If
End Sub

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_DTL_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)), _
								TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTSUBCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CLIENTSUBNAME",Row)))
								
			vntRet = gShowModalWindow("SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CLIENTSUBNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht_DTL, Col,Row
			End If
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BTNHIGH") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"HIGHCUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"COMPANYNAME",Row)))
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"HIGHCUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"COMPANYNAME",Row, vntRet(3,0)
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
		SelectRtn_DTLBinding frmThis.sprSht_CUST.ActiveCol,frmThis.sprSht_CUST.ActiveRow
		
	End If
End Sub

Sub txtCLIENTNAME_onKeyDown
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

Sub txtCOMPANYNAME_onKeyDown
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
		'���� �ŷ�ó �׸���(������, ����ó)
		gSetSheetColor mobjSCGLSpr, .sprSht_CUST	
		mobjSCGLSpr.SpreadLayout .sprSht_CUST, 21, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_CUST, "BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO | KOBACOCUSTCODE | GREATCODE | GREATNAME | ATTR10 | ATTR08 | UUSER"
		mobjSCGLSpr.SetHeader .sprSht_CUST,		  "����ڹ�ȣ|��ȣ��|�ŷ�ó��|�ڵ�|��ǥ��|���|�迭|����|����|�����ȣ|�ּ�1|�ּ�2|��ȭ��ȣ|�ѽ�|���|�ڹ��ڱ������ڵ�|����ó�ڵ�|����ó��|����|���ε�|�Է���"
		mobjSCGLSpr.SetColWidth .sprSht_CUST, "-1", "      13|    20|      18|   7|    10|   5|   7|  10|  10|       0|   15|   15|       0|   0|   0|               0|         0|       0|   4|     5|     6"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CUST, "USE_FLAG | ATTR10 | ATTR08"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | CUSTTYPE | BUSISTAT | BUSITYPE |ADDRESS1 |ADDRESS2", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_CUST, true, "BUSINO | HIGHCUSTCODE | UUSER"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_CUST, "CUSTTYPE", -1, -1, "�迭" & vbTab & "��迭" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.colhidden .sprSht_CUST, "ZIPCODE | TEL | FAX | MEMO | KOBACOCUSTCODE | GREATCODE | GREATNAME",true
		mobjSCGLSpr.SetCellAlign2 .sprSht_CUST, "BUSINO | HIGHCUSTCODE | ZIPCODE | CUSTTYPE | KOBACOCUSTCODE | GREATCODE" ,-1,-1,2,2,false
		
		
		'���� �ŷ�ó �׸���(��, CIC/�����)
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 10, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_DTL, 8, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "GBNFLAG | CLIENTSUBCODE | BTN | CLIENTSUBNAME | CUSTNAME | CUSTCODE | HIGHCUSTCODE | BTNHIGH | COMPANYNAME | USE_FLAG"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "����|�ڵ�|CIC/�����|�ŷ�ó��|�ŷ�ó�ڵ�|û�����ڵ�|û����|���"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "10|   8|2|      20|      20|        10|         9|2|  28|   5"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_DTL,"..", "BTN | BTNHIGH"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "CLIENTSUBCODE | CLIENTSUBNAME | CUSTNAME | CUSTCODE | HIGHCUSTCODE | COMPANYNAME", -1, -1, 200
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "USE_FLAG"
		
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_DTL, "GBNFLAG", -1, -1, "��" & vbTab & "CIC/�����" , 10, 80, FALSE, FALSE
		
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CUSTCODE"
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CLIENTSUBCODE | CUSTCODE | HIGHCUSTCODE",-1,-1,2,2,False
			
		
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
   	Dim strCLIENTNAME, strCOMPANYNAME, strBUSINO
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht_CUST.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'���� �ʱ�ȭ
		strCLIENTNAME = "" : strCOMPANYNAME = "" :  strBUSINO = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value 
		strCOMPANYNAME	= .txtCOMPANYNAME.value
		strBUSINO		= .txtBUSINO.value
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_CUSTHDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strCLIENTNAME, strCOMPANYNAME, strBUSINO)

		If not gDoErrorRtn ("SelectRtn_CUSTHDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_CUST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			For intCnt = 1 To .sprSht_CUST.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"USE_FLAG",intCnt) = "0" Then
					mobjSCGLSpr.SetCellShadow .sprSht_CUST, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
				End If
			Next
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   			
   			Call SelectRtn_DTLBinding(1,1)
   		End If
   	End With
End Sub

'------------------------------------------
' DTL ������ ��ȸ
'------------------------------------------
Sub SelectRtn_DTLBinding(ByVal Col, ByVal Row)
	Dim strHIGHCUSTCODEHRD
	Dim vntData
	Dim i, strCols
	Dim strRows
	Dim intCnt, intCnt2
	
	With frmThis
		'sprSht2�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		
		If mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row) <> "" Then
			strHIGHCUSTCODEHRD = ""
		
			strHIGHCUSTCODEHRD = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row)
				
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			intCnt2 = 1
			
			vntData = mobjSCCOCUSTLIST.SelectRtn_CUSTDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODEHRD )

			If not gDoErrorRtn ("SelectRtn_CUSTDTL") Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			End If	
		
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht_DTL.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"GBNFLAG",intCnt) = "CIC/�����" Then
						If intCnt2 = 1 Then
							strRows = intCnt
						Else
							strRows = strRows & "|" & intCnt
						End If
						intCnt2 = intCnt2 + 1
					End If
					
					If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"USE_FLAG",intCnt) = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
					End If
				Next
				
				mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,strRows,2,4,true
			End If
	   		
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
   	Dim strDIVAMT
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_CUST,"BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO | KOBACOCUSTCODE | GREATCODE | GREATNAME | ATTR10 | ATTR08")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_CUSTHDR(gstrConfigXml,vntData, "A")
	
		If not gDoErrorRtn ("ProcessRtn_CUSTHDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_CUST,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht_CUST.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht_CUST, 1, strRow
			Call SelectRtn_DTLBinding(1,strRow)
   		End If
   		
   	End With
End Sub

'------------------------------------------
' DTL ������ ����
'------------------------------------------
Sub ProcessRtn_CUSTDTL ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	
	With frmThis
		'������ Validation
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_DTL, "CUSTNAME|HIGHCUSTCODE|COMPANYNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� �ŷ�ó��/û�����ڵ�/û�������� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"GBNFLAG | CLIENTSUBCODE | BTN | CLIENTSUBNAME | CUSTNAME | CUSTCODE | HIGHCUSTCODE | BTNHIGH | COMPANYNAME | USE_FLAG")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_CUSTDTL(gstrConfigXml,vntData, "A")
	
		If not gDoErrorRtn ("ProcessRtn_CUSTDTL") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht_CUST.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht_CUST, 1, strRow
			Call SelectRtn_DTLBinding(1,strRow)
   		End If
   		
   	End With
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
												<td class="TITLE">������ ����&nbsp;</td>
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
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCOMPANYNAME,'')"
													width="100">��ȣ��</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCOMPANYNAME" title="�����ָ�" style="WIDTH: 200px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="28" name="txtCOMPANYNAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, '')"
													width="100">�����ָ�</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 200px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="28" name="txtCLIENTNAME"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
													width="100">����ڹ�ȣ</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtBUSINO" title="�ڵ���ȸ" style="WIDTH: 168px; HEIGHT: 22px" type="text"
														maxLength="15" align="left" name="txtBUSINO">
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
												DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="6826">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
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
												DESIGNTIMEDRAGDROP="213" >
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="6853">
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
								<!--Bottom Split End-->
							</TABLE>
							<!--Input Define Table End-->
						</TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 600px; HEIGHT: 600px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
