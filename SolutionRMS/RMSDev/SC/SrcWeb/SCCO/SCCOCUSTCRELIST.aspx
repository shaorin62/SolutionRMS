<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOCUSTCRELIST.aspx.vb" Inherits="SC.SCCOCUSTCRELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ��� ���(ũ������)</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOCUSTCRELIST.aspx
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By KTY
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
			'.txtBUSINO.focus()
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
'�ű�
'-----------------------------------
Sub imgNew_onclick
	DataClean
	call sprSht_CUST_Keydown(meINS_ROW, 0)
End Sub

'-----------------------------
'���߰�
'-----------------------------
Sub imgAddRow_onclick ()
	With frmThis
		Call sprSht_CUST_Keydown(meINS_ROW, 0)
	End With 
End Sub

'-----------------------------------
' ����   
'-----------------------------------
Sub imgSave_onclick ()
	IF frmThis.sprSht_CUST.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn_EXEHDR
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht_CUST
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'����   - ���� �������� ����
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
		If Col = 2 Then
			Set_IframeValue TRIM(mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"BUSINO",Row)) , 1
		End If
		mobjSCGLSpr.CellChanged .sprSht_CUST, Col, Row
	End With
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
Sub sprSht_CUST_DblClick (ByVal Col, ByVal Row)
	with frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CUST, ""
		End If
	End With
End Sub

'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_CUST_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_CUST, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"CUSTTYPE",frmThis.sprSht_CUST.ActiveRow, "��迭"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"USE_FLAG",frmThis.sprSht_CUST.ActiveRow, "1"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_CUST,"BUSINO",frmThis.sprSht_CUST.ActiveRow, "104-86-36968"
		
		'�ڵ����� MC�� ������ SAP���� �����´�.
		Call sprSht_CUST_Change(2, frmThis.sprSht_CUST.ActiveRow)
		mobjSCGLSpr.ActiveCell frmThis.sprSht_CUST, 4,frmThis.sprSht_CUST.ActiveRow
		
		frmThis.txtCLIENTNAME.focus
		frmThis.sprSht_CUST.focus
	End If
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
	set mobjSCCOGET		 = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
	
		gSetSheetColor mobjSCGLSpr, .sprSht_CUST	
		mobjSCGLSpr.SpreadLayout .sprSht_CUST, 17, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_CUST, "CHK | BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO | UUSER"
		mobjSCGLSpr.SetHeader .sprSht_CUST,		  "����|����ڹ�ȣ|��ȣ��|ũ��������|�ڵ�|��ǥ��|���|�迭|����|����|�����ȣ|�ּ�1|�ּ�2|��ȭ��ȣ|�ѽ�|���|�Է���"
		mobjSCGLSpr.SetColWidth .sprSht_CUST, "-1", " 4|        13|    25|        20|   7|    10|   5|   7|  10|  10|       0|   15|   15|       0|   0|   0|     6"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CUST, "CHK | USE_FLAG"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | CUSTTYPE | BUSISTAT | BUSITYPE |ADDRESS1 |ADDRESS2", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_CUST, true, "BUSINO | COMPANYNAME | HIGHCUSTCODE | CUSTOWNER | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO | UUSER"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_CUST, "CUSTTYPE", -1, -1, "�迭" & vbTab & "��迭" , 10, 60, FALSE, FALSE
		mobjSCGLSpr.colhidden .sprSht_CUST, "ZIPCODE|TEL|FAX|MEMO",true
		mobjSCGLSpr.SetCellAlign2 .sprSht_CUST, "CUSTTYPE" ,-1,-1,2,2,false
		
		.sprSht_CUST.style.visibility = "visible"

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
	End With
End Sub

'------------------------------------------
' HDR ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTNAME
   	Dim intCnt
   	
	With frmThis

		'Sheet�ʱ�ȭ
		.sprSht_CUST.MaxRows = 0
		
		'���� �ʱ�ȭ
		strCLIENTNAME = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value 
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_CREHDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strCLIENTNAME)

		If not gDoErrorRtn ("SelectRtn_CREHDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_CUST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			For intCnt = 1 To .sprSht_CUST.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"USE_FLAG",intCnt) = "0" Then
					mobjSCGLSpr.SetCellShadow .sprSht_CUST, -1, -1, intCnt, intCnt,&HB6B6B9, &H000000,False
				End If
			Next
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   		End If
   	End With
End Sub

'------------------------------------------
' HDR ����/���� ó�� 
'------------------------------------------
Sub ProcessRtn_EXEHDR ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	With frmThis
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_CUST, "CUSTNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� �ŷ�ó���� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_CUST,"CHK | BUSINO | COMPANYNAME | CUSTNAME | HIGHCUSTCODE | CUSTOWNER | USE_FLAG | CUSTTYPE | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		intRtn = mobjSCCOCUSTLIST.ProcessRtn_EXEHDR(gstrConfigXml,vntData, "K")
	
		If not gDoErrorRtn ("ProcessRtn_EXEHDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_CUST,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			SelectRtn
   		End If
   	End With
End Sub

'------------------------------------------
'������ ����
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
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",i)
				If strHIGHCUSTCODE = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					vntData = mobjSCCOCUSTLIST.SelectRtn_CountCheck(gstrConfigXml,mlngRowCnt,mlngColCnt, strHIGHCUSTCODE, "K") 
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
			End if
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes Then Exit Sub
		
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		For i = .sprSht_CUST.MaxRows To 1 Step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"CHK",i) = 1 Then
				strHIGHCUSTCODE2 = mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"HIGHCUSTCODE",i)
			
				If strHIGHCUSTCODE2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht_CUST,i
				Else
					intRtn = mobjSCCOCUSTLIST.DeleteRtn_EXE(gstrConfigXml, strHIGHCUSTCODE2)
					
					IF not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht_CUST,i
   					End IF
				End if				
   				intCnt = intCnt + 1
   			End IF
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn") Then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		SelectRtn
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
													<TABLE cellSpacing="0" cellPadding="0" width="82" background="../../../images/back_p.gIF"
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
												<td class="TITLE">ũ������ ����&nbsp;</td>
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
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,'')"
													width="100">&nbsp;ũ��������</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 200px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="28" name="txtCLIENTNAME">
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
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_CUST" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="16378">
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
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 10px; HEIGHT: 10px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
