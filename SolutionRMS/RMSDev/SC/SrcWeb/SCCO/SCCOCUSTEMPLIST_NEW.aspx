<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOCUSTEMPLIST_NEW.aspx.vb" Inherits="SC.SCCOCUSTEMPLIST_NEW" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ�ó ����ڰ���</title>
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
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOCUSTEMP '�����ڵ�, Ŭ���� ccMDCMCUSTEMP
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

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
sub ImgAddRowDTR_onclick ()
	With frmThis
		If .sprSht_CUST.MaxRows = 0 Then
			gErrorMsgBox "����� û���� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		
		If .sprSht_DTL.MaxRows = 0 Then
			gErrorMsgBox "��/��ü ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		
		If mobjSCGLSpr.GetTextBinding( frmThis.sprSht_DTL,"CUSTCODE",frmThis.sprSht_DTL.ActiveRow) = "" Then
			gErrorMsgBox "��/��ü �ŷ�ó������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		call sprSht_EMP_Keydown(meINS_ROW, 0)
		.txtREG_NUM1.focus()
		.sprSht_EMP.focus
	End With 
End sub

'-----------------------------------
' ����   
'-----------------------------------
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
		mobjSCGLSpr.ExportExcelFile .sprSht_CUST
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
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

'-----------------------------
' �ŷ��� �˾� ��ȸ 
'-----------------------------
Sub txtCUST_NAME_onkeydown
	if window.event.keyCode = meEnter then
		SelectRtn
	end if
End Sub

'--------------------------------------------------
' SpreadSheet �̺�Ʈ
'--------------------------------------------------
Sub sprSht_EMP_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_EMP, Col, Row
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

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_EMPBinding Col, Row
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

sub sprSht_EMP_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_EMP, ""
		End If
	End With
End sub
'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_DTL_Keyup(KeyCode, Shift)
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
		SelectRtn_EMPBinding frmThis.sprSht_DTL.ActiveCol,frmThis.sprSht_DTL.ActiveRow
	End If
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

Sub sprSht_EMP_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_EMP, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_EMP,"CUSTCODE",frmThis.sprSht_EMP.ActiveRow, mobjSCGLSpr.GetTextBinding( frmThis.sprSht_DTL,"CUSTCODE",frmThis.sprSht_DTL.ActiveRow) 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_EMP,"USE_YN",frmThis.sprSht_EMP.ActiveRow, "1"
		mobjSCGLSpr.ActiveCell frmThis.sprSht_EMP, 1,frmThis.sprSht_EMP.MaxRows
		
		frmThis.txtREG_NUM1.focus
		frmThis.sprSht_EMP.focus
	End If
End Sub

Sub txtREG_NUM1_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

	
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'----------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'----------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjSCCOCUSTEMP = gCreateRemoteObject("cSCCO.ccSCCOCUSTEMP")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		'���� �ŷ�ó �׸���(������, ����ó)
		gSetSheetColor mobjSCGLSpr, .sprSht_CUST	
		mobjSCGLSpr.SpreadLayout .sprSht_CUST, 13, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_CUST, "BUSINO | GUBUN |COMPANYNAME | HIGHCUSTCODE | CUSTOWNER | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_CUST,		  "����ڹ�ȣ|����|��ȣ��|�ڵ�|��ǥ��|����|����|�����ȣ|�ּ�1|�ּ�2|��ȭ��ȣ|�ѽ�|���"
		mobjSCGLSpr.SetColWidth .sprSht_CUST, "-1", "      13|   7|   16|   6|     7|   7|   8|       8|   15|   15|       8|   8|  15"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CUST, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CUST, "BUSINO | GUBUN | COMPANYNAME | HIGHCUSTCODE | CUSTOWNER | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_CUST, true, "BUSINO | GUBUN | COMPANYNAME | HIGHCUSTCODE | CUSTOWNER | BUSISTAT | BUSITYPE | ZIPCODE | ADDRESS1 | ADDRESS2 | TEL | FAX | MEMO"
		mobjSCGLSpr.SetCellAlign2 .sprSht_CUST, "BUSINO | HIGHCUSTCODE | ZIPCODE" ,-1,-1,2,2,false
		
		
		'���� �ŷ�ó �׸���(��, CIC/�����)
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 2, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CUSTCODE | CUSTNAME"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "�ڵ�|��/��ü��"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 6|      24"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "CUSTCODE | CUSTNAME", -1, -1, 200
		
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CUSTCODE | CUSTNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CUSTCODE",-1,-1,2,2,False
		
		
		'���� �ŷ�ó �׸���(���)
		gSetSheetColor mobjSCGLSpr, .sprSht_EMP
		mobjSCGLSpr.SpreadLayout .sprSht_EMP, 10, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_EMP, "CUSTCODE | SEQ | EMP_NAME | EMP_EMAIL | EMP_HP | EMP_TEL | DEPT_NAME | USE_YN | DEF_GBN | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_EMP,		 "�ڵ�|��ȣ|*����ڸ�|*�̸���|�ڵ���|��ȭ��ȣ|���μ�|���|�⺻|���"
		mobjSCGLSpr.SetColWidth .sprSht_EMP, "-1", " 6|   5|        9|     18|    11|       9|      20|   5|   4|  20"
		mobjSCGLSpr.SetRowHeight .sprSht_EMP, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_EMP, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_EMP, "DEF_GBN | USE_YN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_EMP, "SEQ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_EMP, "CUSTCODE | EMP_NAME | EMP_EMAIL | EMP_HP | EMP_TEL | DEPT_NAME | MEMO", -1, -1, 200
		
		mobjSCGLSpr.SetCellsLock2 .sprSht_EMP, true, "CUSTCODE | SEQ"
		mobjSCGLSpr.SetCellAlign2 .sprSht_EMP, "CUSTCODE",-1,-1,2,2,False
		mobjSCGLSpr.SetCellAlign2 .sprSht_EMP, "EMP_HP | EMP_TEL",-1,-1,0,2,False
		
		.sprSht_CUST.style.visibility = "visible"
		.sprSht_DTL.style.visibility = "visible"
		.sprSht_EMP.style.visibility = "visible"
	
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOCUSTEMP = Nothing
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
   	Dim strCUST_NAME, strREG_NUM, strMEDFLAG
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht_CUST.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		.sprSht_EMP.MaxRows = 0
		
		'���� �ʱ�ȭ
		strCUST_NAME = "" :  strREG_NUM = "" :  strMEDFLAG = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCUST_NAME = .txtCUST_NAME.value 
		strREG_NUM	 = replace(.txtREG_NUM1.value,"-","")
		strMEDFLAG	 = .cmbMEDFLAG1.value

		vntData = mobjSCCOCUSTEMP.SelectRtn_CUSTHDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strCUST_NAME, strREG_NUM, strMEDFLAG)

		If not gDoErrorRtn ("SelectRtn_CUSTHDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_CUST, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
						
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   			
   			Call SelectRtn_DTLBinding(1,1)
   		End If
   	End With
End Sub

'------------------------------------------
' DTL ������ ��ȸ
'------------------------------------------
Sub SelectRtn_DTLBinding(ByVal Col, ByVal Row)
	Dim strCUSTCODEHRD
	Dim vntData
	Dim i, strCols
	Dim strRows
	Dim intCnt
	
	With frmThis
		'sprSht2�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		
		If mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row) <> "" Then
			strCUSTCODEHRD = ""
		
			strCUSTCODEHRD = mobjSCGLSpr.GetTextBinding( .sprSht_CUST,"HIGHCUSTCODE",Row)
				
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			
			vntData = mobjSCCOCUSTEMP.SelectRtn_CUSTDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strCUSTCODEHRD )

			If not gDoErrorRtn ("SelectRtn_CUSTDTL") Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			End If	
	   		
   			Call SelectRtn_EMPBinding(1,1)
   		End If
	End With
End Sub

'------------------------------------------
' DTL ������ ��ȸ
'------------------------------------------
Sub SelectRtn_EMPBinding(ByVal Col, ByVal Row)
	Dim strCUSTCODE
	Dim vntData
	Dim i, strCols
	Dim strRows
	Dim intCnt
	
	With frmThis
		'sprSht2�ʱ�ȭ
		.sprSht_EMP.MaxRows = 0
		
		If mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CUSTCODE",Row) <> "" Then
			strCUSTCODE = ""
		
			strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"CUSTCODE",Row)
				
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOCUSTEMP.SelectRtn_EMPDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strCUSTCODE )

			If not gDoErrorRtn ("SelectRtn_EMPDTL") Then
				mobjSCGLSpr.SetClipbinding .sprSht_EMP, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			End If	
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
   	Dim strtest2
   	Dim strRow
	Dim lngCnt,intCnt
	Dim lngCol, lngRow
	Dim strDataCHK
	Dim strHIGHCUSTCODE
	Dim i
	With frmThis
		'������ Validation
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_EMP, "EMP_NAME | EMP_EMAIL",lngCol, lngRow, False) 
		 
		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ����ڸ�/�̸����� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		End If
		
		strtest2 = 0
		for i=1 to .sprSht_EMP.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_EMP,"DEF_GBN",i) = "1" then
				strtest2 = strtest2 +1
			end if
		next
		if strtest2 >1 then
			gErrorMsgBox "�⺻����ڴ� �ϳ��� ���� �����մϴ�.",""
			Exit Sub
		elseif strtest2 =0 then
			gErrorMsgBox "�⺻����ڸ� �����ϼ���.",""
			Exit Sub
		end if

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_EMP,"CUSTCODE | SEQ | EMP_NAME | EMP_EMAIL | EMP_HP | EMP_TEL | DEPT_NAME | USE_YN | DEF_GBN | MEMO")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		strHIGHCUSTCODE = ""
		strHIGHCUSTCODE = mobjSCGLSpr.GetTextBinding(.sprSht_CUST,"HIGHCUSTCODE",.sprSht_CUST.ActiveRow)
		
		intRtn = mobjSCCOCUSTEMP.ProcessRtn_CUSTDTL(gstrConfigXml,vntData, strHIGHCUSTCODE)
	
		If not gDoErrorRtn ("ProcessRtn_CUSTDTL") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_EMP,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht_DTL.ActiveRow
			mobjSCGLSpr.ActiveCell .sprSht_DTL, 1, strRow
			Call SelectRtn_EMPBinding(1,strRow)
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
											<td class="TITLE">�ŷ�ó ����� ����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
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
											<TD class="SEARCHLABEL" width="60">�ŷ�ó</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtCUST_NAME" title="�ڵ��" style="WIDTH: 299px; HEIGHT: 22px"
													maxLength="100" align="left" size="44" name="txtCUST_NAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
												width="80">����ڹ�ȣ</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREG_NUM1" title="�ڵ���ȸ" style="WIDTH: 152px; HEIGHT: 22px"
													align="left" name="txtREG_NUM1">&nbsp;<SELECT id="cmbMEDFLAG1" title="����" style="WIDTH: 108px" name="cmbMEDFLAG1">
													<OPTION value="" selected>����-��ü</OPTION>
													<OPTION value="A">������</OPTION>
													<OPTION value="B">��ü</OPTION>
												</SELECT></TD>
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
										border="0">
										<TR>
											<TD align="left" width="297" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign=bottom width="292">û����(������/��ü��)
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
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
								<TD class="BODYSPLIT" id="TD1" style="WIDTH: 100%; HEIGHT: 3px" runat="server"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR id="tblBody1">
								<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_CUST" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31802">
											<PARAM NAME="_ExtentY" VALUE="5953">
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
											<TD align="left" width="297" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="bottom" width="292">������Է�
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="22">
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
							<TR id="tblBody2">
								<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 70%" vAlign="top" align="center">
									<table height="100%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
										<tr>
											<td width="25%">
												<DIV id="pnlTab_2" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
													ms_positioning="GridLayout">
													<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
														DESIGNTIMEDRAGDROP="213" >
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="6323">
														<PARAM NAME="_ExtentY" VALUE="5953">
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
											<td width="75%">
												<DIV id="pnlTab_3" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
													ms_positioning="GridLayout">
													<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_EMP" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="25400">
														<PARAM NAME="_ExtentY" VALUE="5953">
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
										</tr>
									</table>
								</TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
