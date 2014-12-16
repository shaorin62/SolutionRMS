<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOBRANDHDRLIST.aspx.vb" Inherits="SC.SCCOBRANDHDRLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��ǥ�귣�����</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOCUSTEXELIST.aspx
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
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOBRANDLIST '�����ڵ�, Ŭ����
Dim mobjSCCOGET
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

'-----------------------------
'���߰�
'-----------------------------
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtCLIENTNAME.focus()
		.sprSht.focus()
	End With 
end sub

'-----------------------------------
' ����   
'-----------------------------------
Sub imgSave_onclick ()
	If frmThis.sprSht.MaxRows = 0 Then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	End if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'����
'-----------------------------------
Sub imgDelete_onclick ()
	Dim i
	If frmThis.sprSht.MaxRows = 0 Then
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

Sub txtCLIENTNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtHIGHSEQNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'--------------------------------------------------
' SpreadSheet �̺�Ʈ
'--------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"HIGHSEQNAME") Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSEQNAME",Row) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",Row) <> "" Then
				If HIGHSEQNAME_Check =False Then Exit Sub
			End If
		End If
		
		If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then
			strCode		= TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row))
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row))
			
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntData(1,1)
						
						.txtCLIENTNAME.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME"), Row
						.txtCLIENTNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
   			
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSEQNAME",Row) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",Row) <> "" Then
				If HIGHSEQNAME_Check =False Then Exit Sub
			End If
		End If
		
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CUSTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSEQNAME",Row) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",Row) <> "" Then
				If HIGHSEQNAME_Check =False Then Exit Sub
			End If
		End If
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtCLIENTNAME.focus
		.sprSht.Focus
	End With
End Sub

'��ܱ׸��� ����� ��ȣ �Է½� ����ڹ�ȣ �ߺ� üũ
Function HIGHSEQNAME_Check ()
	HIGHSEQNAME_Check = false
	Dim vntData
   	Dim i, strCols
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCCOBRANDLIST.HIGHSEQNAME_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												trim(mobjSCGLSpr.GetTextBinding( .sprSht,"HIGHSEQNAME",.sprSht.ActiveRow)), _
												trim(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",.sprSht.ActiveRow)))
		
		If mlngRowCnt > 0 Then
			gErrorMsgBox "��ǥ�귣��� �ߺ��� ��ǥ�귣����� �ֽ��ϴ�.",""
			mobjSCGLSpr.SetTextBinding .sprSht,"HIGHSEQNAME",.sprSht.ActiveRow,""
			Exit Function
   		End if
   	End With
   	HIGHSEQNAME_Check = True
End Function

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End if
	End With
End Sub

'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,false,frmThis.sprSht.ActiveRow,2,2,true
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME.focus
		frmThis.sprSht.focus
	End If
End Sub

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row)))
			vntRet = gShowModalWindow("SCCOCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			
			If mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSEQNAME",Row) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"CUSTCODE",Row) <> "" Then
				If HIGHSEQNAME_Check =False Then Exit Sub
			End If
		End If	
		.txtCLIENTNAME.focus
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	End With
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjSCCOBRANDLIST = gCreateRemoteObject("cSCCO.ccSCCOBRANDLIST")
	set mobjSCCOGET		  = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
	
		gSetSheetColor mobjSCGLSpr, .sprSht	
		mobjSCGLSpr.SpreadLayout .sprSht, 7, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 4, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | HIGHSEQNO | HIGHSEQNAME | CUSTCODE | BTN | CUSTNAME | SEQNAMES"
		mobjSCGLSpr.SetHeader .sprSht,		  "����|��ǥ�귣���ڵ�|��ǥ�귣��|�������ڵ�|�����ָ�|�귣��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  4|             0|        23|         8|2|    20|   60"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "HIGHSEQNAME | CUSTCODE | CUSTNAME | SEQNAMES", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SEQNAMES"
		mobjSCGLSpr.colhidden .sprSht, "HIGHSEQNO",true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTCODE" ,-1,-1,2,2,false
		
		.sprSht.style.visibility = "visible"

    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjSCCOBRANDLIST = Nothing
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
		.sprSht.MaxRows = 0
	End With
End Sub

'------------------------------------------
' HDR ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTNAME, strHIGHSEQNAME
   	Dim intCnt
   	
	With frmThis

		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'���� �ʱ�ȭ
		strCLIENTNAME = ""
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strCLIENTNAME	= .txtCLIENTNAME.value 
		strHIGHSEQNAME  = .txtHIGHSEQNAME.value
		vntData = mobjSCCOBRANDLIST.SelectRtn_HIGHSUBSEQ(gstrConfigXml,mlngRowCnt,mlngColCnt, strCLIENTNAME, strHIGHSEQNAME)

		If not gDoErrorRtn ("SelectRtn_HIGHSUBSEQ") Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   		End if
   	End With
End Sub

'------------------------------------------
' HDR ����/���� ó�� 
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	Dim vntData
	Dim strMasterData
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	Dim lngCol, lngRow
	Dim strDataCHK
	Dim strYEAR
	
	With frmThis
		 strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "HIGHSEQNAME | CUSTNAME",lngCol, lngRow, False) 
		 
		 If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ��ǥ�귣���/�����ִ� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		 End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | HIGHSEQNO | HIGHSEQNAME | CUSTCODE | BTN | CUSTNAME | SEQNAMES")
		
		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If
		
		strYEAR = Mid(gNowDate,3,2)
		
		intRtn = mobjSCCOBRANDLIST.ProcessRtn_HIGHSUBSEQ(gstrConfigXml,vntData, strYEAR)
	
		If not gDoErrorRtn ("ProcessRtn_HIGHSUBSEQ") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht.ActiveRow
			SelectRtn
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		End If
   	End With
End Sub

'------------------------------------------
'������ ����
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i , lngchkCnt
	Dim strHIGHSEQNO
	Dim strHIGHSEQNO2
	Dim intCnt
	Dim strMSG
	
	With frmThis
		For i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strHIGHSEQNO = mobjSCGLSpr.GetTextBinding( .sprSht,"HIGHSEQNO",i)
				
				If strHIGHSEQNO = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNAMES",i) <> "" Then

						gErrorMsgBox i & "���� ��ǥ�귣��� �귣�带 �����ϰ� �ֽ��ϴ�.","�����ȳ�!"
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
		For i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strHIGHSEQNO2 = mobjSCGLSpr.GetTextBinding(.sprSht,"HIGHSEQNO",i)
			
				If strHIGHSEQNO2 = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				Else
					intRtn = mobjSCCOBRANDLIST.DeleteRtn_HDR(gstrConfigXml, strHIGHSEQNO2)
					
					IF not gDoErrorRtn ("DeleteRtn_HIGH") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
   		
   		If not gDoErrorRtn ("DeleteRtn_HIGH") Then
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
													<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
												<td class="TITLE">��ǥ�귣�� ����&nbsp;</td>
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
												<TD class="SEARCHLABEL" width="100" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,'')">�����ָ�</TD>
												<TD class="SEARCHDATA" width = "400"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 256px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="37" name="txtCLIENTNAME"></TD>
												<TD class="SEARCHLABEL" width="100" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtHIGHSEQNAME,'')">��ǥ�귣���</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtHIGHSEQNAME" title="��ǥ�귣���" style="WIDTH: 256px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="37" name="txtHIGHSEQNAME"></TD>
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
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												DESIGNTIMEDRAGDROP="213" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="16087">
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
	</body>
</HTML>
