<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOBATCHLOGLIST.aspx.vb" Inherits="SC.SCCOBATCHLOGLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ڵ�ó�� ����</title>
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
'HISTORY    :1) 2003/04/29 By hwnagducksu
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
		
'�������� ����
Dim mobjSCCOCUSTLIST
Dim mlngRowCnt,mlngColCnt

CONST meTAB = 9

'---------------------------------------------------
' �ű� SAP ���޾ƿ���
'---------------------------------------------------
Sub Set_CustValue (strBUSINO,strBANKTYPE)
	Dim intRtn
	With frmThis
		intRtn = mobjSCCOCUSTLIST.ProcessRtnRFC(gstrConfigXml,strBUSINO,strBANKTYPE, .cmbCUST.value)
		
		if not gDoErrorRtn ("Set_CustValue") then
			mobjSCGLSpr.SetFlag  frmThis.sprSht_BUSINO,meCLS_FLAG
			if intRtn > 0 Then
				gErrorMsgBox "������Ʈ�� �Ϸ� �Ǿ����ϴ�.","������Ʈ �ȳ�"
			Else
				gErrorMsgBox "������Ʈ�� ���� �Ͽ����ϴ�...","����ȳ�"
			End If
		end if

	End With
End Sub
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'--------------------------
'------��ư �̺�Ʈ --------
'--------------------------
'-----------------------------------
'��ȸ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis	
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'--RFC �����Ͽ� �ֽ� �����ͷ� ������Ʈ �Ѵ�.
sub imgRFC_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		if .sprSht_BUSINO.Maxrows = 0 then
			gErrorMsgBox "�����Ͱ� �����ϴ�. �����͸� ��ȸ �Ͻʽÿ�. ","RFC �Լ� ȣ�� �ȳ�!"
			exit sub
		elseif .sprSht_BUSINO.Maxrows = 1 then
			gErrorMsgBox "RFC ȣ���� ��� �����Ͱ� 2�ο� �̻��־�� �մϴ�.","RFC �Լ� ȣ�� �ȳ�!"
			exit sub
		end if 
		PROCESSRTN_RFC_CALL
	end with
	gFlowWait meWAIT_OFF
end sub

'-----------------------------------
'��Ʈ Ŭ��
'-----------------------------------
Sub sprSht_BUSINO_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_DTL Col, Row
			Selectrtn_BANK Col, Row
		End If
	End With
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
sub sprSht_BUSINO_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_BUSINO, ""
		End If
	End With
End sub

sub sprSht_DTL_Click (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprsht, ""
		End If
	End With
End sub

sub sprSht_BANK_Click (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_BANK, ""
		End If
	End With
End sub

Sub sprSht_BUSINO_Keyup(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_DTL frmThis.sprSht_BUSINO.ActiveCol,frmThis.sprSht_BUSINO.ActiveRow
		SelectRtn_BANK frmThis.sprSht_BUSINO.ActiveCol,frmThis.sprSht_BUSINO.ActiveRow
		
	End If
End Sub


Sub cmbCUST_onchange
	SelectRtn
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'����������ü ����	
	set mobjSCCOCUSTLIST = gCreateRemoteObject("cSCCO.ccSCCOCUSTLIST")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
   
    Dim intGBN
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		'-----�⺻ ����� ��ȸ �˻�
		gSetSheetColor mobjSCGLSpr, .sprSht_BUSINO
		mobjSCGLSpr.SpreadLayout .sprSht_BUSINO, 3, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_BUSINO,    "NO | BUSINO | CUSTNAME"
		mobjSCGLSpr.SetHeader .sprSht_BUSINO,		    "���� | ����ڹ�ȣ|����ڸ�"
		mobjSCGLSpr.SetColWidth .sprSht_BUSINO, "-1",  "     4|         14|      20"
		mobjSCGLSpr.SetRowHeight .sprSht_BUSINO, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_BUSINO, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_BUSINO, "NO | BUSINO | CUSTNAME", -1, -1, 200
		mobjSCGLSpr.SetCellAlign2 .sprSht_BUSINO, "NO|BUSINO|CUSTNAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht_BUSINO,true,"NO|BUSINO|CUSTNAME"
		'mobjSCGLSpr.CellGroupingEach .sprSht_BUSINO,""
		
		'-----�⺻ ����� �⺻���� 
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 7, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CUSTNAME | CUSTOWNER | BUSISTAT | BUSITYPE | ADDRESS1 | ADDRESS2 | TEL"
		mobjSCGLSpr.SetHeader .sprSht,		    " ����ڸ�|��ǥ�ڸ�|����|����|�ּ�1|�ּ�2|��ȭ��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "       15|       8|  10|  10|   15|   20|      10"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTNAME | CUSTOWNER | BUSISTAT | BUSITYPE | ADDRESS1 | ADDRESS2 | TEL", -1, -1, 200
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTNAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CUSTNAME | CUSTOWNER | BUSISTAT | BUSITYPE | ADDRESS1 | ADDRESS2 | TEL"
		'mobjSCGLSpr.CellGroupingEach .sprSht_BUSINO,""
		
		'-----BANK TYPE ��ȸ
		gSetSheetColor mobjSCGLSpr, .sprSht_BANK
		mobjSCGLSpr.SpreadLayout .sprSht_BANK, 4, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_BANK,    "BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER"
		mobjSCGLSpr.SetHeader .sprSht_BANK,		    "����Ű|���¹�ȣ|��ũŸ��|����������"
		mobjSCGLSpr.SetColWidth .sprSht_BANK, "-1",  "   10|      18|       8|        25"
		mobjSCGLSpr.SetRowHeight .sprSht_BANK, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_BANK, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_BANK, "BANK_TYPE | BANK_USER", -1, -1, 200
		mobjSCGLSpr.SetCellAlign2 .sprSht_BANK, "BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER",-1,-1,2,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht_BANK,true,"BANK_KEY | BANK_NUM | BANK_TYPE | BANK_USER"
		mobjSCGLSpr.CellGroupingEach .sprSht_BANK,"BANK_USER"
	
		pnlTab1.style.visibility = "visible" 
		pnlTab2.style.visibility = "visible" 
		pnlTab3.style.visibility = "visible" 
		
    End with
    
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtBUSINO.value = ""
		.sprSht_BUSINO.maxrows = 0
		.sprSht_BANK.maxrows = 0
		.sprSht.maxrows = 0
		.txtTO.value = 1
		.txtFROM.value = 2000
	End with
End Sub

Sub EndPage()
	set mobjSCCOCUSTLIST = Nothing
	gEndPage	
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub SelectRtn ()
   	Dim vntData
   	Dim strBUSINO
   	Dim strMEDFLAG
   	Dim lngTO, lngFROM
	'On error resume next
	with frmThis
		strBUSINO = ""
		strMEDFLAG = ""
		lngTO = 0
		lngFROM = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strBUSINO	= replace(.txtBUSINO.value,"-","")
		strMEDFLAG	= .cmbCUST.value
		lngTO		= .txtTO.value
		lngFROM		= .txtFROM.value
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_BUSINO(gstrConfigXml,mlngRowCnt,mlngColCnt,strBUSINO,strMEDFLAG,lngTO,lngFROM)
		
		if not gDoErrorRtn ("SelectRtn_BUSINO") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_BUSINO, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				Call SelectRtn_DTL(1,1)
   				Call SelectRtn_BANK(1,1)
   			Else
   				.sprSht_BUSINO.MaxRows = 0
   				.sprSht_BANK.MaxRows = 0
   				.sprSht.MaxRows = 0
   				gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   			
   			
   		end if
   	end with
End Sub

'------------------------------------------
' DTL ������ ��ȸ(������� �⺻ ���� ��ȸ)
'------------------------------------------
Sub SelectRtn_DTL(ByVal Col, ByVal Row)
	Dim vntData
	Dim i
	Dim strBUSINO, strMEDFLAG
	
	With frmThis
		.sprSht.MaxRows = 0
		strBUSINO = ""
		strMEDFLAG = ""
		
		strBUSINO = replace(mobjSCGLSpr.GetTextBinding( .sprSht_BUSINO,"BUSINO",Row),"-","")
		strMEDFLAG = .cmbCUST.value
			
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strBUSINO, strMEDFLAG)

		If not gDoErrorRtn ("SelectRtn_CUSTDTL") Then
			If mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			Else
   				'.sprSht.MaxRows = 0
			End If
		End If	
	End With
End Sub


'------------------------------------------
' DTL ������ ��ȸ(BANK_TYPE ��ȸ)
'------------------------------------------
Sub SelectRtn_BANK(ByVal Col, ByVal Row)
	Dim vntData
	Dim i
	Dim strBUSINO
	
	With frmThis
		.sprSht_BANK.MaxRows = 0
		strBUSINO = ""
		strBUSINO = replace(mobjSCGLSpr.GetTextBinding( .sprSht_BUSINO,"BUSINO",Row),"-","")
			
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		vntData = mobjSCCOCUSTLIST.SelectRtn_BANK(gstrConfigXml,mlngRowCnt,mlngColCnt, strBUSINO )

		If not gDoErrorRtn ("SelectRtn_BANK") Then
			If mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding frmThis.sprSht_BANK, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			Else
   				.sprSht_BANK.MaxRows = 0
			End If
		End If	
	End With
End Sub

Sub PROCESSRTN_RFC_CALL()
	Dim i
	Dim strBUSINO
	Dim intCNT, intRtn
	with frmThis
		
		intCNT = 0
		strBUSINO = ""
		
		intRtn = gYesNoMsgbox("RFC �� ȣ���Ͽ� �ش� �����͸� �ֽ� ������ UPDATE �մϴ�." & vbCrlf & " �����Ͱ� �������� �ð��� ���� �ɸ��ϴ�.. ������Ʈ �Ͻðڽ��ϱ�? ","ó���ȳ�!")
		IF intRtn <> vbYes then exit Sub
		
		for i = 1 to .sprSht_BUSINO.maxrows
			if i = 1 then
				strBUSINO = replace(mobjSCGLSpr.GetTextBinding( .sprSht_BUSINO,"BUSINO",i),"-","")
			else
				strBUSINO = strBUSINO & + "|" + replace(mobjSCGLSpr.GetTextBinding( .sprSht_BUSINO,"BUSINO",i),"-","")
			end if 
			intCNT = intCNT + 1
		next
		
		Set_IframeValue strBUSINO , intCNT
	end with
end sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
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
												<TABLE cellSpacing="0" cellPadding="0" width="83" background="../../../images/back_p.gIF"
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
											<td class="TITLE">�ڵ�ó�� ����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="50" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 69px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')">����ڹ�ȣ</TD>
											<TD class="SEARCHDATA" style="WIDTH: 196px"><INPUT class="INPUT_L" id="txtBUSINO" title="�ڵ���ȸ" style="WIDTH: 168px; HEIGHT: 22px" type="text"
													maxLength="15" align="left" name="txtBUSINO">
												<asp:textbox id="txtSAPBUSINO" runat="server" Visible="False" Width="8px"></asp:textbox><asp:textbox id="txtCNT" runat="server" Visible="false" Width="8px"></asp:textbox></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 67px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')">�ŷ�ó����</TD>
											<td>
												<SELECT id="cmbCUST" title="����" style="WIDTH: 100px" name="cmbCUST">
													<OPTION value="A" selected>������</OPTION>
													<OPTION value="B">��ü��</OPTION>
													<OPTION value="G">������</OPTION>
													<OPTION value="M">����ó</OPTION>
												</SELECT>
												<INPUT class="INPUT_L" id="txtTO" title="������ȸ" style="WIDTH: 50px; HEIGHT: 22px" type="text"
													maxLength="15" align="left" name="txtTO" accessKey="NUM">~<INPUT class="INPUT_L" id="txtFROM" title="������ȸ" style="WIDTH: 50px; HEIGHT: 22px" type="text"
													maxLength="15" align="left" name="txtFROM" accessKey="NUM">
											</td>
											<TD class="SEARCHDATA" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgRFC" onmouseover="JavaScript:this.src='../../../images/imgRFCOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgRFC.gIF'" height="20" alt="RFC�� �����Ͽ� �ֽ� �����ͷ� update�մϴ�.."
																src="../../../images/imgRFC.gIF" border="0" name="imgRFC"></TD>
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
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							<!--���� �� �׸���-->
							<TR>
								<!--����-->
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
										<TR>
											<td style="WIDTH: 320px; HEIGHT: 100%" vAlign="top" align="left">
												<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 320px; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_BUSINO" height="100%" width="320" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
														VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="8467">
														<PARAM NAME="_ExtentY" VALUE="17489">
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
												<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 50%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="23283">
														<PARAM NAME="_ExtentY" VALUE="8758">
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
												<DIV id="pnlTab3" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 50%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_BANK" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="23283">
														<PARAM NAME="_ExtentY" VALUE="8758">
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
											<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End-->
					</TD>
				</TR>
				<!--Top TR End-->
			</TABLE>
		</FORM>
		<iframe id="frmSapCon" style="DISPLAY: none; WIDTH: 600px; HEIGHT: 600px" name="frmSapCon"
			src="SCCOSAPBUSINO.aspx"></iframe>
	</body>
</HTML>
