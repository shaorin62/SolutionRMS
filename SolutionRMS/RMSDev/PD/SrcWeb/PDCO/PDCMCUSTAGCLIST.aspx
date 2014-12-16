<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCUSTAGCLIST.aspx.vb" Inherits="PD.PDCMCUSTAGCLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ��� ���</title>
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
'HISTORY    :1) 2008/08/25 By hwang duck-su
'			 2) 
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" >
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMCUSTLIST, mobjPDCMCUSTAGCLIST '�����ڵ�, Ŭ����
Dim mobjPDCMGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9

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
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

'-----------------------------
'���߰�
'-----------------------------
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub


'-----------------------------------
'����    -
'-----------------------------------
Sub imgSave_onclick ()
	IF frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------
'����   - ���� �������� ����
'-----------------------------------
Sub imgDelete_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
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

'-----------------------------------
'��Ŭ��
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


'-----------------------------------
'���羲������.
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'-----------------------------------
'��Ʈ���� �˾�
'-----------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""

		IF Col = 5 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CUSTNAME",Row, vntRet(0,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
			
		end if
		
		.sprSht.focus 

	End with
	
End Sub


'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				'Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' ������ ȭ�� ������ �� �ʱ�ȭ 
'----------------------------------------------------------------------
	'����������ü ����	
	set mobjPDCMCUSTAGCLIST = gCreateRemoteObject("cPDCO.ccPDCOCUSTLIST_HWANG")
	'set mobjPDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET") ���� ���� ����
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis

	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 15, 0
	mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
	mobjSCGLSpr.SpreadDataField .sprSht, "CUSTCODE | HIGHCUSTCODE | ACCUSTCODE | CUSTTYPE | BTN |CUSTNAME | COMPANYNAME | CUSTOWNER | MEDFLAG | BUSINO | JURENTNO|ATTR06|HIGHCUSTNAME|ATTR10|MEMO"
	mobjSCGLSpr.SetHeader .sprSht, "�ŷ�ó�ڵ�|û�����ڵ�|ȸ��ŷ�ó�ڵ�|�ŷ�ó����|�ŷ�ó��|��ȣ��|��ǥ�ڼ���|�ŷ�ó����|����ڵ�Ϲ�ȣ|���ε�Ϲ�ȣ|�д㱤���ּ���|�����ŷ�ó��|��뿩��|���"
	mobjSCGLSpr.SetColWidth .sprSht, "-1", " 9|         9|             0|        12|      2|20|    20|        15|        10|            15|          15|             0|           0|0|20"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
	mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CUSTCODE | HIGHCUSTCODE | ACCUSTCODE | CUSTTYPE |  CUSTNAME | COMPANYNAME | CUSTOWNER | MEDFLAG | BUSINO | JURENTNO | ATTR10|MEMO", -1, -1, 200
	mobjSCGLSpr.colhidden .sprSht, "ATTR06|HIGHCUSTNAME|ACCUSTCODE|ATTR10|MEMO",true
	
	

			
	.sprSht.style.visibility = "visible"
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCMCUSTLIST = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis

	'�ʱ� ������ ����
	with frmThis
		.sprSht.MaxRows = 0
	End With
	'���ο� XML ���ε��� ����
	'cmbMEDFLAG_onchange
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'--------------------------------------------------
'txt �ʱ�ȭ
'--------------------------------------------------
Sub DataClean
	with frmThis
		.txtCUSTNAME.value = ""
		.txtBUSINO.value = ""
		.sprSht.MaxRows = 0
	End With
End Sub





'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim Flag
   	Flag = "G"
   	
	with frmThis

		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMCUSTAGCLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, Flag, .cmbMEDDIV1.value, .txtCUSTNAME.value, .txtBUSINO.value )

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
   			
   		end if
   		
   	end with
 
End Sub

'------------------------------------------
'������ ����
'------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCODE
	
	with frmThis
		If .txtPROJECTNO.value = "" Or .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ �ڷᰡ �����ϴ�.","�����ȳ�"
			Exit Sub
		End If
		intSelCnt = 0
		strCODE = Trim(.txtPROJECTNO.value)
		vntData = mobjPONO.GetPONODELSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
		
	
		If mlngRowCnt > 0 Then
			gErrorMsgBox "JOBNO �� ��ϵǾ��ִ� PROJECT �Դϴ�.�����ɼ� �����ϴ�.","�����ȳ�"
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'�ڷ� ����
		intRtn = mobjPONO.DeleteRtn(gstrConfigXml,strCODE)
			
		IF not gDoErrorRtn ("DeleteRtn") then
			'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			msgbox "[" & strCODE & "] PROJECT �� �����Ǿ����ϴ�."
   		End IF
		'���� ���� ����
		SelectRtn
	End with
	err.clear
End Sub


'------------------------------------------
' ������ ����/���� ó�� 
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strJOBNO,strDEMANDAMT,strJOBYEARMON
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	
	
	
	with frmThis
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CUSTCODE | HIGHCUSTCODE | ACCUSTCODE | CUSTTYPE | CUSTNAME | COMPANYNAME | CUSTOWNER | MEDFLAG | BUSINO | JURENTNO|ATTR06|HIGHCUSTNAME|ATTR10|MEMO")
		
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		
		intRtn = mobjPDCMCUSTAGCLIST.ProcessRtn(gstrConfigXml,vntData,strJOBNO)
	
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht.ActiveRow
			SelectRtn
			'Call sprSht_Click(1,strRow)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		end if
   		
   	end with
End Sub


'------------------------------------------
'��ƮŰ�ٿ�
'------------------------------------------

Sub sprSht1_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 13 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
	'	DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
		'	Case meINS_ROW':
		'			DefaultValue
		'	Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML><XML id="xmlBind1"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD style="HEIGHT: 100%">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">�ŷ�ó ����(������)</td>
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
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 17px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="1024" border="0"
											align="left">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 100px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTCODE1, txtCUSTNAME1)">��ü����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 220px"><SELECT id="cmbMEDDIV1" style="WIDTH: 86px" name="cmbMEDDIV1">
														<OPTION value="" selected>��ü</OPTION>
														<OPTION value="1">������TV</OPTION>
														<OPTION value="1">������RD</OPTION>
														<OPTION value="1">������DMB</OPTION>
														<OPTION value="2">CABLE-TV</OPTION>
														<OPTION value="3">�μ�</OPTION>
														<OPTION value="4">���ͳ�</OPTION>
														<OPTION value="5">����</OPTION>
														<OPTION value="0">��Ÿ</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 76px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,'')">�ŷ�ó��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 195px"><INPUT class="INPUT_L" id="txtCUSTNAME" title="�ڵ���ȸ" style="WIDTH: 128px; HEIGHT: 22px"
														type="text" maxLength="15" align="left" size="16" name="txtCUSTNAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')">�ŷ�ó����ڹ�ȣ</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtBUSINO" title="�ڵ���ȸ" style="WIDTH: 128px; HEIGHT: 22px" type="text"
														maxLength="15" align="left" size="16" name="txtBUSINO"></TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
															<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow">
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." 
																	src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
															<TD><IMG id="imgModySave" onmouseover="JavaScript:this.src='../../../images/imgModySaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgModySave.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgModySave.gIF" border="0" name="imgModySave"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
																	name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 95%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 95%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27120">
												<PARAM NAME="_ExtentY" VALUE="13996">
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
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
