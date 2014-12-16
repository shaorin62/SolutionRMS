<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUDCUST.aspx.vb" Inherits="MD.MDCMCLOUDCUST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CGVŬ����� �ڵ� �߰� �� û���� ��Ī</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/������ ��ü�ڵ� ��� ȭ��(MDCMREALMEDCODEMST)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/08 By Ȳ����
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
Dim mobjMDOTCLOUDCUST
Dim mobjMDCOGET
Dim mlngRowCnt,mlngColCnt

Dim intSelectRows 'lock�� �ɱ����� ��ȸ�ؿ� row���� ������ �ִ´�.
Dim mUploadFlag

CONST meTAB = 9
intSelectRows = 0
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'�ű� - �űԽ� �������˾�
Sub imgNew_onclick
	With frmThis
		CALL sprSht_Keydown(meINS_ROW, 0)
	End With
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
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

'-----------------------------------------------------------------------------------------
' ��ü���ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMEDNAME_onkeydown
	if window.event.keyCode = meEnter then
   		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	end if
End Sub

Sub txtBUSINO_onkeydown
	if window.event.keyCode = meEnter then
   		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	end if
End Sub

Sub txtREAL_MED_NAME_onkeydown
	if window.event.keyCode = meEnter then
   		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	end if
End Sub

'-----------------------------------------------------------------------------------------
' onchange�̺�Ʈ
'-----------------------------------------------------------------------------------------
Sub cmbUSE_FLAG_onchange
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
'---------------------------------
' �������� ��Ʈ ����� üũ 
'--------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strRow
	
	with  frmThis
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		
		If KeyCode = meINS_ROW Then
			intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_CODE",frmThis.sprSht.ActiveRow, "B00874"
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REAL_MED_NAME",frmThis.sprSht.ActiveRow, "CJ CGV(������ ��������) (��)"
			
			mobjSCGLSpr.SetTextBinding .sprSht,"USE_FLAG", .sprSht.ActiveRow, "1"
			mobjSCGLSpr.ActiveCell .sprSht, 2, .sprSht.ActiveRow
		End If
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row) = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, trim(vntData(1,1))
												
						.txtBUSINO.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"MEDNAME"), Row
						.txtBUSINO.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_CODE",Row) = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "B")		

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntData(1,1))
												
						.txtBUSINO.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_MED_NAME"), Row
						.txtBUSINO.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF	
	End With

	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	Dim vntRet, vntInParams
    
    With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(sprSht,"MEDNAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding(sprSht,"MEDNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding sprSht,"MEDCODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding sprSht,"MEDNAME",Row, trim(vntRet(1,0))

				mobjSCGLSpr.CellChanged sprSht, Col,Row
				mobjSCGLSpr.ActiveCell sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(sprSht,"REAL_MED_NAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( sprSht,"REAL_MED_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding sprSht,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding sprSht,"REAL_MED_NAME",Row, trim(vntRet(1,0))

				mobjSCGLSpr.CellChanged sprSht, Col,Row
				mobjSCGLSpr.ActiveCell sprSht, Col+2,Row
			End If
		End If
		
		.txtBUSINO.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		sprSht.Focus
	End With
End Sub


Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNMED") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht,"MEDNAME",Row, trim(vntRet(1,0))
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		END IF
		
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNREAL") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",Row, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_NAME",Row, trim(vntRet(1,0))
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		END IF
		
		.txtBUSINO.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.sprSht.Focus
		mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
	End with
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDOTCLOUDCUST = gCreateRemoteObject("cMDOT.ccMDOTCLOUDCUST")
	set mobjMDCOGET		  = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'***Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 2, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "SEQ | MEDCODE | BTNMED | MEDNAME | REAL_MED_CODE | BTNREAL | REAL_MED_NAME | USE_FLAG"
		mobjSCGLSpr.SetHeader .sprSht,		    "����|��ü���ڵ�|��ü���|û�����ڵ�|û������|��뱸��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|         7|2|    20|         7|2|    30|      8    "
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "USE_FLAG"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME", -1, -1, 255
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTNMED | BTNREAL"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SEQ"
	End with

	pnlTab1.style.visibility = "visible" 
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub InitPageData
	gClearAllObject frmThis
End Sub

Sub EndPage()
	set mobjMDOTCLOUDCUST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
	Dim strSEQ
	Dim strBCODE
	Dim strCUSTCODE
	Dim strMEMO
	
	With frmThis
	
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		.sprSht.MaxRows = 0
		
		strMEDNAME		= .txtMEDNAME.value
		strBUSINO		= .txtBUSINO.value
		strREAL_MED_CODE= .txtREAL_MED_NAME.value
		strUSE_FLAG		= .cmbUSER_FLAG.value
		
		vntData = mobjMDOTCLOUDCUST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strMEDNAME, strBUSINO, strREAL_MED_CODE, strUSE_FLAG)
		
		intSelectRows = mlngRowCnt
		
		If not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		End if
   	End With
End Sub

Sub ProcessRtn()
	Dim intRtn
   	Dim vntData
	Dim strDataCHK
   	
	with frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'��ȸ�� row �����ٺ��� �űԵ����͸� validation
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "MEDCODE | REAL_MED_CODE",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ��ü���ڵ�/û�����ڵ�� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		End If
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SEQ | MEDCODE | BTNMED | MEDNAME | REAL_MED_CODE | BTNREAL | REAL_MED_NAME | USE_FLAG")
		
		'ó�� ������ü ȣ��
		intRtn = mobjMDOTCLOUDCUST.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox intRtn & " �� �� ����Ǿ����ϴ�.","����ȳ�"
			SelectRtn
   		end if
   	end with
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="98%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="600" border="0">
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
											<td class="TITLE">CGVŬ����� �ڵ� �߰� �� û���� ��Ī</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 95%" cellSpacing="0" cellPadding="0" width="1040"
							border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME,'')"
												width="50">��ü��
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 180px"><INPUT class="INPUT_L" id="txtMEDNAME" style="WIDTH: 176px; HEIGHT: 22px" type="text" maxLength="100"
													size="24" name="txtMEDNAME" title="��ü��"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
												width="60">����ڹ�ȣ
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 150px"><INPUT class="INPUT_L" id="txtBUSINO" style="WIDTH: 152px; HEIGHT: 22px" type="text" maxLength="20"
													size="20" name="txtBUSINO" title="����ڹ�ȣ"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME,'')"
												width="50">û����
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 180px"><INPUT class="INPUT_L" id="txtREAL_MED_NAME" style="WIDTH: 176px; HEIGHT: 22px" type="text"
													maxLength="100" size="24" name="txtREAL_MED_NAME" title="û����"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbUSER_FLAG,'')"
												width="50">��뱸��
											</TD>
											<TD class="SEARCHDATA"><SELECT id="cmbUSER_FLAG" title="��뱸��" style="WIDTH: 104px" name="cmbUSER_FLAG">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="1">���</OPTION>
													<OPTION value="0">�̻��</OPTION>
												</SELECT>
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 10px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<!--���̺��� �������°��� �����ش�-->
									<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
										<TR>
											<TD align="left" width="100%" height="1"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD id="tblSheet" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31856">
											<PARAM NAME="_ExtentY" VALUE="15319">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR> 
			<!--List End-->
			<!--Bottom Split Start-->
			<!--Bottom Split End--> </TABLE> 
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
