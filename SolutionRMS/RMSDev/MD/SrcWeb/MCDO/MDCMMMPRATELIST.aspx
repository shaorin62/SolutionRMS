<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMMPRATELIST.aspx.vb" Inherits="MD.MDCMMMPRATELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���۴���� �������� ���</title>
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
'HISTORY    :1) 2009/11/25 By kty
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
Dim mobjMDCOMMPRATE
Dim mobjMDCOGET
Dim mlngRowCnt,mlngColCnt
Dim mstrRow

mstrRow= 1
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

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
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

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		.txtCLIENTCODE.focus
		.sprSht.focus
	End With 
end sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' ����� �ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			selectrtn
     	end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
					selectrtn
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' ��Ʈ �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	with frmThis
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim strCode
	Dim strCodeName
	Dim vntData
	
	with frmThis
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"FDATE") Then
			strFDATE = mobjSCGLSpr.GetTextBinding(.sprSht,"FDATE",Row) 
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row) 
			
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			strEDATE = mobjMDCOMMPRATE.SelectRtn_FDATE(gstrConfigXml,mlngRowCnt,mlngColCnt,strFDATE, strCLIENTCODE)
			
			IF strEDATE <> "" THEN
				gErrorMsgBox "�ߺ��� ��¥�Դϴ�.","ó���ȳ�"	
				mobjSCGLSpr.SetTextBinding .sprSht,"FDATE",Row, strEDATE
			END IF 
			
		END IF 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EDATE") Then
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"EDATE",Row) <> "" THEN
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"FDATE",Row) > mobjSCGLSpr.GetTextBinding(.sprSht,"EDATE",Row)  THEN
					gErrorMsgBox "�������� �����Ϻ��� �����������ϴ�.","ó���ȳ�"	
					mobjSCGLSpr.SetTextBinding .sprSht,"EDATE",Row, mobjSCGLSpr.GetTextBinding(.sprSht,"FDATE",Row)
				END IF 
			END IF 
		END IF 
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		mobjSCGLSpr.SetTextBinding .sprSht,"INPUTFLAG",Row, "INPUT"
	End With
	
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		.sprSht.Focus
	End With
End Sub

'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	with frmThis
		If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
		
		If KeyCode = meINS_ROW Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht, cint(KeyCode), cint(Shift), -1, 1)
					
			mobjSCGLSpr.SetCellsLock2 .sprSht,false,"FDATE | EDATE | RATE | BIGO ",.sprSht.MaxRows,.sprSht.MaxRows,false
			
			mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ACTIVEROW, Mid(gNowDate,1,4) & Mid(gNowDate,6, 2)
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ACTIVEROW, mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",.sprSht.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht,"RATE",.sprSht.ACTIVEROW, 0
			mobjSCGLSpr.ActiveCell .sprSht, 1,.sprSht.MaxRows
			.txtCLIENTNAME.focus
			.sprSht.focus
			mobjSCGLSpr.ActiveCell .sprSht, 3,2
		End If
	End with
End Sub


Sub sprSht_Keyup(KeyCode, Shift)
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
	End If
End Sub

'��Ʈ ��ư Ŭ�� �̺�Ʈ
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End If
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col, Row
	
		End If
	End with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDCOMMPRATE	= gCreateRemoteObject("cMDCO.ccMDCOMMPRATE")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    
	gSetSheetDefaultColor
    with frmThis
		'**************************************************
		'*** ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 2, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "YEARMON | CLIENTCODE | BTN | CLIENTNAME | SEQ | FDATE | EDATE | RATE | BIGO | INPUTFLAG | CUSER | UUSER"
		mobjSCGLSpr.SetHeader .sprSht,		    "��ϳ��|�������ڵ�|�����ָ�|����|������|������|��������(%)|���|INPUTFLAG|������|������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "       9|         9|2|    12|   4|     9|     9|         12|  15|        0|     7|     7"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "FDATE | EDATE "
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "YEARMON | CLIENTCODE | SEQ | FDATE | EDATE | RATE | BIGO | CUSER | UUSER"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQ |YEARMON | CLIENTCODE | CUSER | UUSER ",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME | BIGO",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "INPUTFLAG", true
		
	End with

	pnlTab1.style.visibility = "visible" 
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCOMMPRATE = Nothing
	set mobjMDCOGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0

	'.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	End with
End Sub

Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strCLIENTCODE
   	Dim strCLIENTNAME
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
				
		vntData = mobjMDCOMMPRATE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE,strCLIENTNAME)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   			else
   				.sprSht.MaxRows = 0
   			End If
   			
   			For intCnt = 1 To .sprSht.MaxRows 
				If mobjSCGLSpr.GetTextBinding(.sprSht,"EDATE",intCnt) <> "" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HAAF290, &H000000,False 
				ELSE
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
				End If 
				
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"FDATE",intCnt) <> "" AND  mobjSCGLSpr.GetTextBinding( .sprSht,"EDATE",intCnt) <> "" Then
					mobjSCGLSpr.SetCellsLock2 .sprSht,TRUE,"FDATE | EDATE | RATE | BIGO",intCnt,intCnt,false
				ELSE
					mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,"EDATE | RATE | BIGO ",intCnt,intCnt,false
				END IF
				
			Next
   			gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		End If
   	end with
End Sub

Sub ProcessRtn()
	Dim intRtn
   	Dim vntData
   	Dim intCnt
   	Dim strFDATE,strEDATE, strCLIENTCODE
   	
	with frmThis
		mstrRow = .sprSht.ActiveRow
		
		IF .sprSht.MAXROWS = 0 THEN
			gErrorMsgBox "������ ����Ÿ�� �����ϴ�.","����ȳ�"
			exit Sub
		END IF 
		
		IF mobjSCGLSpr.GetTextBinding( .sprSht,"FDATE",.sprSht.ActiveRow) = ""  THEN
			gErrorMsgBox "�������� �ʼ��Դϴ�.","����ȳ�"
			exit Sub
		END IF
		
		for intcnt =1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"INPUTFLAG",intcnt) = "INPUT" AND _
				(mobjSCGLSpr.GetTextBinding( .sprSht,"RATE",intcnt)="" or mobjSCGLSpr.GetTextBinding( .sprSht,"RATE",intcnt)=0 ) then 
			
				intRtnSave = gYesNoMsgbox("�����Ͻ� �����Ϳ� �����ᰡ 0�ΰ��� �ֽ��ϴ�." & vbcrlf & "�̴�� ���� �Ͻðڽ��ϱ�?","ó���ȳ�")
				IF intRtnSave <> vbYes then 
					exit Sub
				END IF 
			end if 
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | CLIENTCODE | CLIENTNAME | SEQ | FDATE | EDATE | RATE | BIGO | INPUTFLAG | CUSER | UUSER")
				
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		'ó�� ������ü ȣ��
		intRtn = mobjMDCOMMPRATE.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "����Ǿ����ϴ�.","����ȳ�"
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
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="152" background="../../../images/back_p.gIF"
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
											<td class="TITLE">MPP ������ �������� ���</td>
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
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="90">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="34" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
										</TR>
									</TABLE>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 1040px; HEIGHT: 10px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD vAlign="middle" align="right" height="28">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																alt="�ڷḦ ���� �߰��մϴ�." src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"></TD>
								<!--���� �� �׸���-->
							</TR>
							<tr>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
										<TR>
											<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="15928">
														<PARAM NAME="_ExtentY" VALUE="10081">
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
											<TD class="BOTTOMSPLIT" id="lblStatus_CLIENT" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
								</TD>
							</tr>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
