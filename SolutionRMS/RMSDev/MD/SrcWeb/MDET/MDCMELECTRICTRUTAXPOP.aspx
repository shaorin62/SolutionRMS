<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRICTRUTAXPOP.aspx.vb" Inherits="MD.MDCMELECTRICTRUTAXPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��������</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/����/�����ڵ� �˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCDOC.aspx
'��      �� : ITEM ��ȸ�� ���� �˾�
'�Ķ�  ���� :ITEM_CODE OR NAME, ��ȸ�߰��ʵ�, ���� ������� �͸� ��ȸ���� ����,
'			  �ڵ� ������, �ڵ�Like���� ����
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Const meTab = 9
Dim mobjMDCMELECTRICTRUTAX
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode,mstrAddWhere
DIm mstrTRANSYEARMON
DIm mstrTRANSNO
DIm mstrCLIENTNAME
DIm mstrREAL_MED_NAME
DIm mstrAMT
DIm mstrVAT
DIm mstrSUMM
DIm mstrPRINTDAY
DIm mstrCLIENTCODE
DIm mstrCLIENTACCODE
DIm mstrCLIENTBISNO
DIm mstrREAL_MED_CODE
DIm mstrREAL_MED_ACCODE
DIm mstrREAL_MED_BISNO
DIm mstrMEDCODE
DIm mstrDEPTCODE
DIm mstrMEDFLAG
DIm mstrTAXYEARMON
DIm mstrSPONSOR
DIm mstrCLIENTOWNER
DIm mstrCLIENTADDR1
DIm mstrCLIENTADDR2
DIm mstrREAL_MEDOWNER
DIm mstrREAL_MEDADDR1
DIm mstrREAL_MEDADDR2
Dim mstrDEMANDDAY

'-----------------------------
' �̺�Ʈ ���ν��� 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

'����
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'�ο��߰�
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
		mobjSCGLSpr.SetTextBinding .sprSht,"SUMM",.sprSht.ActiveRow, "���� ���౤���"
	End With 
end sub

'�ο����
Sub imgDelRow_onclick
	call sprSht_KeyDown(meDEL_ROW, 0)
End Sub

'�������
Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ܰ�
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'�ݾ�
Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMTSUM, IntVATSUM, IntSUMAMTVATSUM
	Dim IntAMT, IntVAT, IntSUMAMTVAT
	Dim lngSUSU
	
	
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMAMTVATSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntVAT = 0
			IntSUMAMTVAT = 0
			
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
			IntSUMAMTVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
			IntVATSUM = IntVATSUM + IntVAT
			IntSUMAMTVATSUM = IntSUMAMTVATSUM + IntSUMAMTVAT
		Next
		mobjSCGLSpr.SetTextBinding .sprSht1,"AMT",1, IntAMTSUM
		mobjSCGLSpr.SetTextBinding .sprSht1,"VAT",1, IntVATSUM
		mobjSCGLSpr.SetTextBinding .sprSht1,"SUMAMTVAT",1, IntSUMAMTVATSUM
	End With
End Sub
'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
    if KeyCode <> meINS_ROW AND KeyCode <> meDEL_ROW  then exit sub
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					'SetDefaultNewRow
			Case meDEL_ROW: DeleteRtn_Row
		End Select
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim vntData
   	Dim i, strCols
   	Dim strAMT, strVAT, strAMTTEMP, strVATTEMP
   	Dim strCNTTEMP, strMODAMT, strCNTTEMP2, strMODVAT
		with frmThis
			If  Col = 2 Then
   				strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   				
   				mobjSCGLSpr.SetTextBinding .sprSht,"VAT",Row, strAMT/10
   				
   				strVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT",Row)
   				
   				mobjSCGLSpr.SetTextBinding .sprSht,"SUMAMTVAT",Row, strAMT+strVAT
   				
   			END IF
   			AMT_SUM
   		end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
With frmThis
	mobjSCGLSpr.SameColWidth .sprSht, .sprSht1	
End with
end sub
'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht1, NewTop, NewLeft
End Sub
'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	dim vntData, vntInParam
	dim intNo,i

	'����������ü ����	
	set mobjMDCMELECTRICTRUTAX = gCreateRemoteObject("cMDET.ccMDETELECTRICTRUTAX")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : mstrTRANSYEARMON = vntInParam(i)
				case 1 : mstrTRANSNO = vntInParam(i)
				case 2 : .txtCLIENTNAME.value = vntInParam(i)
				case 3 : .txtREAL_MED_NAME.value = vntInParam(i)
				case 4 : .txtAMT.value = vntInParam(i)
				case 5 : .txtVAT.value = vntInParam(i)
				case 6 : mstrSUMM = vntInParam(i)
				case 7 : mstrPRINTDAY = vntInParam(i)
				case 8 : mstrCLIENTCODE = vntInParam(i)
				case 9 : mstrCLIENTACCODE = vntInParam(i)
				case 10 : mstrCLIENTBISNO = vntInParam(i)
				case 11 : mstrREAL_MED_CODE = vntInParam(i)
				case 12 : mstrREAL_MED_ACCODE = vntInParam(i)
				case 13 : mstrREAL_MED_BISNO = vntInParam(i)
				case 14 : mstrMEDCODE = vntInParam(i)
				case 15 : mstrDEPTCODE = vntInParam(i)
				case 16 : mstrMEDFLAG = vntInParam(i)
				case 17 : mstrTAXYEARMON = vntInParam(i)
				case 18 : mstrSPONSOR = vntInParam(i)
				case 19 : mstrCLIENTOWNER = vntInParam(i)
				case 20 : mstrCLIENTADDR1 = vntInParam(i)
				case 21 : mstrCLIENTADDR2 = vntInParam(i)
				case 22 : mstrREAL_MEDOWNER = vntInParam(i)
				case 23 : mstrREAL_MEDADDR1 = vntInParam(i)
				case 24 : mstrREAL_MEDADDR2 = vntInParam(i)
				case 25 : mstrDEMANDDAY = vntInParam(i)
			end select
		next
		
		gSetSheetDefaultColor()
		With frmThis
            gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 4, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht, "SUMM | AMT | VAT | SUMAMTVAT"
			mobjSCGLSpr.SetHeader .sprSht, "����|���ް���|�ΰ���|�հ�"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "20 | 12 | 12 | 12"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "SUMM"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMTVAT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUMM", -1, -1, 50
			mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SUMAMTVAT"
			'mobjSCGLSpr.ColHidden .sprSht, "JOBYEARMON | JOBCUST |JOBSEQ  ", true
			mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			
			
			'���� �հ� ǥ�� �׸��� �⺻ȭ�� ����
			gSetSheetColor mobjSCGLSpr, .sprSht1
			mobjSCGLSpr.SpreadLayout .sprSht1, 4, 1, 0,0,1,1,1,false,true,true,1
			mobjSCGLSpr.SpreadDataField .sprSht1, "SUMM | AMT | VAT | SUMAMTVAT"
			mobjSCGLSpr.SetText .sprSht1, 1, 1, "��   ��"
			mobjSCGLSpr.SetScrollBar .sprSht1, 0
			mobjSCGLSpr.SetBackColor .sprSht1,"1",rgb(205,219,215),false
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT | VAT | SUMAMTVAT", -1, -1, 0
			
			mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"	  
			mobjSCGLSpr.SameColWidth .sprSht, .sprSht1
        End With
	end with	
	InitPageData
	'�ڷ���ȸ	
	'SelectRtn
end sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	txtAMT_onblur
	txtVAT_onblur
End Sub

Sub EndPage()
	set mobjMDCMELECTRICTRUTAX = Nothing
	gEndPage
End Sub


'------------------------------------------
' ���縦 �մϴ�.
'------------------------------------------
Sub ProcessRtn ()
   	Dim intRtn
    Dim intRtn2
   	dim vntData, vntData1
	Dim strMasterData
	Dim strTAXYEARMON
	Dim intTAXNO
	Dim strTAXSET
	Dim strSUMM
	'Dim strATTR02FLAG
	Dim intCnt
	Dim strDEMANDDAY,strPRINTDAY
	Dim chkcnt
	Dim intCnt2
	Dim intColFlag
	Dim intMaxCnt
	Dim bsdiv
	Dim strVALIDATION
	Dim strCLIENTNAME
	Dim strREAL_MED_NAME
	with frmThis
		
		
		'�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		'if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SUMM | AMT | VAT | SUMAMTVAT")
		
		'������ �����͸� ���� �´�.
		
		strCLIENTNAME = .txtCLIENTNAME.value
		strREAL_MED_NAME = .txtREAL_MED_NAME.value			
		
		
				
		'ó�� ������ü ȣ��
		intTAXNO = 0
		intRtn = mobjMDCMELECTRICTRUTAX.ProcessRtn_GROUP(gstrConfigXml,vntData, mstrTRANSYEARMON, mstrTRANSNO, mstrPRINTDAY, mstrCLIENTCODE, mstrCLIENTACCODE, mstrCLIENTBISNO, mstrREAL_MED_CODE, mstrREAL_MED_ACCODE, mstrREAL_MED_BISNO, mstrMEDCODE, mstrDEPTCODE, mstrMEDFLAG, mstrTAXYEARMON, mstrSPONSOR, mstrCLIENTOWNER, mstrCLIENTADDR1, mstrCLIENTADDR2, mstrREAL_MEDOWNER, mstrREAL_MEDADDR1, mstrREAL_MEDADDR2, strCLIENTNAME, strREAL_MED_NAME, mstrDEMANDDAY)


		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "���强��","����ȳ�!"
			EndPage
   		end if
   	end with
End Sub

Sub DeleteRtn_Row
	Dim intSelCnt, intRtn, i
	Dim vntData
	With frmThis
			vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
			
			if intSelCnt < 1 then
				gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
				Exit sub
			end if
			
			intRtn = gYesNoMsgbox("�Է��ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
			
			if intRtn <> vbYes then exit sub
			
			'���õ� �ڷḦ ������ ���� ����
			for i = intSelCnt-1 to 0 step -1
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			next
			'���� ���� ����
			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<XML id="xmlBind"></XML>
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 234px" align="left" width="214" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">&nbsp;������ ���Ҽ��ݰ�꼭 ����</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 250px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 88px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="88" border="0">
										<TR>
											<TD width="3"><FONT face="����"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
														width="54" border="0" name="imgSave"></FONT></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="1" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 518px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 518px; HEIGHT: 29px" vAlign="middle" height="29">
									<TABLE class="DATA" height="24" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" width="80"><FONT face="����">������</FONT></TD>
											<TD class="DATA" width="168"><INPUT class="NOINPUTB_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 168px; HEIGHT: 22px"
													type="text" size="22" name="txtCLIENTNAME"></TD>
											<TD class="LABEL" width="80">��ü��</TD>
											<TD class="DATA" width="169"><INPUT class="NOINPUTB_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 168px; HEIGHT: 22px"
													type="text" size="22" name="txtREAL_MED_NAME"></TD>
										</TR>
										<TR>
											<TD class="LABEL">�ݾ�</TD>
											<TD class="DATA"><INPUT class="NOINPUTB_R" id="txtAMT" title="�ݾ�" style="WIDTH: 136px; HEIGHT: 22px" type="text"
													size="16" name="txtAMT"></TD>
											<TD class="LABEL">�ΰ���</TD>
											<TD class="DATA"><INPUT class="NOINPUTB_R" id="txtVAT" title="�ΰ���" style="WIDTH: 136px; HEIGHT: 22px" type="text"
													size="17" name="txtVAT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 518px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD style="HEIGHT: 26px" align="right" width="100%"><IMG id="imgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gIF'" height="20" alt="�� �� �߰�" src="../../../images/imgAddRow.gIF"
										width="54" border="0" name="imgAddRow"><IMG id="imgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gIF'" height="20" alt="�� �� ����" src="../../../images/imgDelRow.gIF"
										width="54" border="0" name="imgDelRow">
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 518px" align="center"><FONT face="����">
										<OBJECT id="sprSht" style="WIDTH: 509px; HEIGHT: 252px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="13467">
											<PARAM NAME="_ExtentY" VALUE="6668">
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
										<OBJECT id="sprSht1" style="WIDTH: 509px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="13467">
											<PARAM NAME="_ExtentY" VALUE="635">
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
											<PARAM NAME="EditEnterAction" VALUE="5">
											<PARAM NAME="EditModePermanent" VALUE="0">
											<PARAM NAME="EditModeReplace" VALUE="0">
											<PARAM NAME="FormulaSync" VALUE="-1">
											<PARAM NAME="GrayAreaBackColor" VALUE="12632256">
											<PARAM NAME="GridColor" VALUE="12632256">
											<PARAM NAME="GridShowHoriz" VALUE="1">
											<PARAM NAME="GridShowVert" VALUE="1">
											<PARAM NAME="GridSolid" VALUE="1">
											<PARAM NAME="MaxCols" VALUE="15">
											<PARAM NAME="MaxRows" VALUE="0">
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
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD height="5"></TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 518px"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM></TR>
		</TABLE>
	</body>
</HTML>
