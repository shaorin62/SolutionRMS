<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGELISTPOP.aspx.vb" Inherits="PD.PDCMCHARGELISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� �̷°���</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ó�� ǥ�� ����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
'****************************************************************************************
-->
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet ���� --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjccPDDCCHARGEEXCOM 
Dim mobjPDCMGET
Dim mobjMDCMGET
Dim mstrCheck
mstrCheck = True
Dim mstrGFLAG
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	if frmThis.txtJOBNO.value = "" OR frmThis.txtOUTSCODE.value = "" OR frmThis.txtREVSEQ.value = "" then
		gErrorMsgBox "���۰�����ȣ/����ó�ڵ�/������ �־��ּ���.!",""
		exit Sub
	end if
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'frmThis.imgSetting.style.visibility = "hidden"
	'frmThis.ImgConfirmCancel.style.visibility = "hidden"
	'����������ü ����	
	Set mobjccPDDCCHARGEEXCOM = gCreateRemoteObject("cPDCO.ccPDDCCHARGEEXCOM")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 6, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, "ITEMNAME|STD|QTY|PRICE|AMT|BIGO"
        mobjSCGLSpr.SetHeader           .sprSht, "�����׸�|�԰�|����|�ܰ�|�ݾ�|���"
        mobjSCGLSpr.SetColWidth .sprSht, "-1","         22|  10|  10|  11|  13|  30"
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, "ITEMNAME|STD|BIGO"     , , ,200
        mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "QTY|PRICE|AMT", -1, -1, 0
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"     
		
		
    End With
    
	'ȭ�� �ʱⰪ ����
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)	
				case 1 : .txtJOBNAME.value = vntInParam(i)
				case 2 : .txtOUTSCODE.value = vntInParam(i)		'���� ������� �͸�
				case 3 : .txtOUTSNAME.value = vntInParam(i)		'�ڵ� ��� ����
				case 4 : .txtREVSEQ.value = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjccPDDCCHARGEEXCOM = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	'�ʱ� ������ ����
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0	
		
	End with
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value, .txtOUTSCODE.value, .txtREVSEQ.value)

		if not gDoErrorRtn ("SelectRtn") then
			'mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			 mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub


-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%"  HEIGHT="100%" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">�������
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 225px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 108px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD style="WIDTH: 126px"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody"  HEIGHT="100%" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="0" cellPadding="0" width="1040" align="LEFT"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" width="60">���۹�ȣ</TD>
											<td class="SEARCHDATA" width="315"><INPUT class="INPUT_L" id="txtJOBNAME" title="�ڵ��" style="WIDTH: 227px; HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="32" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
													align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 64px; HEIGHT: 22px" type="text"
													maxLength="6" size="5" name="txtJOBNO"></td>
											<TD class="SEARCHLABEL" width="60">
											����ó
											<td class="SEARCHDATA" width="317"><INPUT class="INPUT_L" id="txtOUTSNAME" title="�ڵ��" style="WIDTH: 227px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="16" name="txtOUTSNAME"><IMG id="imgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="imgOUTSCODE"><INPUT class="INPUT" id="txtOUTSCODE" title="jobno" style="WIDTH: 64px; HEIGHT: 22px" type="text"
													maxLength="6" size="5" name="txtOUTSCODE"></td>
											<TD class="SEARCHLABEL" width="60">
											����
											<td class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtREVSEQ" title="����" style="WIDTH: 64px; HEIGHT: 22px" type="text"
													maxLength="6" size="5" name="txtREVSEQ"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="����">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 90%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="23918">
											<PARAM NAME="_ExtentY" VALUE="9287">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM></TR>
		</TABLE>
	</body>
</HTML>
