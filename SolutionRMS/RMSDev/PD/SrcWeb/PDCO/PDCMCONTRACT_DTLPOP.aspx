<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT_DTLPOP.aspx.vb" Inherits="PD.PDCMCONTRACT_DTLPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��༭ �󼼳���</title>
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
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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

Dim mobjPDCMCONTRACT

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick()
	EndPage
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intNo, i, vntInParam
	'����������ü ����	
	set mobjPDCMCONTRACT	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "152px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
    with frmThis
		
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		
		for i = 0 to intNo
			select case i
				case 0 : .txtCONTRACTNO.value = vntInParam(i)	
						 
			end select
		next
				
		'Sheet Į�� ����
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout ������
		mobjSCGLSpr.SpreadLayout .sprSht, 9, 0

		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 6, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT"
		mobjSCGLSpr.SetHeader .sprSht,		   "�ڵ�|����ó|�����|JOBNO|JOB��/����|�ݾ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   6|    20|    10|    8|          20|  11"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REGDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "OUTSCODE | OUTSNAME | JOBNO | JOBNAME", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO | OUTSCODE",-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSCODE | OUTSNAME"
		
	End with

	pnlTab1.style.visibility = "visible" 
	'�ϴ���ȸ
	SelectRtn
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	gClearAllObject frmThis
	
	'���ο� XML ���ε��� ����
	frmThis.sprSht.MaxRows = 0
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸMASTER
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strCONTRACTNO
	
	With frmThis
		strCONTRACTNO	= .txtCONTRACTNO.value
		
		IF not SelectRtn_HDR (strCONTRACTNO) Then Exit Sub
		
		'��Ʈ ��ȸ
		If not SelectRtn_DTL(strCONTRACTNO) Then
			gErrorMsgBox "����ȸ���� ��ȸ����","��ȸ�ȳ�!"
			InitPageData
		End If

	End With
	gWriteText lblStatus, "�����Ͻ� ���ݰ�꼼�� �� ���Ͽ� �ڷᰡ �˻�" & mePROC_DONE
End Sub
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸHEADER
'-----------------------------------------------------------------------------------------
Function SelectRtn_HDR(ByVal strCONTRACTNO)
	Dim vntData
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_HDR = false
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMCONTRACT.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCONTRACTNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ ��༭ ��ȣ �� ���Ͽ�" & meNO_DATA, ""
			InitPageData
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)

			SelectRtn_HDR = True 
		End IF
	End IF
End Function
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸDETAIL
'-----------------------------------------------------------------------------------------
Function SelectRtn_DTL (ByVal strCONTRACTNO)
	Dim vntData
	'on error resume next
	SelectRtn_DTL = false
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMCONTRACT.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCONTRACTNO)
	
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		mobjSCGLSpr.SetClipbinding frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_DTL = True
	End IF
End Function





		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 793px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 427px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="98" background="../../../images/back_p.gIF"
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
												<td class="TITLE">��༭&nbsp;�󼼳���</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 375px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 53px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="203" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
								<!--Top Define Table End-->
								<!--Input Define Table End--></TABLE>
							<TABLE id="tblBody" style="WIDTH: 792px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="����">
											<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 80px" width="80">��༭��ȣ</TD>
													<TD class="SEARCHDATA"><INPUT dataFld="CONTRACTNO" class="NOINPUTB" id="txtCONTRACTNO" title="��༭��ȣ" style="WIDTH: 120px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" readOnly maxLength="30" size="14" name="txtCONTRACTNO"></TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1"
											cellPadding="0" align="right" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="90"><FONT face="����">�� �� ó</FONT></TD>
												<TD class="SEARCHDATA" width="173"></FONT><INPUT dataFld="OUTSNAME" class="NOINPUTB_L" id="txtOUTSNAME" title="����ó" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtOUTSNAME">
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="����"> �� �� ��</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="����"><INPUT dataFld="CONTRACTDAY" class="NOINPUTB" id="txtCONTRACTDAY" title="�����" style="WIDTH: 172px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtCONTRACTDAY"></FONT>
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="����">�� ��&nbsp;�� ��</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="����"><INPUT dataFld="AMT" class="NOINPUTB_R" id="txtAMT" title="���ݾ�" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtAMT"></FONT></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">�� �� ��</TD>
												<TD class="SEARCHDATA" colspan="5"><INPUT dataFld="CONTRACTNAME" class="NOINPUTB" id="txtCONTRACTNAME" title="����" style="WIDTH: 435px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="300" size="66" name="txtCONTRACTNAME"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 791px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 302px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 336px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20929">
									<PARAM NAME="_ExtentY" VALUE="8890">
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
									<PARAM NAME="MaxCols" VALUE="19">
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
							</DIV>
						</TD>
					</TR>
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 794px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 794px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
