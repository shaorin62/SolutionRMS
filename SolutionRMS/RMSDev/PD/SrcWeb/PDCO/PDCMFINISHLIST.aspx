<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMFINISHLIST.aspx.vb" Inherits="PD.PDCMFINISHLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��������</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���� ��� ȭ��
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
'�������� ����
Dim mobjPDCMFINISHLIST
Dim mobjMDCMMEDGet
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag

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
Dim vntData

with frmThis

		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMFINISHLIST.SelectRtn_PreCount(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt = 0 Then
			'mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
   			Else
   			'mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			gFlowWait meWAIT_ON
			SelectRtn_Search
			gFlowWait meWAIT_OFF
   			end If
   			
   		end if
End with
	
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub ImgExeConfirm_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub ImgExeConfirmCancel_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn_Cancel
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'����������ü ����	
	Set mobjPDCMFINISHLIST = gCreateRemoteObject("cPDCO.ccPDCOFINISHLIST")
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
	
    Call Grid_Layout()
    frmThis.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht,    "JOBNO|JOBNAME|SEQ|ENDDAY|ADJDAY|PURCHASENO|STD|ADJAMT|DIVAMT|DEMANDAMT|BALANCE|DIVFLAG|DIVFLAGNAME|OUTSCODE|OUTSNAME|RANKTRANS"
		mobjSCGLSpr.SetHeader .sprSht,		    "JOBNO|JOB��|����|�����|������|��ȣ|����|����ݾ�|û�������ݾ�|û���ݾ�|��û���ݾ�|�����ڵ�|��������|����ó�ڵ�|����ó��|������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "10   |20   |0   |10    |10    |10  |25  |10      |11          |11      |11        |0       |8       |0         |25      |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT|DIVAMT|DEMANDAMT|BALANCE", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|SEQ|ENDDAY|PURCHASENO|DIVFLAGNAME",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "STD|OUTSNAME|JOBNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"JOBNO|JOBNAME|SEQ|ENDDAY|ADJDAY|PURCHASENO|STD|ADJAMT|DIVAMT|DEMANDAMT|BALANCE|DIVFLAG|DIVFLAGNAME|OUTSCODE|OUTSNAME"
		mobjSCGLSpr.ColHidden .sprSht, "DIVFLAG|OUTSCODE|SEQ|RANKTRANS", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"JOBNO|JOBNAME|DIVAMT|DEMANDAMT|BALANCE"
	End with
	'DateClean
	pnlTab1.style.visibility = "visible" 
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'�˻����� ������
Sub imgFrom_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtFrom,.imgFrom,"txtFrom_onchange()"
		gSetChange
	end with
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub

'�˻����� ������
Sub imgTo_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtTo,.imgTo,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtYEARMON_onchange
	gSetChange
End Sub
Sub txtTo_onchange
	gSetChange
End Sub


Sub SelectRtn ()

   	Dim vntData
   	Dim i, strCols
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMFINISHLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
			Next	
			.ImgExeConfirm.disabled = false
			.ImgExeConfirmCanCel.disabled = true
			.ImgExeConfirm.src = "../../../images/ImgExeConfirmOn.gIF"
			.ImgExeConfirmCanCel.src = "../../../images/ImgExeConfirmCanCel.gif"
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			Else
   			initpageData
   			gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   			
   		end if
   	end with
End Sub
Sub SelectRtn_Search ()

   	Dim vntData
   	Dim i, strCols
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCMFINISHLIST.SelectRtn_End(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value )
		
		if not gDoErrorRtn ("SelectRtn_End") then
			if mlngRowCnt > 0 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
			Next	
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			.ImgExeConfirm.disabled = true
			.ImgExeConfirmCanCel.disabled = false
			.ImgExeConfirm.src = "../../../images/ImgExeConfirm.gIF"
			.ImgExeConfirmCanCel.src = "../../../images/ImgExeConfirmCanCelOn.gif"
   			Else
   			initpageData
   			gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   			
   		end if
   	end with
End Sub

Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFrom.value = date1
		.txtTo.value = date2
	End With
End Sub
Sub EndPage()
	set mobjPDCMFINISHLIST = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
		.ImgExeConfirm.disabled = true
		.ImgExeConfirmCanCel.disabled = true
	End with
End Sub
'-----------------------------------------------------------------------------------------
' Ȯ�� Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn ()
  	Dim intRtn
  
	Dim intConRtn
	with frmThis
	'On error resume next
		intConRtn = gYesNoMsgbox("[" & .txtYEARMON.value & "] �� ����ó�� �� �Ͻðڽ��ϱ�?","����ó��Ȯ��")
		IF intConRtn <> vbYes then exit Sub
		intRtn = mobjPDCMFINISHLIST.ProcessRtn(gstrConfigXml,.txtYEARMON.value)
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " ���긶�� ó����" & mePROC_DONE,"����ó���ȳ�" 
			SelectRtn_Search
  		end if
 	end with
End Sub
'-----------------------------------------------------------------------------------------
' Ȯ�� Proc
'-----------------------------------------------------------------------------------------
Sub ProcessRtn_Cancel ()
  	Dim intRtn
  
	Dim intConRtn
	with frmThis
	'On error resume next
		intConRtn = gYesNoMsgbox("[" & .txtYEARMON.value & "] �� ������� ó���� �Ͻðڽ��ϱ�?","�������ó��Ȯ��")
		IF intConRtn <> vbYes then exit Sub
		intRtn = mobjPDCMFINISHLIST.ProcessRtn_Cancel(gstrConfigXml,.txtYEARMON.value)
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG	
			gErrorMsgBox " ���긶�� ���ó����" & mePROC_DONE,"����ó���ȳ�" 
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
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%"  height="100%"border="0">
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
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;���� ����</td>
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
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="50" border="0">
										<TR>
											<TD></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0"
							 border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 1040px" vAlign="top" align="center" colSpan="2">
										<TABLE id="tblKey" cellSpacing="1" cellPadding="0" width="1040" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
													width="90">&nbsp;������
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 886px"><INPUT class="INPUT" id="txtYEARMON" title="������" style="WIDTH: 80px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="8" name="txtYEARMON"></TD>
												<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 30px"></TD>
								</TR>
								<!--�߰�-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="����"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">
																&nbsp;����</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50"
														border="0">
														<TR>
															<TD><IMG id="ImgExeConfirm" style="CURSOR: hand" height="20" alt="����ó���� �մϴ�." src="../../../images/ImgExeConfirm.gIF"
																	border="0" name="ImgExeConfirm"></TD>
															<TD><IMG id="ImgExeConfirmCanCel" style="CURSOR: hand" height="20" alt="������ ����մϴ�." src="../../../images/ImgExeConfirmCanCel.gIF"
																	border="0" name="ImgExeConfirmCanCel"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																	name="imgExcel"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
										<!--�׽�Ʈ ��--></TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 2px"></TD>
								<!--���� �� �׸���-->
								<TR vAlign="top" align="left">
									<!--����-->
									<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 95%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 95%"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="17463">
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
				</TR>
				<!--BodySplit End-->
				<!--List Start-->
				</TD></TR> 
				<!--List End-->
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
				</TR>
				<!--BodySplit End-->
				<!--Brench Start-->
				<TR>
					<TD class="BRANCHFRAME" style="WIDTH: 1040px"><FONT face="����" color="#666666" size="3"></FONT>
						<!--<INPUT class="BUTTON" id="btn1" style="WIDTH: 123px; HEIGHT: 16pt" type="button" value="�б��ư"
											name="Button">--></TD>
				</TR>
				<!--Brench End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" style="WIDTH: 1040px" id="lblstatus"><FONT face="����"></FONT></TD>
				</TR>
				<!--Bottom Split End-->
			</TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
