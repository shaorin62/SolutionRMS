<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMOUTSETTLE.aspx.vb" Inherits="PD.PDCMOUTSETTLE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����ó�� ������Ȳ</title>
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
Dim mobjPDCMSEARCH
Dim mobjPDCMGET
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
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF

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
	Set mobjPDCMSEARCH = gCreateRemoteObject("cPDCO.ccPDCOSEARCH")
	Set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'mobjPDCMGET
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
	
    Call Grid_Layout()
    'frmThis.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
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
		mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "GBN|PURCHASENO|OUTSCODE|OUTSNAME|JOBNO|JOBNAME|ITEMNAME|ADJAMT|ADJDAY|RANKTRANS|VOCHNO"
		mobjSCGLSpr.SetHeader .sprSht,		    "����|�����ȣ|����ó�ڵ�|����ó��|JOBNO|JOB��|�����׸�|����ݾ�|������|Į��|��ǥȮ��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "0   |12      |0         |23      |10   |25   |22      |12      |10    |0   |8"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|PURCHASENO|VOCHNO",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|OUTSNAME|ITEMNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PURCHASENO|OUTSCODE|OUTSNAME|ITEMNAME|JOBNO|JOBNAME|ADJAMT|ADJDAY|RANKTRANS|VOCHNO"
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE|RANKTRANS|GBN", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"PURCHASENO|OUTSNAME|JOBNO|JOBNAME"
	End with
	DateClean
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
    Dim strGUBN
    Dim strFROM 
    Dim strTO
    
	'On error resume next
	with frmThis
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .txtOUTSCODE.value = "" and .txtOUTSNAME.value = "" Then
		strGUBN = "DIV"
		Else
		strGUBN = "ALL"
		End If
		vntData = mobjPDCMSEARCH.SelectRtn_SETTLE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtPURCHASENO.value,.txtOUTSCODE.value,.txtOUTSNAME.value,strGUBN ,strFROM,strTO)
		
		if not gDoErrorRtn ("SelectRtn_SETTLE") then
			if mlngRowCnt > 1 Then
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows 

				If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
				Else
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "2" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
				End If 
			Next
			'msgbox "11111"
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			Else
   			initpageData
   			gWriteText lblStatus,  "0 ���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   			
   		end if
   	end with
End Sub




Sub EndPage()
	set mobjPDCMSEARCH = Nothing
	set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	Dim vntData
	with frmThis
		.sprSht.maxrows = 0
	End with
End Sub
'-----------------------------------------------------------------------------------------
' ����ó ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub imgOUTSCODE_onclick
	Call SEARCHOUT_POP()
End Sub

'���� ������List ��������
Sub SEARCHOUT_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtOUTSNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHOUT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
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
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%"  border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%"height="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;�������</td>
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
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE id="tblBody" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<tr>
					<TD class="BODYSPLIT" style="WIDTH: 1040px;HEIGHT: 10px"></TD>
				</tr>
				<TR>
					<TD style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center" colSpan="2"><FONT face="����">
							<TABLE id="tblKey" cellSpacing="1" cellPadding="0" width="1040" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()" width="90">&nbsp;����Ⱓ
									</TD>
									<TD class="SEARCHDATA" style="WIDTH: 220px"><INPUT class="INPUT" id="txtFROM" title="�������ڽ���" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="8" size="6" name="txtFROM"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="������������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="8" size="6" name="txtTO"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgTo"></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPURCHASENO, '')"
										width="90">&nbsp;�����ȣ
									</TD>
									<TD class="SEARCHDATA" style="WIDTH: 121px"><INPUT class="INPUT" id="txtPURCHASENO" title="�����ȣ" style="WIDTH: 120px; HEIGHT: 22px"
											accessKey="NUM" type="text" maxLength="10" size="14" name="txtPURCHASENO"></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)"
										width="90">����ó</TD>
									<TD class="SEARCHDATA" style="WIDTH: 398px"><INPUT class="INPUT_L" id="txtOUTSNAME" title="�ڵ��" style="HEIGHT: 22px" type="text" maxLength="255"
											align="left" size="38" name="txtOUTSNAME"><IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
											width="23" align="absMiddle" border="0" name="ImgOUTSCODE"><INPUT class="INPUT_L" id="txtOUTSCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="5" name="txtOUTSCODE"></TD>
									<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
											src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
								</TR>
							</TABLE>
						</FONT>
					</TD>
				</TR>
				<!--Input End-->
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 30px"></TD>
				</TR>
				<!--�߰�-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%">
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left"  height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;����ó�� ������Ȳ</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50"
										border="0">
										<TR>
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
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 5px"><FONT face="����"></FONT></TD>
				</TR>
				<!--�߰���-->
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
				<!--BodySplit End-->
				<!--List Start--> </TD></TR> 
				<!--List End-->
				<!--BodySplit Start--> </TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
				</TR>
				<!--BodySplit End-->
				<!--Brench Start-->
				<TR>
					<TD class="BRANCHFRAME" style="WIDTH: 1040px"><FONT face="����" color="#666666" size="3"></FONT>
					</TD>
				</TR>
				<!--Brench End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
				<!--Bottom Split End-->
			</TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
