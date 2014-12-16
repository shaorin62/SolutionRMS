<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMTRANSTIMPOP.aspx.vb" Inherits="MD.MDCMTRANSTIMPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�� ��ȸ</title> 
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMTRANSTIMPOP.aspx
'��      �� : �ŷ���������ȸ
'�Ķ�  ���� : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , ��ȸ�߰��ʵ�, ���� ������� �͸� ��ȸ���� ����,
'			  �ڵ� ������, �ڵ�Like���� ����
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/05 By Kim Tae Yub
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mobjMDCMGET 
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mtranscommiflag, mtransTblflag
Dim msponsor
'-----------------------------
' �̺�Ʈ ���ν��� 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

sub imgQuery_onclick ()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
end sub

Sub txtCUSTCODE_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick ()
	if frmThis.sprSht.ActiveRow > 0 then
		sprSht_DblClick frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	else
		call Window_OnUnload()
	end if
end sub

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

sub sprSht_DblClick (Col,Row)
	'���õ� �ο� ��ȯ
	With frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			window.returnvalue = mobjSCGLSpr.GetClip (.sprSht,1,.sprSht.ActiveRow,.sprSht.MaxCols,1,1)
			call Window_OnUnload()
		end if
	End With
end sub

Sub sprSht_Keydown(KeyCode, Shift)
    if KeyCode <> meCR then exit sub
	'��Ʈ���� ���ͽ� Ȯ�� ó��
	Call sprSht_DblClick (frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow)		
End Sub
'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'����������ü ����	
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
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
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : .txtCUSTCODE.value = vntInParam(i)
				case 2 : .txtCUSTNAME.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 3 : .txtTIMCODE.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 4 : .txtTIMNAME.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 5 : mtranscommiflag = vntInParam(i)		'���� ������� �͸�
				case 6 : mtransTblflag = vntInParam(i)		'�ڵ� ��� ����
			end select
		next
		'SpreadSheet ������
		gSetSheetDefaultColor()

        SheetColChange
	end with	
	'�ڷ���ȸ	
	SelectRtn
end sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   	Dim strCOMMITCHECK

	'On error resume next
	with frmThis
		IF .txtYEARMON.value ="" THEN
			gErrorMsgBox "����� �ݵ�� �Է��ϼ���.",""
			EXIT SUB
		END IF
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		IF .chkGUBUN.checked = TRUE THEN
			strCOMMITCHECK = "COMMIT"
		ELSE
			strCOMMITCHECK = ""
		END IF
		
		SheetColChange

		vntData = mobjMDCMGET.GetTRANSTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtCUSTCODE.value,.txtCUSTNAME.value, .txtTIMCODE.value, .txtTIMNAME.value, strCOMMITCHECK, mtranscommiflag, mtransTblflag)

		if not gDoErrorRtn ("GetTRANSTIMCODE") then
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)

   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			
   			if mlngRowCnt <> 0 then
   				.sprSht.focus()
   			else
   				.sprSht.MaxRows = 0
   				.txtTIMNAME.focus()
   			end if 
   		end if
   	end with
end sub

Sub SheetColChange
	with frmThis
		if mtranscommiflag = "trans" then
			IF .chkGUBUN.checked = TRUE THEN
				gSetSheetColor mobjSCGLSpr, .sprSht
				mobjSCGLSpr.SpreadLayout .sprSht, 6, 0, 0, 0,2
				mobjSCGLSpr.SpreadDataField .sprSht, "TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE"
				mobjSCGLSpr.SetHeader .sprSht,        "�����|����|�����ָ�|��������|�������ڵ�|���ڵ�"
				mobjSCGLSpr.SetColWidth .sprSht, "-1", "      7|  15|      15|         8|         0|    0"
				mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
				mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
				
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE", -1, -1, 200
				mobjSCGLSpr.SetCellsLock2 .sprSht, true, "TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE"
				mobjSCGLSpr.SetCellAlign2 .sprSht, "TRANSYEARMON|GBN",-1,-1,2,2,false
				mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE | TIMCODE", true
				mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			else
				gSetSheetColor mobjSCGLSpr, .sprSht
				mobjSCGLSpr.SpreadLayout .sprSht, 6, 0, 0, 0,2
				mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE"
				mobjSCGLSpr.SetHeader .sprSht,		"û�����|����|�����ָ�|��������|�������ڵ�|���ڵ�"
				mobjSCGLSpr.SetColWidth .sprSht, "-1", "    7|  15|      15|         8|         0|    0"
				mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
				mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "YEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE", -1, -1, 200
				mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE"
				mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON|GBN",-1,-1,2,2,false
				mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE | TIMCODE", true
				mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			end if
		end if
	end with
End Sub

-->
		</script>
	</HEAD>
	<body class="base"  bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" valign=bottom>
												��&nbsp;��ȸ
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 168px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="168" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgConfirmOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirm.gif'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgConfirm.gif" width="54" border="0"
													name="imgConfirm"></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</tr>
						</table>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="����">
										<TABLE class="SEARCHDATA" id="tblKey" style="WIDTH: 392px" cellSpacing="0" cellPadding="0" width="392"
											align="right" border="0">
											<TBODY>
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">
														���</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" type="text" name="txtYEARMON" style=" WIDTH: 66px; HEIGHT: 22px"
															size="5">&nbsp;</TD>
													<TD class="SEARCHLABEL">
														�ŷ������Ϸᱸ��</TD>
													<TD class="SEARCHDATA" align="left"><INPUT style="WIDTH: 24px; HEIGHT: 20px" type="checkbox" id="chkGUBUN" name="chkGUBUN">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</TD>
												</TR>
												<TR>
													<TD class="SEARCHLABEL" style=" WIDTH: 68px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTCODE,txtCUSTNAME)">
														�ڵ�</TD>
													<TD class="SEARCHDATA" style="WIDTH: 52px"><INPUT class="INPUT" id="txtCUSTCODE" type="text" name="txtCUSTCODE" style=" WIDTH: 66px; HEIGHT: 22px"
															size="5">&nbsp;</TD>
													<TD class="SEARCHLABEL" style=" WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTNAME,txtCUSTCODE)">
														�����ָ�</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtCUSTNAME" style=" WIDTH: 160px; HEIGHT: 22px" type="text" name="txtCUSTNAME"
															tabIndex="1" size="21"></TD>
												</TR>
												<TR>
													<TD class="SEARCHLABEL" style=" WIDTH: 68px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMCODE,txtTIMNAME)">
														�ڵ�</TD>
													<TD class="SEARCHDATA" style="WIDTH: 52px"><INPUT class="INPUT" id="txtTIMCODE" type="text" name="txtTIMCODE" style=" WIDTH: 66px; HEIGHT: 22px"
															size="5">&nbsp;</TD>
													<TD class="SEARCHLABEL" style=" WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMCODE,txtTIMNAME)">
														����</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtTIMNAME" style=" WIDTH: 160px; HEIGHT: 22px" type="text" name="txtTIMNAME"
															tabIndex="1" size="21"></TD>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="����">
										<OBJECT id="sprSht" style="WIDTH: 392px; HEIGHT: 274px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="10372">
											<PARAM NAME="_ExtentY" VALUE="7250">
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
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
