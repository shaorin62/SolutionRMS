<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCUSTSEQPOP.aspx.vb" Inherits="PD.PDCMCUSTSEQPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�귣�� ��ȸ</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/����/�����ڵ� �˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPOP1.aspx
'��      �� : JOBNO ��ȸ�� ���� �˾�
'�Ķ�  ���� : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , ��ȸ�߰��ʵ�, ���� ������� �͸� ��ȸ���� ����,
'			  �ڵ� ������, �ڵ�Like���� ����
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
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

Dim mobjPDCMGET
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrPD_JOB_REG '�������...1


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

Sub txtCodeName_onkeydown
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
	'msgbox Col & Row
	window.returnvalue = mobjSCGLSpr.GetClip (.sprSht,1,.sprSht.ActiveRow,.sprSht.MaxCols,1,1)
	call Window_OnUnload()
	end if
	End With
end sub

'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'����������ü ����	
	set mobjPDCMGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		mstrPD_JOB_REG = ""
		
		for i = 0 to intNo
			select case i
				case 0 : .txtCustCode.value = vntInParam(i)	
				case 1 : .txtCustCodeName.value = vntInParam(i)
				case 2 : .txtCode.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 3 : .txtCodeName.value = vntInParam(i)		'���� ������� �͸�
				case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
		'txtCustCodeName,txtCustCode
		'SpreadSheet ������
		gSetSheetDefaultColor	'Sheet �⺻Color ����
		gSetSheetColor mobjSCGLSpr, .sprSht	'Sheet Į�� ����
		mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht, "CODE| SEQNO | SEQNAME| CUSTCODE | CUSTNAME |DEPT_CD| DEPTNAME | CLIENTSUBCODE | CLIENTSUBNAME "
		'mobjSCGLSpr.SpreadDataField .sprSht, "SEQNO | SEQNAME| DEPT_CD | DEPTNAME |ATTR02  "
		'mobjSCGLSpr.SetHeader .sprSht, "�ڵ�|�ڵ��|�μ��ڵ�|�μ���|�귣���"
		mobjSCGLSpr.SetHeader .sprSht, "�ڵ�|�귣���ڵ�|�귣���|�������ڵ�|�����ָ�|�μ��ڵ�|�μ���|������ڵ�|�����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "9|0|13|0|12|0|8|0|0"
		mobjSCGLSpr.SetRowHeight .sprSht,"-1","13"
		mobjSCGLSpr.SetRowHeight .sprSht,"0","15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"SEQNO"	'�߾� ����
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"SEQNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CUSTCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CUSTNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"DEPT_CD"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"DEPTNAME"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CLIENTSUBCODE"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht,"CLIENTSUBNAME"
		mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
		mobjSCGLSpr.ColHidden .sprSht, "SEQNO|CUSTCODE|DEPT_CD| CLIENTSUBCODE | CLIENTSUBNAME ", true
		
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CODE",-1,-1,2,2,false
		
	end with	
	'�ڷ���ȸ	
	SelectRtn
end sub

Sub EndPage()
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCode.value,.txtCodeName.value,.txtCustCode.value,.txtCustCodeName.value)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
end sub

-->
		</script>
	</HEAD>
	<body class="base" style="BACKGROUND-IMAGE: url(../../../images/imgBodyBg.gif)" bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" HEIGHT="100%" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/PopupIcon.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												�귣�� �˻�
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
											<TD><FONT face="����"></FONT></TD>
											<TD width="3"><FONT face="����"></FONT></TD>
											<TD><FONT face="����"></FONT></TD>
											<TD width="3"><FONT face="����"></FONT></TD>
											<TD style="WIDTH: 1px"><FONT face="����"></FONT></TD>
											<TD width="3"><FONT face="����"></FONT></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gif'"
													height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gif" width="54" border="0"
													name="imgQuery"></TD>
											<TD width="3"><FONT face="����"></FONT></TD>
											<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgConfirmOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirm.gif'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgConfirm.gif" width="54" border="0"
													name="imgConfirm"></TD>
											<TD width="3"><FONT face="����"></FONT></TD>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCancelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCancel.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgCancel.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="20" style="WIDTH: 20px"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" HEIGHT="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 382px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" height="21" style="WIDTH: 382px" align="RIGHT"><FONT face="����">
										<TABLE class="KEY" id="tblKey" cellSpacing="0" cellPadding="0" width="392" align="right"
											border="0">
											<TBODY>
												<TR>
													<TD class="KEY" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCustCode,txtCustCodeName)">
														�������ڵ�&nbsp;</TD>
													<TD class="KEY"><INPUT class="INPUT" id="txtCustCode" type="text" size="9" name="txtCustCode" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="KEY" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCustCodeName,txtCustCode)">�����ָ�</TD>
													<TD class="KEY"><INPUT class="INPUT" id="txtCustCodeName" style="WIDTH: 140px; HEIGHT: 22px" type="text"
															size="18" name="txtCustCodeName" tabIndex="1"></TD>
												</TR>
												<TR>
													<TD class="KEY" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCode,txtCodeName)">
														�귣���ڵ�&nbsp;</TD>
													<TD class="KEY"><INPUT class="INPUT" id="txtCode" type="text" size="9" name="txtCode" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
													<TD class="KEY" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCodeName,txtCode)">�귣���</TD>
													<TD class="KEY"><INPUT class="INPUT" id="txtCodeName" style="WIDTH: 140px; HEIGHT: 22px" type="text" size="18"
															name="txtCodeName" tabIndex="1"></TD>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 382px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD align="center" style="WIDTH: 100%"><FONT face="����">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31882">
											<PARAM NAME="_ExtentY" VALUE="16351">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 382px"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
