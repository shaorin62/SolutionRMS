<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMATTER.aspx.vb" Inherits="MD.MDCMMATTER" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� �ϰ�����</title>
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
'�������� ����
Dim mobjMDCMPREMATTER
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
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
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
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
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
	Set mobjMDCMPREMATTER = gCreateRemoteObject("cMDET.ccMDETPREMATTER")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	Dim strComboList 
	strComboList =  "���" & vbTab & "�̻��"
	gSetSheetDefaultColor
    with frmThis
		
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		'CC_CODE,CC_NAME,OC_CODE,OC_NAME,USE_YN,STDATE,EDATE
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0, 0
		'mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,    "CLIENTNAME|KOBACOCODE|MATTER|MATTERDIV|CMLAN|SISCUNO|DISCUDATE|MATTERVIEW|SUBSEQNAME|WASTECODE|LCODE|LNAME|MCODE|MNAME|SNAME|TELECASTLIMIT|TELECASTTIME|CUSER|CDATE|CLIENTCODE|ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht,		    "�����ָ�|�������ڵ�|����|���籸��|�ʼ�|���ǹ�ȣ|��������|���纸��|ǰ��|�ڵ�|������з��ڵ�|������з���|�����ߺз��ڵ�|�����ߺз���|�����Һз���|�濵����|�ð�|�Է���|�Է�����|CLIENTCODE|��������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "10      |10        |10  |10      |10  |10      |10      |10      |10  |10  |12            |10          |12            |10          |10          |10      |10  |10    |10      |10         |10"         
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		'mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "SDATE|EDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME|KOBACOCODE|MATTER|MATTERDIV|CMLAN|SISCUNO|DISCUDATE|MATTERVIEW|SUBSEQNAME|WASTECODE|LCODE|LNAME|MCODE|MNAME|SNAME|TELECASTLIMIT|TELECASTTIME|CUSER|CDATE|CLIENTCODE|ERRMSG", -1, -1,200
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,2,2,false '�߾�����
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,0,2,false '��������
		'mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CC_CODE|CC_NAME|OC_CODE|OC_NAME"
		'mobjSCGLSpr.SetCellTypeComboBox .sprSht,6,6,,,strComboList
		'mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE", true
	End with

	pnlTab1.style.visibility = "visible" 
End Sub

Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .chkAll.checked = True Then
		strCHK = ""
		Else
		strCHK = "All"
		End if
		
		vntData = mobjMDCMPREMATTER.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTCODE.value,.txtDEPTNAME.value,.cmbYN.value,strCHK)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtDEPTCODE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	
End Sub
'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
End Sub







Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'������� ��Ʈ ��ư Ŭ��

'Validation
Function DataValidation ()
	DataValidation = false	
	With frmThis
		
	End With
	DataValidation = True
End Function
'�������

Sub ProcessRtn()
	Dim intRtn
   	dim vntData
	Dim intCnt
	Dim intDelCnt
	Dim strCODE
	Dim intErrCnt
	Dim intRtnYN
		with frmThis
   		'������ Validation ����
		'if DataValidation =false then exit sub
		'DataErrorValidation
		'On error resume next
		'�ܿ�Row ���� ó��
		intRtnYN = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtnYN <> vbYes then exit Sub
		for intDelCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,1, intDelCnt) = "" Then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End IF
		Next
		
		'�������ڵ� ��������
		strCODE = ""
		for intCnt = 1 To .sprSht.MaxRows 
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"KOBACOCODE", intCnt)
		vntData = mobjMDCMPREMATTER.GetCLIENTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
			if mlngRowCnt > 0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",intCnt, vntData(0,1)
			Else
			mobjSCGLSpr.SetTextBinding .sprSht,"ERRMSG",intCnt, "�������ڵ忡��"
   			End if
		next
		
		for intErrCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"ERRMSG", intErrCnt) <> "" Then
				gErrorMsgbox "���������� Ȯ���Ͻʽÿ�.","����ȳ�!"
				exit Sub 
			end if
		Next
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"MATTER|SUBSEQNAME|CLIENTCODE")
		'ó�� ������ü ȣ��
		intRtn = mobjMDCMPREMATTER.ProcessRtn(gstrConfigXml,vntData)
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if intRtn > 1 Then
			gErrorMsgBox intRtn & " �� �� ����Ǿ����ϴ�." & vbcrlf & "�귣�� �ڵ�� ���� �ڵ尡 �ڵ������Ǿ�����," & vbcrlf & "�귣�� �ڵ� �� MC �μ��� �����Ͽ� �ֽð�," & vbcrlf & "�����ڵ� �� ������ �� ���� �Ͽ� �ֽʽÿ�.","����ȳ�"
			End If
			'SelectRtn
   		end if
   	end with
End Sub
Sub EndPage()
	set mobjMDCMPREMATTER = Nothing
	'set mobjMDCMMEDGet = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 2000
	End with
End Sub

sub DeleteRtn

	
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
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
											<td class="TITLE">&nbsp;���� �ϰ�����</td>
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
									<TABLE id="tblButton" style="WIDTH: 183px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="183" border="0">
										<TR>
											<TD></TD>
											<TD width="54"></TD>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD width="54"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
											<!--<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>--></TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 1040px; HEIGHT: 32px" cellSpacing="0" cellPadding="0"
							width="1040" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center"
									colSpan="2"><FONT face="����">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 14px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCODE,'')">&nbsp;</TD>
												<TD class="SEARCHDATA" style="WIDTH: 172px"><INPUT class="INPUTL" id="txtDEPTCODE" style="WIDTH: 96px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="8" size="10" name="txtDEPTCODE"></TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 2px"></TD>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD class="DATAFRAME">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 683px"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 683px"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27490">
											<PARAM NAME="_ExtentY" VALUE="18071">
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
				<!--List Start--></TABLE>
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
				<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
			</TR>
			<!--Bottom Split End--> </TABLE> 
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
