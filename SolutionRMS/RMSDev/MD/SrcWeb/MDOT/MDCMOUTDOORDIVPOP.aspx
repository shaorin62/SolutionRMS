<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORDIVPOP.aspx.vb" Inherits="MD.MDCMOUTDOORDIVPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� û�� ������ ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : ������ �����˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMEXCUTIONPOP.aspx
'��      �� : JOBNO ��ȸ�� ���� �˾�
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 20120326 by OH SE HOON
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

Dim mobjMDOTOUTDOOR
Dim mlngRowCnt, mlngColCnt
'-----------------------------
' �̺�Ʈ ���ν��� 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgClose_onclick()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub

sub imgDelRow_onclick ()
	With frmThis
		DeleteRtn
	End With 
end sub

Sub sprSht_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW  and KeyCode <> meCR then exit sub  
	
	With frmThis
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",	.sprSht.ActiveRow, .txtYEARMON.value
		mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",		.sprSht.ActiveRow, "0"
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"TITLE",		.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"TITLE",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"TOTALAMT",	.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"TOTALAMT",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMI_RATE",.sprSht.ActiveRow-1)
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",		.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"OUT_AMT",	.sprSht.ActiveRow, 0
	
	End With
End Sub

'-----------------------------
' Spread Sheet Event
'-----------------------------	
Sub sprSht_change(ByVal Col,ByVal Row)
	Dim intAMT
	Dim intOUT_AMT
	Dim intCOMMISSION
	Dim intCOMMI_RATE
	
	with frmThis
		intAMT = 0
		intOUT_AMT = 0
		
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUT_AMT") OR Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") then
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht, "AMT",Row)
			intOUT_AMT = mobjSCGLSpr.GetTextBinding(.sprSht, "OUT_AMT",Row)
			
			intCOMMISSION = intAMT - intOUT_AMT
			
			IF intAMT = 0 THEN
				intCOMMI_RATE = 0
			ELSE 
				intCOMMI_RATE = intCOMMISSION / intAMT * 100
			END IF
			
			mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION", Row, intCOMMISSION
			mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE", Row, intCOMMI_RATE
		end if
	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
End Sub	

'��Ʈ ����Ŭ�� 
sub sprSht_DBLClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		ELSE
		
		end if
	end with
end sub

'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	Dim intNo,i,vntInParam
	Dim strATTR01
	
	set mobjMDOTOUTDOOR	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR")
	
	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		
		.txtATTR01.value = vntInParam(0)
		
		strATTR01 = split(vntInParam(0),"-")
		
		.txtYEARMON.value = strATTR01(0)
		.txtSEQ.value = strATTR01(1)
		
		gSetSheetDefaultColor()
			
        gSetSheetColor mobjSCGLSpr, .sprSht 
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | CLIENTNAME | TITLE | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MEMO | COMMI_TRANS_NO | TRU_VOCH_NO | ATTR01"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|��ȣ|������|����|�Ѱ��ݾ�|��û����|�����޾�|������|������|���|�ŷ���ǥ|������ǥ|ATTR01"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   0|    18|    20|        11|      11|      10|     4|    10|  20|         0|       0|     0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME | TITLE | MEMO | COMMI_TRANS_NO | TRU_VOCH_NO"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOTALAMT | AMT | OUT_AMT | COMMISSION | ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, " YEARMON | SEQ | CLIENTNAME | COMMI_TRANS_NO | TRU_VOCH_NO "
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | COMMI_TRANS_NO | TRU_VOCH_NO", true
	
		.sprSht.focus
	End With
    
	SelectRtn
end sub

Sub EndPage()
	set mobjMDOTOUTDOOR = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		vntData = mobjMDOTOUTDOOR.SelectRtn_OUTDOORDIV(gstrConfigXml,mlngRowCnt,mlngColCnt, .txtATTR01.value)

		if not gDoErrorRtn ("SelectRtn_OUTDOORDIV") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
   		end if
   	end with
end sub

Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	
	with frmThis
   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		
		'����� ��ο� ������ ����
   		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht, "AMT",intCnt) = "0" AND mobjSCGLSpr.GetTextBinding(.sprSht, "OUT_AMT",intCnt) = "0" then
			mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End If
		Next
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTNAME | TITLE | TOTALAMT | AMT | OUT_AMT | COMMI_RATE | COMMISSION | MEMO")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		intRtn = mobjMDOTOUTDOOR.ProcessRtn_DIV(gstrConfigXml,vntData, .txtATTR01.value)
	
		if not gDoErrorRtn ("ProcessRtn_DIV") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox intRtn & "���� �ڷᰡ ����" & mePROC_DONE , "����ȳ�!"
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  		for intCnt = 1 to .sprSht.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) = 0  AND mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",intCnt) = 0 Then 
					gErrorMsgBox intCnt & " ��° ���� �Է³��� �� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
		next
   	End with
	DataValidation = true
End Function

'---------------------
'----������ ����------
'---------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim lngchkCnt
		
	lngchkCnt = 0
	strSEQFLAG = False
	With frmThis
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_VOCH_NO",i) <> "" Then
					gErrorMsgBox "�����Ͻ� " & i & "���� �ڷ�� �ŷ���ǥ/������ǥ�� ���� �մϴ�.  " & vbcrlf & "  ���� �ŷ���ǥ/������ǥ�� ���� �Ͻʽÿ�!","�����ȳ�!"
					exit Sub
				elseIF mobjSCGLSpr.GetTextBinding(.sprSht,"ATTR01",i) = "" THEN  
					gErrorMsgBox "�����Ͻ� " & i & "���� �ڷ�� ���ҽ� ���� �����ߴ� ���� ������ �Դϴ�.  " & vbcrlf & "  ���� �����ʹ� ������ �� �����ϴ�.!","�����ȳ�!"
					exit Sub
				ELSE
					lngchkCnt = lngchkCnt +1
				End If
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjMDOTOUTDOOR.DeleteRtn(gstrConfigXml,strYEARMON,dblSEQ)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
   		End If

		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn
			If .sprSht.MaxRows = 0 Then
				Window_OnUnload
			End If 
		End If
	End With
	err.clear	
End Sub

-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="573" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td style="WIDTH: 300px" align="left" width="300" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" vAlign="bottom">���� û�� ������ ����
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 225px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
												height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0" name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="HEIGHT: 340px" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="����">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="0" cellPadding="0" width="100%" align="right"
										border="0">
										<TBODY>
											<TR>
												<TD class="SEARCHLABEL" width="60">û���ȣ
												</TD>
												<td class="SEARCHDATA"><INPUT class="NOINPUT" id="txtYEARMON" style="WIDTH: 80px; HEIGHT: 22px" readOnly size="8"
														name="txtYEARMON"><INPUT class="NOINPUT" id="txtSEQ" style="WIDTH: 56px; HEIGHT: 22px" readOnly size="4"
														name="txtSEQ">
												</td>
											</TR>
										</TBODY>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD align = "right"><INPUT id="txtATTR01" tabIndex="1" type="hidden" name="txtATTR01"><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
									style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0"
									name="imgAddRow"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
									onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
									width="54" border="0" name="imgSave"><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
									style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'" alt="�� �� ����" src="../../../images/imgDelRow.gif"
									width="54" border="0" name="imgDelRow">
								</TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="����">
										<OBJECT style="WIDTH: 574px; HEIGHT: 251px" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15187">
											<PARAM NAME="_ExtentY" VALUE="6641">
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
					</FORM>
				</TD>
			</TR>
		</TABLE>
	</body>
</HTML>
