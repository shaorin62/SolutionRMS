<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOSMSPOP.aspx.vb" Inherits="SC.SCCOSMSPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��� ��ȸ</title> 
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

Dim mobjSCCOGET 
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrFromUserName,mstrFromUserEmail,mstrFromUserPhone
Dim mstrCheck
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

Sub txtDEPTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

sub imgConfirm_onclick ()
	Dim intCnt2
	Dim intColSum
	Dim i
	Dim intRtn
	Dim vntData
	Dim dblID
	dblID = 0
	with frmThis
		intColSum = 0
  		If .sprSht2.MaxRows = 0 Then
  			gErrorMsgBox "�������� �۽ŵ����Ͱ� �����ϴ�.","ó���ȳ�"
  			Exit Sub
  		End If
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht2,"EMPNO|CELLPHONE")
	

		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� " & meNO_DATA,"ó���ȳ�"
			exit sub
		End If
		
		intRtn = mobjSCCOGET.ProcessRtn_SMS(gstrConfigXml,vntData,mstrFromUserName,mstrFromUserPhone,.txtMSG.value,dblID)	
		
		if not gDoErrorRtn ("ProcessRtn") then
			'msgbox "Popupâ�� �޾ƿ� ��ȣ" & dblID
			
			window.returnvalue = dblID
			call Window_OnUnload()	
		Else
			gErrorMsgBox "SMS �߼۸�� ������ ������ �߻� �Ͽ����ϴ�." & vbcrlf & "�����ڿ��� ���� �Ͻʽÿ�.","ó���ȳ�"				
			Exit Sub
		End If
	
	
	End With
	
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
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		'mstrFromUserName,mstrFromUserEmail,mstrFromUserPhone
		for i = 0 to intNo
			select case i
				case 0 : .txtDEPTCD.value = vntInParam(i)	
				case 1 : .txtDEPTNAME.value = vntInParam(i)
				case 2 : .txtEMPNO.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 3 : .txtEMPNAME.value = vntInParam(i)			'���� ������� �͸�
				case 4 : .txtMSG.value = vntInParam(i)				'�����޼���
				case 5 : mstrFromUserName = vntInParam(i)				'�����޼���
				case 6 : mstrFromUserEmail = vntInParam(i)				'�����޼���
				case 7 : mstrFromUserPhone = vntInParam(i)				'�����޼���
				
				
			case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
		'SpreadSheet ������
		gSetSheetDefaultColor()
        With frmThis
            gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 6, 0
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK|EMPNO | EMP_NAME |CC_CODE | CC_NAME|CELLPHONE"
			mobjSCGLSpr.SetHeader .sprSht, "����|����ڵ�|�����|�μ��ڵ�|�μ���|��ȭ��ȣ"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|8|10|0|27|10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "EMPNO"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "EMP_NAME"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "CC_CODE"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "CC_NAME"
			mobjSCGLSpr.ColHidden .sprSht, "CC_CODE|CELLPHONE",true
			mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			mobjSCGLSpr.SetCellAlign2 .sprSht, "EMPNO",-1,-1,2,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht, "EMP_NAME|CC_NAME",-1,-1,0,2,false
			
			
			gSetSheetColor mobjSCGLSpr, .sprSht2
			mobjSCGLSpr.SpreadLayout .sprSht2, 6, 0
			mobjSCGLSpr.SpreadDataField .sprSht2, "CHK|EMPNO | EMP_NAME |CC_CODE | CC_NAME|CELLPHONE"
			mobjSCGLSpr.SetHeader .sprSht2, "����|����ڵ�|�����|�μ��ڵ�|�μ���|��ȭ��ȣ"
			mobjSCGLSpr.SetColWidth .sprSht2, "-1", "4|8|10|0|27|10"
			mobjSCGLSpr.SetRowHeight .sprSht2, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht2, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht2, "CHK"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht2, "EMPNO"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht2, "EMP_NAME"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht2, "CC_CODE"
			mobjSCGLSpr.SetCellTypeStatic2 .sprSht2, "CC_NAME"
			mobjSCGLSpr.ColHidden .sprSht2, "CC_CODE|CELLPHONE",true
			mobjSCGLSpr.SetScrollBar .sprSht2,2,False,0,-1
			mobjSCGLSpr.SetCellAlign2 .sprSht2, "EMPNO",-1,-1,2,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht2, "EMP_NAME|CC_NAME",-1,-1,0,2,false
		
	
        End With
	end with	
	
	'�ڷ���ȸ	
	SelectRtn
	frmThis.sprSht2.MaxRows= 0
	'frmThis.sprSht.focus()
end sub


Sub EndPage()
	set mobjSCCOGET = Nothing
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

		vntData = mobjSCCOGET.GetSMS(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A",.txtDEPTCD.value,.txtDEPTNAME.value)

		if not gDoErrorRtn ("GetSMS") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			' mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			If mlngRowCnt > 0 Then
			
			Else
				.sprSht.MaxRows = 0 
			End If
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
end sub
Sub imgRight_onclick
	Dim intCnt
	Dim intLeftCnt
	Dim intCHK
	Dim intRtn
	Dim intCnt2
	Dim intColSum
	with frmThis 
		for intCnt2 = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "������ ���õ� �����Ͱ� �����ϴ�.","�̵��ȳ�"
			exit Sub
		End If
		
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				intCHK = 0
				For intLeftCnt = 1 To .sprSht2.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",intCnt) = mobjSCGLSpr.GetTextBinding(.sprSht2,"EMPNO",intLeftCnt) Then
						intCHK = 1
					end If
				Next
				If intCHK = 0 Then
					intRtn = mobjSCGLSpr.InsDelRow(.sprSht2, cint(meINS_ROW), 0, -1, 1)
					
					mobjSCGLSpr.SetTextBinding .sprSht2,"EMPNO",.sprSht2.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"EMPNO",intCnt) 
					mobjSCGLSpr.SetTextBinding .sprSht2,"EMP_NAME",.sprSht2.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"EMP_NAME",intCnt) 
					mobjSCGLSpr.SetTextBinding .sprSht2,"CC_CODE",.sprSht2.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"CC_CODE",intCnt) 
					mobjSCGLSpr.SetTextBinding .sprSht2,"CC_NAME",.sprSht2.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"CC_NAME",intCnt) 
					mobjSCGLSpr.SetTextBinding .sprSht2,"CELLPHONE",.sprSht2.ActiveRow, mobjSCGLSpr.getTextBinding(.sprSht,"CELLPHONE",intCnt) 
					
				End If
			end If
		Next
		
	End with
End Sub
Sub imgLeft_onclick
	Dim intCnt2
	Dim intColSum
	Dim i
	with frmThis
		intColSum = 0
  		for intCnt2 = 1 to .sprSht2.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht2,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "�������� ���õ� �����Ͱ� �����ϴ�.","�̵��ȳ�"
			exit Sub
		End If
	
		for i = .sprSht2.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht2,"CHK",i) = 1 then				
					mobjSCGLSpr.DeleteRow .sprSht2,i
			End If
   		next
   	End With
End Sub

Sub sprSht2_click(ByVal Col, ByVal Row)

Dim intcnt,intCnt2
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht2, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		End If
	End with
End Sub

-->
		</script>
	</HEAD>
	<body class="base"  bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<FORM id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="470" border="0">
				<TR>
					<TD>
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
											<td class="TITLE" id="objTitle">
												���&nbsp;��ȸ
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
											<td class="TITLE">Massage:&nbsp;</td>
											<TD><INPUT class="INPUT_L" id="txtMSG" style="WIDTH: 360px; HEIGHT: 22px" type="text" size="54"
													name="txtMSG"></TD>
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
											<TD width="15"><FONT face="����"></FONT></TD>
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
						<TABLE  id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD  style="HEIGHT: 20px" vAlign="middle" height="20">
									<TABLE  class="SEARCHDATA"  id="tblKey" style="WIDTH: 392px" cellSpacing="0" cellPadding="0" width="392"
										align="left" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">
												�μ� �ڵ�</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtDEPTCD" type="text" size="9" name="txtDEPTCD" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">
												�μ�&nbsp;��&nbsp;</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtDEPTNAME" style="WIDTH: 140px; HEIGHT: 22px" type="text" size="18"
													name="txtDEPTNAME" tabIndex="1"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
												��� �ڵ�</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtEMPNO" type="text" size="9" name="txtEMPNO" style="WIDTH: 90px; HEIGHT: 22px">&nbsp;</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEMPNO,txtEMPNAME)">
												���&nbsp;��&nbsp;</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtEMPNAME" style="WIDTH: 140px; HEIGHT: 22px" type="text" size="18"
													name="txtEMPNAME" tabIndex="1"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<td>
									<table border="0" cellpadding="0" cellspacing="0" width="100%" height="92%">
										<tr>
											<TD align="center" width="48%">
												<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="20373">
														<PARAM NAME="_ExtentY" VALUE="16272">
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
											</TD>
											<td align="center" height="95%" width="4%">
												<img src="../../../images/imgRight.gif" style="CURSOR: hand" id="imgRight" border="0"><BR>
												<BR>
												<img src="../../../images/imgLeft.gif" style="CURSOR: hand" id="imgLeft" border="0">
											</td>
											<TD align="center" width="48%">
												<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht2" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="20346">
														<PARAM NAME="_ExtentY" VALUE="16272">
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
											</TD>
										</tr>
									</table>
								</td>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
