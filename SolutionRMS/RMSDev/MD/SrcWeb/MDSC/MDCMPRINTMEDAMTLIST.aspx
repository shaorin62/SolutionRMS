<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTMEDAMTLIST.aspx.vb" Inherits="MD.MDCMPRINTMEDAMTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��ü�� ���೻��</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/�׷챤�� �д�� �Է�/��ȸ ȭ��(MDCMGROUP)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMGROUP.aspx.aspx
'��      �� : �׷챤�� �д�� �� ��ȸ/�Է� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/01/09 By Kim Tae Yub
'			 2) 
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
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET, mobjMDSRREPORTLIST'�����ڵ�, Ŭ����
Dim mstrClientcode

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
' ���� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	
	if frmThis.txtYEARMON.value = ""  then
		gErrorMsgBox "����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
'	if frmThis.txtCLIENTCODE.value = ""  then
'		gErrorMsgBox "�������ڵ带 �Է��Ͻÿ�","��ȸ�ȳ�"
'		exit Sub
'	end if
	
	gFlowWait meWAIT_ON
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

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i
	Dim strYEARMON
	Dim strCLIENTNAME
	Dim strCLIENTCODE
	dim chkflag
   	dim strLIST
   	Dim strClientLIST
   	Dim intSUBRow
	
	with frmThis
		
		strLIST = ""
		chkflag = 1
				gErrorMsgBox "��¹��� �������Դϴ�..",""
			Exit Sub
'		if frmThis.sprSht.MaxRows = 0 then
'			gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
'			Exit Sub
'		end if
		
'		strClientLIST = split(mstrClientcode, "|")
		
'		intSUBRow = UBound(strClientLIST, 1)
'		FOR i = 0 to intSUBRow
'			IF chkflag = 1 then
'				strLIST = "'" & strClientLIST(i) & "'"
'				chkflag = 2
'			else
'				strLIST = strLIST & ",'" & strClientLIST(i) & "'"
'			end if 
'		Next
		
'		ModuleDir = "MD"
'		ReportName = "MDCMCLIENTSUBSEQMEDLIST.rpt"
'		
'		strYEARMON		= .txtYEARMON.value
'		strCLIENTNAME	= .txtCLIENTNAME.value
'		
'		Params = strYEARMON & ":" & strLIST & ":" & strCLIENTNAME
'		
'		Opt = "A"
'		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

Sub imgCUSTPOP_onclick
	Call CLIENTCODECHK_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODECHK_POP
	Dim vntRet
	Dim vntInParams
	mstrClientcode = ""
	'InitPage
	With frmThis
		vntInParams = array("", "MED") '<< �޾ƿ��°�� '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTCHKPOP.aspx",vntInParams , 413,435)
		if vntRet <> "" then
			mstrClientcode = vntRet
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
		end if
	End With
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	mstrClientcode = ""
	
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			mstrClientcode = trim(vntRet(0,0))
     	end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		mstrClientcode = ""
		
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "B")
			
			if not gDoErrorRtn ("GetHIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					mstrClientcode = vntData(0,1)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CLIENTNAME | MEDFLAGNAME | MEDCODE | A1 | A2 |  A3 |  A4 |  A5 |  A6 |  A7 |  A8 |  A9 |  A10 |  A11 |  A12 | SUMAMT"
		mobjSCGLSpr.SetHeader .sprSht,        "������|����|��ü|1��|2��|3��|4��|5��|6��|7��|8��|9��|10��|11��|12��|���հ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   10|  12|  10| 10| 10| 10| 10| 10| 10| 10| 10| 10|  10|  10|  10|   12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "A1 | A2 |  A3 |  A4 |  A5 |  A6 |  A7 |  A8 |  A9 |  A10 |  A11 |  A12 | SUMAMT", -1, -1,0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CLIENTNAME | MEDFLAGNAME | MEDCODE | A1 | A2 |  A3 |  A4 |  A5 |  A6 |  A7 |  A8 |  A9 |  A10 |  A11 |  A12 | SUMAMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME | MEDCODE | MEDFLAGNAME",-1,-1,2,2,false
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME",-1,-1,0,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht, "CLIENTNAME | MEDFLAGNAME"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDSRREPORTLIST = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtYEARMON.value = Mid(gNowDate,1,4)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEARMON.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	Dim chkflag
   	Dim strLIST
   	Dim strClientLIST
   	Dim intSUBRow
   	Dim strFROMMON, strTOMON, strMONCNT
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		strLIST = ""
		chkflag = 1
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strClientLIST = split(mstrClientcode, "|")
		
		intSUBRow = UBound(strClientLIST, 1)
		FOR i = 0 to intSUBRow
			IF chkflag = 1 then
				strLIST = "'" & strClientLIST(i) & "'"
				chkflag = 2
			else
				strLIST = strLIST & ",'" & strClientLIST(i) & "'"
			end if 
		Next
		strFROMMON = .cmbFROMMON.value
		strTOMON = .cmbTOMON.value
		strMONCNT = (strTOMON - strFROMMON) +1
		vntData = mobjMDSRREPORTLIST.SelectRtn_PRINTMEDAMTLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, strLIST, strFROMMON, strTOMON, strMONCNT)

		if not gDoErrorRtn ("SelectRtn_CLIENTSUBPRINTLIST") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME",intCnt),2) = "�Ұ�" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		end if
		If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME",intCnt),2) = "�հ�" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		End If
	Next 
	End With
End Sub

		</script>
	</HEAD>
		<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF"
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
											<td class="TITLE">��ü�� ���೻��&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="110" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table Start-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="95%" border="0"> <!--TopSplit Start->
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="�⵵�������մϴ�." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">��&nbsp; 
												��
											</TD>
											<TD class="SEARCHDATA" width="424" style="WIDTH: 424px"><INPUT class="INPUT" id="txtYEARMON" title="�⵵���Է��ϼ���" style="WIDTH: 100px; HEIGHT: 22px"
													type="text" maxLength="4" size="14" name="txtYEARMON" accessKey="NUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
													<SELECT id="cmbFROMMON" title="���" style="WIDTH: 80px" name="cmbFROMMON">
														<OPTION value="1" selected>1��</OPTION>
														<OPTION value="2">2��</OPTION>
														<OPTION value="3">3��</OPTION>
														<OPTION value="4">4��</OPTION>
														<OPTION value="5">5��</OPTION>
														<OPTION value="6">6��</OPTION>
														<OPTION value="7">7��</OPTION>
														<OPTION value="8">8��</OPTION>
														<OPTION value="9">9��</OPTION>
														<OPTION value="10">10��</OPTION>
														<OPTION value="11">11��</OPTION>
														<OPTION value="12">12��</OPTION>
													</SELECT>&nbsp;~
													<SELECT id="cmbTOMON" title="���" style="WIDTH: 80px" name="cmbTOMON">
														<OPTION value="1">1��</OPTION>
														<OPTION value="2">2��</OPTION>
														<OPTION value="3">3��</OPTION>
														<OPTION value="4">4��</OPTION>
														<OPTION value="5">5��</OPTION>
														<OPTION value="6">6��</OPTION>
														<OPTION value="7">7��</OPTION>
														<OPTION value="8">8��</OPTION>
														<OPTION value="9">9��</OPTION>
														<OPTION value="10">10��</OPTION>
														<OPTION value="11">11��</OPTION>
														<OPTION value="12" selected>12��</OPTION>
													</SELECT>
											</TD>
											<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 207px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
											</TD>
											<TD class="SEARCHDATA" width="100"><IMG id="imgCUSTPOP" onmouseover="JavaScript:this.src='../../../images/imgMedMultiChkOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgMedMultiChk.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgMedMultiChk.gIF" border="0" name="imgCUSTPOP"></TD>
										</TR>
									</TABLE>									
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="����"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; height: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; height:100%; POSITION: relative; "
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 100%"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27490">
											<PARAM NAME="_ExtentY" VALUE="20320">
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
							</TR>
							<!--List End-->
							<!--Bottom Split Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
								</TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
			</TABLE>
			</FORM>
		</TR>
	</TABLE>
	</body>
</HTML>