<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMONMEDLIST.aspx.vb" Inherits="MD.MDCMMONMEDLIST" %>
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
		<SCRIPT id="clientEventHandlersVBS" language="vbscript">
		
'�������� ����
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMGET, mobjMDSRREPORTLIST'�����ڵ�, Ŭ����
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
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtYEARMON.value = ""  then
		gErrorMsgBox "����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
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
	
	With frmThis
		vntInParams = array("", "CLIENT") '<< �޾ƿ��°�� '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTCHKPOP.aspx",vntInParams , 413,435)
		if vntRet <> "" then
			mstrClientcode = vntRet
			SelectRtn
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

	with frmThis
		vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME.value) '<< �޾ƿ��°��
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
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					
					mstrClientcode = trim(vntData(0,1))
					SELECTRTN
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'üũü����
Sub chkALL_onclick
	CheckCleanField
End Sub

Sub chkMEDFLAG1_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG2_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG3_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG4_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG5_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG6_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG7_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG8_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chkMEDFLAG9_onclick
	frmThis.chkALL.checked = False
	chk_chk
End Sub

Sub chk_chk
	with frmThis
		If .chkMEDFLAG1.checked = false and .chkMEDFLAG2.checked = false and .chkMEDFLAG3.checked = false and  _
		   .chkMEDFLAG4.checked = false and .chkMEDFLAG5.checked = false and .chkMEDFLAG6.checked = false and  _
		   .chkMEDFLAG7.checked = false and .chkMEDFLAG8.checked = false and .chkMEDFLAG9.checked = false then
		.chkALL.checked = True
		Else
		.chkALL.checked = false
		end If
	end with
End Sub

Sub CheckCleanField
	with frmThis
		.chkALL.checked = True
		.chkMEDFLAG1.checked = False
		.chkMEDFLAG2.checked = False
		.chkMEDFLAG3.checked = False
		.chkMEDFLAG4.checked = False
		.chkMEDFLAG5.checked = False
		.chkMEDFLAG6.checked = False
		.chkMEDFLAG7.checked = False
		.chkMEDFLAG8.checked = False
		.chkMEDFLAG9.checked = False
	End with
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
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

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
	set mobjMDCMGET = Nothing
	set mobjMDSRREPORTLIST = Nothing
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
		.txtYEARMON.value = Mid(gNowDate2,1,4)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEARMON.focus()
		
	End with
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
   	Dim strMEDFLAGALL,strMEDFLAG1,strMEDFLAG2,strMEDFLAG3,strMEDFLAG4,strMEDFLAG5,strMEDFLAG6,strMEDFLAG7,strMEDFLAG8,strMEDFLAG9
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		strLIST = ""
		chkflag = 1
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
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
		
		If .chkALL.checked = True Then strMEDFLAGALL = "1" Else strMEDFLAGALL = "0" 
		If .chkMEDFLAG1.checked = True Then strMEDFLAG1 = "1" Else strMEDFLAG1 = "0"  
		If .chkMEDFLAG2.checked = True Then strMEDFLAG2 = "1" Else strMEDFLAG2 = "0"
		If .chkMEDFLAG3.checked = True Then strMEDFLAG3 = "1" Else strMEDFLAG3 = "0"
		If .chkMEDFLAG4.checked = True Then strMEDFLAG4 = "1" Else strMEDFLAG4 = "0"
		If .chkMEDFLAG5.checked = True Then strMEDFLAG5 = "1" Else strMEDFLAG5 = "0"
		If .chkMEDFLAG6.checked = True Then strMEDFLAG6 = "1" Else strMEDFLAG6 = "0"
		If .chkMEDFLAG7.checked = True Then strMEDFLAG7 = "1" Else strMEDFLAG7 = "0"
		If .chkMEDFLAG8.checked = True Then strMEDFLAG8 = "1" Else strMEDFLAG8 = "0"
		If .chkMEDFLAG9.checked = True Then strMEDFLAG9 = "1" Else strMEDFLAG9 = "0"
		
		vntData = mobjMDSRREPORTLIST.SelectRtn_MONMEDLIST2(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value, _
														   strLIST, strFROMMON, strTOMON, strMONCNT, strMEDFLAGALL, _
														   strMEDFLAG1, strMEDFLAG2, strMEDFLAG3, strMEDFLAG4, strMEDFLAG5, _
														   strMEDFLAG6, strMEDFLAG7, strMEDFLAG8, strMEDFLAG9)

		if not gDoErrorRtn ("SelectRtn_MONMEDLIST2") then
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
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
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
												<TABLE cellSpacing="0" cellPadding="0" width="146" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���� ���� ��ü�� �����&nbsp;</td>
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
									<TABLE id="tblButton" cellSpacing="0" cellPadding="0" border="0">
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
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
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
											<TD class="SEARCHDATA" style="WIDTH: 424px" width="424"><INPUT class="INPUT" id="txtYEARMON" title="�⵵���Է��ϼ���" style="WIDTH: 100px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="4" size="14" name="txtYEARMON">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
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
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="80">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 192px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="26" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
											</TD>
											<TD class="SEARCHDATA" width="100"><IMG id="imgCUSTPOP" onmouseover="JavaScript:this.src='../../../images/imgCustMultiChkOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCustMultiChk.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgCustMultiChk.gIF" border="0" name="imgCUSTPOP"></TD>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand">��ü����
											</TD>
											<TD class="SEARCHDATA" colSpan="4"><INPUT id="chkALL" type="checkbox" CHECKED name="chkALL">&nbsp;��ü&nbsp;
												<INPUT id="chkMEDFLAG3" type="checkbox" name="chkMEDFLAG3">&nbsp;TV 
												<INPUT id="chkMEDFLAG4" type="checkbox" name="chkMEDFLAG4">&nbsp;Radio
												<INPUT id="chkMEDFLAG5" type="checkbox" name="chkMEDFLAG5">&nbsp;������DMB 
												<INPUT id="chkMEDFLAG6" type="checkbox" name="chkMEDFLAG6">&nbsp;CATV
												<INPUT id="chkMEDFLAG9" type="checkbox" name="chkMEDFLAG9">&nbsp;���������
												<INPUT id="chkMEDFLAG1" type="checkbox" name="chkMEDFLAG1">&nbsp;�Ź� 
												<INPUT id="chkMEDFLAG2" type="checkbox" name="chkMEDFLAG2">&nbsp;����
												<INPUT id="chkMEDFLAG7" type="checkbox" name="chkMEDFLAG7">&nbsp;���ͳ� 
												<INPUT id="chkMEDFLAG8" type="checkbox" name="chkMEDFLAG8">&nbsp;����</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 2px"><FONT face="����"></FONT></TD>
							</TR> <!--BodySplit End--> <!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="16722">
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
							</TR> <!--List End--> <!--Bottom Split Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
								</TD>
							</TR> <!--Bottom Split End-->
						</TABLE> 
						<!--Input Define Table End-->
					</TD>
				</TR> <!--Top TR End-->
			</TABLE>
		</FORM>
	</body>
</HTML>
