<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMMERGETAXDTL.aspx.vb" Inherits="PD.PDCMMERGETAXDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ۺ� ���ݰ�꼭</title>
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
Dim mobjPDCMMERGETAX
Dim mobjPDCMGET
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
Sub imgQuery_onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub



Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim strMTAXYEARMON
	Dim strMTAXNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim VATFLAG
	Dim FLAG
	Dim strUSERID
	
	IF frmThis.sprSht.MaxRows = 0 then
		gFlowWait meWAIT_ON
		with frmThis		
			ModuleDir = "PD"
			ReportName = "TAXNO_BLACK.rpt"
						
			IF .cmbFLAG.value = "receipt" THEN
				FLAG = "Y"
			ELSE
				FLAG = "N"
			END IF
						
			Params = FLAG
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
		end with
		gFlowWait meWAIT_OFF
	else
		
		gFlowWait meWAIT_ON
		with frmThis
			'�μ��ư�� Ŭ���ϱ� ���� PD_tax_temp���̺��� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			'PD_trans_temp���� ����
			intRtn = mobjPDCMMERGETAX.DeleteRtn_TEMP(gstrConfigXml)
			'PD_trans_temp���� ��
			
			ModuleDir = "PD"
			'������/���޹޴��� �������� ���忡 �ٺ����ְų� ���޹޴��� �����븸 �����ִ� ��
			ReportName = "MERGETAX.rpt"
			
			mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
			strMTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"MTAXYEARMON",1)
			strMTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"MTAXNO",1)
			IF .txtVAT.value = 0 OR .txtVAT.value = "" THEN
				VATFLAG = "N"
			ELSE
				VATFLAG = "Y"
			END IF
			
			FLAG = "Y"
			strUSERID = ""
			
			vntDataTemp = mobjPDCMMERGETAX.ProcessRtn_TEMP(gstrConfigXml,strMTAXYEARMON, strMTAXNO, VATFLAG, FLAG, i, strUSERID)
			
			Params = "PD_MERGETAX_TEMP" & ":" & strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	



'����� �Ϸ���� PD_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMMERGETAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
Dim intNo,i,vntInParam
	'����������ü ����	
	set mobjPDCMMERGETAX	= gCreateRemoteObject("cPDCO.ccPDCOMERGETAX")	
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "152px"
	'pnlTab1.style.height ="300px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
    with frmThis
		
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		
		for i = 0 to intNo
			select case i
				case 0 : .txtMTAXYEARMON.value = vntInParam(i)	
				case 1 : .txtMTAXNO.value = vntInParam(i)
			end select
		next
		
		'Sheet Į�� ����
	    gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0
	    mobjSCGLSpr.SpreadDataField .sprSht, "MTAXYEARMON | MTAXNO | MEDFLAG | MEDFLAGNAME | CLIENTNAME | CLIENTBISNO | AMT | VAT | SUMAMT | DEPT_NAME"
		mobjSCGLSpr.SetHeader .sprSht,       "���|��ȣ|����|����|������|�����ֻ���ڹ�ȣ|�ݾ�|�ΰ���|��|���μ�",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "0 |   0|   0|   5|    20|              14|  10|     9|10|     12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MTAXYEARMON | MTAXNO | MEDFLAG | MEDFLAGNAME | CLIENTNAME | CLIENTBISNO | DEPT_NAME ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"MTAXYEARMON | MTAXNO | MEDFLAG | MEDFLAGNAME | CLIENTNAME | CLIENTBISNO | AMT | VAT | SUMAMT | DEPT_NAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDFLAGNAME | CLIENTBISNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "MTAXYEARMON | MTAXNO | MEDFLAG",true
	
	End with

	pnlTab1.style.visibility = "visible" 
	'�ϴ���ȸ
	SelectRtn
	'ȭ�� �ʱⰪ ����
	'InitPageData	
End Sub
'=================================�հ转Ʈ ó�� ����
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub


'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub AMTSUM
	Dim lngCnt
	Dim lngAMT, lngVAT
	Dim lngAMTSUM,lngVATSUM
	With frmThis
		lngAMTSUM = 0
		lngVATSUM = 0

		For lngCnt = 1 To .sprSht.MaxRows
			lngAMT = 0
			lngVAT = 0
			
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			lngVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
			IngAMTSUM = IngAMTSUM + lngAMT
			lngVATSUM = lngVATSUM + lngVAT
		Next

	End With
End Sub

'=================================�հ转Ʈ ó�� ��
'���ް��� �ݾ�ó��
Sub txtAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub
Sub txtAMT_onblur
	
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub
'�ΰ��� �ݾ�ó��
Sub txtVAT_onfocus
	with frmThis
		.txtVAT.value = Replace(.txtVAT.value,",","")
	end with
End Sub
Sub txtVAT_onblur
	with frmThis
		.txtSUMAMT.value = cdbl(.txtVAT.value) + cdbl(.txtAMT.value) 
		call gFormatNumber(.txtVAT,0,true)
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
'�հ�ݾ� ó��
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸMASTER
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strMTAXYEARMON
	Dim strMTAXNO
	Dim intCnt
	Dim strCNT
	With frmThis
		
		strMTAXYEARMON	= .txtMTAXYEARMON.value
		strMTAXNO	= .txtMTAXNO.value
		IF strMTAXYEARMON = "" OR strMTAXNO = ""  THEN
			gErrorMsgBox "�˻����ǿ� ���ݰ�꼭 ��ȣ�� �ݵ�� �����ž� �մϴ�.","��ȸ�ȳ�!"
			
			If strMTAXYEARMON = "" AND strMTAXNO = "" Then
				.strMTAXYEARMON.focus
			Elseif strMTAXYEARMON = "" And strMTAXNO <> "" Then
				.strMTAXYEARMON.focus
			Elseif strMTAXYEARMON <> "" And strMTAXNO = "" Then
				.txtMTAXNO.focus
			End If
			
			Exit Sub
			
		End If
	End With 
	
	IF not SelectRtn_HDR (strMTAXYEARMON, strMTAXNO) Then Exit Sub
	
	'��Ʈ ��ȸ
	'Call SelectRtn_DTL 
	If not SelectRtn_DTL(strMTAXYEARMON, strMTAXNO) Then
		gErrorMsgBox "����ȸ���� ��ȸ����","��ȸ�ȳ�!"
		InitPageData
		Exit Sub
	Else
		AMTSUM
	End If
	with frmThis
	If .txtVOCHNO.value <> ""  Then
	.txtVAT.className = "NOINPUTB_R"
	.txtSUMM.className = "NOINPUTB"
	.txtVAT.readOnly = true
	.txtSUMM.readOnly = true
	Else
	.txtVAT.className = "INPUT_R"
	.txtSUMM.className = "INPUTB"
	.txtVAT.readOnly = false
	.txtSUMM.readOnly = false	
	End If
	End With
	'SHEET1_SUM
	gWriteText lblStatus, "�����Ͻ� ���ݰ�꼼�� �� ���Ͽ� �ڷᰡ �˻�" & mePROC_DONE
End Sub
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸHEADER
'-----------------------------------------------------------------------------------------
Function SelectRtn_HDR(ByVal strMTAXYEARMON, ByVal strMTAXNO)
	dim vntData
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMMERGETAX.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strMTAXYEARMON,strMTAXNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ ���ݰ�꼼�� ��ȣ �� ���Ͽ�" & meNO_DATA, ""
			InitPageData
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMT_onblur
			SelectRtn_HDR = True 
			'gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
		End IF
	End IF
End Function
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸDETAIL
'-----------------------------------------------------------------------------------------
Function SelectRtn_DTL (ByVal strMTAXYEARMON, ByVal strMTAXNO)
	Dim vntData
	Dim lngCnt
	'on error resume next
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMMERGETAX.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strMTAXYEARMON,strMTAXNO)
	
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		mobjSCGLSpr.SetClipbinding frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_DTL = True
	End IF
End Function


Sub ProcessRtn
	Dim intRtn
	Dim strMTAXYEARMON
	Dim strMTAXNO
	Dim strSUMM
	Dim lngVAT
	Dim strVOCHNO
	with frmThis
	'��ǥ����� �����Ұ�
	If .txtVOCHNO.value <> "" Then
		strVOCHNO = .txtVOCHNO.value 
		gErrorMsgBox "��ǥ��ȣ" & strVOCHNO & " �� �����մϴ�.��ǥ����� �����Ͻʽÿ�.","����ȳ�!"
		Exit Sub
	'ELSEif NOT VOCHNO_CHECKED (.txtMTAXYEARMON.value,.txtMTAXNO.value) then
	'	gErrorMsgBox .txtMTAXYEARMON.value & "-" & .txtMTAXNO.value & " �� ���Ͽ�" &vbcrlf & "��ǥó�� �������� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
	'	Exit Sub		
	End IF
		If .txtMTAXYEARMON.value = "" Or .txtMTAXNO.value = "" Or .txtSUMM.value = "" Then
			gErrorMsgBox "���ݰ�꼭 ��� �� ��ȣ�� �Է��Ͽ��ּ���","����ȳ�!"
			Exit Sub
		End If
		strMTAXYEARMON = .txtMTAXYEARMON.value
		strMTAXNO = .txtMTAXNO.value
		strSUMM = .txtSUMM.value
		lngVAT = Replace(.txtVAT.value,",","")
		
		intRtn = mobjPDCMMERGETAX.ProcessRtn_SUMM(gstrConfigXml,strMTAXYEARMON,strMTAXNO,strSUMM,lngVAT)
		if not gDoErrorRtn ("ProcessRtn_SUMM") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "���ݰ�꼭 [" & strMTAXYEARMON & "-" & strMTAXNO & "] �� ����" & mePROC_DONE,"����ȳ�" 
			SelectRtn
   		end if
	end with
End Sub

Sub EndPage()
	set mobjPDCMMERGETAX = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
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

Function VOCHNO_CHECKED (ByRef strMTAXYEARMON, ByRef strMTAXNO)
	Dim vntData
	Dim intCnt
	Dim strCOUNT
	'on error resume next

	'�ʱ�ȭ
	VOCHNO_CHECKED = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMGET.COMMIVOCHNO_CHECKED(gstrConfigXml,mlngRowCnt,mlngColCnt, strMTAXYEARMON,strMTAXNO)
	
	IF mlngRowCnt >0 THEN
		VOCHNO_CHECKED = false
	ELSE
		VOCHNO_CHECKED = TRUE	
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
													<TABLE cellSpacing="0" cellPadding="0" width="120" background="../../../images/back_p.gIF"
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
												<td class="TITLE">���ռ��ݰ�꼭 ��</td>
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
										<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="203" border="0">
											<TR>
												<TD></TD>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="���丸 ���� �����մϴ�."
														src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
												<TD></TD>
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
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
													<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(strMTAXYEARMON,txtMTAXNO)"
														width="80">��꼭��ȣ</TD>
													<TD class="SEARCHDATA"><INPUT dataFld="MTAXYEARMON" class="INPUT" id="txtMTAXYEARMON" title="���ݰ�꼭���" style="WIDTH: 56px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtMTAXYEARMON">&nbsp;-
														<INPUT dataFld="MTAXNO" class="INPUT" id="txtMTAXNO" title="���ݰ�꼭��ȣ" style="WIDTH: 48px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="4" size="2" name="txtMTAXNO">&nbsp;<INPUT dataFld="VOCHNO" id="txtVOCHNO" style="WIDTH: 144px; HEIGHT: 21px" dataSrc="#xmlBind"
															type="hidden" size="18" name="txtVOCHNO"></TD>
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
												<TD class="SEARCHLABEL" width="90"><FONT face="����">�� �� ��</FONT></TD>
												<TD class="SEARCHDATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME">
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="����"> ��</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="����"><INPUT dataFld="TIMNAME" class="NOINPUTB_L" id="txtTIMNAME" title="����" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtTIMNAME"></FONT>
												</TD>
												<TD class="SEARCHLABEL" width="90"><FONT face="����">û������</FONT></TD>
												<TD class="SEARCHDATA" width="173"><FONT face="����"><INPUT dataFld="DEMANDDAY" class="NOINPUTB" id="txtDEMANDDAY" title="û����" style="WIDTH: 172px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtDEMANDDAY"></FONT></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL"><FONT face="����">���ް���</FONT></TD>
												<TD class="SEARCHDATA"><FONT face="����"><INPUT dataFld="AMT" class="NOINPUTB_R" id="txtAMT" title="���ް���" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtAMT"></FONT>
												</TD>
												<TD class="SEARCHLABEL"><FONT face="����">�ΰ�����</FONT></TD>
												<TD class="SEARCHDATA"></FONT></FONT><INPUT dataFld="VAT" class="INPUTB_R" id="txtVAT" title="�ΰ���" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtVAT"></TD>
												<TD class="SEARCHLABEL"><FONT face="����">�հ�ݾ�</FONT></TD>
												<TD class="SEARCHDATA"></FONT></FONT><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="�հ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtSUMAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">����ڹ�ȣ</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="BUSINO" class="NOINPUTB" id="txtBUSINO" title="����ڹ�ȣ" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtBUSINO"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUMM,'')"><FONT face="����">��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
														��</FONT></TD>
												<TD class="SEARCHDATA" colSpan="3"><INPUT dataFld="SUMM" class="INPUT_L" id="txtSUMM" title="����" style="WIDTH: 441px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="100" size="66" name="txtSUMM">
												</TD>
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