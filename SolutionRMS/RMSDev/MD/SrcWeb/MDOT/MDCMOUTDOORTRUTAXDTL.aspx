<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORTRUTAXDTL.aspx.vb" Inherits="MD.MDCMOUTDOORTRUTAXDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ݰ�꼭 �󼼳���</title>
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
Dim mobjMDCMTRUTAX
Dim mstrMEDFLAG
Dim mobjMDCMOUTDOORTRUTAX
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

Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	
	gFlowWait meWAIT_ON
	with frmThis
		strTAXYEARMON = .txtTAXYEARMON.value
		strTAXNO = .txtTAXNO.value
		IF strTAXYEARMON = "" THEN
			gErrorMsgBox "���ݰ�꼭����� �Է��ϼ���",""
			EXIT Sub
		END IF
		
		IF strTAXYEARMON = "" THEN
			gErrorMsgBox "���ݰ�꼭��ȣ�� �Է��ϼ���",""
			EXIT Sub
		END IF
		
		'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCMOUTDOORTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "TRANSTAX_BLACK_NEW.rpt"
			
		vntDataTemp = mobjMDCMOUTDOORTRUTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO)

	
		Params = ""
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "printSetTimeout", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMOUTDOORTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	Dim intNo,i,vntInParam
	'����������ü ����	
	set mobjMDCMTRUTAX  	 = gCreateRemoteObject("cMDCO.ccMDCOTRUTAX")		'���ݰ�꼭��ȸ
	set mobjMDCMOUTDOORTRUTAX	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOORTRUTAX")
	

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "176px"
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
				case 0 : .txtTAXYEARMON.value = vntInParam(i)	
				case 1 : .txtTAXNO.value = vntInParam(i)
				case 2 : mstrMEDFLAG = vntInParam(i)
			end select
		next
		'ȭ���� �������� �����ϱ� ����(Tab�� ���� ó���� ǥ�õǴ� �͸� ��)
		'.sprSht.style.visibility = "hidden"
		
		'**************************************************
		'***ù��° Sheet ������
		'**************************************************
		
		'Sheet Į�� ����
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout ������
		mobjSCGLSpr.SpreadLayout .sprSht, 8, 0
		
		'Binding Field ����
	    mobjSCGLSpr.SpreadDataField .sprSht, "TAXYEARMON|TAXNO|TAXNOSEQ|TRANSNO|MEDNAME|MEDCODE|AMT|VAT"
		'Header ������
		mobjSCGLSpr.SetHeader .sprSht,       "���|��ȣ|����|�ŷ�������ȣ|��ü��|��ü�ڵ�|�ݾ�|�ΰ�����",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "0 |0   |0   |20            |35    |0       |20  |15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"TAXYEARMON|TAXNO|TAXNOSEQ|TRANSNO|MEDNAME|MEDCODE|AMT|VAT"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "TRANSNO|MEDNAME",-1,-1,0,2,false
		mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON|TAXNO|TAXNOSEQ|MEDCODE",true
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 8, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "TAXYEARMON|TAXNO|TAXNOSEQ|TRANSNO|MEDNAME|MEDCODE|AMT|VAT"
		mobjSCGLSpr.SetText .sprSht_SUM, 4, 1, "��   ��"
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"4",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|VAT", -1, -1, 0
	    mobjSCGLSpr.SetCellAlign2 .sprSht_SUM, "TRANSNO",-1,-1,2,2,false
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
		
	
	End with

	pnlTab1.style.visibility = "visible" 
	'�ϴ���ȸ
	SelectRtn
	'ȭ�� �ʱⰪ ����
	'InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMTRUTAX = Nothing
	set mobjMDCMOUTDOORTRUTAX = Nothing
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

'=================================�հ转Ʈ ó�� ����
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM	
	End with
end sub

'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_SUM, NewTop, NewLeft
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
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, IngAMTSUM
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, lngVATSUM
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
		call gFormatNumber(.txtVAT,0,true)
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
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim intCnt
	Dim strCNT
	With frmThis
		strTAXYEARMON		= .txtTAXYEARMON.value
		strTAXNO	= .txtTAXNO.value
		
		IF strTAXYEARMON = "" OR strTAXNO = ""  THEN
			gErrorMsgBox "�˻����ǿ� ���ݰ�꼭 ��ȣ�� �ݵ�� �����ž� �մϴ�.","��ȸ�ȳ�!"
			If strTAXYEARMON = "" AND strTAXNO = "" Then
			.txtTAXYEARMON.focus
			Elseif strTAXYEARMON = "" And strTAXNO <> "" Then
			.txtTAXYEARMON.focus
			Elseif strTAXYEARMON <> "" And strTAXNO = "" Then
			.txtTAXNO.focus
			End If
			Exit Sub
		End If
	End With 
	
	IF not SelectRtn_HDR (strTAXYEARMON, strTAXNO) Then Exit Sub
	
	'��Ʈ ��ȸ
	'Call SelectRtn_DTL 
	If not SelectRtn_DTL(strTAXYEARMON, strTAXNO) Then
		gErrorMsgBox "����ȸ���� ��ȸ����","��ȸ�ȳ�!"
		InitPageData
		Exit Sub
	Else
		AMTSUM
	End If
	'SHEET1_SUM
	gWriteText lblStatus, "�����Ͻ� ���ݰ�꼼�� �� ���Ͽ� �ڷᰡ �˻�" & mePROC_DONE
End Sub
'-----------------------------------------------------------------------------------------
' ���ݰ�꼭��ȸHEADER
'-----------------------------------------------------------------------------------------
Function SelectRtn_HDR(ByVal strTAXYEARMON, ByVal strTAXNO)
	dim vntData
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
	vntData = mobjMDCMTRUTAX.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strTAXYEARMON,strTAXNO,mstrMEDFLAG)

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
Function SelectRtn_DTL (ByVal strTAXYEARMON, ByVal strTAXNO)
	Dim vntData
	Dim lngCnt
	'on error resume next
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMTRUTAX.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strTAXYEARMON,strTAXNO,mstrMEDFLAG)
	
	IF not gDoErrorRtn ("SelectRtn_TRANSNO_DTL") then
		mobjSCGLSpr.SetClipbinding frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		SelectRtn_DTL = True
	End IF
End Function

Sub ProcessRtn
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strSUMM
	with frmThis
		If .txtTAXYEARMON.value = "" Or .txtTAXNO.value = "" Or .txtSUMM.value = "" Then
			gErrorMsgBox "���ݰ�꼭 ��� �� ��ȣ�� �Է��Ͽ��ּ���","����ȳ�!"
			Exit Sub
		End If
		strTAXYEARMON = .txtTAXYEARMON.value
		strTAXNO = .txtTAXNO.value
		strSUMM = .txtSUMM.value
		strVAT = .txtVAT.value
		
		intRtn = mobjMDCMTRUTAX.ProcessRtn_SUMM(gstrConfigXml,strTAXYEARMON,strTAXNO,strSUMM, strVAT)
		if not gDoErrorRtn ("ProcessRtn_SUMM") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "���ݰ�꼭 [" & strTAXYEARMON & "-" & strTAXNO & "] �� �ΰ��� �� ���䰡 ����" & mePROC_DONE,"����ȳ�" 
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
			<TABLE id="tblForm" style="WIDTH: 793px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 427px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">���ܱ��� ����Ź ���ݰ�꼭</td>
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
														name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�."
														src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"></TD>
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
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="LABEL" style="WIDTH: 77px; CURSOR: hand" width="78" onclick="vbscript:Call gCleanField(txtTAXYEARMON,txtTAXNO)">��꼭��ȣ</TD>
													<TD class="DATA"><INPUT dataFld="TAXYEARMON" class="INPUT" id="txtTAXYEARMON" title="���ݰ�꼭���" style="WIDTH: 56px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtTAXYEARMON">&nbsp;-
														<INPUT dataFld="TAXNO" class="INPUT" id="txtTAXNO" title="���ݰ�꼭��ȣ" style="WIDTH: 48px; HEIGHT: 22px"
															accessKey="NUM" dataSrc="#xmlBind" type="text" maxLength="4" size="2" name="txtTAXNO"></TD>
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
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="right" border="0">
											<TR>
												<TD class="LABEL" width="78" style="WIDTH: 78px"><FONT face="����"> ������</FONT></TD>
												<TD class="DATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME">
												</TD>
												<TD class="LABEL" width="90"><FONT face="����"> ��ü��</FONT></TD>
												<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="REAL_MED_NAME" class="NOINPUTB_L" id="txtREAL_MED_NAME" title="û������" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtDEPT_NAME"></FONT>
												</TD>
												<TD class="LABEL" width="90"><FONT face="����">û������</FONT></TD>
												<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="DEMANDDAY" class="NOINPUTB" id="txtDEMANDDAY" title="û����" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtDEMANDDAY" accessKey="DATE"></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 82px"><FONT face="����">�����ֻ����</FONT></TD>
												<TD class="DATA"><FONT face="����"><INPUT dataFld="CLIENTBISNO" class="NOINPUTB" id="txtCLIENTBISNO" title="���ް���" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtCLIENTBISNO"></FONT>
												</TD>
												<TD class="LABEL"><FONT face="����">��ü������</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="REAL_MED_BISNO" class="NOINPUTB" id="txtREAL_MED_BISNO" title="�ΰ���" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtREAL_MED_BISNO"></TD>
												<TD class="LABEL"><FONT face="����">���μ�</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="DEPTNAME" class="NOINPUTB" id="txtDEPTNAME" title="�հ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtDEPTNAME"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 82px"><FONT face="����"> ���ް���</FONT></TD>
												<TD class="DATA"><FONT face="����"><INPUT dataFld="AMT" class="NOINPUTB_R" id="txtAMT" title="���ް���" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtAMT"></FONT>
												</TD>
												<TD class="LABEL"><FONT face="����">�ΰ�����</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="VAT" class="INPUT_R" id="txtVAT" title="�ΰ���" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="100" size="22" name="txtVAT"></TD>
												<TD class="LABEL"><FONT face="����">�հ�ݾ�</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="�հ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="22" name="txtSUMAMT"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 82px">��ǥ��ȣ</TD>
												<TD class="DATA"><INPUT class="NOINPUTB" id="txtVOCHNO" title="����ڹ�ȣ" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="8" name="txtVOCHNO" dataFld="VOCHNO"></TD>
												<TD class="LABEL" style="CURSOR:hand" onclick="vbscript:Call gCleanField(txtSUMM,'')"><FONT face="����">
														����</FONT></TD>
												<TD class="DATA" colspan="3"><INPUT dataFld="SUMM" class="INPUT_L" id="txtSUMM" title="����" style="WIDTH: 442px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" name="txtSUMM" size="67">
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
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 312px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 288px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20929">
									<PARAM NAME="_ExtentY" VALUE="7620">
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
								<OBJECT id="sprSht_SUM" style="WIDTH: 100%; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20929">
									<PARAM NAME="_ExtentY" VALUE="635">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� 
							�� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="�ݾ�" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<!--
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 794px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					-->
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
