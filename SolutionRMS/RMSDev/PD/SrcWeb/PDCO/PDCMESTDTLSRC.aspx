<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMESTDTLSRC.aspx.vb" Inherits="PD.PDCMESTDTLSRC" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� ��������</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ������������ ȭ��(PDCMESTDTL)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPREESTDTL.aspx
'��      �� : ������ ���� ��� �� Ȯ��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/16 By Tae Ho Kim
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
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMPREESTDTL '�����ڵ�, Ŭ����
Dim mstrPROCESS
Dim mstrPROCESS2 '��ȸ�����̸� true �űԻ����̸� false
Dim mstrCheck
Dim mobjMDLOGIN
Dim mobjMDCMEMP
Dim mobjPDCMGET
CONST meTAB = 9
mstrPROCESS = TRUE
mstrPROCESS2 = TRUE
mstrCheck = True

'=============================
' �̺�Ʈ ���ν��� 
'=============================
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
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


'Sub imgDelete_onclick
'	gFlowWait meWAIT_ON
'	DeleteRtn
'	gFlowWait meWAIT_OFF
'End Sub




Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'���������� ���� ���� Ŭ�� �� �߻�
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


'=============================
' UI���� ���ν��� 
'=============================
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub
Sub txtSUMAMT_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtSUMAMT.value,",","")
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	dim vntInParam
	dim intNo,i
	
	set mobjPDCMPREESTDTL	= gCreateRemoteObject("cPDCO.ccPDCOPREESTDLT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "232px"
	pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
				case 1 : frmThis.txtJOBNO.value = vntInParam(i)
			end select
		next
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis

		
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT"
		mobjSCGLSpr.SetHeader .sprSht,		  "��������ȣ|����|��з�|�ߺз�|�����׸��ڵ�|�����׸��|����|Ŀ�̼�|����|�ܰ�|�ݾ�|������ݾ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|   0|     8|    12|          10|        18|  28|     6|  12|  13|15  |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMCODENAME|STD", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT"
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|ITEMCODESEQ|SUSUAMT|ITEMCODESEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE|ITEMCODESEQ",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME|CLASSNAME",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0
	End With
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	SelectRtn
End Sub

Sub EndPage()
	set mobjPDCMPREESTDTL = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub
'-----------------------------
' Ȯ�� �� Ȯ����� ó��
'-----------------------------	
Sub imgSetting_onclick
	Dim intRtnConfirm
	Dim intRtn
	intRtnConfirm = gYesNoMsgbox("�ڷḦ Ȯ�� �Ͻðڽ��ϱ�?","�ڷ�Ȯ�� Ȯ��")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_Confirm(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_Confirm") then
				gErrorMsgBox " �ڷᰡ Ȯ�� �Ǿ����ϴ�.","Ȯ���ȳ�" 
			End If
			ESTCONFIRM_Search
	End with
End Sub

Sub ImgConfirmCancel_onclick
	Dim intRtnConfirm
	Dim intRtn
	intRtnConfirm = gYesNoMsgbox("�ڷḦ Ȯ����� �Ͻðڽ��ϱ�?","�ڷ�Ȯ����� Ȯ��")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
		intRtn = mobjPDCMPREESTDTL.ProcessRtn_ConfirmCancel(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmCancel") then
			gErrorMsgBox " �ڷᰡ Ȯ����� �Ǿ����ϴ�.","Ȯ����Ҿȳ�" 
		End If
		ESTCONFIRM_Search	
	End with
End Sub
'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtPRINTDAY.value = gNowDate
		.sprSht.MaxRows = 0
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'û��Ȯ�� ��ȸ
Sub ESTCONFIRM_Search
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
				.imgSetting.disabled = true
				.ImgConfirmCancel.disabled = false
			Else
				.imgSetting.disabled = false
				.ImgConfirmCancel.disabled = true
			End if
   		end if
	end with
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	Dim strJOBCODE
	With frmThis
		strCODE = .txtPREESTNO.value
		strJOBCODE = .txtJOBNO.value
		IF strCODE = ""  THEN
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = TRUE
			IF not SelectRtn_HeadLess (strJOBCODE) Then Exit Sub
			
		Else
			IF not SelectRtn_Head (strCODE) Then Exit Sub

			'��Ʈ ��ȸ
			CALL SelectRtn_Detail (strCODE)
			txtSUSUAMT_onblur
			txtCOMMITION_onblur
			txtSUMAMT_onblur
			txtNONCOMMITION_onblur
		End If
	End With
End Sub
'���������� ���� ��� ��ȸ
Function SelectRtn_HeadLess (ByVal strJOBCODE)
	Dim vntData
	'on error resume next

	'�ʱ�ȭ
	SelectRtn_HeadLess = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDRLESS(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBCODE)
	
	IF not gDoErrorRtn ("SelectRtn_HeadLess") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ ������ ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_HeadLess = True
		End IF
	End IF
End Function
'���������� ���� ��� ��ȸ
Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next

	'�ʱ�ȭ
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ �������� ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function


'���� ���̺� ��ȸ
Function SelectRtn_Detail (ByVal strCODE)
	dim vntData
	Dim intCnt
	Dim strRows
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		End with
	End IF
End Function

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strTBRDSTDATE,strTBRDEDDATE, strCAMPAIGN_CODE, strCAMPAIGN_NAME, strCLIENTCODE, strCLIENTNAME)
	With frmThis
		.txtTBRDSTDATE1.value = strTBRDSTDATE
		.txtTBRDEDDATE1.value = strTBRDEDDATE
		.txtCAMPAIGN_CODE1.value = strCAMPAIGN_CODE
		.txtCAMPAIGN_NAME1.value = strCAMPAIGN_NAME
		.txtCLIENTCODE1.value = strCLIENTCODE
		.txtCLIENTNAME1.value = strCLIENTNAME
	End With
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%"cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;JOB �󼼳���</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0"
							width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">���� �ڵ�</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="PREESTNO" class="NOINPUT_L" id="txtPREESTNO" title="�������ڵ�" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtPREESTNO"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">������</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="PREESTNAME" class="NOINPUT_L" id="txtPREESTNAME" title="��������" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtPREESTNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAGREEYEARMON,'')">����Ȯ����</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="AGREEYEARMON" class="NOINPUT" id="txtAGREEYEARMON" title="����������" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="10" name="txtAGREEYEARMON"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">���� �Ǹ�</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="���۰Ǹ�" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtJOBNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">��ü�ι�</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="��ü�ι�" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtJOBGUBN"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">��ü�з�</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="��ü�з�" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtCREPART"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">������</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="������" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">�����</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="�����" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCLIENTSUBNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">�귣��</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="�귣��" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUBSEQNAME"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 72px" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left"  height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;��������</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<tr height="5">
											<td></td>
										</tr>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="1040" border="0" align="LEFT">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">���ۼ�����</TD>
											<TD class="DATA" width="230"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="���ۼ�����" style="WIDTH: 224px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="32" name="txtSUSUAMT">
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand" align="right" width="94">Commition</TD>
											<TD class="DATA" width="260"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="commition ��" style="WIDTH: 256px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand" align="right" width="80">�հ�</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="���հ�ݾ�" style="WIDTH: 272px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUMAMT"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">��������</TD>
											<TD class="DATA"><INPUT dataFld="SUSURATE" class="NOINPUT_R" id="txtSUSURATE" style="WIDTH: 200px; HEIGHT: 22px"
													accessKey=",NUM,M" dataSrc="#xmlBind" readOnly type="text" size="28" name="txtSUSURATE">&nbsp;(%)
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px">Non Commition</TD>
											<TD class="DATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="noncommition ��"
													style="WIDTH: 256px; HEIGHT: 22px" accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37"
													name="txtNONCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" width="80">������ ���</TD>
											<TD class="DATA"><INPUT class="NOINPUT" id="txtPRINTDAY" title="������������" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" readOnly type="text" maxLength="10" size="10" name="txtPRINTDAY">&nbsp;&nbsp;
												<IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="������ �� �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" align="absMiddle"
													border="0" name="imgPrint">&nbsp;<INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtJOBNO"><INPUT dataFld="CREDAY" id="txtCREDAY" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCREDAY"><INPUT dataFld="CLIENTSUBCODE" id="txtCLIENTSUBCODE" style="WIDTH: 16px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="hidden" size="1" name="txtCLIENTSUBCODE"><INPUT dataFld="CLIENTCODE" id="txtCLIENTCODE" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCLIENTCODE"><INPUT dataFld="SUBSEQ" id="txtSUBSEQ" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtSUBSEQ"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27464">
											<PARAM NAME="_ExtentY" VALUE="12515">
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
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
