<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_CFINPUT.aspx.vb" Inherits="PD.PDCMJOBMST_CFINPUT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���������</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBMST_SUBITEM.aspx
'��      �� : JOBMST�� �ι�° �� PDCMJOBMST_ESTDTL �� ������ó�� ��ư�� Ŭ���Ͽ����� ó�� 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/28 By KimTH
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
		<script language="vbscript" id="clientEventHandlersVBS">
Dim mobjPDCOCFINPUT
Dim mstrPREESTNO



'DIVNAME,CLASSNAME,ITEMCODENAME,ITEMCODE,IMESEQ,SAVEFLAG
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	with frmThis
	
	End with
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub


Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i

									  
	set mobjPDCOCFINPUT = gCreateRemoteObject("cPDCO.ccPDCOCFINPUT")

	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue

	'gSetSheetDefaultColor
	
	
	with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		.txtPREESTNO.style.visibility = "hidden"
		.txtSAVEFLAG.style.visibility = "hidden"
		for i = 0 to intNo
			select case i
				case 0 : mstrPREESTNO = vntInParam(i)				'������ȣ�� �����´�.
			end select
		next								  
	End With
	
	InitpageData
	SelectRtn
	
End Sub

Sub InitpageData
	with frmThis
		.txtPRODUCTIONNAME.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub imgRowAdd_onclick ()
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

Sub imgDelete_onclick()
	Dim intRtn
	
	intRtn = mobjPDCOCFINPUT.DeleteRtn(gstrConfigXml,mstrPREESTNO)
		if not gDoErrorRtn ("DeleteRtn") then
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"�����ȳ�" 
			SelectRtn
		End If
End SUb
'================================================================
'UI
'================================================================

Sub imgCalEndarDATE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtDATE,frmThis.imgCalEndarDATE,"txtDATE_onchange()"
		gSetChange
	end with
End Sub
Sub txtDATE_onchange
	gSetChange
End Sub

Sub imgCalEndarMEETINGDATE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtMEETINGDATE,frmThis.imgCalEndarMEETINGDATE,"txtMEETINGDATE_onchange()"
		gSetChange
	end with
End Sub
Sub txtMEETINGDATE_onchange
	gSetChange
End Sub


Sub imgCalEndarSHOOTDATE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtSHOOTDATE,frmThis.imgCalEndarSHOOTDATE,"txtDATE_onchange()"
		gSetChange
	end with
End Sub
Sub txtSHOOTDATE_onchange
	gSetChange
End Sub

Sub EndPage
	Set mobjPDCOCFINPUT = Nothing
	gEndPage
End Sub


'=============================================================
'��ȸ
'=============================================================

Sub SelectRtn
	Dim vntData
	'on error resume next

	'�ʱ�ȭ

	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCOCFINPUT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
	If not gDoErrorRtn ("SelectRtn") Then
			'��� �÷��� Ŭ����
		If mlngRowCnt > 0  Then
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
		End If		
	End If
		
End Sub

Sub processRtn
	Dim intRtn
	Dim intCnt 
	
	Dim strDATE
	Dim strMEETINGDATE
	Dim strSHOOTDATE
	
	with frmThis
		
		strMasterData = gXMLGetBindingData (xmlBind)
		'��Ʈ�� ����� �����͸� �����´�.
		
		If  Not gXMLIsDataChanged (xmlBind) Then
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		strDATE = MID(.txtDATE.value,1,4) & MID(.txtDATE.value,6,2) & MID(.txtDATE.value,9,2)
		strMEETINGDATE = MID(.txtMEETINGDATE.value,1,4) & MID(.txtMEETINGDATE.value,6,2) & MID(.txtMEETINGDATE.value,9,2)
		strSHOOTDATE = MID(.txtSHOOTDATE.value,1,4) & MID(.txtSHOOTDATE.value,6,2) & MID(.txtSHOOTDATE.value,9,2)
		
		intRtn = mobjPDCOCFINPUT.ProcessRtn(gstrConfigXml,strMasterData,Trim(.txtPREESTNO.value),strDATE,strMEETINGDATE,strSHOOTDATE)
		

		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
		End If
	End with
End Sub


		</script>
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="68" background="../../../images/back_p.gIF"
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
								<td class="TITLE">CF���ֳ���</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="SEARCHDATA" width="100%">
				<tr>
					<td class="SEARCHDATA" colSpan="7">&nbsp;CLIENT <INPUT class="NOINPUTB_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 224px; HEIGHT: 20px"
							accessKey=",NUM" type="text" maxLength="10" size="24" name="txtCLIENTNAME" dataFld="CLIENTNAME" dataSrc="#xmlBind">&nbsp;PRODUCT
						<INPUT class="INPUT_L" id="txtPRODUCT" title="���۰Ǹ�" style="WIDTH: 224px; HEIGHT: 20px"
							accessKey=",NUM" type="text" maxLength="15" size="26" name="txtPRODUCT" dataFld="PRODUCT"
							dataSrc="#xmlBind"> &nbsp;PROJECT <INPUT class="INPUT_L" id="txtPROJECT" title="������Ʈ��" style="WIDTH: 224px; HEIGHT: 20px"
							type="text" maxLength="255" size="32" name="txtPROJECT" dataFld="PROJECT" dataSrc="#xmlBind">&nbsp;<INPUT dataFld="PREESTNO" class="INPUT" id="txtPREESTNO" title="������" style="WIDTH: 16px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="1" name="txtPREESTNO"><INPUT dataFld="SAVEFLAG" class="INPUT" id="txtSAVEFLAG" title="���屸��" style="WIDTH: 16px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="1" name="txtSAVEFLAG"></td>
					<td align="right" width="54"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="ȭ���� �ݽ��ϴ�."
							src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
				</tr>
			</table>
			</TABLE>
			<BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">���λ���&nbsp;
					</td>
					<TD align="right" width="600"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
							onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
							align="absMiddle" border="0" name="imgSave">&nbsp;<IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF"
							border="0" name="imgDelete" align="absMiddle">&nbsp;
					</TD>
				</tr>
			</table>
			<table class="SEARCHDATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100">PRODUCTION</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="PRODUCTIONNAME" class="INPUT_L" id="txtPRODUCTIONNAME" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtPRODUCTIONNAME"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >DATE</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="DATE" class="INPUT" id="txtDATE" title="����������" style="WIDTH: 136px; HEIGHT: 22px"
							accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="17" name="txtDATE"><IMG id="imgCalEndarDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
							name="imgCalEndarDATE"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >DIRECTOR</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="DIRECTORNAME" class="INPUT_L" id="txtDIRECTORNAME" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtDIRECTORNAME"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >PRE-PRODUCTION 
						MEETING DATE</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="MEETINGDATE" class="INPUT" id="txtMEETINGDATE" title="����������" style="WIDTH: 136px; HEIGHT: 22px"
							accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="17" name="txtMEETINGDATE"><IMG id="imgCalEndarMEETINGDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndarMEETINGDATE"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >EDIT</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="EDIT" class="INPUT_L" id="txtEDIT" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtEDIT"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >SHOOT 
						DATE</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="SHOOTDATE" class="INPUT" id="txtSHOOTDATE" title="����������" style="WIDTH: 136px; HEIGHT: 22px"
							accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="17" name="txtSHOOTDATE"><IMG id="imgCalEndarSHOOTDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
							name="imgCalEndarSHOOTDATE"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >CG(2D,3D)</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="CG" class="INPUT_L" id="txtCG" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtCG"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >STAGE 
						SHOOT DAY</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="DAYS" class="INPUT" id="txtDAYS" title="����������" style="WIDTH: 56px; HEIGHT: 22px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="4" name="txtSDAYS">&nbsp;DAYS
						<INPUT dataFld="HOURS" class="INPUT" id="txtHOURS" title="����������" style="WIDTH: 40px; HEIGHT: 22px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="1" name="txtHOURS">&nbsp;HOURS</TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >TELECINE</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="TELECINE" class="INPUT_L" id="txtTELECINE" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtTELECINE"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >SPOT 
						TITLES</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="TITLE" class="INPUT_L" id="txtTITLE" title="�����ڵ�" style="WIDTH: 248px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="36" name="txtTITLE"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >RECORDING</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="RECORDING" class="INPUT_L" id="txtRECORDING" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtRECORDING"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" >LENGTHS</TD>
					<TD class="SEARCHDATA"><INPUT dataFld="LENGTHS" class="INPUT_L" id="txtLENGTHS" title="�����ڵ�" style="WIDTH: 248px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="36" name="txtLENGTHS"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >CM-SONG</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="CMSONG" class="INPUT_L" id="txtCMSONG" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtCMSONG"></TD>
					<TD class="SEARCHLABEL" style="WIDTH: 212px; CURSOR: hand" width="212" rowSpan="3" >COMMENTS</TD>
					<TD class="SEARCHDATA" rowSpan="3"><TEXTAREA dataFld="COMMENTS" id="txtCOMMENT" style="WIDTH: 443px; HEIGHT: 70px" dataSrc="#xmlBind"
							name="txtCOMMENT" rows="249" wrap="hard" cols="53"></TEXTAREA></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >STUDIO</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="STUDIO" class="INPUT_L" id="txtSTUDIO" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtSTUDIO"></TD>
				</TR>
				<TR>
					<!--����-->
					<TD class="SEARCHLABEL" style="WIDTH: 150px; CURSOR: hand" width="100" >MODELAGENCY</TD>
					<TD class="SEARCHDATA" width="200"><INPUT dataFld="MODELAGENCY" class="INPUT_L" id="txtMODELAGENCY" title="�����ڵ�" style="WIDTH: 228px; HEIGHT: 22px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="32" name="txtMODELAGENCY"></TD>
				</TR>
			</table>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 101.05%"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 101.05%"><FONT face="����"></FONT></TD>
				</TR>
			</table>
		</form>
	</body>
</HTML>
