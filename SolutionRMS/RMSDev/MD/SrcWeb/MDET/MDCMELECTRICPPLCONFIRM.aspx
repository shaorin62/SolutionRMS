<%@ Page CodeBehind="MDCMELECTRICPPLCONFIRM.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRICPPLCONFIRM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����/�������� ����ȭ��</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ó�� ǥ�� ����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
'****************************************************************************************
-->
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<!-- StyleSheet ���� --><LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script id="clientEventHandlersVBS" language="vbscript">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOGET
Dim mobjMDETELECTRICPPLLIST 
Dim mstrCheck
Dim mstrConfirmGBN

mstrConfirmGBN = "Y"
mstrCheck = True

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
	gFlowWait meWAIT_ON
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�",""
		exit Sub
	end if
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'���ι�ư Ŭ��
Sub imgAgree_onclick
	gFlowWait meWAIT_ON
		mstrConfirmGBN = "Y"
		ProcessRtn_Confirm(mstrConfirmGBN)
	gFlowWait meWAIT_OFF
End Sub

'������� ��ư Ŭ��
Sub imgAgreeCancel_onclick
	gFlowWait meWAIT_ON
		mstrConfirmGBN = "N"
		ProcessRtn_Confirm(mstrConfirmGBN)
	gFlowWait meWAIT_OFF
End Sub

'������ư Ŭ��
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'�ݱ��ư Ŭ��
Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
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

'���� ����� ��Ʈ�� �˾� 
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'�ڵ�� ǥ��
			gSetChangeFlag .txtEXCLIENTCODE
		End If
	end With
End Sub


Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)

			vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code�� ����
					.txtEXCLIENTNAME.value = trim(vntData(2,1))	'�ڵ�� ǥ��
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub



Sub rdT_onclick
	rdChecked
	SelectRtn
End Sub

Sub rdF_onclick
	rdChecked
	SelectRtn
End Sub

Sub rdChecked
	with frmThis
		If .rdT.checked = True Then
			.imgAgreeCanCel.style.display = "none"
			.imgAgree.style.display = "inline"
		Else
			.imgAgree.style.display = "none"
			.imgAgreeCanCel.style.display = "inline"
		End If
	End with
End sub
'****************************************************************************************
' ��Ʈ ����Ŭ�� �̺�Ʈ
'****************************************************************************************

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
	
		If Row = 0 and Col = 1  then 'AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "N"
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) <> "" Then
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
				
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
				
				for intcnt = 1 to .sprSht.MaxRows
					sprSht_Change 1, intcnt
					
				next
				For intCnt = 1 To .sprSht.MaxRows
					If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
						'����ƽ
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					'Else
						'üũ
					'	mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
					End If			
				Next
			End IF
		end if
	end with
End Sub  	


sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	
	'����������ü ����	
	set mobjMDETELECTRICPPLLIST = gCreateRemoteObject("cMDET.ccMDETELECTRICPPLLIST")
	set mobjMDCOGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
       gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 26, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO "
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|����|�������ڵ�|�����ָ�|���μ��ڵ�|���μ���|ä��|���α׷�|����|û���۽�����|û����������|��Ƚ��|���Ƚ��|�ܿ�Ƚ��|��ü��ܰ�|���Ѹ�ü��|���Ѽ�����|��Ʈ���ڵ�|��Ʈ�ʸ�|��Ʈ��ȸ���ü����|���Ѹ�üû����|��ǥ��ȣ|������|������|���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   0|   4|         8|      18|           0|        10|	8|      15|   5|            12|            12|     7|       7|       7|        10|        10|        10|         8|      15|                15|            13|       0|     0|     0|  20"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "TBRDDAY", -1, -1, "��" & vbTab & "ȭ" & vbTab & "��" & vbTab & "��" & vbTab & "��" & vbTab & "��" & vbTab & "��"  , 10, 40, False, False
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | CNT_AMT | EXSUSU ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDFDATE | TBRDTDATE | CONFIRM_DATE ", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, " CLIENTCODE | CLIENTNAME | DEPT_CD | MEDNAME | PROGRAM | EXCLIENTCODE | EXCLIENTNAME | VOCHNO | CONFIRM_USER ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE |  EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO "
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | DEPT_CD | VOCHNO | EXCLIENTCODE |  CONFIRM_USER | CONFIRM_DATE", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | YEARMON | SEQ | PROGRAM ",-1,-1,2,2,False  '���
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEMO",-1,-1,0,2,false
		.sprSht.style.visibility = "visible"
    End With

	InitPageData
	'SelectRtn	
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	set mobjMDETELECTRICPPLLIST = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	with frmThis
		.txtYEARMON.value =  Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
		
		.imgAgreeCanCel.style.display = "none"
		.imgAgree.style.display = "inline"
			
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	dim vntData
	Dim strYEARMON, strCLIENTCODE, strEXCLIENTCODE
   	Dim intCnt, strGBN, strEMPNO
	on error resume next
	with frmThis
		
		.sprSht.MaxRows = 0
		
		strYEARMON			= .txtYEARMON.value
		strCLIENTCODE		= .txtCLIENTCODE.value
		strEXCLIENTCODE		= .txtEXCLIENTCODE.value
		strEMPNO			= gstrEmpNo
		
		if .rdF.checked = TRUE then
			strGBN = "Y"
		ELSE
			strGBN = "N"
		end if
		

		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		vntData = mobjMDETELECTRICPPLLIST.SelectRtn_confirm(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE , strEXCLIENTCODE,strGBN,strEMPNO)
		
		IF not gDoErrorRtn ("SelectRtn_confirm") then
			'��ȸ�� �����͸� ���ε�
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				
			Next
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If 	mlngRowCnt < 1 Then
			.sprSht.MaxRows= 0
			End If
		End IF
		gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE	
		
	end with
End Sub

'------------------------------------------
' ����/��� �������
'------------------------------------------
Sub ProcessRtn_Confirm(strCONFIRMFLAG)
	Dim intRtn, intRtnChk
   	Dim vntData
   	Dim lngCHK , intCnt
	
	with frmThis
		'On error resume next
   		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = lngCHK + 1
			END IF
		Next
		
		If lngCHK = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		IF strCONFIRMFLAG = "Y" THEN
			intRtnChk = gYesNoMsgbox("�����Ͻ� �ڷḦ ���� �Ͻðڽ��ϱ�?","���ξȳ�")
			If intRtnChk <> vbYes then 
				exit sub
			End If
		ELSE 
			intRtnChk = gYesNoMsgbox("�����Ͻ� �ڷḦ ������� �Ͻðڽ��ϱ�?" & vbcrlf & "��������Ͻø� �����Ͱ� �ݷ��˴ϴ�.","���� ��� �ȳ�")
			If intRtnChk <> vbYes then 
				exit sub
			End If
		END IF 
	
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | DEPT_CD | DEPT_NAME | MEDNAME | PROGRAM | TBRDDAY | TBRDFDATE | TBRDTDATE | TOT_CNT | CNT | CHARGE_CNT | PRICE | AMT | COMMISSION | EXCLIENTCODE | EXCLIENTNAME | CNT_AMT | EXSUSU | VOCHNO | CONFIRM_USER | CONFIRM_DATE | MEMO")
		
		intRtn = mobjMDETELECTRICPPLLIST.ProcessRtn_ConfirmOK(gstrConfigXml,vntData,strCONFIRMFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOK") then 
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			If strCONFIRMFLAG = "Y" Then
				msgbox lngCHK & " ���� �ڷᰡ ����" & mePROC_DONE
			else
				msgbox lngCHK & " ���� �ڷᰡ ���� ���" & mePROC_DONE
			end if
			
			SelectRtn
   		end if
   	end with
End Sub


-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE style="WIDTH: 100%; HEIGHT: 98%" id="tblForm" border="0" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<TD height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="160" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">������ ���������� ����</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--���̺��� �������°��� �����ش�-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%">
							<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 75px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')">�� 
												��</TD>
											<TD style="WIDTH: 87px" class="SEARCHDATA"><INPUT accessKey="NUM" id="txtYEARMON" class="INPUT" title="�����ȸ" maxLength="6" size="10"
													name="txtYEARMON"></TD>
											<TD style="WIDTH: 51px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">������
											</TD>
											<TD style="WIDTH: 239px" class="SEARCHDATA" width="239"><INPUT style="WIDTH: 150px; HEIGHT: 22px" id="txtCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="�����ָ�" dataFld="CLIENTNAME" size="32" name="txtCLIENTNAME"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 64px; HEIGHT: 22px" id="txtCLIENTCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="�������ڵ�" dataFld="CLIENTCODE" size="5" name="txtCLIENTCODE"></TD>
											<TD style="HEIGHT: 22px; CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)"
												width="70">��Ʈ��</TD>
											<TD class="DATA"><INPUT style="WIDTH: 150px; HEIGHT: 22px" id="txtEXCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="���ۻ��" dataFld="EXCLIENTNAME" maxLength="100" size="30" name="txtEXCLIENTNAME">
												<IMG style="CURSOR: hand" id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgEXCLIENTCODE"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT style="WIDTH: 55px; HEIGHT: 22px" id="txtEXCLIENTCODE" dataSrc="#xmlBind" class="INPUT_L"
													title="���ۻ��ڵ�" dataFld="EXCLIENTCODE" maxLength="10" size="4" name="txtEXCLIENTCODE"></TD>
											<TD class="SEARCHDATA"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ �˻��մϴ�."
													align="right" src="../../../images/imgQuery.gIF" height="20"></TD>
										</TR>
										<tr>
											<TD style="CURSOR: hand" class="SEARCHLABEL" title="�ڷḦ ���� �ϰų� ����մϴ�." width="75">�۾�����</TD>
											<TD class="SEARCHDATA" colSpan="6">&nbsp;<INPUT id="rdT" title="��û������ȸ" value="rdT" CHECKED type="radio" name="rdGBN">&nbsp;��û���� 
												��ȸ <INPUT id="rdF" title="���γ�����ȸ" value="rdF" type="radio" name="rdGBN">&nbsp;���γ�����ȸ&nbsp;</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 25px" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%" class="BODYSPLIT">
						<!--�׽�Ʈ ����-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" vAlign="middle" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton1" border="0" cellSpacing="0" cellPadding="2"
										width="50">
										<TR>
											<TD><IMG style="CURSOR: hand" id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'" border="0" name="imgAgree"
													alt="������ ���� �����մϴ�." align="absMiddle" src="../../../images/imgAgree.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'" border="0"
													name="imgAgreeCanCel" alt="������ ���� ������� �մϴ�." align="absMiddle" src="../../../images/imgAgreeCanCel.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!-- �߰� �����γ�--></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 1040px; HEIGHT: 3px" class="BODYSPLIT"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" class="LISTFRAME" vAlign="top" align="center">
						<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%" id="pnlTab1" ms_positioning="GridLayout">
							<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31802">
								<PARAM NAME="_ExtentY" VALUE="11853">
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
				<TR>
					<TD id="lblStatus" class="BOTTOMSPLIT"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
