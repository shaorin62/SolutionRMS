<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECCOMMIALL.aspx.vb" Inherits="MD.MDCMELECCOMMIALL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ������ �ŷ���/���ݰ�꼭 ��ü����</title>
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECCOMMIALL, mobjMDCMGET
Dim mstrCheck
mstrCheck=True
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
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgNew_onclick
	InitPageData
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht1
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------
Sub imgCalDemandday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û�����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' �˻����� ����� MON ������ �����ֱ� ����
'-----------------------------------------------------------------------------------------
Sub txtYEARMON_onblur
	With frmThis
		If .txtYEARMON.value <> "" AND Len(.txtYEARMON.value) = 6 Then DateClean
	End With
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
sub sprSht1_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		end if
	end with
end sub

Sub sprSht1_Change(ByVal Col, ByVal Row)
	Dim intAMT,intADJAMT,intBALANCE,intCalCul	
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col, Row  
End Sub

Sub sprSht1_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCol
	Dim strColFlag
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
		strCOLUMN = ""
		strColFlag = 0
		If .sprSht1.MaxRows >0 Then
			If .sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht1,"AMT") or .sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht1,"SUSU") OR _
				.sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht1,"SUSUVAT") OR .sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht1,"SUMSUSUVAT") Then
				If .sprSht1.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt1)
					
					FOR i = 0 TO intSelCnt -1
						If vntData_col(i) <> "" Then
							strColFlag = strColFlag + 1
							strCol = vntData_col(i)
						End If 
					Next
					
					If strColFlag <> 1 Then 
						.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht1,strCol,vntData_row(j))
						End If
					Next
					
					.txtSELECTAMT.value = strSUM
				End If
				
			else
				.txtSELECTAMT.value = 0
			End If
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	'����������ü ����	
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECCOMMIALL	= gCreateRemoteObject("cMDET.ccMDETELECCOMMIALL")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "125px"
	pnlTab2.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	    
    '*********************************
    '�������Ʈ
    '*********************************
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 18, 0
		mobjSCGLSpr.SpreadDataField .sprSht1,   "REAL_MED_NAME | CLIENTNAME | INPUT_MEDFLAG | INPUT_MEDNAME | AMT |SUSURATE| SUSU | SUSUVAT | SUMSUSUVAT | CLIENTCODE | MEDCODE | REAL_MED_CODE | DEPTCD | TRANSRANK | CLIENTBISNO | REAL_MED_BISNO | YEARMON | ATTR05"
		mobjSCGLSpr.SetHeader .sprSht1,		   "��ü��|������|INPUT_MEDFLAG|��ü����|����ݾ�|��������|������|�ΰ���|��|CLIENTCODE|MEDCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|CLIENTBISNO|REAL_MED_BISNO|YEARMON|����"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "   34|    34|           10|      12|      15|      12|    15|     0| 0|         0|      0|            0|     0|        0|          0|             0|      0|   8"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, " CLIENTNAME | REAL_MED_NAME| ATTR05", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "AMT | SUSU | SUSUVAT | SUMSUSUVAT | SUSURATE | INPUT_MEDFLAG", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT | SUSU | SUSUVAT | SUMSUSUVAT", -1, -1, 0
		
		mobjSCGLSpr.ColHidden .sprSht1, "CLIENTCODE | MEDCODE | REAL_MED_CODE | DEPTCD | TRANSRANK | SUSUVAT | SUMSUSUVAT | CLIENTBISNO | REAL_MED_BISNO | YEARMON", true
		
	
    End With
    
	pnlTab2.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	DateClean
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECCOMMIALL = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
	.txtYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
	'.txtDEMANDDAY.value = gNowDate
	.txtPRINTDAY.value  = gNowDate
	.sprSht1.MaxRows = 0

	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON.value,1,4) & "-" & MID(frmThis.txtYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData, vntData1
	Dim strMasterData
	Dim strTRANSYEARMON, strCOMMIYEARMON
	Dim intTRANSNO, intCOMMINO
	Dim intRANKTRANS
	Dim intCnt,bsdiv, bsdiv1
	Dim intColFlag, intColFlag1
	Dim strDESCRIPTION
	with frmThis
	strDESCRIPTION = ""
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtDEMANDDAY.value = "" Then
			msgbox "û������ �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If
		
		If .txtPRINTDAY.value = "" Then
			msgbox "�������� �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If

		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'�׷� �ִ밪 ����
		intColFlag = 0
		
		For intCnt = 1 To .sprSht1.MaxRows
		'�ִ밪
			bsdiv1 = cint(mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSRANK",intCnt))
			IF intColFlag1 < bsdiv1 THEN
				intColFlag1 = bsdiv1
			END IF
		Next
		
   		'������ Validation
   		If .sprSht1.MaxRows = 0 Then
   			msgbox "���׸� �� �����ϴ�."
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData1 = mobjSCGLSpr.GetDataRows(.sprSht1,"REAL_MED_NAME | CLIENTNAME | INPUT_MEDFLAG | INPUT_MEDNAME | AMT | SUSU | SUSUVAT | SUMSUSUVAT | SUSURATE  | CLIENTCODE | MEDCODE | REAL_MED_CODE | DEPTCD | TRANSRANK | CLIENTBISNO | REAL_MED_BISNO | YEARMON")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intCOMMINO = 0
		strCOMMIYEARMON = .txtYEARMON.value
		
		intRtn = mobjMDCMELECCOMMIALL.ProcessRtn(gstrConfigXml,strMasterData,vntData1, intCOMMINO,strCOMMIYEARMON,intColFlag1,strDESCRIPTION)

		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			InitPageData
			If strDESCRIPTION <> ""  Then
			gErrorMsgBox "�������ȣ ��Ź�ڷ� ��Ī�� �����Ͽ����ϴ�.","����ȳ�!"
			Else
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			End If
   		end if
   	end with
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	'On error resume next
	with frmThis
'		intColSum = 0
' 		for intCnt = 1 to .sprSht1.MaxRows
'			if mobjSCGLSpr.GetTextBinding(.sprSht1,"CHK",intCnt) = 1  Then 
'				intColSum = intColSum + 1
'			End if
'		next
'		If intColSum = 0 Then exit Function
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData1
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngsusuColCnt, IngsusuRowCnt
   
	'On error resume next
	with frmThis
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		'Sheet�ʱ�ȭ
		.sprSht1.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		IngsusuColCnt=clng(0)
		IngsusuRowCnt=clng(0)
		
		strYEARMON	= .txtYEARMON.value
		
		vntData1 = mobjMDCMELECCOMMIALL.SelectRtn_SUSU(gstrConfigXml,IngsusuRowCnt,IngsusuColCnt,strYEARMON)
		
		if not gDoErrorRtn ("SelectRtn") then
			if IngsusuRowCnt > 0 then
				mobjSCGLSpr.SetClipbinding .sprSht1, vntData1, 1, 1, IngsusuColCnt, IngsusuRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
				gWriteText lblStatus, "������ " & IngsusuRowCnt & " �� �� �ڷᰡ �˻�" & mePROC_DONE
				
				PreSearchFiledValue strYEARMON
   				AMT_SUM
   			else
   				gErrorMsgBox strYEARMON & "���� Ȯ���� �����Ͱ� ���ų�, ������ŷ����� ������ �Ϸ�Ǿ����ϴ�." & vbcrlf & "�Ϸ᳻���� ������ŷ����� ��ȸ ���α׷����� ��ȸ�Ͻʽÿ�.","��ȸ�ȳ�!"
   				gWriteText lblStatus, "������ " & IngsusuRowCnt & " �� �� �ڷᰡ �˻�" & mePROC_DONE
   				'InitPageData
   				PreSearchFiledValue strYEARMON
   			end if
   		end if
   	end with
End Sub

Sub PreSearchFiledValue (strYEARMON)
	frmThis.txtYEARMON.value = strYEARMON
End Sub


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht1.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht1.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="183" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�ڹ��� ������ �ŷ����� ����</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 332px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<TABLE height="95%" id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE class="SEARCHDATA" id="tblDATA1" style="WIDTH: 100%" cellSpacing="1" cellPadding="0"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" width="90" title="�����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">���</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 89px; HEIGHT: 22px" type="text"
														maxLength="6" size="9" name="txtYEARMON" accessKey="NUM" onchange="vbscript:Call gYearmonCheck(txtYEARMON)"></TD>
												<TD class="SEARCHLABEL" width="90" title="�����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')">û������</TD>
												<TD class="SEARCHDATA" width="200"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="û������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday"></TD>
												<TD class="SEARCHLABEL" width="90" title="�����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')">��������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="12" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalPrintday">
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" border="0" align="right" name="imgQuery"></td>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 25px"></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%">
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" width="400" height="20">
										<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td class="TITLE">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
													<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/ImgTransCreOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTransCre.gIF'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/ImgTransCre.gIF" border="0" name="imgSave"></td>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31829">
									<PARAM NAME="_ExtentY" VALUE="13282">
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
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
					</TR>
				</TBODY>
			</TABLE>
			</TD></TR></TBODY></TABLE>
		</FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
