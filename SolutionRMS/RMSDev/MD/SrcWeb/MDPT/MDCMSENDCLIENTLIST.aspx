<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMSENDCLIENTLIST.aspx.vb" Inherits="MD.MDCMSENDCLIENTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�μ�����Ź ���ݰ�꼭 �뷮 ���ε�� ��ȸ</title>
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
'�������� ����
Dim mobjMDCOSENDTRUTAX
Dim mlngRowCnt,mlngColCnt
Dim mstrCLIENTCODE

CONST meTAB = 9
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
'��ȸ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' �ݱ�
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub GetCLIENTLIST()
	Dim vntData
   	Dim i, strCols
   	Dim strHTML
   	Dim intCNT
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strHTML = "" 
		intCNT = 1 
		vntData = mobjMDCOSENDTRUTAX.GetCLIENTLIST(gstrConfigXml,mlngRowCnt,mlngColCnt)
		if not gDoErrorRtn ("GetCLIENTLIST") then
			If mlngRowCnt > 0 Then		
				For i = 0 to mlngRowCnt-1
					IF intCNT = 1 then
						strHTML = strHTML & "<INPUT id='chkCLIENTLIST"& i&"' type='radio' name='chkCLIENTLIST' checked value='"& vntData(0,i) & "' onclick='vbscript:Call SelectRtn()'>" & vntData(1,i) & "&nbsp;&nbsp;"
						intCNT = 2
					else
						strHTML = strHTML & "<INPUT id='chkCLIENTLIST"& i&"' type='radio' name='chkCLIENTLIST' value='"& vntData(0,i) & "' onclick='vbscript:Call SelectRtn()'>" & vntData(1,i) & "&nbsp;&nbsp;"
					End if
				next
			Else
				strHTML = ""
			End If
			
			document.getElementById("tdCLIENTLIST").innerHTML = strHTML
   		end if
   	end with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjMDCOSENDTRUTAX		= gCreateRemoteObject("cMDCO.ccMDCOSENDTRUTAX")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
     gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 0, 0, 0, 0,5
		
		pnlTab1.style.visibility = "visible" 
		
		'ȭ�� �ʱⰪ ����
		InitPageData
    End With
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis

	'�ʱ� ������ ����
	With frmThis
		.sprSht.MaxRows = 0
		
		.txtTAXYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		
		Call GetCLIENTLIST ()
	End With
End Sub

Sub EndPage()
	set mobjMDCOSENDTRUTAX = Nothing
	gEndPage	
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "MON"
		mobjSCGLSpr.SetHeader .sprSht,		 "MON"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MON", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MON"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MON",-1,-1,2,2,false
	End With
End Sub

Sub SetChangeLayout () 
	Dim strID
	Dim i
	
	gInitComParams mobjSCGLCtl,"MC"
	
	With frmThis
		mstrCLIENTCODE = ""
		for i=0 to document.getElementsByName("chkCLIENTLIST").length -1
			strID = "chkCLIENTLIST" + cstr(i)
			if document.getElementById(strID).checked then
				mstrCLIENTCODE = document.getElementById(strID).value
				exit for
			end if		
		Next

		Call Grid_init()
		
		if mstrCLIENTCODE = "A00003" or mstrCLIENTCODE = "A00212" Then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 47, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht,    "SEQ | MEDFLAGNAME | TIMNAME | BILLFLAG | MATTERNAME | MEMO | REAL_MED_NAME | REAL_MED_BISNO | DEMANDDAY | AMT | VAT | EMP_NAME | EMP_HP | EMP_EMAIL | BIGO | SUMM_DATE1 | SUMM1 | SUMM_STD1 | SUMM_QTY1 | SUMM_PRICE1 | SUMM_AMT1 | SUMM_VAT1 | SUMM_MEMO1 | SUMM_DATE2 | SUMM2 | SUMM_STD2 | SUMM_QTY2 | SUMM_PRICE2 | SUMM_AMT2 | SUMM_VAT2 | SUMM_MEMO2 | SUMM_DATE3 | SUMM3 | SUMM_STD3 | SUMM_QTY3 | SUMM_PRICE3 | SUMM_AMT3 | SUMM_VAT3 | SUMM_MEMO3 | SUMM_DATE4 | SUMM4 | SUMM_STD4 | SUMM_QTY4 | SUMM_PRICE4 | SUMM_AMT4 | SUMM_VAT4 | SUMM_MEMO4"
			mobjSCGLSpr.SetHeader .sprSht,		    "��ȣ|��ü����|����|��꼭|�����|���|��ü���|����ڹ�ȣ|û����|���ް���|����|�ŷ�ó������̸�|�ŷ�ó������ڵ���|�ŷ�ó������̸���|���|ǰ��1-��������|ǰ��1-ǰ���|ǰ��1-�԰�|����1-����|ǰ��1-�ܰ�|ǰ��1-���ް���|ǰ��1-����|ǰ��1-���|ǰ��2-��������|ǰ��2-ǰ���|ǰ��2-�԰�|����2-����|ǰ��2-�ܰ�|ǰ��2-���ް���|ǰ��2-����|ǰ��2-���|ǰ��3-��������|ǰ��3-ǰ���|ǰ��3-�԰�|����3-����|ǰ��3-�ܰ�|ǰ��3-���ް���|ǰ��3-����|ǰ��3-���|ǰ��4-��������|ǰ��4-ǰ���|ǰ��4-�԰�|����4-����|ǰ��4-�ܰ�|ǰ��4-���ް���|ǰ��4-����|ǰ��4-���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|       6|  15|     6|    18|  18|      18|        13|     8|      10|  10|               8|                10|                15|   8|             8|          25|         6|         6|         6|            10|        10|         8|             8|          25|         6|         6|         6|            10|        10|         8|             8|          25|         6|         6|         6|            10|        10|         8|             8|          25|         6|         6|         6|            10|        10|         8"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT | VAT | SUMM_PRICE1 | SUMM_AMT1 | SUMM_VAT1 | SUMM_PRICE2 | SUMM_AMT2 | SUMM_VAT2 | SUMM_PRICE3 | SUMM_AMT3 | SUMM_VAT3 | SUMM_PRICE4 | SUMM_AMT4 | SUMM_VAT4", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDFLAGNAME | TIMNAME | BILLFLAG | MATTERNAME | MEMO | REAL_MED_NAME | REAL_MED_BISNO | DEMANDDAY | EMP_NAME | EMP_HP | EMP_EMAIL | BIGO | SUMM_DATE1 | SUMM1 | SUMM_STD1 | SUMM_MEMO1 | SUMM_DATE2 | SUMM2 | SUMM_STD2 | SUMM_MEMO2 | SUMM_DATE3 | SUMM3 | SUMM_STD3 | SUMM_MEMO3 | SUMM_DATE4 | SUMM4 | SUMM_STD4 | SUMM_MEMO4", -1, -1, 200
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SEQ | MEDFLAGNAME | TIMNAME | BILLFLAG | MATTERNAME | MEMO | REAL_MED_NAME | REAL_MED_BISNO | DEMANDDAY | AMT | VAT | EMP_NAME | EMP_HP | EMP_EMAIL | BIGO | SUMM_DATE1 | SUMM1 | SUMM_STD1 | SUMM_QTY1 | SUMM_PRICE1 | SUMM_AMT1 | SUMM_VAT1 | SUMM_MEMO1 | SUMM_DATE2 | SUMM2 | SUMM_STD2 | SUMM_QTY2 | SUMM_PRICE2 | SUMM_AMT2 | SUMM_VAT2 | SUMM_MEMO2 | SUMM_DATE3 | SUMM3 | SUMM_STD3 | SUMM_QTY3 | SUMM_PRICE3 | SUMM_AMT3 | SUMM_VAT3 | SUMM_MEMO3 | SUMM_DATE4 | SUMM4 | SUMM_STD4 | SUMM_QTY4 | SUMM_PRICE4 | SUMM_AMT4 | SUMM_VAT4 | SUMM_MEMO4"
		else
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 0, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht,    "SEQ | REAL_MED_BISNO | EMP_EMAIL | TITLE | BILLFLAG | REGDATE | DEMANDDAY | SUMM | STD | QTY | AMT | VAT"
			mobjSCGLSpr.SetHeader .sprSht,		    "����|����Ź ��ü"& vbCrlf &"����� ��ȣ|����Ź"& vbCrlf &"�̸���|����|����(����)"& vbCrlf &"1---���ݰ�꼭"& vbCrlf &"2-------��꼭"& vbCrlf &"3---------�鼼|��������|��������|ǰ��|�԰�|����|�ܰ�|����"
			mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   0|								  15|					   18|  12|																						 12|       8|       8|  18|   4|   4|   9|   9"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "35"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REGDATE | DEMANDDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " SEQ | BILLFLAG | QTY | AMT | VAT ", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "REAL_MED_BISNO | EMP_EMAIL | TITLE | REGDATE | DEMANDDAY | SUMM | STD  ", -1, -1, 200
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SEQ | REAL_MED_BISNO | EMP_EMAIL | TITLE | BILLFLAG | REGDATE | DEMANDDAY | SUMM | STD | QTY | AMT | VAT"
			mobjSCGLSpr.ColHidden .sprSht, "SEQ", True
			mobjSCGLSpr.SetCellAlign2 .sprSht, "REAL_MED_BISNO | REGDATE | DEMANDDAY",-1,-1,2,2,False
			mobjSCGLSpr.SetCellAlign2 .sprSht, "STD",-1,-1,2,2,False
		End if 		
   	End With
End Sub

Sub SelectRtn ()
   	Dim vntData
	Dim strTAXYEARMON
	Dim strMEDFLAG
	Dim strGUBUN
	Dim strCLIENTCODE
	
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		SetChangeLayout

		.sprSht.MaxRows = 0
		
		strTAXYEARMON = .txtTAXYEARMON.value
		strMEDFLAG	  = .cmbMED_FLAG1.value
		strGUBUN	  = "PRINT"
		strCLIENTCODE = mstrCLIENTCODE
		
		if strCLIENTCODE = "A00003" or strCLIENTCODE = "A00212" then
			vntData = mobjMDCOSENDTRUTAX.Get_SENDED_CUST_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strMEDFLAG, strGUBUN,strCLIENTCODE)
		else
			vntData = mobjMDCOSENDTRUTAX.Get_SENDED_CLIENT_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strMEDFLAG, strGUBUN,strCLIENTCODE)
		end if
		
		if not gDoErrorRtn ("SelectRtn") then
		
			mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True

   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub


		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
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
												<TABLE cellSpacing="0" cellPadding="0" width="285" background="../../../images/back_p.gIF"
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
											<td class="TITLE">�μ�����Ź ���ݰ�꼭 �뷮 ���ε�� ��ȸ</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%; HEIGHT: 58px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCODESEARCH,'')">���
											</TD>
											<TD class="SEARCHDATA" width="110"><INPUT class="INPUT" id="txtTAXYEARMON" style="WIDTH: 96px; HEIGHT: 22px" type="text" maxLength="8"
													size="10" name="txtTAXYEARMON"></TD>
											<TD class="SEARCHLABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCODESEARCH,'')">����
											</TD>
											<TD class="SEARCHDATA" width="100"><SELECT id="cmbMED_FLAG1" title="��������" style="WIDTH: 96px" name="cmbMED_FLAG1">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="B">�Ź�</OPTION>
													<OPTION value="C">����</OPTION>
												</SELECT></TD>
											<TD class="SEARCHLABEL" width="640">
											</TD>
											<TD align="right" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20"
													alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="HEIGHT: 24px;">������
											</TD>
											<TD id="tdCLIENTLIST" class="DATA" colspan="5">
											</TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
				</TR>
				<tr>
					<td>
						<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
						</table>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="40" background="../../../images/back_p.gIF"
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
											<td class="TITLE">����Ʈ&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
										</TR>
									</TABLE>
									<!--Common Button End-->
								</TD>
							</TR>
						</TABLE>
					</td>
				</tr>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31803">
								<PARAM NAME="_ExtentY" VALUE="11721">
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
								<PARAM NAME="MaxCols" VALUE="11">
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
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> 
			</TABLE></TR></TABLE></FORM>
	</body>
</HTML>

