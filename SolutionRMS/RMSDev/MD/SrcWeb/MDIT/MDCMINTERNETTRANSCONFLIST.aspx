<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNETTRANSCONFLIST.aspx.vb" Inherits="MD.MDCMINTERNETTRANSCONFLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ�����</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ����� ��� ȭ��(MDCMPRINTTRANS1.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ����Ź�ŷ����� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/16 By Kim Tae Yub
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDITINTERNETTRANS, mobjMDCOGET
Dim mstrCheck
Dim mALLCHECK
CONST meTAB = 9
mALLCHECK = TRUE
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
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	mobjSCGLSpr.ExportExcelFile frmThis.sprSht_HDR
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	mobjSCGLSpr.ExportExcelFile frmThis.sprSht_DTL
	gFlowWait meWAIT_OFF
End Sub

Sub imgAgree_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn_CONFIRM
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		elseif Row > 0 AND Col > 1 then
			SelectRtn_DTL Col, Row
		end if
	end with
End Sub  

Sub sprSht_HDR_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
End Sub

sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	'����������ü ����	
	set mobjMDITINTERNETTRANS	= gCreateRemoteObject("cMDIT.ccMDITINTERNETTRANS")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'�ŷ����� ��� �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 13, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "����|����|��꼭|������|��ü����|���ް���|�ΰ�����|�հ�ݾ�|û����|������|�ŷ����|��ȣ|���"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "  4|   4|     6|    18|       8|      12|      11|      12|     9|     9|       8|   5|  21"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY | PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "TRANSNO", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT | VAT | SUMAMTVAT ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | TRANSYEARMON | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | MEMO"
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | MED_FLAGNAME" ,-1,-1,2,2,false
		
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 14, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CLIENTNAME | TIMNAME | SUBSEQNAME | MATTERNAME | MEDNAME | REAL_MED_NAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | DEMANDDAY | PRINTDAY "
		mobjSCGLSpr.SetHeader .sprSht_DTL,		"������|CIC/��|�귣��|�����|��ü��|��ü��|��ü����|���ް���|�ΰ���|��|��������|������|û����|������|������"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 15|	15|	   11|    13|    11|	15|       7|      10|    10|10|       7|    10|     8|     8| 	 8"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY | PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "AMT | VAT | SUMAMTVAT | COMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "CLIENTNAME | TIMNAME | SUBSEQNAME | MATTERNAME | MEDNAME | REAL_MED_NAME | MED_FLAGNAME ", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CLIENTNAME | TIMNAME | SUBSEQNAME | MATTERNAME | MEDNAME | REAL_MED_NAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | COMMI_RATE | COMMISSION | DEMANDDAY | PRINTDAY " 
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "MED_FLAGNAME",-1,-1,2,2,false

		.sprSht_HDR.style.visibility = "visible"
		.sprSht_DTL.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDITINTERNETTRANS = Nothing
	set mobjMDCOGET = Nothing
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
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
	End with
	SelectRtn
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON, strDEMANDYEARMON
	Dim strCLIENTCODE, strTIMCODE
   	Dim strMED_FLAG
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjMDITINTERNETTRANS.SelectRtn_TransList(gstrConfigXml,mlngRowCnt,mlngColCnt)

		If not gDoErrorRtn ("SelectRtn_TransList") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				CALL SelectRtn_DTL(1,1)
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TRANSNO", -1, -1, 0
		vntData = mobjMDITINTERNETTRANS.SelectRtn_TransList_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
															 strTRANSYEARMON, strTRANSNO)
													
		If not gDoErrorRtn ("SelectRtn_TransList_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   		End If
   	end with
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn_CONFIRM ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intCnt
	Dim chkcnt
	Dim strCONFIRMGBN
	chkcnt = 0
	
	with frmThis
		For intCnt = 1 To .sprSht_HDR.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if
		
		'���� ����
		strCONFIRMGBN = "1"

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO")
		
		intRtn = mobjMDITINTERNETTRANS.ProcessRtn_Confirm(gstrConfigXml, vntData, strCONFIRMGBN)
   		
   		if not gDoErrorRtn ("ProcessRtn_CIC") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
			gOkMsgBox "������ �ŷ������� ���εǾ����ϴ�.","Ȯ��"
			selectRtn
   		end if
   	end with
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
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="162" background="../../../images/back_p.gIF"
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
												<td class="TITLE">û����� - �ŷ����� ����</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE id="tblBody" height="93%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
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
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gif'"
																	alt="CIC���� �ŷ������� �����մϴ�." src="../../../images/imgAgree.gif" border="0" name="imgAgree"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="4498">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="22"></TD>
												<TD vAlign="middle" align="right" height="22">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="9208">
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
									<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
								</TR>
								<TR>
									<TD></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>
