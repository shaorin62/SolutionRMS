<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOORBATCH.aspx.vb" Inherits="MD.MDCMOUTDOORBATCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� û�� ������ ��ȸ �� ����</title>
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
Dim mobjMDCMOUTDOOR, mobjMDCMGET
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
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
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onClick()
	Window_OnUnload
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 Then
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
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjMDCMOUTDOOR	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "70px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 25, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,   "YEARMON|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|TITLE|PROGNAME|TBRDSTDATE|TBRDEDDATE|TOTALAMT|AMT|COMMI_RATE|COMMISSION|MED_GBN|CLIENTCODE|MEDCODE|REAL_MED_CODE|DEPT_CD|SUBSEQ|CLIENTSUBCODE|OUT_AMT|LOCATION|CONTIDX|CYEAR|CMONTH"
		
		mobjSCGLSpr.SetHeader .sprSht,		   "���|������|�����|��ü��|�귣��|����|�����|��������|���������|�ѱ���ݾ�|����ݾ�|��������|������|��������|�������ڵ�|��ü�ڵ�|��ü���ڵ�|�μ��ڵ�|�귣���ڵ�|������ڵ�|���ֺ�|���|CONTIDX|CYEAR|CMONTH"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   0|    13|	13|	   13|    13|    15|    15|         8|         8|        10|      10|       6|    10|      10|         0|       0|         0|       0|         0|         0|     0|  10|      0|    0|     0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|TOTALAMT|COMMISSION|OUT_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|TITLE|PROGNAME|TBRDSTDATE|TBRDEDDATE|TOTALAMT|AMT|COMMI_RATE|COMMISSION|CLIENTCODE|MEDCODE|REAL_MED_CODE|DEPT_CD|SUBSEQ|CLIENTSUBCODE|OUT_AMT|MED_GBN|LOCATION|CONTIDX|CYEAR|CMONTH"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|MEDCODE|REAL_MED_CODE|DEPT_CD|SUBSEQ|CLIENTSUBCODE|CONTIDX|CYEAR|CMONTH", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME|REAL_MED_NAME",-1,-1,0,2,false
		
    
		pnlTab1.style.visibility = "visible"
		
		InitPageData	
		
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : mstrFields = vntInParam(i)
				case 2 : mstrFields = vntInParam(i)			'��ȸ�߰��ʵ�
				case 3 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
		SelectRtn
	End With    
End Sub

Sub EndPage()
	set mobjMDCMOUTDOOR = Nothing
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
		.txtYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		.sprSht.MaxRows = 0	
		
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strYEARMON
	
	with frmThis
		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|TITLE|PROGNAME|TBRDSTDATE|TBRDEDDATE|TOTALAMT|AMT|COMMI_RATE|COMMISSION|MED_GBN|CLIENTCODE|MEDCODE|REAL_MED_CODE|DEPT_CD|SUBSEQ|CLIENTSUBCODE|OUT_AMT|LOCATION|CONTIDX|CYEAR|CMONTH")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
		
		intRtn = mobjMDCMOUTDOOR.ProcessRtn_BATCH(gstrConfigXml,strMasterData,vntData,strYEARMON)
   		
   		if not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOKMsgbox strYEARMON & "���� �����Ͱ� �����Ǿ����ϴ�.", ""
			
			window.returnvalue = strYEARMON
			call Window_OnUnload()
			'SelectRtn
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
' 		for intCnt = 1 to .sprSht.MaxRows
'			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1  Then 
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
	Dim strSPONSOR
   	Dim i, strCols
   
	'On error resume next
	with frmThis
		If .txtYEARMON.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.","��ȸ�ȳ�"
			Exit SUb
		End If

		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON	= .txtYEARMON.value
			
		vntData = mobjMDCMOUTDOOR.SelectRtn_BATCH(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
		
		if not gDoErrorRtn ("SelectRtn") Then
			if mlngRowCnt > 0 Then
				Call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end if
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
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TBODY>
									<TR>
										<TD style="WIDTH: 400px" align="left" width="427" height="28">
											<table cellSpacing="0" cellPadding="0" width="100%" border="0">
												<tr>
													<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
													<td align="left" height="4"><FONT face="����"></FONT></td>
												</tr>
												<tr>
													<td class="TITLE">&nbsp;���� û�� ������ ��ȸ �� ����</td>
												</tr>
											</table>
										</TD>
										<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
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
											<TABLE id="tblButton" style="WIDTH: 153px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
												width="153" border="0">
												<TBODY>
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
																name="imgQuery"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
														<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
																name="imgClose"></TD>
													</TR>
												</TBODY>
											</TABLE>
											<!--Common Button End--></TD>
									</TR>
									<!--Top Define Table End-->
									<!--Input Define Table End--></TBODY></TABLE>
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center"><FONT face="����">
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, txtTRANSNO)"
														width="91">�� ��</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="�ŷ������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="MON"
															type="text" maxLength="6" size="6" name="txtYEARMON"></TD>
												</TR>
											</TABLE>
										</FONT>
									</TD>
								</TR>
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 510px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 508px"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 508px"
									width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27490">
									<PARAM NAME="_ExtentY" VALUE="13441">
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
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
