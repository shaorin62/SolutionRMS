<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTDBLLIST.aspx.vb" Inherits="MD.MDCMPRINTDBLLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�����ֺ� ��ü�纰 �˻�(����)</title>
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
Dim mobjMDCMGET, mobjEXECUTE'�����ڵ�, Ŭ����
Dim mstrClientcode

Dim mintCnt
Dim mintCnt2
Dim mintCnt3
Dim mvntData3
Dim mstrField
Dim mintCntExist
Dim mstrFieldExist
Dim mvntDataCust
Dim mvntDataMed
Dim mvntDataCustCNT
Dim mvntDataMedCNT

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
' ���� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	if frmThis.txtYEAR.value = "" then
		gErrorMsgBox "�⵵�� �Է��Ͻÿ�",""
		exit Sub
	end if
	Call CLIENTCODE_POP()
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

sub imgPrint_onclick ()
	gFlowWait meWAIT_ON
	mobjSCGLSpr.SSPrint  frmThis.sprSht,window.document.title,"",0,0,0,0, true,false,true, 2
	gFlowWait meWAIT_OFF                              
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
'Sub ImgCLIENTCODE_onclick
'	Call CLIENTCODE_POP()
'End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	mstrClientcode = ""
	InitPage
	With frmThis
		vntInParams = array(trim(.txtYEAR.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMPRINTDBLPOP.aspx",vntInParams , 580,415)
		if vntRet <> "" then
			mstrClientcode = vntRet
			SelectRtn
		end if
	End With
	gSetChange
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	set mobjEXECUTE	= gCreateRemoteObject("cMDCO.ccMDCOEXECUTE")
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "75px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "MON"
											  '       1|
		mobjSCGLSpr.SetHeader .sprSht,		 "����"
											   '  1|
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6"
   												'1|
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MON", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MON"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MON",-1,-1,2,2,false
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjEXECUTE = Nothing
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
		.txtYEAR.value = mid(gNowDate,1,4)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEAR.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'��ȸ
Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
   	Dim intCnt
	Dim strSEQ
	Dim intRtn
	Dim strSPONSOR
	Dim strCOMMIT
	Dim strClientAndMed
	
	With frmThis
		'�����ȸ�� ��� üũ
		If .txtYEAR.value = ""  Then
			gErrorMsgbox "��ȸ����� �����ϼ���","��ȸ�ȳ�"
			Exit Sub
		End If
		'�׸��� ����� 
		SetChangeLayout
		Dim intLayOutCnt
		intLayOutCnt = (mvntDataCustCNT+1) * mvntDataMedCNT
		strClientAndMed = split(mstrClientcode, "��")
		
		'EXIT SUB
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjEXECUTE.SelectRtn_ClientGroup2(gstrConfigXml,mlngRowCnt,mlngColCnt, mvntDataMed,mvntDataCust, mvntDataCustCNT, intLayOutCnt, .txtYEAR.value, strClientAndMed(0), strClientAndMed(1))
		
		If not gDoErrorRtn ("SelectRtn_ClientGroup2") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			'SUMCLEAN
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		End if
   		Layout_change
   	End With
End Sub

Sub SetChangeLayout () 
	Dim strYEAR
	Dim strCLIENTCODE
	Dim intAddCnt,intAddHeadCnt,intAddWith,intFieldSetting,intHide,intFloat,intAddCnt2 'For �� Count����
	Dim vntData
	Dim strAddHead
	Dim lngRowReal
	Dim lngColReal
	Dim strStartHead
	Dim strEndHead
	
	Dim strClientAndMed
	Dim i
	
	mvntDataCustCNT = ""
	mvntDataMedCNT = ""
	mvntDataCust = ""
	mvntDataMed = ""
	gInitComParams mobjSCGLCtl,"MC"
	With frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		lngRowReal=clng(0)
		lngColReal=clng(0)
		
		strClientAndMed = split(mstrClientcode, "��")
		strYEAR = .txtYEAR.value
		
		mvntDataCust = mobjEXECUTE.GetCLIENTCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEAR, strClientAndMed(0))
		
		mvntDataMed = mobjEXECUTE.GetMED_CNT(gstrConfigXml,lngRowReal,lngColReal,strYEAR, strClientAndMed(1))
		mvntDataCustCNT = mlngRowCnt
		mvntDataMedCNT = lngRowReal
		
		If not gDoErrorRtn ("GetCLIENTCNT") then
			If mlngRowCnt > 0 Then 
				'�ʵ� ����������
				Dim strField
				strField = "MON"
				
				'�ʵ� ���������� [�������ڵ�]
				Dim strAddField
				strAddField = ""
				For intAddCnt = 1 To (mvntDataCustCNT+1) * mvntDataMedCNT
					strAddField = strAddField & "|A" & intAddCnt
				Next
				
				'�ʵ� ������ [��]
				mstrField = strField & strAddField & "|SUMAMT"
				
				'��� ����������
				Dim strHead
				strHead = "����"
				'��� ����������
				Dim strHeadCLIENT
				Dim strHeadMED
				Dim lngSUBCNT
				lngSUBCNT =1
				strHeadCLIENT = ""
				strHeadMED = ""
				strStartHead = ""
				strEndHead = ""
				For intAddHeadCnt = 1 To  ((mvntDataCustCNT+1) * mvntDataMedCNT)
					IF mvntDataMedCNT = 1 THEN
						IF intAddHeadCnt = 1 THEN
							strHeadMED = strHeadMED & "|" & TRIM(mvntDataMed(0,1))
						ELSE
							strHeadMED = strHeadMED & "|"
						END IF
					ELSE
						IF intAddHeadCnt MOD (mvntDataCustCNT+1) = 1 THEN 
							strHeadMED = strHeadMED & "|" & TRIM(mvntDataMed(0,lngSUBCNT))
							lngSUBCNT = lngSUBCNT +1
						ELSE 
							strHeadMED = strHeadMED & "|"
						END IF	
					END IF
					
					IF intAddHeadCnt MOD (mvntDataCustCNT+1) = 0 THEN
						strHeadCLIENT   = strHeadCLIENT & "|��" 
					ELSE 
						strHeadCLIENT   = strHeadCLIENT & "|" & TRIM(mvntDataCust(0,intAddHeadCnt MOD (mvntDataCustCNT+1)))
					END IF
				Next
				strStartHead = strHead & strHeadMED & "|��"
				strEndHead =  strHeadCLIENT & "|"
				
				'���� ����������
				Dim strWith
				strWith = "6"
				'���� ����������
				Dim strAddWith
				Dim strEndWith
				strAddWith = ""
				For intAddWith = 1 To (mvntDataCustCNT+1) * mvntDataMedCNT
					strAddWith = strAddWith & "|13"
				Next
				strEndWith = strWith & strAddWith & "|13"
				
				
				'���÷�����
				Dim intLayOutCnt
				intLayOutCnt = 1 + ((mvntDataCustCNT+1)* mvntDataMedCNT) + 1
				'������� ������
				
				gSetSheetColor mobjSCGLSpr, .sprSht
	    
				'Sheet Layout ������
				mobjSCGLSpr.SpreadLayout .sprSht, intLayOutCnt, 0, 0, 0, , 2, 1, , , True
				mobjSCGLSpr.SpreadDataField .sprSht, mstrField 
				mobjSCGLSpr.SetHeader .sprSht,       strStartHead ,0,1,true
				mobjSCGLSpr.SetHeader .sprSht,       strEndHead ,SPREAD_HEADER + 1,1,true
				
				mobjSCGLSpr.AddCellSpan .sprSht, 1, SPREAD_HEADER + 0, 1    , 2      , -1 , true
				mobjSCGLSpr.AddCellSpan .sprSht, 2, SPREAD_HEADER + 0, (mvntDataCustCNT+1)    , 1      , -1 , true
				'                                 20��° ����            ����6���� 1���� 3�������� ������
				mobjSCGLSpr.AddCellSpan .sprSht, intLayOutCnt, SPREAD_HEADER + 0, 1    , 2      , -1 , true
				'                                 ������ Ǯ���°� �� 44��°�̰� 2���� ���Ķ� -1 ��ü
				mobjSCGLSpr.SetColWidth .sprSht, "-1", strEndWith
				mobjSCGLSpr.SetCellTypeFloat2 .sprSht, mstrField, -1, -1, 0
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MON", , , 50, , ,0
				mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
				mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,strField
				mobjSCGLSpr.SetCellAlign2 .sprSht, "MON",-1,-1,2,2,false
			ELSE
				'Sheet �⺻Color ����
				gSetSheetDefaultColor() 
				
				With frmThis
					gSetSheetColor mobjSCGLSpr, .sprSht
					mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
					mobjSCGLSpr.SpreadDataField .sprSht, "MON"
					mobjSCGLSpr.SetHeader .sprSht,		 "MON"
															'  1|
					mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6"
   															'1|
					mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
					mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
					mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MON", -1, -1, 20
					mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MON"
					mobjSCGLSpr.SetCellAlign2 .sprSht, "MON",-1,-1,2,2,false
					
				End With
			End If
   		End if
   	End With
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If mobjSCGLSpr.GetTextBinding(.sprSht,"MON",intCnt) = "��" Then
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		End If
	Next 
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD style="HEIGHT: 54px">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">�����ֺ� ��ü�纰 �˻�(����)</td>
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
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 115px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="115" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center"><FONT face="����">
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TBODY>
													<TR>
														<TD class="SEARCHLABEL" width="98" style="WIDTH: 98px">�⵵</TD>
														<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEAR" title="�ڵ���ȸ" style="WIDTH: 96px; HEIGHT: 22px" type="text"
																maxLength="4" align="left" size="10" name="txtYEAR" accessKey="NUM">
														</TD>
													</TR>
												</TBODY>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 608px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 99.77%; POSITION: relative; HEIGHT: 608px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 1038px; HEIGHT: 608px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27464">
												<PARAM NAME="_ExtentY" VALUE="16087">
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
								<!--List End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--Brench Start-->
								<TR>
									<TD class="BRANCHFRAME" style="WIDTH: 1040px"><FONT face="����" color="#666666" size="3"></FONT>
										<!--<INPUT class="BUTTON" id="btn1" style="WIDTH: 123px; HEIGHT: 16pt" type="button" value="�б��ư"
											name="Button">--></TD>
								</TR>
								<!--Brench End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>