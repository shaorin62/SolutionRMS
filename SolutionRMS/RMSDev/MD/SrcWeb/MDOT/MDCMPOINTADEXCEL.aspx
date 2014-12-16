<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPOINTADEXCEL.aspx.vb" Inherits="MD.MDCMPOINTADEXCEL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����Ʈ ģ�� AD ���� ���ε�</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMPOINTADEXCEL.aspx
'��      �� : �ŷ����� ������ ���� ���� ���ε� 
'�Ķ�  ���� : 
'Ư��  ���� : ���� ������ ���ε� �Ͽ� POINT AD ���α׷��� �����͸� �����Ѵ�.
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2012/08/01 By OH Se Hoon
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script id="clientEventHandlersVBS" language="vbscript">
		
Dim mobjMDOTPOINTADCOMMI
Dim mstrTRANSYEARMON
Dim mstrTRANSNO
Dim mCAMPAIGN_CODE

Dim mlngRowCnt, mlngColCnt
'������ �ʵ带 ������� ��������
Dim sprSht_DataFields
Dim sprSht_DisplayFields

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
'�˾� �ݱ� ��ư 
Sub imgClose_onclick()
	EndPage
End Sub

'�ʱ�ȭ ��ư Ŭ��
Sub imgFind_onclick
	gFlowWait meWAIT_ON
	EXCEL_UPLOAD
	gFlowWait meWAIT_OFF
End sub

'��ȸ��ư 
Sub imgQuery_onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�����ư Ŭ��
Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'���� ��� ��ư
Sub imgExcel_onclick ()
	with frmThis
		gFlowWait meWAIT_ON
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
		gFlowWait meWAIT_OFF
	end with
End Sub

'��Ʈ ���� 
Sub sprSht_Change(ByVal Col, ByVal Row)
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
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
	
	set mobjMDOTPOINTADCOMMI  = gCreateRemoteObject("cMDOT.ccMDOTPOINTADCOMMI")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    gSetSheetDefaultColor
    with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)

		'�⺻�� ����
		for i = 0 to intNo
			select case i
				case 0 : mstrTRANSYEARMON = vntInParam(i)	
				case 1 : mstrTRANSNO = vntInParam(i)
				case 2 : mCAMPAIGN_CODE = vntInParam(i)
			end select
		next
		
		gSetSheetDefaultColor
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* ���� ���ε带 ���� �ʱ�ȭ ��ư�� ���� �ֽñ� �ٶ��ϴ�.."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "110"
		
		.txtTRANSYEARMON.value = mstrTRANSYEARMON
		.txtTRANSNO.value = mstrTRANSNO
		.txtCAMPAIGN_CODE.value = mCAMPAIGN_CODE
		
		if mstrTRANSYEARMON = "" or mstrTRANSNO = "" or mCAMPAIGN_CODE = "" then
			gErrorMsgBox "�󼼳����� Ȯ���ϴµ� �ʿ��� ������ ������� �ʽ��ϴ�. �����ڿ��� �����ϼ���.","�󼼳��� ����!"
			EndPage
		end if 
	
	End with
	pnlTab1.style.visibility = "visible" 
End Sub

Sub EndPage()
	set mobjMDOTPOINTADCOMMI = Nothing
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


'-----------------------------------------------------------------------------------------
'���� �Է��� ���� �ʱ�ȭ
'-----------------------------------------------------------------------------------------
Sub EXCEL_UPLOAD
	with frmThis
		
		'�⺻ �������� �ٽ� �׷��� ��Ʈ�� �ٿ� �ְų� �Է��� �޵��� �Ѵ�.
		makePageData
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'���� ��ȸ�� �����Ͱ� ���� �ϸ� �����͸� �����ְ� ���� ���� ������ �Է��� �����Ѵ�.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			intRtn = gYesNoMsgBox("�̹� ���� �ŷ������� �� ������ ���� �մϴ�. ���ðڽ��ϱ�?" & vbCrlf & "(��:�ٽú���,�ƴϿ�:�ڷ����)","�ڷ���� Ȯ��")
			IF intRtn = vbYes then
			
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG 
				mobjSCGLSpr.SetFlag  .sprSht,meINS_FLAG
				
				'��ȸ�Ѵ�.
				SelectRtn
			elseif intRtn = vbNo then
				'����ڰ� ������ ������� �����͸� ���� �Ѵ�.
				DeleteRtn
			end if
		ELSE
			RowNum = 500
			mobjSCGLSpr.SetMaxRows .sprSht, RowNum
			gOKMsgbox "�����͸� �Է��� �غ� �Ǿ����ϴ�. Excel Data�� �ٿ��־� �ֽʽÿ�.[�ִ� 500 ���� �������Է��� ���� �մϴ�.]", " EXCEL UPLOAD"
		end if
		
	End with
End sub


'-----------------------------------------------------------------------------------------
'���� ���� ��ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	with frmThis
		
		'�⺻ �������� �ٽ� �׷��� ��Ʈ�� �ٿ� �ְų� �Է��� �޵��� �Ѵ�.
		makePageData
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'���� ��ȸ�� �����Ͱ� ���� �ϸ� �����͸� �����ְ� ���� ���� ������ �Է��� �����Ѵ�.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			'��ȸ�� �Աݿ��θ� üũ�ڽ��� �����Ѵ�.
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "PAY_YN"
			Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			
			mobjSCGLSpr.SetCellsLock2 .sprSht,True,-1,1,16,True
			
			
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
		ELSE
			.sprSht.MaxRows = 0
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
		end if
		
	End with
End Sub


'======================================
'��Ʈ�� �ٽ� �׸���
'======================================
Sub makePageData
     
     With frmThis
        .sprSht.MaxRows = 0
        sprSht_DataFields    = "POINTNO | CHN | CLIENTNAME | GACODE | EXCLIENTNAME | ADEXCLIENTNAME | CLIENT_TYPE | ADCLIENTCODE | TITLE | TDATE | EDATE | SAND_STATUS | SAND_DATE | PAY_YN | AMT | CDATE"
        sprSht_DisplayFields = "��ȣ|ä��|�����ָ�|�������ڵ�|�����|������|��������|�����ڵ�|�����|�̺�Ʈ������|�̺�Ʈ������|�߼ۻ���|�߼�����|�Աݿ���|����ܰ�|�������"	
  
        gSetSheetDefaultColor
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 16, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
        mobjSCGLSpr.SetHeader           .sprSht, sprSht_DisplayFields
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
        mobjSCGLSpr.SetCellTypeFloat2	.sprSht, "AMT", -1, -1, 0
        
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetColWidth         .sprSht, "-1", 10
    End With
End Sub


'------------------------------------------------
'���� ���ε� ���� 
'------------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim intCnt
	Dim vntData
	
	with frmThis
	
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'�ش� �ŷ����� ������ �� ������ �ִ��� ��ȸ�Ѵ�.
		vntData = mobjMDOTPOINTADCOMMI.SelectRtn_EXCEL(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		IF mlngRowCnt >0 THEN 
			gErrorMsgBox "����� �����Ͱ� �ֽ��ϴ� Ȯ���Ͻð� �ٽ� ������ �ֽʽÿ�.!","�������"
			exit sub
		end if 
	
		'���� Rows ����ó��
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"POINTNO",intCnt) = ""  then
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			END IF
		Next

		mobjSCGLSpr.SetFlag  .sprSht,meINS_FLAG
		
		'����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
	
 	    if not IsArray(vntData) then 
		    gErrorMsgBox "����� " & meNO_DATA,"�������"
		    exit sub
        end if
		
		intRtn = mobjMDOTPOINTADCOMMI.ProcessRtn_EXCEL(gstrConfigXML, vntData, sprSht_DataFields, mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
	   	    mobjSCGLSpr.SetMaxRows frmThis.sprSht, 0 
	   	    gOKMsgbox "�����͸� ���������� UPLOAD �Ͽ����ϴ�.", "���� ���ε� �ȳ�!" 
	   	    '���ε��� ��ȸ�Ѵ�.
	   	    SelectRtn
	   	 END IF

	end with
End Sub

'------------------------------------------------
'UPLOAD �ߴ� EXCEL �����͸�  �ϰ� ���� �մϴ�.
'------------------------------------------------
Sub DeleteRtn()
	Dim vntData
	Dim intRtn, i

	'On error resume next
	with frmThis
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		if intRtn <> vbYes then exit sub
		
		intRtn = mobjMDOTPOINTADCOMMI.DeleteRtn_EXCEL(gstrConfigXml,mstrTRANSYEARMON, mstrTRANSNO, mCAMPAIGN_CODE)
		'������ �ʱ�ȭ�Ѵ�.
		InitPage
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
	end with
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="880">
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../images/TitleBG.gIF"
							height="28">
							<TR>
								<TD style="WIDTH: 427px" height="28" width="427" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td rowSpan="2" width="14" align="left"><IMG src="../../../images/TitleIcon.gIF" width="14" height="28"></td>
											<td height="4" align="left"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;����Ʈ ģ�� AD &nbsp;���� ���ε�</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 282px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
													src="../../../images/Waiting.GIF" height="23">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="880">
							<TR>
								<TD style="WIDTH: 880px" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 880px" class="KEYFRAME" vAlign="middle" align="center">
									<TABLE id="tblKey" class="DATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTAXYEARMON,txtTAXNO)"
												width="100">�ŷ�������ȣ</TD>
											<TD style="WIDTH: 124px" class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 56px; HEIGHT: 22px" id="txtTRANSYEARMON" class="NOINPUT_L"
													title="�ŷ��������" readOnly maxLength="6" size="4" name="txtTRANSYEARMON">&nbsp;-
												<INPUT accessKey="NUM" style="WIDTH: 48px; HEIGHT: 22px" id="txtTRANSNO" class="NOINPUT_L"
													title="�ŷ�������ȣ" readOnly maxLength="4" size="2" name="txtTRANSNO"></TD>
											<td class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 122px; HEIGHT: 22px" id="txtCAMPAIGN_CODE" class="NOINPUT_L"
													title="ķ���� �ڵ�" readOnly maxLength="4" size="2" name="txtCAMPAIGN_CODE">
											</td>
											<TD align = "right"><IMG style="CURSOR: hand" id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgFind"
													alt="Loading" src="../../../images/imgCho.gif" width="64" height="20">
													<IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery"
													alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgSave"
													alt="���丸 ���� �����մϴ�." src="../../../images/imgSave.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" height="20">
													<IMG style="CURSOR: hand" id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" border="0" name="imgClose"
													alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" height="20">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 880px; HEIGHT: 3px" class="TOPSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				<TR>
					<TD style="WIDTH: 880px" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" class="LISTFRAME" vAlign="top" align="center">
						<DIV style="POSITION: relative; WIDTH: 100%; VISIBILITY: hidden" id="pnlTab1" ms_positioning="GridLayout">
							<OBJECT style="WIDTH: 100%; HEIGHT: 550px" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="23256">
								<PARAM NAME="_ExtentY" VALUE="14552">
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
					<TD style="WIDTH: 880px" id="lblStatus" class="BOTTOMSPLIT"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
