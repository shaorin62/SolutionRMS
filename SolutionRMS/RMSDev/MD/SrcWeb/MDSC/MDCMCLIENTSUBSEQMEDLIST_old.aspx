<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLIENTSUBSEQMEDLIST_old.aspx.vb" Inherits="MD.MDCMCLIENTSUBSEQMEDLIST_old" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�����ֺ� ��ü�纰 �˻�</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/�׷챤�� �д�� �Է�/��ȸ ȭ��(MDCMGROUP)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMGROUP.aspx.aspx
'��      �� : �׷챤�� �д�� �� ��ȸ/�Է� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/01/09 By Kim Tae Yub
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
'�������� ����
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMGET, mobjMDSRREPORTLIST'�����ڵ�, Ŭ����
Dim mClientsubcode
Dim mstrField
Dim mstrHEADField
Dim mstrEndWith
Dim mmoncnt
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
	
	if frmThis.txtYEAR.value = ""  then
		gErrorMsgBox "����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	if frmThis.txtCLIENTCODE.value = ""  then
		gErrorMsgBox "�������ڵ带 �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME.value) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
     	end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value,.txtCLIENTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(0,0)
					.txtCLIENTNAME.value = vntData(1,0)
					
					Call GetCLIENTSUBLIST (.txtCLIENTCODE.value)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


Sub GetCLIENTSUBLIST (strCLIENTCODE)
	Dim vntData
   	Dim i, strCols
   	Dim strHTML
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strHTML = "" 
		mClientsubcode = ""
		vntData = mobjMDSRREPORTLIST.GetCLIENTSUBLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE)
		if not gDoErrorRtn ("GetCLIENTSUBLIST") then
			If mlngRowCnt > 0 Then
				For i = 0 to mlngRowCnt-1
					strHTML = strHTML & "<INPUT id='"& vntData(0,i) & "' type='checkbox' name='"&  vntData(0,i) & "' checked>" & vntData(1,i) & "&nbsp;&nbsp;"
					IF i = 0 THEN
						mClientsubcode = mClientsubcode & vntData(0,i)
					ELSE
						mClientsubcode = mClientsubcode & "��" & vntData(0,i)
					END IF
				next
			Else
				strHTML = ""
			End If
			document.getElementById("tdCLIENTSUB").innerHTML = strHTML
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
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST")
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
     With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 0, 0, 0, 0,5

    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDSRREPORTLIST = Nothing
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
		.txtYEAR.value = Mid(gNowDate,1,4)
		
		.txtFROMMON.value = "1"
		.txtTOMON.value = MID(gNowDate,6,2)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEAR.focus()
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strSPONSOR
   	dim chkflag
   	dim strSUBLIST
   	Dim strCLIENTSUBLIST
   	Dim intSUBRow
   	Dim strFROMMON
   	Dim strTOMON
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		strSUBLIST = ""
		chkflag = 1
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCLIENTSUBLIST=""
		
		SetChangeLayout
		
'[		exit sub
		
'		IF .txtFROMMON.value ="" THEN
'			strFROMMON = "01"
'		ELSE
'			IF LEN(.txtFROMMON.value) = 1 THEN
'				strFROMMON = "0" & .txtFROMMON.value
'			ELSE
'				strFROMMON = .txtFROMMON.value
'			END IF
'		END IF
'		
'		IF .txtTOMON.value ="" THEN
'			strTOMON = "12"
'		ELSE
'			IF LEN(.txtFROMMON.value) = 1 THEN
'				strTOMON = "0" & .txtTOMON.value
'			ELSE
'				strTOMON = .txtTOMON.value
'			END IF
'		END IF
		
		strCLIENTSUBLIST = 	split(mClientsubcode,"��")
		
		intSUBRow = UBound(strCLIENTSUBLIST, 1)
		FOR i = 0 to intSUBRow
			IF document.getElementById(strCLIENTSUBLIST(i)).checked = true then
				IF chkflag = 1 then
					strSUBLIST = "'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
					chkflag = 2
				else
					strSUBLIST = strSUBLIST & ",'" & document.getElementById(strCLIENTSUBLIST(i)).id & "'"
				end if 
			end if
		Next
		MsgBox .txtYEAR.value
		MsgBox .txtCLIENTCODE.value
		MsgBox strSUBLIST
		vntData = mobjMDSRREPORTLIST.SelectRtn_CLIENTSUBSEQMEDLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEAR.value, .txtCLIENTCODE.value, strSUBLIST, mmoncnt, .txtFROMMON.value, .txtTOMON.value)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			mobjSCGLSpr.ColHidden .sprSht,strCols,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		Layout_change
   	end with
End Sub

Sub SetChangeLayout () 
	Dim strYEAR
	Dim strFROMMON, strTOMON
	Dim strCLIENTCODE
	Dim intAddCnt,intAddHeadCnt,intAddWith,intFieldSetting,intHide,intFloat,intAddCnt2 'For �� Count����
	Dim lngYEARCNT
	Dim strLASTYEAR
	Dim strLASTMON
	Dim lngMONCNT
	Dim i
	
	
	gInitPageSetting mobjSCGLCtl,"MD"
	With frmThis
		'��ü������ ������ ���� ����
		strYEAR		  = .txtYEAR.value
		if .txtFROMMON.value = "" then
			strFROMMON	  = "1"
		else
			strFROMMON	  = .txtFROMMON.value
		end if
		
		if .txtTOMON.value = "" then
			strTOMON	  = "12"
		else
			strTOMON	  = .txtTOMON.value
		end if
		
		strCLIENTCODE = .txtCLIENTCODE.value
		
		
        '���ۿ�, ����� ����
        If strFROMMON <> "" Then
            If strTOMON <> "" Then
				lngMONCNT = CDBL(replace(.txtTOMON.value,"0","")) - CDBL(replace(.txtFROMMON.value,"0",""))
            End If
        Else
            If strTOMON <> "" Then
				lngMONCNT = CDBL(replace(.txtTOMON.value,"0","")) - 1
            Else
                lngMONCNT = 11
            End If
        End If
        mmoncnt = lngMONCNT
				
				
		'CLIENTSUBCODE,  SUBSEQ, INPUT_MEDFLAG, MEDNAME
		'�ʵ� ����������
		Dim strField
		strField = "SUBSEQ|INPUT_MEDFLAG|MEDNAME"
		
		'�ʵ� ���������� [�������ڵ�]
		Dim strAddField
		strAddField = ""
		For i = 0 To lngMONCNT
			strAddField = strAddField & "|A" & i
		Next
		'�ʵ� ������ [��]
		mstrField = strField & strAddField & "|SUMAMT"
		
		
		'OK
		'��� ����������
		Dim strHead
		strHead = "�귣��|����|��ü��"
		
		Dim strADDHead
		strADDHead = ""
		'��� ����������
		For i = 0 To lngMONCNT				
			strADDHead = strADDHead & "|" & CDBL(replace(.txtFROMMON.value,"0","")) + i & "��"
		Next
		
		mstrHEADField = strHead & strADDHead & "|���հ�"
			
		'���� ����������
		Dim strWith
		strWith = "15|15|15"
		'���� ����������
		Dim strAddWith
		'Dim strEndWith ��¹��� ���� ��������ȭ ���Ѻ���
		strAddWith = ""
		For i = 1 To (lngMONCNT+1)
			strAddWith = strAddWith & "|9"
		Next
		mstrEndWith = strWith & strAddWith & "|10"
		
		'���÷�����
		Dim intLayOutCnt
		intLayOutCnt = 3 + (lngMONCNT+1) + 1
		
		gSetSheetColor mobjSCGLSpr, .sprSht
	
			'Sheet Layout ������
			mobjSCGLSpr.SpreadLayout .sprSht, intLayOutCnt, 0,2
			mobjSCGLSpr.SpreadDataField .sprSht, mstrField 
			mobjSCGLSpr.SetHeader .sprSht,       mstrHEADField ,0,1,true
			
			mobjSCGLSpr.SetColWidth .sprSht, "-1", mstrEndWith
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, mstrField, -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUBSEQ|INPUT_MEDFLAG|MEDNAME", , , 50, , ,0
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,mstrField
			mobjSCGLSpr.SetCellAlign2 .sprSht, "SUBSEQ|INPUT_MEDFLAG|",-1,-1,2,2,false
			mobjSCGLSpr.CellGroupingEach .sprSht, "SUBSEQ|INPUT_MEDFLAG"
			
   	End With
End Sub


Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,mlngEvenRowBackColor, &H000000,False
		If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"INPUT_MEDFLAG",intCnt),3) = " ���" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
		END IF
		If RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",intCnt),3) = " �հ�" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		elseIf RIGHT(mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",intCnt),3) = "���հ�" Then
			mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H8876F4, &H000000,False
		End If
	Next 
	End With
End Sub

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
												<td class="TITLE">
													&nbsp;�Ҽӻ纰 ������ ������</td>
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
										<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="110" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
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
									<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 15px" vAlign="top" align="center"><FONT face="����">
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" title="�⵵�������մϴ�." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEAR,'')">��&nbsp; 
														��
													</TD>
													<TD class="SEARCHDATA" width="424" style="WIDTH: 424px"><INPUT class="INPUT" id="txtYEAR" title="�⵵���Է��ϼ���" style="WIDTH: 100px; HEIGHT: 22px"
															type="text" maxLength="4" size="14" name="txtYEAR" accessKey="NUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
														<INPUT class="INPUT" id="txtFROMMON" title="���۳���� �Է��Ͻÿ�" style="WIDTH: 32px; HEIGHT: 22px"
															type="text" maxLength="2" size="1" name="txtFROMMON">&nbsp;- <INPUT class="INPUT" id="txtTOMON" title="������� �Է��Ͻÿ�" style="WIDTH: 32px; HEIGHT: 22px"
															type="text" maxLength="2" size="1" name="txtTOMON">
													</TD>
													<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)">������
													</TD>
													<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 207px; HEIGHT: 22px"
															type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
															border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
															type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE">
													</TD>
												</TR>
												<tr>
													<TD class="SEARCHLABEL" style="WIDTH: 80px">�����
													</TD>
													<TD id="tdCLIENTSUB" class="SEARCHDATA" colspan="3">
													</TD>
												</tr>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 2px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 770px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 768px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 768px"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27490">
												<PARAM NAME="_ExtentY" VALUE="20320">
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
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
								</TR>
								<TR>
									<TD>
									</TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>