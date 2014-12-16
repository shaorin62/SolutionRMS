<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMMERGELIST.aspx.vb" Inherits="PD.PDCMMERGELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����û�� ������</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : MD/OUTDOORLIST û�����ȭ��
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMOUTDOORLIST.aspx
'��      �� : 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/23 By Hwang Duck su
			:2) 2009/09/28 By Kim Tae Yub
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
Dim mobjPDCOMERGE, mobjPDCOGET , mobjSCCMGET
Dim mstrPROCESS	'�ű��̸� True ��ȸ�� False
Dim mstrCheck
Dim mstrCheck2

CONST meTAB = 9
mstrPROCESS = False

mstrCheck = True
mstrCheck2 = True

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
	if frmThis.txtTAXYEARMON.value = "" then
		gErrorMsgBox "���ݰ�꼭����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub


Sub imgMergeSave_onclick
	Call ProcessRtn_MergeSave()
End Sub

Sub imgMergeDelete_onclick
	Call ProcessRtn_MergeDelete()
End Sub


'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��		
			SelectRtn
     	end if
	End with
	SelectRtn
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
   		SelectRtn
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	SelectRtn
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	SelectRtn
End Sub

'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		If Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK")  then 
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"),,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			NEXT
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Keyup(KeyCode, Shift)
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
	
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
		   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")  Then
		   
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
				strCOLUMN = "VAT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				strCOLUMN = "SUMAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	End With
End Sub

Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or  _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				If .sprSht.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
					
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
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
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

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����									
	set mobjPDCOMERGE		= gCreateRemoteObject("cPDCO.ccPDCOMERGE")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCMGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXYEARMON | TAXNO | MEDFLAG | CLIENTNAME | CLIENTBISNO | AMT | VAT | SUMAMT | SUMM"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|��ȣ|����|������|����ڹ�ȣ|�ݾ�|�ΰ���|�հ�|����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|  10|   6|   0|    20|        15|  11|    11|  11|  25"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TAXNO | AMT | VAT | SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXYEARMON | MEDFLAG | CLIENTNAME | CLIENTBISNO | SUMM", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "TAXYEARMON | TAXNO | MEDFLAG | CLIENTNAME | CLIENTBISNO | AMT | VAT | SUMAMT | SUMM"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXYEARMON | CLIENTBISNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "MEDFLAG", true
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjPDCOMERGE = Nothing
	set mobjPDCOGET = Nothing
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
		.txtTAXYEARMON.value = MID(gNowDate2,1,4) & MID(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		.txtCLIENTNAME1.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		ELSE
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next

			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim strTAXYEARMON, strTAXNO, strCLIENTCODE,strCLIENTNAME
   	Dim strGUBUN, strMEDFLAG
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strTAXYEARMON	 = .txtTAXYEARMON.value
		strTAXNO	 = .txtTAXNO.value
		strCLIENTCODE = .txtCLIENTCODE1.value
		strCLIENTNAME	 = ""
		strGUBUN  = ""
		
		If .rdT.checked = True Then
			strGUBUN = "MERGE"
		END IF
		
		vntData = mobjPDCOMERGE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											strTAXYEARMON, strTAXNO, _
											strCLIENTCODE, strCLIENTNAME, _
											strGUBUN)
		
		

		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
   			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			
			AMT_SUM
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE	
   		end if
   	end with
End Sub

'------------------------------------------
' �ջ� �������
'------------------------------------------
Sub ProcessRtn_MergeSave
	Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .rdT.checked = True Then
			gErrorMsgBox "�̿�û ���¿��� ����û�� �����Ⱑ �����մϴ�.","����û���ȳ�!"
			Exit Sub
		End If
		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ����û���� �Ұ��� �մϴ�.","����û���ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "����û��û���� �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		
		'���⼭ ���� ����
		'if DataValidation =false then exit sub
	    '������ Validation End
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | TAXYEARMON | TAXNO | MEDFLAG | CLIENTNAME | CLIENTBISNO | AMT | VAT | SUMAMT | SUMM")
		
		intRtn = mobjPDCOMERGE.ProcessRtn_MERGE(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then 
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ ����û�� ����" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub




'------------------------------------------
' �ջ���� �������
'------------------------------------------
Sub ProcessRtn_MergeDelete
    Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
				
	with frmThis
		
		If .rdF.checked = True Then
			gErrorMsgBox "û����û���¿��� ����û�� ������ �����մϴ�.","����û���ȳ�!"
			Exit Sub
		End If
		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ��������� �Ұ��� �մϴ�.","������Ҿȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				
				vntData = mobjPDCOMERGE.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, mobjSCGLSpr.GetTextBinding( .sprSht,"TAXYEARMON",intCnt), mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDFLAG",intCnt)) 
				
				If mlngRowCnt > 0 Then
					gErrorMsgBox intCnt & "�� �ջ�û���� ����� �� �Դϴ�.","�����ȳ�!"
					Exit Sub
				End If
			
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "��������� �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | TAXYEARMON | TAXNO | MEDFLAG ")
		
		intRtn = mobjPDCOMERGE.ProcessRtn_MERGE_CANCEL(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn_MERGE_CANCEL") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ �ջ�û�� ����" & mePROC_DONE
			SelectRtn
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
													<TABLE cellSpacing="0" cellPadding="0" width="96" background="../../../images/back_p.gIF"
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
												<td class="TITLE">
													����û�� ������</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
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
										<!--Common Button End--></TD>
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
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
													width="90">���ݰ�꼭��ȣ</TD>
												<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtTAXYEARMON" title="���ݰ�꼭�����ȸ" style="WIDTH: 96px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="10" name="txtTAXYEARMON">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT class="INPUT" id="txtTAXNO" title="���ݰ�꼭��ȣ��ȸ" style="WIDTH: 64px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="5" name="txtTAXNO"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="50">������</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 173px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
													<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL">����û�� ����
												</TD>
												<TD class="SEARCHDATA" colspan='4'><INPUT id="rdT" title="����û����û����" type="radio" value="rdT" name="rdGBN">
													&nbsp;����û����û����&nbsp; &nbsp; <INPUT id="rdF" title="�̿�û����" type="radio" value="rdF" name="rdGBN" CHECKED>
													&nbsp;�̿�û����</TD>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD align="left" width="400" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
																<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgMergeSave" onmouseover="JavaScript:this.src='../../../images/ImgMergeSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMergeSave.gIF'"
																	height="20" alt="�ջ�û�������մϴ�." src="../../../images/ImgMergeSave.gIF" border="0" name="ImgMergeSave"></TD>
															<td><IMG id="ImgMergeDelete" onmouseover="JavaScript:this.src='../../../images/ImgMergeDeleteOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMergeDelete.gIF'"
																	height="20" alt="�ջ�û�������մϴ�." src="../../../images/ImgMergeDelete.gif" border="0"
																	name="ImgMergeDelete"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
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
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31856">
												<PARAM NAME="_ExtentY" VALUE="15663">
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
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
