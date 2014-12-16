<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUDCONFIRMLIST.aspx.vb" Inherits="MD.MDCMCLOUDCONFIRMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����û�� ����/��ȸ</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
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
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjCLOUDCONFIRM
Dim mobjMDCMGET
Dim mstrCheck

CONST meTAB = 9
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


Sub imgClose_onclick ()
	'Window_OnUnload
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------

'���ι�ư
Sub imgSetting_onclick
	Call ProcessRtn_ConfirmOK()
End Sub

'�������
Sub imgSettingCancel_onclick
	Call ProcessRtn_ConfirmCancel()
End Sub

'��ȸ��ư Ŭ����
Sub imgQuery_onclick
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	end if
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'������ư Ŭ����
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
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
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") OR _
		   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
		   
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") Then
				strCOLUMN = "EXSUSU"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
				strCOLUMN = "CGV_AMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU")) OR _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") or  _
			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CGV_AMT") Then
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
   	With frmThis
	End With
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
	set mobjCLOUDCONFIRM	= gCreateRemoteObject("cMDOT.ccMDOTCLOUDCONFIRM")
	set mobjMDCMGET		    = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		'��� û�� ����
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 17, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|��������|����|���|��ȣ|����ڵ�|����|������|���μ�|��û����|������ڵ�|������|���������|���������޾�|������|û������|���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|       5|   4|   5|   4|       8|    15|    15|      15|       8|         8|      12|         8|           8|    10|      12|  10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT | EXSUSU | CGV_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONFIRMGBN | GBN | YEARMON | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | CGV_NAME | CUSTNAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | EXCLIENTCODE | EXCLIENTNAME | CGV_NAME | CUSTNAME | MEMO",-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"CONT_CODE"
		
    End With

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjCLOUDCONFIRM = Nothing
	set mobjMDCMGET = Nothing
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
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		.txtCLIENTNAME.focus()
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON
	Dim strCLIENTNAME
	Dim strCLIENTCODE
	Dim strCONT_CODE
	Dim strCONT_NAME

	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCONT_CODE	= .txtCONT_CODE.value
		strCONT_NAME	= .txtCONT_NAME.value
		

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjCLOUDCONFIRM.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											strYEARMON, strCLIENTCODE, strCLIENTNAME , _
											strCONT_CODE, strCONT_NAME)

		if not gDoErrorRtn ("SelectRtn") then
   			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE	
   		end if
   		AMT_SUM
   	end with
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'------------------------------------------
'���� ������ Ȯ��
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strUSER
	
	with frmThis
			strUSER = gstrEmpNo

   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� Ȯ���� �Ұ��� �մϴ�.","Ȯ���ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) = "Ȯ��" THEN
   					gErrorMsgBox "��Ȯ�������� �����͸� Ȯ���� �����մϴ�.","����ȳ�!"
					Exit Sub
   				END IF
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "Ȯ���� �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If


		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO")
		
		intRtn = mobjCLOUDCONFIRM.ProcessRtn(gstrConfigXml,vntData,strUSER,"CONFIRM")
		
		if not gDoErrorRtn ("ProcessRtn_ConfirmOK") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ Ȯ��" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' ���� ������ Ȯ�� ���
'------------------------------------------
Sub ProcessRtn_ConfirmCancel
    Dim intRtn
   	Dim vntData
	Dim strYEARMON,strSEQ
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strUSER

	
	with frmThis
		strUSER = gstrEmpNo
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ��������� �Ұ��� �մϴ�.","������Ҿȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) = "��Ȯ��" THEN
   					gErrorMsgBox "Ȯ�������� �����͸� Ȯ����Ұ� �����մϴ�.","Ȯ�� ��� �ȳ�!"
					Exit Sub					
   				END IF
   				
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "��������� �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		'On error resume next
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONFIRMGBN | GBN | YEARMON | SEQ | CONT_CODE | CONT_NAME | CLIENTNAME | DEPT_NAME | AMT | EXCLIENTCODE | EXCLIENTNAME | EXSUSU | CGV_AMT | CGV_NAME | CUSTNAME | MEMO")
		
		intRtn = mobjCLOUDCONFIRM.ProcessRtn(gstrConfigXml,vntData,strUSER,"CANCEL")
		
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ Ȯ�����" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<td style="WIDTH: 400px" height="28" width="400" align="left">
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
											<td id="tblTitleName" class="TITLE">CGV Ŭ���� û��Ȯ��</td>
										</tr>
									</table>
								</td>
								<TD height="28" vAlign="middle" width="640" align="right">
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 350px"
										id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
										<TR>
											<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
													src="../../../images/Waiting.GIF" height="23"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1024" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
							<TR>
								<TD style="HEIGHT: 10px" class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="left">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="45">���</TD>
											<TD style="WIDTH: 100px" class="SEARCHDATA"><INPUT accessKey="NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtYEARMON" class="INPUT"
													title="�����ȸ" maxLength="6" size="10" name="txtYEARMON"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="50">������</TD>
											<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 170px; HEIGHT: 22px" id="txtCLIENTNAME" class="INPUT_L" title="�ڵ��"
													maxLength="100" align="left" size="22" name="txtCLIENTNAME"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE" class="INPUT_L" title="�ڵ���ȸ"
													maxLength="6" align="left" name="txtCLIENTCODE"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONT_NAME,'')"
												width="50">����</TD>
											<TD style="WIDTH: 130px" class="SEARCHDATA"><INPUT style="WIDTH: 160px; HEIGHT: 22px" id="txtCONT_NAME" class="INPUT_L" title="����"
													maxLength="100" size="22" name="txtCONT_NAME">
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONT_CODE,'')"
												width="60">��� �ڵ�</TD>
											<TD style="WIDTH: 86px" class="SEARCHDATA"><INPUT style="WIDTH: 80px; HEIGHT: 22px" id="txtCONT_CODE" class="INPUT_L" title="����ڵ�"
													maxLength="6" size="6" name="txtCONT_CODE"></TD>
											<TD align="right"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ ��ȸ�մϴ�."
													src="../../../images/imgQuery.gIF" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<tr>
								<td>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="13">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 25px" class="TOPSPLIT"></TD>
										</TR>
									</TABLE>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28">
										<TR>
											<td style="WIDTH: 100%" class="DATA">�հ� : <INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 20px" id="txtSUMAMT" class="NOINPUTB_R"
													title="�հ�ݾ�" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT style="WIDTH: 120px; HEIGHT: 20px" id="txtSELECTAMT" class="NOINPUTB_R" title="���ñݾ�"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
											<TD style="WIDTH: 100%" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'" border="0" name="imgSetting"
																alt="�ڷḦȮ��ó���մϴ�." src="../../../images/imgSetting.gIF" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgSettingCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gIF'" border="0"
																name="ImgConfirmCancel" alt="Ȯ���� ����մϴ�." src="../../../images/ImgConfirmCancel.gif"
																height="20"></td>
														<td><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
																alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></td>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							<!--BodySplit Start-->
							<TR>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="17012">
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
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
