<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMITEMLIST.aspx.vb" Inherits="PD.PDCMITEMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� �ڵ����</title> 
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
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet ���� --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjPDCMGET
Dim mobjPDCMCODETR
Dim mstrCheck
mstrCheck = True
Dim mstrGFLAG
Const meTab = 9
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
Sub imgItemPop_onclick
	dim vntRet
	Dim vntInParams
	with frmThis	
		vntInParams = ""
		vntRet = gShowModalWindow("PDCMCLASSLIST.aspx",vntInParams , 1062,900)
		SelectRtn
	End with
End Sub
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ImgSave_onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF	
End Sub

Sub imgNew_onclick
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

sub imgDelRow_onclick ()
		call sprSht_Keydown(meDEL_ROW, 0)
end sub

Sub CmbSetting
	with frmThis
	.cmbDIV.selectedIndex = 0
	End with
End Sub


'-----------------------------
' ,�����׸��ڵ� ��ȸ 
'-----------------------------
Sub ImgITEMCODE_onclick
	Call ImgITEM_POP()
End Sub

Sub ImgITEM_POP
	Dim vntRet, vntInParams
	with frmThis
		vntInParams = array(trim(.txtITEMNAME.value),.txtCLASSNM.value,.cmbDIV.value )
		vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtITEMCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtITEMNAME.value = trim(vntRet(3,0))	'�ڵ�� ǥ��
			.cmbDIV.value = trim(vntRet(5,0))
			.txtCLASSCD.value = trim(vntRet(6,0))
			.txtCLASSNM.value =  trim(vntRet(2,0))
			'gSetChangeFlag .txtCPDEPTCD
		end if
	end with
End Sub

Sub txtITEMNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
		
			vntData = mobjPDCMGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",.txtITEMNAME.value)
	
			if not gDoErrorRtn ("GetITEMCODE") then
				If mlngRowCnt = 1 Then
					.txtITEMCODE.value = trim(vntData(0,0))
					.txtITEMNAME.value = trim(vntData(3,0))
					.cmbDIV.value = trim(vntData(5,0))
					.txtCLASSCD.value = trim(vntData(6,0))
					.txtCLASSNM.value =  trim(vntData(2,0))
					'.txtCPEMPNAME.focus()
				Else
					Call ImgITEM_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------
' �ߺз��ڵ� ��ȸ 
'-----------------------------
Sub ImgCLASSCD_onclick
	Call ImgCLASS_POP()
End Sub

Sub ImgCLASS_POP
	Dim vntRet, vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLASSCD.value),trim(.txtCLASSNM.value))
		vntRet = gShowModalWindow("PDCMITEMCLASSPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCLASSCD.value = trim(vntRet(1,0))	'Code�� ����
			.txtCLASSNM.value = trim(vntRet(2,0))	'�ڵ�� ǥ��
			.cmbDIV.value = trim(vntRet(3,0))
			'gSetChangeFlag .txtCPDEPTCD
		end if
	end with
End Sub

Sub txtCLASSNM_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
		
			vntData = mobjPDCMGET.GetDIVCLASS(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLASSCD.value,.txtCLASSNM.value )
	
			if not gDoErrorRtn ("GetITEMCODE") then
				If mlngRowCnt = 1 Then
					.txtCLASSCD.value = trim(vntData(0,1))
					.txtCLASSNM.value = trim(vntData(1,1))
					.cmbDIV.value = trim(vntData(2,1))
					'.txtCPEMPNAME.focus()
				Else
					Call ImgCLASS_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub



'--------------------------------------------------------------------------------------------------------------------
' ���������Ʈ 
'--------------------------------------------------------------------------------------------------------------------


sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub


Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	with frmThis
	If Cint(KeyCode) = 13 then exit Sub 
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
		
		Select Case intRtn
				Case meINS_ROW:		
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,.sprSht.activeRow,1,2,true
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,.sprSht.activeRow,4,4,true
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,.sprSht.activeRow,6,-1,true
				Case meDEL_ROW: DeleteRtn
		End Select

		
	End with
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	with frmThis
	
		IF Col = 7 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",Row))
			vntRet = gShowModalWindow("PDCMITEMCLASSPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASS",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIV",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(0,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtITEMNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
			
		end if
	End with
End Sub

Sub sprSht_change(ByVal Col,ByVal Row)
Dim strCode
Dim strCodeName
Dim vntData
	with frmThis
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
		IF Col = 6 Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",.sprSht.ActiveRow)
			vntData = mobjPDCMGET.GetDIVCLASS(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASS",Row, vntData(0,1)       '
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIV",Row, vntData(2,1)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(3,1)
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
			Else
				mobjSCGLSpr_ClickProc .sprSht, 7, .sprSht.ActiveRow
			End If
			.txtITEMNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		END IF
	End With
	
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = 7 Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"CLASSNAME",Row))
			vntRet = gShowModalWindow("PDCMITEMCLASSPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASS",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIV",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(0,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtCLASSNM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub


'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjPDCMCODETR	= gCreateRemoteObject("cPDCO.ccPDCOCODETR") 
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	'�� ��ġ ���� �� �ʱ�ȭ
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "207px"
	'pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	GridLayout
    'Sheet �⺻Color ����
    with frmThis
	mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
	End with
	InitPageData	
End Sub
Sub GridLayout

gSetSheetDefaultColor()
	With frmThis
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 0,0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "DIV|CLASS|ITEM|ITEMCODE|DIVNAME|CLASSNAME|BTN|ITEMNAME|DETAIL_YN"
		mobjSCGLSpr.SetHeader .sprSht,		"��з��ڵ�|�ߺз��ڵ�|�Һз��ڵ�|�����׸��ڵ�|��з���|�ߺз���|�����׸��|�ι�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","      12|        12|      10  |        15  |      12|    25|2|        25|   5"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "DETAIL_YN "
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMNAME", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "ITEMCODE|DIV|CLASS|ITEM|DIVNAME|CLASSNAME|BTN|ITEMNAME"
		'mobjSCGLSpr.ColHidden .sprSht, "DETAIL_YN", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEM|DIVNAME|CLASSNAME|ITEMNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE|DIV|CLASS|ITEM",-1,-1,2,2,false
		
	End With
End Sub

Sub InitPageData
	call SUBCOMBO_TYPE()
End SUb

Sub EndPage()
	set mobjPDCMCODETR = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' SUBCOMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()
	Dim vntDIVNM
	
	With frmThis   
		
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
       	
       	vntDIVNM = mobjPDCMGET.GetDataType_DIVNM(gstrConfigXml, mlngRowCnt, mlngColCnt)
		If not gDoErrorRtn ("GetDataTypeChange") Then 
			 gLoadComboBox .cmbDIV, vntDIVNM, False
   		End If  
   		gSetChange
   	end With   
End Sub



Sub SelectRtn
	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMCODETR.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtITEMCODE.value,.txtITEMNAME.value,.cmbDIV.value,.txtCLASSCD.value,.txtCLASSNM.value )

		if not gDoErrorRtn ("SelectRtn") then
			'mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 0,0,2
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,.sprSht.activeRow,1,8,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
   		end if
   	end with
End Sub


'-----------------------------
' ����ó��
'-----------------------------
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim lngCnt,intCnt,intCnt2
	Dim strSUMAMT 
	with frmThis
   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"DIV|CLASS|ITEM|BTN|ITEMCODE|DIVNAME|CLASSNAME|ITEMNAME|DETAIL_YN")
	
		if .sprSht.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		intRtn = mobjPDCMCODETR.ProcessRtn(gstrConfigXml,vntData)
	
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			SelectRtn
   		end if
   		
   	end with
End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  	
  		for intCnt = 1 to .sprSht.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"DIV",intCnt) = "" Then 
					gErrorMsgBox "��з��ڵ� �� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			  if mobjSCGLSpr.GetTextBinding(.sprSht,"CLASS",intCnt) = "" Then 
					gErrorMsgBox "�ߺз��ڵ� �� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNAME",intCnt) = "" Then 
					gErrorMsgBox "��з����� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			  if mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNAME",intCnt) = "" Then 
					gErrorMsgBox "�ߺз����� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			  if mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMNAME",intCnt) = "" Then 
					gErrorMsgBox "�����׸���� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			
		next
   	End with
	DataValidation = true
End Function



'------------------------------------------
' ����ó��
'------------------------------------------

Sub DeleteRtn
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	Dim strJOBNO
	Dim intRtn2
	Dim strITEMCODE
	Dim strERRMSG
	'On error resume next
	
	with frmThis
		'�� �Ǿ� ������ ���
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)

		if gDoErrorRtn ("DeleteRtn") then exit sub
		if intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit sub
		end if
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		if intRtn <> vbYes then exit sub
		
		strITEMCODE = ""
		strERRMSG = ""
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			strITEMCODE= mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",vntData(i))
			if mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",vntData(i)) <> ""  then
				intRtn2 = mobjPDCMCODETR.DeleteRtn(gstrConfigXml,strITEMCODE,strERRMSG)
				If strERRMSG <>  "" Then
					gErrorMsgBox strERRMSG,"�����ȳ�"
					Exit Sub
				End If
				
			end if
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			end if
		next
		If intRtn2 = 0 Then
   		Else
			SelectRtn
		End If
		gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�"
		mobjSCGLSpr.DeselectBlock .sprSht
		mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		
	end with
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<P dir="ltr" style="MARGIN-RIGHT: 0px">
				<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0">
					<TR>
						<TD>
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
								border="0">
								<TR>
									<td align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF"
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
												<td class="TITLE">���۰���</td>
											</tr>
										</table>
									</td>
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
										<TABLE id="tblButton" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											border="0">
											<TR>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<!--Top Define Table Start-->
							<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0"
											align="left">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CmbSetting()" width="90">��з��ڵ�</TD>
												<TD class="SEARCHDATA" style="WIDTH: 122px"><SELECT id="cmbDIV" title="��з��ڵ�" style="WIDTH: 120px" name="cmbDIV">
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLASSNM, txtCLASSCD)"
													width="90">�ߺз��ڵ�</TD>
												<TD class="SEARCHDATA" style="WIDTH: 256px"><INPUT class="INPUT_L" id="txtCLASSNM" title="�����׸��" style="WIDTH: 168px; HEIGHT: 22px"
														type="text" maxLength="255" size="22" name="txtCLASSNM"> <IMG id="ImgCLASSCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCLASSCD"> <INPUT class="INPUT_L" id="txtCLASSCD" title="�����׸��ڵ�" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" size="5" name="txtCLASSCD"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtITEMNAME, txtITEMCODE)"
													width="90">�����׸��ڵ�</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtITEMNAME" title="�����׸��" style="WIDTH: 216px; HEIGHT: 22px"
														type="text" maxLength="255" size="30" name="txtITEMNAME"> <IMG id="ImgITEMCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgITEMCODE"> <INPUT class="INPUT_L" id="txtITEMCODE" title="�����׸��ڵ�" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" size="5" name="txtITEMCODE"></TD>
												<TD class="SEARCHDATA2" width="54"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" width="54" align="absMiddle" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 25px"><FONT face="����"></FONT></TD>
					</TR>
					<!--�������-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%">
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0" align="left"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="80" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�����ڵ����</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50"
											border="0">
											<TR>
												<TD><IMG id="imgItemPop" onmouseover="JavaScript:this.src='../../../images/imgItemPopOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgItemPop.gIF'"
														height="20" alt="�����׸�з��ڵ� �� �����մϴ�." src="../../../images/imgItemPop.gIF" width="107"
														border="0" name="imgItemPop"></TD>
												<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
														src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
												<TD><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'"
														alt="�� �� ����" src="../../../images/imgDelRow.gif" width="54" align="absMiddle" border="0"
														name="imgDelRow">
												</TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
							</TABLE>
							<!--�׽�Ʈ ��--></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="POSITION: relative;HEIGHT: 100%;vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31803">
									<PARAM NAME="_ExtentY" VALUE="12965">
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
						<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
					</TR>
				</TABLE>
			</P>
		</form>
	</body>
</HTML>
