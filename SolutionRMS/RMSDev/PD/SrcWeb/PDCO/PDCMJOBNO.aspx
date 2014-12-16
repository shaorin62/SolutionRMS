<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBNO.aspx.vb" Inherits="PD.PDCMJOBNO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���۰�����ȣ ���</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/���۰�����ȣ ��� ȭ��
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBNO.aspx
'��      �� : ���۰�����ȣ C/D/U/R
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/19 By Kim Tae Ho
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
Dim mlngRowCnt, mlngColCnt		'��ȸ�� �ο�,�÷� ������ ��ȯ
Dim mlngRowCnt2,mlngColCnt2		'��ȸ�� �ο�,�÷� ������ ��ȯ
Dim mobjPDCMJOBNO, mobjPDCMACTUALRATE, mobjSCCOGET, mobjPDCMGET '���(JOBNO-CRUD, ACTUALRATE-CRUD, ��ü����, ���۰���)
Dim mstrFlag					'�Է½� Insert,Update ����
Dim mstrBindCHK					'COMBO �� SUBCOMBO ȣ���ʿ伺 üũ
Dim mstrHIDDEN					'�Է��ʵ��� �����
Dim mstrNoClick					'��üüũ�� �ߵ� �ϳ�, �ش� ȭ�鿡�� �̻��


Const meTab = 9
mstrFlag = "SELECT"
mstrNoClick = False
mstrBindCHK = False
mstrHIDDEN = 0

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'=========================================================================================
' ��ɹ�ư
'=========================================================================================
'��ȸ��ư
Sub imgQuery_onclick
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
	
End Sub

'�űԹ�ư
Sub imgNew_onclick
	NewRegNo
End Sub

'�����ư
Sub imgSave_onclick ()
	If frmThis.cmbENDFLAG.value = "PF01" Or  frmThis.cmbENDFLAG.value = "PF02" Then
	Else
		gErrorMsgBox "������°� �Ƿ� �� ���� �� �ƴѰ��� �����ɼ� �����ϴ�.","����ȳ�"
		SelectRtn
		exit Sub
	End If
			
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'������ư
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'�ݱ��ư
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'������ư
Sub imgDelete_onclick()
	with frmThis 
	
	End with 
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'�߰���ư
sub imgAddRow_onclick ()
	Dim strREG_NUM
	
	With frmThis
	
		strREG_NUM	= .txtREG_NUM.value
		
		IF strREG_NUM = "" THEN
			gErrorMsgBox "�߰��� �� �����ϴ�.","�߰��ȳ�!"
			Exit Sub
		Else
			call sprSht_Keydown(meINS_ROW, 0)
		End if
	End With 
end sub

'�μ�������Ϲ�ư ���� Ŭ��
Sub ImgDivamtPop_onclick
	Call ACTUALRATE_POP()
	SelectRtn
End Sub

'�μ�������� ��ư ��������
Sub ACTUALRATE_POP
	Dim vntRet, vntInParams
	Dim strJOBNO , strJOBNAME
	with frmThis
		
		strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.ActiveRow)
		strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow)
		
		If .sprSht.MaxRows = 0  Then
			gErrorMsgBox "���õȵ����Ͱ� �����ϴ�.","ó���ȳ�"
			Exit Sub
		End If
		
		If  mstrFlag = "NEW" Then
			gErrorMsgBox "����� �����й����� �Է��Ҽ� �ֽ��ϴ�.","ó���ȳ�"
			Exit Sub
		End If
		
		vntInParams = array(trim(strJOBNO),trim(strJOBNAME))
		vntRet = gShowModalWindow("PDCMACTUALRATEPOP.aspx",vntInParams , 1060,800)
		
		if isArray(vntRet) then
		    .txtDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
			.txtDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtEMPNAME.focus()
			gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub


'�ű� ���
Sub NewRegNo
	mstrNoClick = True
	mstrFlag = "NEW"
	Dim vntRet
	Dim vntInParams
	dim intRtn
	DataClean
	
	with frmThis
	
		vntInParams = array(trim(.txtPROJECTNO.value), trim(.txtPROJECTNM.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			.cmbJOBGUBN.disabled = false
			.cmbCREPART.disabled = false
			
			if .txtPROJECTNO.value = vntRet(0,0) and .txtPROJECTNM.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(2,0))  ' �ڵ�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(4,0))  ' �ڵ�� ǥ��
			.txtGROUPGBN.value = trim(vntRet(5,0))  ' �ڵ�� ǥ��	
			.txtCREDAY.value = trim(vntRet(6,0))  ' �ڵ�� ǥ��
			.txtCPDEPTNAME.value = trim(vntRet(7,0))  ' �ڵ�� ǥ��
			.txtCPEMPNAME.value = trim(vntRet(8,0))  ' �ڵ�� ǥ��
			.txtCLIENTTEAMNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
			.txtMEMO.value = trim(vntRet(9,0))
			If mstrHIDDEN = 0 Then
				.txtJOBNAME.focus()					' ��Ŀ�� �̵�
			End If
			call sprSht_Keydown(meINS_ROW, 0)
			DataFill
			 mobjSCGLSpr.SetCellsLock2 .sprSht,false,"JOBNAME|JOBGUBN|REQDAY|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|BIGO|CREPART|CREDEPTNAME|CREEMPNAME|EXCLIENTCODE|EXCLIENTNAME",1,1,false
     	end if
	End with
	
End Sub
'=========================================================================================
' Į������ư �� ��Ÿ ��ư
'=========================================================================================

'�����Է� �ʵ� �����
Sub Set_SELECTTBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblSelectBody").style.display = "inline"
		Else
			document.getElementById("tblSelectBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub


' �����Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�Է��ʵ� ����' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblBody2").style.display = "inline"
			document.getElementById("spacebar1").style.display = "inline"
			document.getElementById("spacebar2").style.display = "inline"
			
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='�Է��ʵ� ����' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblBody2").style.display = "none"
			document.getElementById("spacebar1").style.display = "none"
			document.getElementById("spacebar2").style.display = "none"
	
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

'=========================================================================================
' Sub ���� ȣ���ϴ� Sub
'=========================================================================================
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		
		.txtHOPEENDDAY.value = gNowDate
		.txtREQDAY.value = gNowDate
		.sprSht.MaxRows = 0
		.txtFROM.focus
		DateClean
		.txtFROM.value = ""
	
	Call COMBO_TYPE()
	Call SUBCOMBO_TYPE()
	Call SEARCHCOMBO_TYPE()
	End with
	DataNewClean
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub EndPage()
	set mobjPDCMJOBNO = Nothing
	set mobjPDCMGET = Nothing
	set mobjPDCMACTUALRATE = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub

Sub DataNewClean
	with frmThis
		.cmbJOBGUBN.selectedIndex = -1
		.cmbCREPART.selectedIndex = -1
		.cmbCREGUBN.selectedIndex = -1
		.cmbENDFLAG.selectedIndex = -1 
		.cmbJOBBASE.selectedIndex = -1
		.txtREQDAY.value = ""
		.txtHOPEENDDAY.value = ""	
	End with
End Sub

Sub imgEndChange_onclick

Dim vntData2
Dim strCODE
Dim intRtnSave
Dim intRtn
Dim intCnt
Dim intCode
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCODE = .txtJOBNO.value

		If strCODE = "" Then
		gErrorMsgBox "�켱 JOBNO �� ��ȸ�Ͻʽÿ�.","ó���ȳ�"
		Exit Sub
		End If

		If .cmbENDFLAG.value <> "PF02" Then
		gErrorMsgBox "�Ϸᱸ�� ������ '����' �ϰ�츸 �����մϴ�.","ó���ȳ�"
		Exit Sub
		End If
		vntData2 = mobjPDCMJOBNO.GetJOBNOSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
		
		If mlngRowCnt = 0 Then
			intRtnSave = gYesNoMsgbox("�Ϸᱸ���� '�Ƿ�'���� �� �����Ͻðڽ��ϱ�?","ó���ȳ�")
			IF intRtnSave <> vbYes then exit Sub
			intRtn = mobjPDCMJOBNO.ProcessRtn_ENDFLAG(gstrConfigXml,strCODE)
			if not gDoErrorRtn ("ProcessRtn_ENDFLAG") then
				gErrorMsgBox "JOBNO [" & strCODE & " ]�Ϸᱸ���� '�Ƿ�' ���·� ����Ǿ����ϴ�.","ó���ȳ�" 
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
						intCode = intCnt 
						Exit For
					End If
				Next
				mobjSCGLSpr.ActiveCell .sprSht, 1,intCode
			end if
		Else
			gErrorMsgBox "�ش� JOBNO �� �������곻���� Ȯ���Ͻʽÿ�","ó���ȳ�"
		End If

	End with
End Sub


Sub DataFill
	with frmThis
	mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTNO",.sprSht.ActiveRow, .txtPROJECTNO.value
	mobjSCGLSpr.SetTextBinding .sprSht,"PROJECTNM",.sprSht.ActiveRow, .txtPROJECTNM.value
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBBASE",.sprSht.ActiveRow, .cmbJOBBASE.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"JOBBASENAME",.sprSht.ActiveRow, .cmbJOBBASE(.cmbJOBBASE.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CREGUBN",.sprSht.ActiveRow, .cmbCREGUBN.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"CREGUBNNAME",.sprSht.ActiveRow, .cmbCREGUBN(.cmbCREGUBN.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CREPART",.sprSht.ActiveRow, .cmbCREPART.value
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBGUBN",.sprSht.ActiveRow, .cmbJOBGUBN.value
	mobjSCGLSpr.SetTextBinding .sprSht,"ENDFLAG",.sprSht.ActiveRow, .cmbENDFLAG.value
	'mobjSCGLSpr.SetTextBinding .sprSht,"ENDFLAGNAME",.sprSht.ActiveRow, .cmbENDFLAG(.cmbENDFLAG.selectedIndex).text
	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
	mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQNAME",.sprSht.ActiveRow, .txtSUBSEQNAME.value
	mobjSCGLSpr.SetTextBinding .sprSht,"REQDAY",.sprSht.ActiveRow, gNowDATE
	mobjSCGLSpr.SetTextBinding .sprSht,"HOPEENDDAY",.sprSht.ActiveRow, gNowDATE
	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTTEAMNAME",.sprSht.ActiveRow, .txtCLIENTTEAMNAME.value 
	End with
End Sub


'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT", lngCnt)
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

'Ŭ����
Sub CleanField (objField1, objField2)
	If frmThis.sprSht.MaxRows > 0 Then
			if isobject(objField1) then 
				objField1.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField1.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			end if
			if isobject(objField2) then 
				objField2.value = ""
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,objField2.dataFld,frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol, frmThis.sprSht.ActiveRow
			End If
	End If
End Sub

'ȭ��Ŭ����
Sub DataClean
	with frmThis
		
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = "" 
		.txtCLIENTNAME.value =  ""
		.txtCPDEPTNAME.value =  ""
		.txtCREDAY.value = ""
		.txtCPEMPNAME.value =  ""
		.txtGROUPGBN.value = ""
		.txtSUBSEQNAME.value =  ""
		.txtCLIENTTEAMNAME.value = ""
		.txtMEMO.value =  ""
		.txtJOBNAME.value =  ""
		.txtJOBNO.value =  ""
		.cmbJOBGUBN.selectedIndex = 0 
		.txtDEPTNAME.value =  ""
		.txtDEPTCD.value =  ""
		.txtREQDAY.value = gNowDate
		'.cmbCREPART.selectedIndex = 0  '���⼭ 0�̾ƴ� ���п� ������ ������
		SUBCOMBO_TYPE
		.txtEMPNAME.value =  "" 
		.txtEMPNO.value =  ""
		.txtHOPEENDDAY.value = gNowDate 
		.cmbCREGUBN.selectedIndex = 0 
		.cmbJOBBASE.selectedIndex = 0 
		.txtCREDEPTNAME.value = "" 
		.txtCREDEPTCD.value =  ""
		.cmbENDFLAG.selectedIndex = 0 
		.txtCREEMPNAME.value =  ""
		.txtCREEMPNO.value =  ""
		.txtAGREEYEARMON.value = "" 
		.txtDEMANDYEARMON.value =  ""
		.txtSETYEARMON.value =  ""
		.txtBUDGETAMT.value =  ""
		.txtBIGO.value =  ""
		.txtEXCLIENTCODE.value = ""
		.txtEXCLIENTNAME.value = ""
		
		.sprSht.MaxRows = 0
	End With
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub InitPage()
	'����������ü ����	
	
	set mobjPDCMJOBNO = gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjPDCMACTUALRATE = gCreateRemoteObject("cPDCO.ccPDCOACTUALRATE")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
   
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 40, 0, 3, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 15, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 22, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 24, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 27, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "JOBNAME|JOBNO|CLIENTNAME|CLIENTTEAMNAME|SUBSEQNAME|JOBGUBN|REQDAY|ENDFLAG|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|BTN_DEPT|EMPNAME|BTN_EMP|HOPEENDDAY|BUDGETAMT|CREPART|CREGUBN|JOBBASE|CREDEPTNAME|BTN_CDEPT|CREEMPNAME|BTN_CEMP|EXCLIENTCODE|EXCLIENTNAME|BTN_EXCLIENTCODE|BIGO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CREDAY|CLIENTTEAMCODE|EMPNO|CREDEPTCD|CREEMPNO"
		mobjSCGLSpr.SetHeader .sprSht,        "JOB��|JOBNO|������|��|�귣��|��ü�ι�|�Ƿ���|�Ϸᱸ��|���ǿ�|û����|����|�μ��ڵ�|�����|�����|�ϷΌ����|����ݾ�|��ü�з�|�űԱ���|������|���۴����|���۴����|ũ���ڵ�|ũ������|���|������Ʈ�ڵ�|������Ʈ��|�׷챸��|CP�μ�|CP�����|�޸�|�����|���ڵ�|�Ƿ���NO|���ۺμ��ڵ�|���ۻ��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   20|    9|    20|20|    20|      12|    10|       8|     8|     8|     8|       0|    12|2|   8|2|       9|      11|      10|       8|      10|      10|2|       8|2|0       |15|    2|  10|           0|         0|       0|    0|        0|   0|     0|     0|       0|           0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "EXCLIENTCODE|EXCLIENTNAME|BIGO", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|HOPEENDDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "BUDGETAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_DEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_EMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_CDEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_CEMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN_EXCLIENTCODE"
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "JOBNO|JOBNAME|CLIENTNAME|SUBSEQNAME|JOBGUBN|REQDAY|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|ENDFLAG|CREDEPTCD|CREDEPTNAME|CREEMPNO|CREEMPNAME|BIGO|CREPART|CLIENTTEAMNAME|EXCLIENTCODE|EXCLIENTNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|CLIENTTEAMNAME|SUBSEQNAME|DEPTNAME|EMPNAME|CREDEPTNAME|CREEMPNAME|BIGO|EXCLIENTNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|EXCLIENTCODE",-1,-1,2,2,false '���
		mobjSCGLSpr.colhidden .sprSht, "DEPTCD|EMPNO|CREDEPTCD|CREEMPNO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CLIENTTEAMCODE|EXCLIENTCODE",true
		.sprSht.style.visibility = "visible"
		'If .cmbENDFLAG.value = "PF01" Or .cmbENDFLAG.value = "PF02" Then 
		'	.cmbENDFLAG.disabled = false
		'Else 
			.cmbENDFLAG.disabled = true
		'End If
		.cmbPOPUPTYPE.value=2
		


    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
	
End Sub
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()

	mstrNoClick = False
	mstrFlag = "SELECT"
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
    Dim intCnt
    Dim strCODE
    Dim vntDataSubCombo
	'On error resume next
	with frmThis
	
	
	
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
	
		
		vntData = mobjPDCMJOBNO.SelectRtn_PROJECTORJOB(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtSEARCHCLIENTSUBCODE.value),Trim(.txtSEARCHCLIENTSUBNAME.value),Trim(.txtSEARCHCLIENTCODE.value),Trim(.txtSEARCHCLIENTNAME.value),.cmbSEARCHJOBGUBN.value,.cmbSEARCHENDFLAG.value,Trim(.txtPROJECTNO1.value),Trim(.txtPROJECTNM1.value),Trim(.cmbPOPUPTYPE.value))
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
				DATACLEAN	
				DataNewClean
			Else
				If .txtBUDGETAMT.value <> "" Then
				txtBUDGETAMT_onblur
				End If
				'SetCellsLock ����
				for intCnt = 1 to .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",intCnt) = "PF01" Or  mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",intCnt) = "PF02" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"JOBNAME|JOBGUBN|REQDAY|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREGUBN|JOBBASE|BIGO|CREPART|CREDEPTNAME|CREEMPNAME|EXCLIENTCODE|EXCLIENTNAME",intCnt,intCnt,false
					End If
					strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",intCnt)
					'Call Get_SUBCOMBO_VALUE(strCODE)
					mlngRowCnt2=clng(0)
					mlngColCnt2=clng(0)
					'���ѷ��� �ɸ��� ���� Combo Setting [���� ��ȸ�� ��ü�κ��� �� �ش�з��� ���� �ʹٸ� (��ü�κ��� �������� �ʰ� �ش�з��� ����ʹٸ� �Ʒ��� ������ �߰��Ͽ��� �Ѵ�.- �ӵ��� ������ ������]
					'vntDataSubCombo = mobjPDCMJOBNO.GetDataType_SubCode(gstrConfigXml, mlngRowCnt2, mlngColCnt2, strCODE)					
									
					'mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",intCnt,intCnt,vntDataSubCombo,,77			
					'mobjSCGLSpr.TypeComboBox = True 			
   				
				Next
			End If
			
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
			sprShtToFieldBinding 1,1
			
			
	
			
		End If		
		.cmbJOBGUBN.disabled = true
		.cmbCREPART.disabled = true
		.txtSELECTAMT.value = 0

	'��ȸ�Ϸ�޼���
	AMT_SUM
	
	End With
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub


'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
  	Dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strJOBNO
	Dim intCnt
	Dim intCode,intEDITCODE
	Dim strEDITJOBNO
	Dim intRtnSave
	with frmThis
	'On error resume next
		strJOBNO = ""
		
  		'������ Validation
  		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNAME|JOBNO|SUBSEQNAME|JOBGUBN|REQDAY|AGREEYEARMON|DEMANDYEARMON|SETYEARMON|DEPTCD|DEPTNAME|EMPNO|EMPNAME|HOPEENDDAY|BUDGETAMT|CREPART|CREGUBN|JOBBASE|ENDFLAG|CREDEPTCD|CREDEPTNAME|CREEMPNO|CREEMPNAME|BIGO|PROJECTNO|PROJECTNM|GROUPGBN|CPDEPTNAME|CPEMPNAME|MEMO|CLIENTNAME|CREDAY|EXCLIENTCODE")
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		if DataValidation =false then exit sub
		'strCODE = .txtPROJECTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
		
		
	
		'ó�� ������ü ȣ��
		strMasterData = gXMLGetBindingData (xmlBind)
		
		if .txtJOBNO.value = "" then
			strSEQFlag = "new"
			intRtn = mobjPDCMJOBNO.ProcessRtn(gstrConfigXml,strMasterData, strSEQFlag,strJOBNO)
		else
			intRtn = mobjPDCMJOBNO.ProcessRtnSheet(gstrConfigXml,vntData)
		end if
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if strSEQFlag = "new" then
				
				'gErrorMsgBox " �ڷᰡ �ű�����" & mePROC_DONE,"����ȳ�"
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
					intCode = intCnt 
					Exit For
					End If
				Next
				mobjSCGLSpr.ActiveCell .sprSht, 1,intCode
				sprShtToFieldBinding 1,intCode		
				intRtnSave = gYesNoMsgbox("�ڷᰡ ���� �Ǿ����ϴ�. �����й����� �Է� �Ͻðڽ��ϱ�?","ó���ȳ�")
				IF intRtnSave <> vbYes then exit Sub
				Call ImgDivamtPop_onclick()	
				sprSht_Click 1,intCode					

			else
				gErrorMsgBox " �ڷᰡ" & intRtn & " �� ��������" & mePROC_DONE,"����ȳ�" 
				strEDITJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",.sprSht.activeRow)
				SelectRtn
				For intCnt = 1 To .sprSht.MaxRows 
					If strEDITJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt) Then
						intEDITCODE = intCnt 
						Exit For
					End If
				Next

				mobjSCGLSpr.ActiveCell .sprSht, 1,intEDITCODE
				sprShtToFieldBinding .sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
			
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
   	
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		If .cmbCREPART.value = "PC01" Or .cmbCREPART.value = "PR01" Then
			If .txtEXCLIENTCODE.value = "" Then
				gErrorMsgBox "��ü�з��� TV-CF �Ǵ� Radio-CM �϶� ũ������ �Է��� �ʼ� �Դϴ�.","�Է¾ȳ�"
				.txtEXCLIENTNAME.focus()
				Exit Function
			End If
		End If
   	
   	End with
	DataValidation = true
End Function


'------------------------------------------
' ������ �Է½� SHEET BINDING onchange EVENT
'------------------------------------------
'�����κ�
Sub txtJOBNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBNAME",frmThis.sprSht.ActiveRow, frmThis.txtJOBNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtJOBNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBNO",frmThis.sprSht.ActiveRow, frmThis.txtJOBNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbJOBGUBN_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		SUBCOMBO_TYPE
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBGUBN",frmThis.sprSht.ActiveRow, frmThis.cmbJOBGUBN.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBGUBNNAME",frmThis.sprSht.ActiveRow, frmThis.cmbJOBGUBN(frmThis.cmbJOBGUBN.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtDEPTCD_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPTCD",frmThis.sprSht.ActiveRow, frmThis.txtDEPTCD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtREQDAY_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REQDAY",frmThis.sprSht.ActiveRow, frmThis.txtREQDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbCREPART_onchange
	if frmThis.sprSht.ActiveRow >0 AND mstrBindCHK = False Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREPART",frmThis.sprSht.ActiveRow, frmThis.cmbCREPART.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREPARTNAME",frmThis.sprSht.ActiveRow, frmThis.cmbCREPART(frmThis.cmbCREPART.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtEMPNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EMPNAME",frmThis.sprSht.ActiveRow, frmThis.txtEMPNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtEMPNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EMPNO",frmThis.sprSht.ActiveRow, frmThis.txtEMPNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtHOPEENDDAY_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"HOPEENDDAY",frmThis.sprSht.ActiveRow, frmThis.txtHOPEENDDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'�������
Sub cmbCREGUBN_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREGUBN",frmThis.sprSht.ActiveRow, frmThis.cmbCREGUBN.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREGUBNNAME",frmThis.sprSht.ActiveRow, frmThis.cmbCREGUBN(frmThis.cmbCREGUBN.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'cmbJOBBASE
Sub cmbJOBBASE_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBBASE",frmThis.sprSht.ActiveRow, frmThis.cmbJOBBASE.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"JOBBASENAME",frmThis.sprSht.ActiveRow, frmThis.cmbJOBBASE(frmThis.cmbJOBBASE.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtCREDEPTNAME
Sub txtCREDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREDEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCREDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCREDEPTCD_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREDEPTCD",frmThis.sprSht.ActiveRow, frmThis.txtCREDEPTCD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbENDFLAG_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDFLAG",frmThis.sprSht.ActiveRow, frmThis.cmbENDFLAG.value
		'mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDFLAGNAME",frmThis.sprSht.ActiveRow, frmThis.cmbENDFLAG(frmThis.cmbENDFLAG.selectedIndex).text
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtCREEMPNAME
Sub txtCREEMPNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREEMPNAME",frmThis.sprSht.ActiveRow, frmThis.txtCREEMPNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCREEMPNO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CREEMPNO",frmThis.sprSht.ActiveRow, frmThis.txtCREEMPNO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtBUDGETAMT_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUDGETAMT",frmThis.sprSht.ActiveRow, frmThis.txtBUDGETAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtBIGO
Sub txtBIGO_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BIGO",frmThis.sprSht.ActiveRow, frmThis.txtBIGO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'PROJECT,JOBNO ����
Sub cmbPOPUPTYPE_onchange
	with frmThis
		.txtPROJECTNM1.value = ""
		.txtPROJECTNO1.value = ""
	End with
	gSetChange
End Sub


Sub txtEXCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEXCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub




'****************************************************************************************
' UI ����
'****************************************************************************************
'�Է¿�
'-----------------------------------------------------------------------------------------
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
	Dim strCODE
    With frmThis   

		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB���� ȣ��
		vntJOBBASE = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBBASE")  'û������ ȣ��	
		vntCREGUBN = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREGUBN")  '�ű�/���� ȣ��
		vntENDFLAG = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '���ۻ��� ȣ��
		vntCREPART = mobjPDCMJOBNO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"CREPART")  
		if not gDoErrorRtn ("COMBO_TYPE") then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "JOBGUBN",,,vntJOBGUBN,,95 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "JOBBASE",,,vntJOBBASE,,77
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREGUBN",,,vntCREGUBN,,65 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "ENDFLAG",,,vntENDFLAG,,65 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",,,vntCREPART,,77
			
			mobjSCGLSpr.TypeComboBox = True 
			 gLoadComboBox .cmbENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbJOBGUBN, vntJOBGUBN, False
			 gLoadComboBox .cmbJOBBASE, vntJOBBASE, False
			 gLoadComboBox .cmbCREGUBN, vntCREGUBN, False 
			 gLoadComboBox .cmbCREPART, vntCREPART, False 
			 'strCODE = vntJOBGUBN(0,1)
			 'Call Get_SUBCOMBO_VALUE(strCODE)	
   		end if    
   		'cmbJOBGUBN_onchange   		
   	end with     	
End Sub
'��ȸ��

'Dynamic Combo
Sub Get_SUBCOMBO_VALUE(strCODE,strPos)							
	Dim vntData					
	With frmThis   					
		On error resume Next				
		'Long Type�� ByRef ������ �ʱ�ȭ				
		mlngRowCnt=clng(0)				
		mlngColCnt=clng(0)				

       	vntData = mobjPDCMJOBNO.GetDataType_SubCode(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)					
		If not gDoErrorRtn ("GetDataType_SubCode") Then 				
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",strPos,strPos,vntData,,77		
			
			mobjSCGLSpr.TypeComboBox = True 			
   		End If  				
   		gSetChange				
   	end With   					
End Sub		
'-----------------------------------------------------------------------------------------
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()
	
	Dim vntJOBGUBN
   	Dim vntCREGUBN
   	Dim vntCREPART
   	Dim vntJOBBASE
	Dim vntENDFLAG  
    With frmThis   

		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntJOBGUBN = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  'JOB���� ȣ��
		vntENDFLAG = mobjPDCMJOBNO.GetDataType_search(gstrConfigXml, mlngRowCnt, mlngColCnt,"ENDFLAG")  '���ۻ��� ȣ��
		if not gDoErrorRtn ("SEARCHCOMBO_TYPE") then 
			 gLoadComboBox .cmbSEARCHENDFLAG, vntENDFLAG, False
			 gLoadComboBox .cmbSEARCHJOBGUBN, vntJOBGUBN, False
			 
			' mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CREPART",,,vntCREPART,,77
   		end if    				   		
   	end with     
   		
End Sub
'-----------------------------------------------------------------------------------------
' SUBCOMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SUBCOMBO_TYPE()

	Dim vntCREPART
   	Dim vntCREGUBN
   
	With frmThis   

		'On error resume next
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
       	
	       	vntCREPART = mobjPDCMJOBNO.GetDataTypeChange(gstrConfigXml, mlngRowCnt, mlngColCnt,.cmbJOBGUBN.value,"K")  '�������� ȣ��	

		if not gDoErrorRtn ("SUBCOMBO_TYPE") then 
			 gLoadComboBox .cmbCREPART, vntCREPART, False
   		end if  
   		cmbCREPART_onchange		   		
   	end with   
End Sub



Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub
Sub txtBUDGETAMT_onfocus
	with frmThis
		.txtBUDGETAMT.value = Replace(.txtBUDGETAMT.value,",","")
	end with
End Sub
Sub txtBUDGETAMT_onblur
	with frmThis
		call gFormatNumber(.txtBUDGETAMT,0,true)
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub







Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtHOPEENDDAY,frmThis.imgCalEndar,"txtHOPEENDDAY_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarREQ_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtREQDAY,frmThis.imgCalEndar,"txtREQDAY_onchange()"
		gSetChange
	end with
End Sub

'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'���� ������List ��������
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array( trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO.value = trim(vntData(0,0))
					.txtJOBNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
		SELECTRTN
	end if
	
End Sub


'-----------------------------------------------------------------------------------------
' ������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------

'�̹�����ư Ŭ����
Sub ImgSEARCHCLIENTSUBCODE_onclick
	with frmThis
			Call SEARCHTEAM_POP()
	End with
	
End Sub
'������ ��ȸ�˾�
Sub SEARCHTEAM_POP
	Dim vntRet, vntInParams
	with frmThis
		'�������ڵ�,�����ָ�,���ڵ�,����
		
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value) , trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value) , trim(.txtSEARCHCLIENTSUBNAME.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
			.txtSEARCHCLIENTSUBCODE.value = trim(vntRet(0,0))
			.txtSEARCHCLIENTSUBNAME.value = trim(vntRet(1,0))
			.txtSEARCHCLIENTCODE.value = trim(vntRet(4,0))	'Code�� ����
			.txtSEARCHCLIENTNAME.value = trim(vntRet(5,0))	'�ڵ�� ǥ��
			
		 
			.txtSEARCHCLIENTNAME.focus()
		end if
	end with

End SUb

'���� ������List ��������
Sub SEARCHCLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value), trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value), trim(.txtSEARCHCLIENTSUBNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCLIENTSUBPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtSEARCHCLIENTSUBCODE.value = vntRet(0,0) and .txtSEARCHCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtSEARCHCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtSEARCHCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtSEARCHCLIENTCODE.value = trim(vntRet(3,0))
			.txtSEARCHCLIENTNAME.value = trim(vntRet(4,0))
			
			.txtSEARCHCLIENTNAME.focus()					' ��Ŀ�� �̵�
			'gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtSEARCHCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
				vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSEARCHCLIENTCODE.value),trim(.txtSEARCHCLIENTNAME.value),trim(.txtSEARCHCLIENTSUBCODE.value),trim(.txtSEARCHCLIENTSUBNAME.value))
			
				if not gDoErrorRtn ("txtSEARCHCLIENTSUBNAME_onkeydown") then
					If mlngRowCnt = 1 Then
						.txtSEARCHCLIENTSUBCODE.value = trim(vntData(0,1))
						.txtSEARCHCLIENTSUBNAME.value = trim(vntData(1,1))
						.txtSEARCHCLIENTCODE.value = trim(vntData(4,1))
						.txtSEARCHCLIENTNAME.value = trim(vntData(5,1))
						.txtSEARCHCLIENTNAME.focus()
					Else
						Call SEARCHTEAM_POP()
					End If
   				end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgSEARCHCLIENTCODE_onclick
	Call SEARCHCLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub SEARCHCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtSEARCHCLIENTCODE.value), trim(.txtSEARCHCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtSEARCHCLIENTCODE.value = vntRet(0,0) and .txtSEARCHCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtSEARCHCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtSEARCHCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��	
			.txtPROJECTNM1.focus()					' ��Ŀ�� �̵�
			'gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
     	
	End with
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtSEARCHCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSEARCHCLIENTCODE.value),trim(.txtSEARCHCLIENTNAME.value),"A")
			
			if not gDoErrorRtn ("txtSEARCHCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtSEARCHCLIENTCODE.value = trim(vntData(0,1))
					.txtSEARCHCLIENTNAME.value = trim(vntData(1,1))
					
				Else
					Call SEARCHCLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------
' ���μ� ��ȸ 
'-----------------------------
Sub ImgDEPTCD_onclick
	Call DEPT_POP()
End Sub

Sub DEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
			.txtDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtEMPNAME.focus()
			gSetChangeFlag .txtDEPTCD
		end if
	end with
End Sub

Sub txtDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtDEPTNAME.value)
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = trim(vntData(0,0))
					.txtDEPTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtEMPNAME.focus()
				Else
					Call DEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub
'-----------------------------------------------------------------------------------------
' Project �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'ProjectNO ��ȸ�˾�
Sub ImgPROJECTNO1_onclick
	with frmThis
		'1�� PROJECT ��ȸ   2�� JOBNO��ȸ
		IF .cmbPOPUPTYPE.value = "1" then
			Call PONO_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	
	End with
End Sub
'���� ������List ��������
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			'.txtCLIENTNAME1.focus()					' ��Ŀ�� �̵�
     	end if
	End with
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
		if .cmbPOPUPTYPE.value = "1" Then '������Ʈ �ڵ� ���
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
		Else
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		End If
   		
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' ����ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'���� ������List ��������
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtDEPTCD.value), trim(.txtDEPTNAME.value), trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtDEPTCD.value = trim(vntRet(2,0))  ' Code�� ����
			.txtDEPTNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			
			if .sprSht.ActiveRow >0 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",.sprSht.ActiveRow, .txtEMPNO.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",.sprSht.ActiveRow, .txtEMPNAME.value
				
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtCREDEPTNAME.focus()
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtEMPNAME
			gSetChangeFlag .txtDEPTCD
			gSetChangeFlag .txtDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A",.txtDEPTCD.value,.txtDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					.txtDEPTCD.value = trim(vntData(2,1))
					.txtDEPTNAME.value = trim(vntData(3,1))
					
					if .sprSht.ActiveRow >0 Then
			
						mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",.sprSht.ActiveRow, .txtEMPNO.value
						mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",.sprSht.ActiveRow, .txtEMPNAME.value
						
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtCREDEPTNAME.focus()
					'.txtMEMO.focus()
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------
' ���ۺμ� ��ȸ 
'-----------------------------
Sub ImgCREDEPTCD_onclick
	Call CREDEPT_POP()
End Sub

Sub CREDEPT_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtCREDEPTNAME.value))
		vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtCREDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
			.txtCREDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then	
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtCREEMPNAME.focus()
			gSetChangeFlag .txtCREDEPTCD
		end if
	end with
End Sub

Sub txtCREDEPTNAME_onkeydown
	If window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols

		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREDEPTNAME.value)
			' mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCodeName.value,strCHK)
			
			if not gDoErrorRtn ("GetCC") then
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = trim(vntData(0,0))
					.txtCREDEPTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtCREEMPNAME.focus()
				Else
					Call CREDEPT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ���ۻ���ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCREEMPNO_onclick
	Call CREEMP_POP()
End Sub

'���� ������List ��������
Sub CREEMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCREDEPTCD.value), trim(.txtCREDEPTNAME.value), trim(.txtCREEMPNO.value), trim(.txtCREEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCREEMPNO.value = vntRet(0,0) and .txtCREEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCREDEPTCD.value = trim(vntRet(2,0))  ' Code�� ����
			.txtCREDEPTNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
			.txtCREEMPNO.value = trim(vntRet(0,0))
			.txtCREEMPNAME.value = trim(vntRet(1,0))
			
			if .sprSht.ActiveRow >0 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			
			.txtBUDGETAMT.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCREEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtCREEMPNAME
			gSetChangeFlag .txtCREDEPTCD
			gSetChangeFlag .txtCREDEPTNAME
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCREEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCREEMPNO.value, .txtCREEMPNAME.value,"A",.txtCREDEPTCD.value,.txtCREDEPTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCREEMPNO.value = trim(vntData(0,1))
					.txtCREEMPNAME.value = trim(vntData(1,1))
					.txtCREDEPTCD.value = trim(vntData(2,1))
					.txtCREDEPTNAME.value = trim(vntData(3,1))
					
					if .sprSht.ActiveRow >0 Then
			
						mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
						
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
						mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
						
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtBUDGETAMT.focus()
					gSetChangeFlag .txtCREEMPNO
				Else
					Call CREEMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' ����� �ڵ��˾� ��ư
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	With frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value), "") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtEXCLIENTCODE.value = trim(vntRet(1,0))  ' Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))  ' �ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, .txtEXCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, .txtEXCLIENTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			.txtBIGO.focus() 
			gSetChangeFlag .txtCREEMPNO	
     	End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value), "")
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))
					.txtEXCLIENTNAME.value = trim(vntData(2,1))
					if .sprSht.ActiveRow >0 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",.sprSht.ActiveRow, .txtEXCLIENTCODE.value
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow, .txtEXCLIENTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					.txtBIGO.focus()
					gSetChangeFlag .txtEXCLIENTCODE
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = false
		window.event.cancelBubble = true
	End If
End Sub


Sub JOBGUBNClean
	with frmThis
		.cmbSEARCHJOBGUBN.selectedIndex = 0
	End with
End Sub
Sub ENDFLAGClean
	with frmThis
		.cmbSEARCHENDFLAG.selectedIndex = 0
	End With
End Sub

Sub DeleteRtn
	Dim vntData
	Dim intSelCnt, intRtn, i , intCnt
	Dim strCODE , strENDFLAG
	Dim intSubRtn
	
	with frmThis
		
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		

		If intSelCnt = 0 Or .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ �ڷᰡ �����ϴ�.","�����ȳ�"
			Exit Sub
		End If
		
		'�����й��� ������ ���� �Ұ���
		'If .chkDEPT.checked = TRUE or .chkRATE.checked =TRUE Then
		'	gErrorMsgBox "�����й������ �ֽ��ϴ�.","�����ȳ�"
		'	Exit Sub
		'End if
		
		for i = intSelCnt-1 to 0 step -1
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strENDFLAG = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",vntData(i))
		
				If strENDFLAG <> "PF01" Then
					gErrorMsgBox "[" & i & "��] ��������°� �Ƿڰ� �ƴѰ��� �����ϽǼ� �����ϴ�.","�����ȳ�!"
					Exit Sub
				End If
			End IF
		next
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			'Insert Transaction�� �ƴ� ��� ���� ������ü ȣ��
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",vntData(i))
				'�ڷ� ����
				intRtn = mobjPDCMJOBNO.DeleteRtn(gstrConfigXml,strCODE)
				'�����й������� �ٷλ���
				intSubRtn =mobjPDCMACTUALRATE.DeleteRtn_DTL_JOBNODEPT_JOBNO(gstrConfigXml,strCODE)
				intSubRtn =mobjPDCMACTUALRATE.DeleteRtn_DTL_ACTUALRATE_JOBNO(gstrConfigXml,strCODE)
			End IF
		next
		

		IF not gDoErrorRtn ("DeleteRtn_DTL_ACTUALRATE_JOBNO") then
			'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			gWriteText "", "������ ������" & intSelCnt & "���� ����" & mePROC_DONE
   		End IF
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear
End Sub



'------------------------------------------
' SHEET EVENT
'------------------------------------------

Sub sprSht_Keydown(KeyCode, Shift)
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
	
		Select Case intRtn
				Case meINS_ROW:
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
	If mstrNoClick = True Then Exit Sub
		if Row > 0 and Col > 0 then		
			
			sprShtToFieldBinding Col,Row
		End If
		'if Col = 20 Then
		'msgbox "�ȴ�."
		'End If
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim strJOBGUBN
	Dim strCode
	Dim strCodeName
	Dim strDeptCodeName
	Dim vntData
	with frmThis
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		.txtBIGO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row)
		.txtBUDGETAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT",Row)
		.cmbCREPART.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREPART",Row)
		.cmbCREGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREGUBN",Row)
		.cmbENDFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",Row)
		.cmbJOBBASE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBBASE",Row)
		.cmbJOBGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row)
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		'��ü�ι� ����� Subtype ����
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"JOBGUBN") Then 
			strJOBGUBN = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row)
			
			Call Get_SUBCOMBO_VALUE(strJOBGUBN,Row)
			SUBCOMBO_TYPE
			
		'��ü�з� Subtype ���� Validation	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREPART") Then '20
			If .cmbCREPART.value = "" Then
			gErrorMsgBox "�����Ͻ� �з��� �ش�ι��� �з������� �ƴմϴ�.","ó���ȳ�!"
			sprSht_Change mobjSCGLSpr.CnvtDataField(.sprSht,"JOBGUBN"),.sprSht.activeRow 
			End If
		
		'���ۻ��			
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREEMPNAME") Then '25
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CREEMPNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, ""
				.txtCREEMPNO.value = ""
				.txtCREEMPNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow)
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREEMPNAME",.sprSht.ActiveRow)
				
				vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = vntData(0,1)  ' Code�� ����
					.txtCREDEPTNAME.value = vntData(1,1)  ' �ڵ�� ǥ��
					.txtCREEMPNO.value = vntData(2,1)
					.txtCREEMPNAME.value = vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntData(3,1)
					mobjSCGLSpr.CellChanged .sprSht,38,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش� �̰ż�
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
			
		'���ۺμ�	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '23
			 
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, ""
				.txtCREDEPTCD.value = ""
				.txtCREDEPTNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDEPTNAME",.sprSht.ActiveRow)
				
				
				vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtCREDEPTCD.value = vntData(0,0)  ' Code�� ����
					.txtCREDEPTNAME.value = vntData(1,0)  ' �ڵ�� ǥ��
					'msgbox vntData(0,0) 
					
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht,37,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
			
		'���μ�	
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPTNAME") Then '14
		
			If mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, ""
				.txtDEPTCD.value = ""
				.txtDEPTNAME.value = ""
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow)
				
				
				vntData = mobjPDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,0)  ' Code�� ����
					.txtDEPTNAME.value = vntData(1,0)  ' �ڵ�� ǥ��
					'msgbox vntData(0,0) 
					
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht,13,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If	
		'�����		
		Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EMPNAME") Then '16
		
			If mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, ""
				.txtEMPNAME.value = ""
				.txtEMPNO.value = ""
				
			Else
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"DEPTNAME",.sprSht.ActiveRow)
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EMPNAME",.sprSht.ActiveRow)
				
				vntData = mobjPDCMGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				If mlngRowCnt = 1 Then
					.txtDEPTCD.value = vntData(0,1)  ' Code�� ����
					.txtDEPTNAME.value = vntData(1,1)  ' �ڵ�� ǥ��
					.txtEMPNO.value = vntData(2,1)
					.txtEMPNAME.value = vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntData(3,1)
					mobjSCGLSpr.CellChanged .sprSht,36,frmThis.sprSht.ActiveRow
				Else
					mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش� �̰ż�
				.sprSht.Focus	
				If Row <> .sprSht.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
				End IF
			End If
		
		'ũ������
		ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then
			If mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",.sprSht.ActiveRow) = "" Then 
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, ""
				.txtEXCLIENTNAME.value = ""
				.txtEXCLIENTCODE.value = ""
				
			Else
				strCode		= ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
				
				If strCode = "" AND strCodeName <> "" Then			
					vntData = mobjSCCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName, "")

					If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
						If mlngRowCnt = 1 Then
							.txtEXCLIENTCODE.value = vntData(1,1)
							.txtEXCLIENTNAME.value = vntData(2,1)	
							mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(1,1)
							mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(2,1)			
							.txtFROM.focus
							.sprSht.focus
							If Row <> .sprSht.MaxRows Then
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
							Else
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
							End IF
						Else
							mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
							.txtFROM.focus
							.sprSht.focus 
							If Row <> .sprSht.MaxRows Then
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
							Else
								mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
							End IF
						End If
   					End If
   				End If
   			End If
		End If
	End with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
	
		'���۴����
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '25
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"CREDEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(sprSht,"CREEMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtCREDEPTCD.value = vntRet(2,0)  ' Code�� ����
				.txtCREDEPTNAME.value = vntRet(3,0)  ' �ڵ�� ǥ��
				.txtCREEMPNO.value = vntRet(0,0)
				.txtCREEMPNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
		
		'���۴��μ�
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CREDEPTNAME") Then '23
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding(sprSht,"CREDEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtCREDEPTCD.value = vntRet(0,0)
				.txtCREDEPTNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
		
		'���μ�
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPTNAME") Then '14
			vntInParams = array(mobjSCGLSpr.GetTextBinding(sprSht,"DEPTNAME",Row))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtDEPTCD.value = vntRet(0,0)
				.txtDEPTNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntRet(1,0)	
					
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		'�����
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EMPNAME") Then '16
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"DEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(sprSht,"EMPNAME",Row))
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				.txtDEPTCD.value = vntRet(2,0)  ' Code�� ����
				.txtDEPTNAME.value = vntRet(3,0)  ' �ڵ�� ǥ��
				.txtEMPNO.value = vntRet(0,0)
				.txtEMPNAME.value = vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		'ũ������
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXCLIENTNAME") Then '28
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				.txtEXCLIENTCODE.value = vntRet(1,0)
				.txtEXCLIENTNAME.value = vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)	
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			If Row <> .sprSht.MaxRows Then
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row -1
			Else
				mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
			End If
			
		End if
		
	End With
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	with frmThis
	
		
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_DEPT")  Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_DEPT") then exit Sub
			vntInParams = array(trim(.txtDEPTNAME.value))
			vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
			if isArray(vntRet) then
				.txtDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
				.txtDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				if .sprSht.ActiveRow >0 Then	
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				.txtFROM.focus()
				gSetChangeFlag .txtDEPTCD
			end if
		
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EMP") then exit Sub
		
			vntInParams = array(trim(.txtDEPTCD.value), trim(.txtDEPTNAME.value), trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< �޾ƿ��°��
			
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				.txtDEPTCD.value = trim(vntRet(2,0))  ' Code�� ����
				.txtDEPTNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
				.txtEMPNO.value = trim(vntRet(0,0))
				.txtEMPNAME.value = trim(vntRet(1,0))
				
				if .sprSht.ActiveRow >0 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNO",.sprSht.ActiveRow, .txtEMPNO.value
					mobjSCGLSpr.SetTextBinding .sprSht,"EMPNAME",.sprSht.ActiveRow, .txtEMPNAME.value
					
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",.sprSht.ActiveRow, .txtDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",.sprSht.ActiveRow, .txtDEPTNAME.value
					
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				
				.txtFROM.focus()
				'.txtMEMO.focus()					' ��Ŀ�� �̵�
				gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
				gSetChangeFlag .txtEMPNAME
				gSetChangeFlag .txtDEPTCD
				gSetChangeFlag .txtDEPTNAME
     		end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CDEPT") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CDEPT") then exit Sub
		
				vntInParams = array(trim(.txtCREDEPTNAME.value))
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
			if isArray(vntRet) then
				.txtCREDEPTCD.value = trim(vntRet(0,0))	'Code�� ����
				.txtCREDEPTNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				if .sprSht.ActiveRow >0 Then	
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				.txtFROM.focus()
				gSetChangeFlag .txtCREDEPTCD
			end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CEMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_CEMP") then exit Sub
		
			vntInParams = array(trim(.txtCREDEPTCD.value), trim(.txtCREDEPTNAME.value), trim(.txtCREEMPNO.value), trim(.txtCREEMPNAME.value)) '<< �޾ƿ��°��
		
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if .txtCREEMPNO.value = vntRet(0,0) and .txtCREEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				.txtCREDEPTCD.value = trim(vntRet(2,0))  ' Code�� ����
				.txtCREDEPTNAME.value = trim(vntRet(3,0))  ' �ڵ�� ǥ��
				.txtCREEMPNO.value = trim(vntRet(0,0))
				.txtCREEMPNAME.value = trim(vntRet(1,0))
				
				if .sprSht.ActiveRow >0 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNO",.sprSht.ActiveRow, .txtCREEMPNO.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREEMPNAME",.sprSht.ActiveRow, .txtCREEMPNAME.value
					
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTCD",.sprSht.ActiveRow, .txtCREDEPTCD.value
					mobjSCGLSpr.SetTextBinding .sprSht,"CREDEPTNAME",.sprSht.ActiveRow, .txtCREDEPTNAME.value
					
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				end if
				
				.txtFROM.focus()					' ��Ŀ�� �̵�
				gSetChangeFlag .txtCREEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
				gSetChangeFlag .txtCREEMPNAME
				gSetChangeFlag .txtCREDEPTCD
				gSetChangeFlag .txtCREDEPTNAME
     		end if
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_EXCLIENTCODE") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOEXEALLPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				.txtEXCLIENTCODE.value = vntRet(1,0)	
				.txtEXCLIENTNAME.value = vntRet(2,0)	
				if .sprSht.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(2,0)	
					mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
				End If		
			
			End If
			.txtFROM.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtEXCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtEXCLIENTNAME
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		End if
	
		
		
	End with
	
End Sub


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
	
	'Ű�� �����϶� ���ε�
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT") Then
				strCOLUMN = "BUDGETAMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"BUDGETAMT")  Then
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


Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	Dim vntData_DEPT , vntData_RATE
   	Dim strJOBNO , strRow
   	
	mstrBindCHK = True
	with frmThis
	
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		
		.txtPROJECTNM.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNM",Row)
		.txtPROJECTNO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PROJECTNO",Row) 
		.txtCLIENTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtCPDEPTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CPDEPTNAME",Row)
		.txtCREDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREDAY",Row)
		.txtCPEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CPEMPNAME",Row)
		.txtGROUPGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"GROUPGBN",Row)
		.txtSUBSEQNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtCLIENTTEAMNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTTEAMNAME",Row)
		.txtJOBNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		.txtJOBNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",Row)
		.cmbJOBGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row) 
		'.txtMEMO.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",Row) 
		Call SUBCOMBO_TYPE()
		.txtDEPTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEPTNAME",Row)
		.txtDEPTCD.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEPTCD",Row)
		.txtREQDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REQDAY",Row)
		.cmbCREPART.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREPART",Row) 
		'Call Get_SUBCOMBO_VALUE(mobjSCGLSpr.GetTextBinding(.sprSht,"JOBGUBN",Row))
		 
		.txtEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNAME",Row) 
		.txtEMPNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EMPNO",Row)
		.txtHOPEENDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"HOPEENDDAY",Row)
		.cmbCREGUBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREGUBN",Row) 
		.cmbJOBBASE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBBASE",Row) 
		.txtCREDEPTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CREDEPTNAME",Row) 
		.txtCREDEPTCD.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREDEPTCD",Row)
		.cmbENDFLAG.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDFLAG",Row) 
		.txtCREEMPNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREEMPNAME",Row)
		.txtCREEMPNO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"CREEMPNO",Row)
		.txtAGREEYEARMON.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AGREEYEARMON",Row) 
		.txtDEMANDYEARMON.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDYEARMON",Row)
		.txtSETYEARMON.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"SETYEARMON",Row)
		.txtBUDGETAMT.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"BUDGETAMT",Row)
		.txtBIGO.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row)
		.txtEXCLIENTCODE.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",Row)
		.txtEXCLIENTNAME.value =  mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
  		If .txtBUDGETAMT.value <> "" Then
			txtBUDGETAMT_onblur
		End If
		'If .cmbENDFLAG.value = "PF01" Or .cmbENDFLAG.value = "PF02" Then 
		'	.cmbENDFLAG.disabled = false
		'Else 
			.cmbENDFLAG.disabled = true
		'End If
		
		
		'�й���� ���θ� �̸� �˷��ش�.
		strJOBNO = .txtJOBNO.value		
		.chkDEPT.checked = FALSE
		.chkRATE.checked = FALSE
		
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData_DEPT = mobjPDCMACTUALRATE.SelectRtn_DTL_JOBNODEPT(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,"")
		IF mlngRowCnt > 0 then
			.chkDEPT.checked = TRUE
		end If
		
		
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData_RATE = mobjPDCMACTUALRATE.SelectRtn_DTL_ACTUALRATE(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO,"")
		If mlngRowCnt > 0  then
			.chkRATE.checked = TRUE
		end IF
		
		.txtFROM.focus()
		.sprSht.focus()	
			
   	end with
  
	mstrBindCHK = False
End Function
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML><XML id="xmlBind1"></XML>
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
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;����&nbsp;���� <!--<span id="spnSELECTHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_SELECTTBL_HIDDEN ()">
														(�����)</span>--></td>
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
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblSelectBody" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="top">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand; HEIGHT: 22px" onclick="vbscript:Call DateClean()"
													width="112">�Ƿ���&nbsp;�˻�</TD>
												<TD class="SEARCHDATA" style="WIDTH: 230px; HEIGHT: 16.72pt" width="230"><INPUT class="INPUT" id="txtFROM" title="�Ƿ��� �˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
														accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="�Ƿ��� �˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
														width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 22px" onclick="vbscript:Call gCleanField(txtSEARCHCLIENTSUBNAME, txtSEARCHCLIENTSUBCODE)"
													width="90">��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 229px; HEIGHT: 16.72pt" width="229"><INPUT class="INPUT_L" id="txtSEARCHCLIENTSUBNAME" title="����θ� ��ȸ" style="WIDTH: 149px; HEIGHT: 22px"
														type="text" maxLength="100" size="18" name="txtSEARCHCLIENTSUBNAME"><IMG id="ImgSEARCHCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgSEARCHCLIENTSUBCODE"><INPUT class="INPUT" id="txtSEARCHCLIENTSUBCODE" title="������ڵ� ��ȸ" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtSEARCHCLIENTSUBCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 88px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtSEARCHCLIENTCODE, txtSEARCHCLIENTNAME)"
													width="88">������</TD>
												<TD class="SEARCHDATA" style="HEIGHT: 18.24pt"><INPUT class="INPUT_L" id="txtSEARCHCLIENTNAME" title="�����ָ� ��ȸ" style="WIDTH: 131px; HEIGHT: 22px"
														type="text" maxLength="100" size="16" name="txtSEARCHCLIENTNAME"><IMG id="ImgSEARCHCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgSEARCHCLIENTCODE"><INPUT class="INPUT" id="txtSEARCHCLIENTCODE" title="�������ڵ� ��ȸ" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtSEARCHCLIENTCODE"></TD>
												<td class="SEARCHDATA" style="HEIGHT: 18.24pt" width="53"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
														align="right" border="0" name="imgQuery"></td>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 112px; CURSOR: hand" onclick="vbscript:Call JOBGUBNClean()"
													width="112">��ü�ι�</TD>
												<TD class="SEARCHDATA" style="WIDTH: 230px" width="230"><SELECT id="cmbSEARCHJOBGUBN" title="��ü�ι���ȸ" style="WIDTH: 223px" name="cmbSEARCHJOBGUBN"></SELECT></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call ENDFLAGClean()"
													width="90">�Ϸᱸ��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 229px" width="229"><SELECT id="cmbSEARCHENDFLAG" title="�Ϸᱸ��" style="WIDTH: 227px" name="cmbSEARCHENDFLAG"></SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 88px; CURSOR: hand" width="88"><SELECT id="cmbPOPUPTYPE" title="������Ʈ,JOBNO����" style="WIDTH: 88px" name="cmbPOPUPTYPE">
														<OPTION value="1" selected>PROJECT</OPTION>
														<OPTION value="2">JOBNO</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtPROJECTNM1" title="���۰Ǹ� ��ȸ" style="WIDTH: 192px; HEIGHT: 22px"
														type="text" maxLength="100" size="26" name="txtPROJECTNM1"><IMG id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgPROJECTNO1"><INPUT class="INPUT" id="txtPROJECTNO1" title="���۹�ȣ ��ȸ" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="7" size="4" name="txtPROJECTNO1"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<tr>
						<td>
							<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 25px"></TD>
								</TR>
							</table>
							<TABLE style="WIDTH: 100%; HEIGHT: 8px" height="8" cellSpacing="0" cellPadding="0" width="100%"
								background="../../../images/TitleBG.gIF" border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table style="WIDTH: 640px; HEIGHT: 26px" cellSpacing="0" cellPadding="0" width="640" border="0">
											<tr>
												<td class="TITLE"><span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG id="imgTableUp" style="CURSOR: hand" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgTableUp.gif"
															align="absMiddle" border="0" name="imgTableUp"></span> &nbsp;JOB 
													����&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; �հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">&nbsp;&nbsp; 
													�����հ� : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
														src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
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
						<TD class="BODYSPLIT" id="spacebar1" style="WIDTH: 100%; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%" vAlign="top" align="left">
							<TABLE class="DATA" id="tblBody1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="GROUP" width="20" rowSpan="3">��<BR>
										��<BR>
										��<BR>
										��
									</TD>
									<TD class="LABEL" width="90">������Ʈ��</TD>
									<TD class="DATA" width="230"><INPUT dataFld="PROJECTNM" class="NOINPUTB_L" id="txtPROJECTNM" title="������Ʈ��" style="WIDTH: 160px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="21" name="txtPROJECTNM">&nbsp;<INPUT dataFld="PROJECTNO" class="NOINPUTB" id="txtPROJECTNO" title="������Ʈ��" style="WIDTH: 65px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="6" name="txtPROJECTNO"></TD>
									<TD class="LABEL" width="90">�귣��</TD>
									<TD class="DATA" width="230"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_L" id="txtSUBSEQNAME" title="�귣��" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="24" name="txtSUBSEQNAME"></TD>
									<TD class="LABEL" width="90">���μ� [CP]</TD>
									<TD class="DATA"><INPUT dataFld="CPDEPTNAME" class="NOINPUTB_L" id="txtCPDEPTNAME" title="���μ� CP" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCPDEPTNAME"></TD>
								<TR>
									<TD class="LABEL">�����</TD>
									<TD class="DATA"><INPUT dataFld="CREDAY" class="NOINPUTB" id="txtCREDAY" title="�����" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCREDAY"></TD>
									<TD class="LABEL">��</TD>
									<TD class="DATA"><INPUT dataFld="CLIENTTEAMNAME" class="NOINPUTB_L" id="txtCLIENTTEAMNAME" title="��" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTTEAMNAME"></TD>
									<TD class="LABEL">����� [CP]</TD>
									<TD class="DATA"><INPUT dataFld="CPEMPNAME" class="NOINPUTB_L" id="txtCPEMPNAME" title="�����CP" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCPEMPNAME"></TD>
								</TR>
								<TR>
									<TD class="LABEL">�׷챸��</TD>
									<TD class="DATA"><INPUT dataFld="GROUPGBN" class="NOINPUTB_L" id="txtGROUPGBN" title="�׷챸��" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtGROUPGBN"></TD>
									<TD class="LABEL">������</TD>
									<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 229px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
									<TD class="LABEL">���</TD>
									<TD class="DATA"><INPUT dataFld="MEMO" class="NOINPUTB_L" id="txtMEMO" title="�޸�" style="WIDTH: 266px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="32" name="txtMEMO"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" id="spacebar2" style="WIDTH: 100%; HEIGHT: 3px"></TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%" vAlign="top" align="left">
							<TABLE class="DATA" id="tblBody2" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="GROUP" width="20" rowSpan="7"><BR>
										��<BR>
										��<BR>
										��<BR>
										��<BR>
									</TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtJOBNAME, '')"
										width="90">JOB��</TD>
									<TD class="DATA" width="230"><INPUT dataFld="JOBNAME" id="txtJOBNAME" title="���۰Ǹ�" style="WIDTH: 164px; HEIGHT: 21px"
											accessKey=",M" dataSrc="#xmlBind" type="text" size="21" name="txtJOBNAME"><INPUT dataFld="JOBNO" class="NOINPUT" id="txtJOBNO" title="���۰���ȣ�ڵ�" style="WIDTH: 65px; HEIGHT: 21px"
											dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBNO"></TD>
									<TD class="LABEL" width="90">��ü�ι�</TD>
									<TD class="DATA" width="230"><SELECT dataFld="JOBGUBN" id="cmbJOBGUBN" title="��ü����" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBGUBN"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtDEPTNAME, txtDEPTCD)"
										width="90">�����</TD>
									<TD class="DATA"><INPUT dataFld="DEPTNAME" class="INPUT_L" id="txtDEPTNAME" title="���μ���" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtDEPTNAME"><IMG id="ImgDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgDEPTCD"><INPUT dataFld="DEPTCD" class="INPUT_L" id="txtDEPTCD" title="���μ��ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="6" name="txtDEPTCD"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtREQDAY, '')">�Ƿ���</TD>
									<TD class="DATA"><INPUT dataFld="REQDAY" class="INPUT" id="txtREQDAY" title="�Ƿ���" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtREQDAY"><IMG id="imgCalEndarREQ" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
											name="imgCalEndarREQ"></TD>
									<TD class="LABEL">��ü�з�</TD>
									<TD class="DATA"><SELECT dataFld="CREPART" id="cmbCREPART" title="��ü�з�" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREPART"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEMPNAME, txtEMPNO)">�����</TD>
									<TD class="DATA"><INPUT dataFld="EMPNAME" class="INPUT_L" id="txtEMPNAME" title="����ڸ�" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtEMPNAME"><IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgEMPNO"><INPUT dataFld="EMPNO" class="INPUT_L" id="txtEMPNO" title="����ڻ��" style="WIDTH: 70px; HEIGHT: 22px"
											accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtEMPNO"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtHOPEENDDAY, '')">�ϷΌ����</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><INPUT dataFld="HOPEENDDAY" class="INPUT" id="txtHOPEENDDAY" title="�ϷΌ����" style="WIDTH: 112px; HEIGHT: 22px"
											accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="13" name="txtHOPEENDDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndar"></TD>
									<TD class="LABEL" style="HEIGHT: 24px">�ű�/���� ����</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><SELECT dataFld="CREGUBN" id="cmbCREGUBN" title="�ű�/���� ����" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbCREGUBN"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call CleanField(txtCREDEPTNAME,txtCREDEPTCD)">���۴����</TD>
									<TD class="DATA" style="HEIGHT: 18.24pt"><INPUT dataFld="CREDEPTNAME" class="INPUT_L" id="txtCREDEPTNAME" title="���۴���ںμ���" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtCREDEPTNAME"><IMG id="ImgCREDEPTCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
											name="ImgCREDEPTCD"><INPUT dataFld="CREDEPTCD" class="INPUT_L" id="txtCREDEPTCD" title="���۴���ںμ��ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCREDEPTCD"></TD>
								</TR>
								<TR>
									<TD class="LABEL">�Ϸᱸ��</TD>
									<TD class="DATA"><SELECT dataFld="ENDFLAG" id="cmbENDFLAG" title="�Ϸᱸ��" style="WIDTH: 112px" dataSrc="#xmlBind"
											name="cmbENDFLAG"></SELECT>&nbsp;&nbsp;<IMG id="imgEndChange" onmouseover="JavaScript:this.src='../../../images/imgEndChangeOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgEndChange.gIF'" height="20" alt="������� �� �Ƿڻ��·� �����մϴ�."
											src="../../../images/imgEndChange.gIF" align="absMiddle" border="0" name="imgEndChange"></TD>
									<TD class="LABEL">������</TD>
									<TD class="DATA"><SELECT dataFld="JOBBASE" id="cmbJOBBASE" title="������" style="WIDTH: 224px" dataSrc="#xmlBind"
											name="cmbJOBBASE"></SELECT></TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCREEMPNAME, txtCREEMPNO)">���۴����</TD>
									<TD class="DATA"><INPUT dataFld="CREEMPNAME" class="INPUT_L" id="txtCREEMPNAME" title="���۴����" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="23" name="txtCREEMPNAME"><IMG id="ImgCREEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0"
											name="ImgCREEMPNO"><INPUT dataFld="CREEMPNO" class="INPUT_L" id="txtCREEMPNO" title="���۴����" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="6" size="4" name="txtCREEMPNO"></TD>
								</TR>
								<TR>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtBUDGETAMT, '')">����</TD>
									<TD class="DATA"><INPUT dataFld="BUDGETAMT" class="INPUT_R" id="txtBUDGETAMT" title="����ݾ�" style="WIDTH: 224px; HEIGHT: 21px"
											accessKey="NUM" dataSrc="#xmlBind" type="text" size="32" name="txtBUDGETAMT"></TD>
									<TD class="LABEL" style="CURSOR: hand">�����й���</TD>
									<TD class="DATA">
										<TABLE class="NOINPUTB" id="Table1" style="WIDTH: 224px; HEIGHT: 27px" cellSpacing="1"
											cellPadding="0">
											<TR>
												<TD class="DATA" style="WIDTH: 54px" align="left" width="54">���<INPUT id="chkDEPT" disabled tabIndex="0" type="checkbox" name="chkDEPT"></TD>
												<TD class="DATA" style="WIDTH: 46px" align="left" width="46">�й�<INPUT id="chkRATE" disabled type="checkbox" name="chkRATE"></TD>
												<TD class="DATA"><IMG id="ImgDivamtPop" onmouseover="JavaScript:this.src='../../../images/ImgDivamtPopOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgDivamtPop.gIF'" height="20"
														alt="�ű��ڷḦ �ۼ��մϴ�." src="../../../images/ImgDivamtPop.gif" border="0" name="ImgDivamtPop"></TD>
											</TR>
										</TABLE>
									</TD>
									<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtEXCLIENTNAME,txtEXCLIENTCODE)">ũ������</TD>
									<TD class="DATA"><INPUT dataFld="EXCLIENTNAME" class="INPUT_L" id="txtEXCLIENTNAME" title="�������" style="WIDTH: 173px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="100" size="24" name="txtEXCLIENTNAME"><IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE"><INPUT dataFld="EXCLIENTCODE" class="INPUT_L" id="txtEXCLIENTCODE" title="�������ڵ�" style="WIDTH: 70px; HEIGHT: 22px"
											dataSrc="#xmlBind" type="text" maxLength="8" size="6" name="txtEXCLIENTCODE">
									</TD>
								</TR>
								<TR>
									<TD class="LABEL">���ǿ�</TD>
									<TD class="DATA"><INPUT dataFld="AGREEYEARMON" class="NOINPUTB" id="txtAGREEYEARMON" title="���ǿ�" style="WIDTH: 224px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="32" name="txtAGREEYEARMON"></TD>
									<TD class="LABEL">û����</TD>
									<TD class="DATA"><INPUT dataFld="DEMANDYEARMON" class="NOINPUTB" id="txtDEMANDYEARMON" title="û����" style="WIDTH: 224px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="32" name="txtDEMANDYEARMON"></TD>
									<TD class="LABEL">����</TD>
									<TD class="DATA"><INPUT dataFld="SETYEARMON" class="NOINPUTB" id="txtSETYEARMON" title="����" style="WIDTH: 264px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="38" name="txtSETYEARMON"></TD>
								</TR>
								<TR>
									<TD class="LABEL" onclick="vbscript:Call CleanField(txtBIGO, '')">���</TD>
									<TD class="DATA" colSpan="5"><INPUT dataFld="BIGO" id="txtBIGO" title="�ΰ����� ���" style="WIDTH: 920px; HEIGHT: 21px" dataSrc="#xmlBind"
											type="text" size="148" name="txtBIGO"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<!--Input End-->
					<!--BodySplit Start-->
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 98%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 98%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="42413">
									<PARAM NAME="_ExtentY" VALUE="6826">
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
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
