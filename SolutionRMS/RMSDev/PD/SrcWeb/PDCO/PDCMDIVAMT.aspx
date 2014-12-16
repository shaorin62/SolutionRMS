<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDIVAMT.aspx.vb" Inherits="PD.PDCMDIVAMT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>û������ ����ó��</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/���۰�����ȣ ��� ȭ��
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMDIVAMT.aspx
'��      �� : ��������Ȯ���п� ���� ���۰�����ȣ ���� ó�� �ɼ� �ֵ��� ó��
'�Ķ�  ���� : 
'Ư��  ���� : �ش� �ϳ��� û��ó �� �������� �����ϸ�, �ϳ��� �Ź�ȣ�� �ι�ȣ�� �ο��Ѵ�.
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
		<script language="vbscript" id="clientEventHandlersVBS">
'�������� ����
Dim mobjPDCMDIVAMT
Dim mobjPDCMGET
Dim mlngRowCnt,mlngColCnt
Dim mlngRowCnt1,mlngColCnt1
Dim mUploadFlag

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

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------


'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	
	'����������ü ����	
	Set mobjPDCMDIVAMT = gCreateRemoteObject("cPDCO.ccPDCODIVAMT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Call Grid_Layout()
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
Sub Grid_Layout()
	Dim intGBN
	Dim strComboList 
	strComboList =  "���" & vbTab & "�̻��"
	gSetSheetDefaultColor
with frmThis
		
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 13, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht,    "PREESTNO|YEARMON|JOBNO|JOBNAME|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|DIVAMT|ADJAMT|INYN|CREDAY|INYNNM"
		mobjSCGLSpr.SetHeader .sprSht,		    "������ȣ|���ǿ�|JOBNO|JOB��|�������ڵ�|�����ָ�|������ڵ�|����θ�|����Ȯ���ݾ�|û���ݾ�|�Ϸᱸ��|����������|�Է±���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "10      |10    |10   |22   |0         |18      |0         |18      |12          |12      |0       |0         |10"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|CLIENTSUBCODE|CREDAY|INYN", true
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PREESTNO|YEARMON|JOBNO|INYN|INYNNM",-1,-1,2,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|CLIENTSUBNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PREESTNO|YEARMON|JOBNO|JOBNAME|DIVAMT|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|INYNNM"
		.sprSht.MaxRows = 1
		
		'**************************************************
		'***�󼼳��� Sheet ������
		'**************************************************	
			
        gSetSheetColor mobjSCGLSpr, .sprSht1 
		mobjSCGLSpr.SpreadLayout .sprSht1, 18, 0
		mobjSCGLSpr.AddCellSpan  .sprSht1, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,10, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht1,13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht1, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT|ATTR02"
		mobjSCGLSpr.SetHeader .sprSht1,         "������ȣ|����|���۹�ȣ|���|����������|�귣��|�귣���|�����|����θ�|������|�����ָ�|���ұݾ�|JOB��|û���ݾ�|û�����"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "0       |0   |0       |0   |10        |6   |2|14      |6     |2|18    |6     |2|18    |10      |28   |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN0"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht1,"..", "BTN2"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "CREDAY", -1, -1, 10
		mobjSCGLSpr.ColHidden .sprSht1, "PREESTNO|SEQ|JOBNO|YEARMON|ADJAMT|ATTR02", true
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht1, "CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|JOBNAME|SUBSEQ|SUBSEQNAME", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "DIVAMT|ADJAMT", -1, -1, 0
		'**************************************************
		'***�󼼳��� Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprShtSum
		mobjSCGLSpr.SpreadLayout .sprShtSum, 18, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|BTN0|SUBSEQNAME|CLIENTSUBCODE|BTN|CLIENTSUBNAME|CLIENTCODE|BTN2|CLIENTNAME|DIVAMT|JOBNAME|ADJAMT|ATTR02"
		mobjSCGLSpr.AddCellSpan  .sprShtSum, 2, 1, 2, 1
		mobjSCGLSpr.SetText .sprShtSum, 2, 1, "�� ��"
		mobjSCGLSpr.SetScrollBar .sprShtSum, 0
		mobjSCGLSpr.SetBackColor .sprShtSum,"1|2",rgb(205,219,215),false
		mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "DIVAMT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprShtSum, "PREESTNO|SEQ|JOBNO|YEARMON|ATTR02", true
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
		mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "15"
		.sprSht1.focus
			
		.txtPREESTNO.style.visibility = "hidden"
		.txtYEARMONPOP.style.visibility = "hidden"
		.txtCREDAY.style.visibility = "hidden"
		.txtJOBNOPOP.style.visibility = "hidden"
		.txtDIVAMT.style.visibility = "hidden"
		
	End with

	
	'pnlTab1.style.visibility = "visible" 
	
End Sub
'-----------------------------------------------------------------------------------------
' �󼼳��� �� �����հ� �׸��� Change �� ó��
'-----------------------------------------------------------------------------------------
'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht1_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
	End with
end sub
'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub


Sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols
    Dim strCHK
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
				
		vntData = mobjPDCMDIVAMT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtJOBNAME.value,.txtJOBNO.value,.cmbYN.value)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.ColHidden .sprSht,strCols,true
   				Call sprSht_Click(1,1)
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim intRtn
	with frmThis
		IF Col = 4 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_CODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OC_NAME",Row)))
			vntRet = gShowModalWindow("MDCMDEPTPOP.aspx",vntInParams , 413,425)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_CODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OC_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtDEPTCODE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		
		end if
	End with
End Sub
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	
End Sub
'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)

end sub

'������� ��Ʈ ��ư Ŭ��

'Validation
Function DataValidation ()
	DataValidation = false	
	With frmThis
		'IF not gDataValidation(frmThis) then exit Function	
	End With
	DataValidation = True
End Function
'�������


Sub EndPage()
	set mobjPDCMDIVAMT = Nothing
	set mobjPDCMGET = Nothing
	gEndPage	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
	.sprSht.maxrows = 0
	.txtYEARMON.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2)
	End with
	
End Sub

sub DeleteRtn

	
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
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtJOBNO.value = trim(vntRet(0,0))  ' Code�� ����
			.txtJOBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
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
	end if
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
  		If mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",1) = "" Then
  			gErrorMsgBox "ù��° ���� ���۰Ǹ��� �ݵ�� �Է��ϼž� �մϴ�.","�Է¿���"
  			Exit Function
  		End if
  		for intCnt = 1 to .sprSht1.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " ��° ���� �������ڵ带 Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTSUBCODE",intCnt) = "" Then 
					gErrorMsgBox intCnt & " ��° ���� ������ڵ带 Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			 if mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",intCnt) = 0 Then 
					gErrorMsgBox intCnt & " ��° ���� ���ұݾ��� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
		next
		
   	End with
	DataValidation = true
End Function
Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, .txtJOBNOPOP.value 		
		mobjSCGLSpr.SetTextBinding .sprSht,"CREDAY",.sprSht.ActiveRow, .txtCREDAY.value  
	End with
End Sub
sub imgAddRow_onclick ()
	With frmThis
		call sprSht1_Keydown(meINS_ROW, 0)
	End With 
end sub
sub imgDelRow_onclick ()
	With frmThis
		call sprSht1_Keydown(meDEL_ROW, 0)
	End With 
end sub

Sub sprSht1_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 13 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub
'����ó��
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strJOBNO,strDEMANDAMT,strJOBYEARMON
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
   	Dim strRow
	Dim lngCnt,intCnt,intCnt2
	
	with frmThis
   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		'ȸ�ǰ�� �޶� ����ɼ� ����.. �д�ݾ��� û���ݾ׺��� ũ�ٸ� ����,,
		'���� �۴ٸ� �ٷ����� û���ݾ��� ���꿡�� ���� �Ǵ� �谨 �Ǹ� ���� �д� PD_GROUP_DIVAMT �� ���� ���� 
		If CDBL(.txtDIVAMT.value) < strSUMDEMANDAMT Then
   			msgbox "���ұݾ��� ���� û���ݾ��� ������ �����ϴ�."
   			Exit Sub
   		End IF
		
		'���۰Ǹ� ó���� �ο�� ��ġ ��Ű��
		For intCnt2 = 2 To .sprSht1.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",intCnt2) = "" Then
				mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNAME",intCnt2, mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",1)  
			end if
		Next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"PREESTNO|SEQ|JOBNO|YEARMON|CREDAY|SUBSEQ|CLIENTSUBCODE|CLIENTSUBNAME|CLIENTCODE|CLIENTNAME|DIVAMT|JOBNAME")
		
		if .sprSht1.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		intRtn = mobjPDCMDIVAMT.ProcessRtn(gstrConfigXml,vntData,.txtCUSTCODEHRD.value )
	
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			gOkMsgBox  intRtn & "���� �ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			strRow = .sprSht.ActiveRow
			SelectRtn
			Call sprSht_Click(1,strRow)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
   		end if
   		
   	end with
End Sub

'��������
Sub DeleteRtn_DTL
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	dim strJOBNO,strCUST,strSEQ
	Dim lngSUMAMT,lngSUMAMT2
	Dim strPREESTNO
	Dim dblSEQ
	Dim strRow
	Dim strGUBN
	'On error resume next
	
	with frmThis
		'�� �Ǿ� ������ ���
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht1,intSelCnt)

		if gDoErrorRtn ("DeleteRtn_Dtl") then exit sub

		if intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		if intRtn <> vbYes then exit sub
		
		strJOBNO = ""
		strCUST = ""
		strSEQ = 0
		lngSUMAMT = 0
		lngSUMAMT2 = 0
		'�հ谡 �´��� ���ΰ˻�
		'��������Ǿ� �ִ� �ݾ�
		
		strGUBN = ""
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			strJOBNO = Trim(.txtJOBNOPOP.value) 
			strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"PREESTNO",vntData(i))	
			dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))	
			'Insert Transaction�� �ƴ� ��� ���� ������ü ȣ��
			if cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "" AND cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "1" then
				If cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"ATTR02",vntData(i))) <> "" Then
					gErrorMsgBox "�ŷ����� �ۼ������� �����ɼ� �����ϴ�.","��������"
					Exit Sub
				End If
				intRtn = mobjPDCMDIVAMT.DeleteRtn(gstrConfigXml,strJOBNO,strPREESTNO,dblSEQ)
				strGUBN = "T"
			Elseif cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) = "1" Then
				gErrorMsgBox "���ʻ��� ���������� �����ɼ� �����ϴ�.","��������"
				Exit Sub
			Elseif cstr(mobjSCGLSpr.GetTextBinding(.sprSht1,"SEQ",vntData(i))) <> "" Then
				strGUBN = "F"
			end if
			
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht1,vntData(i)
				'�հ�����
				
   			end if
		next
		'ProcessRtn
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht1
		mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
		'gWriteText lblStatus,"�ڷᰡ ����" & mePROC_DONE
		If strGUBN = "T" Then
			strRow = .sprSht.ActiveRow
			SelectRtn
			Call sprSht_Click(1,strRow)
			mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		End If
		
	end with
End Sub
sub SelectRtn_DTL ()
   	Dim vntData
   	Dim i, strCols
	Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjPDCMDIVAMT.SelectRtn_DIV(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNOPOP.value)

		if not gDoErrorRtn ("SelectRtn_DIV") then
			mobjSCGLSpr.SetClipBinding .sprSht1, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			If mlngRowCnt < 1 Then
			frmThis.sprSht1.MaxRows = 0 
			
			Else
				'�ŷ����� �ۼ��� �� ���Ͽ� ���� �Ұ��� �ϵ��� ó���Ͽ���
				For intCnt = 1 To .sprSht1.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht1,"ATTR02",intCnt) = "" Then
						If intCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
						mobjSCGLSpr.SetCellsLock2 .sprSht1,false,intCnt,-1,-1,true 
					Else
						mobjSCGLSpr.SetCellShadow .sprSht1, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						
						mobjSCGLSpr.SetCellsLock2 .sprSht1,true,intCnt,-1,-1,true
					End If
				Next
			End If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht1
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			'Call SUM_AMT ()
   		end if
   	end with
end sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	with frmThis
		.txtPREESTNO.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row)
		.txtYEARMONPOP.value = mobjSCGLSpr.GetTextBinding( .sprSht,"YEARMON",Row)
		.txtJOBNOPOP.value = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",Row)
		.txtCREDAY.value = mobjSCGLSpr.GetTextBinding( .sprSht,"CREDAY",Row)
		.txtDIVAMT.value =mobjSCGLSpr.GetTextBinding( .sprSht,"DIVAMT",Row)
		SelectRtn_DTL
		SUM_AMT
	End with
End Sub
'------------------------------------------
' �󼼳��� �׸��� ó��
'------------------------------------------
Sub sprSht1_change(ByVal Col,ByVal Row)
	
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName,strCodeName2
   	Dim strQTY,strPRICE,strAMT 
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		IF  Col = 11 Then
			
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)
			
			if not gDoErrorRtn ("GetCUSTNO_HIGHCUSTCODE") then
			
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(5,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(6,0)			
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+4,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 10, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
   		ElseIF  Col = 14 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
				
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(1,0)			
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 13, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
   		ElseIF  Col = 8 Then
		
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row)
			strCodeName2 = mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row)
			vntData = mobjPDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName,"",strCodeName2)

			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntData(1,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntData(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntData(3,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntData(4,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntData(7,0)
					mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntData(8,0)		
					'mobjSCGLSpr.CellChanged .sprSht1, frmThis.sprSht1.ActiveCol-1,frmThis.sprSht1.ActiveRow
					.txtYEARMON.focus
					.sprSht1.focus 
					mobjSCGLSpr.ActiveCell .sprSht1, Col+7,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht1, 7, Row
					.txtYEARMON.focus
					.sprSht1.focus 
				End If
   			end if
		end if
   	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht1, Col,Row
	SUM_AMT
End Sub	

Sub sprSht1_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Or KeyCode = meTab Then
		if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = 16 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		DefaultValue
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					DefaultValue
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub

Sub DefaultValue
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht1,"PREESTNO",.sprSht1.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht1,"JOBNO",.sprSht1.ActiveRow, .txtJOBNOPOP.value 		
		mobjSCGLSpr.SetTextBinding .sprSht1,"CREDAY",.sprSht1.ActiveRow, .txtCREDAY.value  
	End with
End Sub

Sub sprSht1_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""
		IF Col = 10 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(6,0)				
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+5,Row
			End IF
		elseIF Col = 13 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
			End IF
		elseIF Col = 7 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN0") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(8,0)		
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+8,Row
			End IF
		
		end if
		.txtYEARMON.focus
		.sprSht1.focus 

	End with
	
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht1, Col, Row)
dim vntRet, vntInParams
	with frmThis
		IF Col = 10 Then			
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN1") then exit Sub
			
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTSUBNAME",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			
			vntRet = gShowModalWindow("PDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(1,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(5,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(6,0)		
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+4,Row
			End IF
		elseIF Col = 13 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+2,Row
			End IF
		elseIF Col = 7 Then
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht1,"BTN2") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht1,"CLIENTNAME",Row),"", mobjSCGLSpr.GetTextBinding( .sprSht1,"SUBSEQNAME",Row))
			vntRet = gShowModalWindow("PDCMCUSTSEQPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQ",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"SUBSEQNAME",Row, vntRet(2,0)	
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTCODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTNAME",Row, vntRet(4,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBCODE",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht1,"CLIENTSUBNAME",Row, vntRet(8,0)			
				mobjSCGLSpr.CellChanged .sprSht1, Col,Row
				.txtYEARMON.focus
				.sprSht1.focus 
				mobjSCGLSpr.ActiveCell .sprSht1, Col+7,Row
			End IF
		
		end if
		.txtYEARMON.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.sprSht1.Focus
	end with
End Sub


Sub SUM_AMT()
	Dim lngCnt
	Dim strSUMDEMANDAMT
	Dim strDIVAMT
	strSUMDEMANDAMT = 0
	With frmThis
		For lngCnt = 1 To .sprSht1.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"DIVAMT",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		
		mobjSCGLSpr.SetTextBinding .sprShtSum,"DIVAMT",1, strSUMDEMANDAMT
	End With
End Sub
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD >
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
											<td class="TITLE">
												&nbsp;û������</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="LEFT" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 101; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start-->
								</TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" width="100%" height="100%"  cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TBODY>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="LEFT" colSpan="2">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0" align="LEFT">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
													width="90">���
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 102px; HEIGHT: 22px" type="text"
														maxLength="6" size="11" name="txtYEARMON" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME,txtJOBNO)"
													width="90">JOB��
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 378px"><INPUT class="INPUT_L" id="txtJOBNAME" title="�ڵ��" style="WIDTH: 256px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="37" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23"
														align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="�ڵ���ȸ" style="WIDTH: 65px; HEIGHT: 22px" type="text"
														maxLength="8" align="left" size="3" name="txtJOBNO"></TD>
												<TD class="SEARCHLABEL" width="90">�Ϸᱸ��
												</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbYN" title="��뱸��" style="WIDTH: 104px" name="cmbYN">
														<OPTION value="" selected>��ü</OPTION>
														<OPTION value="Y">�Ϸ�</OPTION>
														<OPTION value="N">�̿Ϸ�</OPTION>
													</SELECT>
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></td>
											</TR>
										</TABLE>
										<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 1040px; HEIGHT: 25px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;JOB ����Ʈ</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
									<!--���� �� �׸���-->
								</TR>
								<TR>
									<!--����-->
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 40%" vAlign="top" align="left">
										<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27517">
												<PARAM NAME="_ExtentY" VALUE="9604">
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
									<TD>
										<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px" id="lblstatus"><FONT face="����"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left"  height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="����"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">
																&nbsp;û������ ����</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD>
																<!--Hidden Control Start-->
																<INPUT class="NOINPUT" id="txtPREESTNO" style="WIDTH: 8px; HEIGHT: 22px" tabIndex="1" type="text"
																	size="1" name="txtPREESTNO"> <INPUT class="NOINPUT" id="txtYEARMONPOP" style="WIDTH: 16px; HEIGHT: 22px" tabIndex="1"
																	type="text" size="1" name="txtYEARMONPOP"> <INPUT class="NOINPUT" id="txtCREDAY" style="WIDTH: 13px; HEIGHT: 22px" tabIndex="1" type="text"
																	size="1" name="txtCREDAY"> <INPUT class="NOINPUT" id="txtJOBNOPOP" style="WIDTH: 32px; HEIGHT: 22px" readOnly type="text"
																	size="1" name="txtJOBNOPOP"> <INPUT class="NOINPUT" id="txtDIVAMT" style="WIDTH: 40px; HEIGHT: 22px" tabIndex="1" readOnly
																	type="text" size="1" name="txtDIVAMT"> 
																<!--Hidden Control End-->
															</TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
															<td><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
																	alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" border="0" name="imgAddRow"></td>
															<TD><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'"
																	alt="�� �� ����" src="../../../images/imgDelRow.gif" width="54" border="0" name="imgDelRow"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
										<TABLE id="tblBody1" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
											border="0">
											<TR>
												<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
										<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 90%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27517">
											<PARAM NAME="_ExtentY" VALUE="6376">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
											
										<OBJECT id="sprShtSum" style="WIDTH: 100%; HEIGHT: 8%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27517">
											<PARAM NAME="_ExtentY" VALUE="609">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
											<PARAM NAME="Protect" VALUE="-1">
											<PARAM NAME="ReDraw" VALUE="-1">
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
									</div>
										</FONT>
									</TD>
								</TR>
				</TR>
				<!--BodySplit End-->
				<!--List Start--></TABLE>
			</TD></TR>
			<TR>
				<TD class="BOTTOMSPLIT" id="lblStatus2" style="WIDTH: 1040px"></TD>
			</TR>
			<!--Bottom Split End--> </TBODY></TABLE> 
			<!--Input Define Table End-->
			</TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
