<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCLASSLIST.aspx.vb" Inherits="PD.PDCMCLASSLIST" %>
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
Sub imgQuery_onclick
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



'-----------------------------
' ,�����׸��ڵ� ��ȸ 
'-----------------------------
Sub ImgITEMCODE_onclick
	Call ImgITEM_POP()
End Sub

Sub ImgITEM_POP
	Dim vntRet, vntInParams
	with frmThis
		vntInParams = array(trim(.txtITEMNAME.value))
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

Sub sprSht_change(ByVal Col,ByVal Row)
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
End Sub

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
		
		if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW then exit sub
		
		Select Case intRtn
				Case meINS_ROW:		
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,.sprSht.activeRow,2,3,true
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,.sprSht.activeRow,5,6,true
				
				Case meDEL_ROW: DeleteRtn
		End Select
	End with
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF Col = 3 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVCD",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNM",Row))
			vntRet = gShowModalWindow("PDCMITEMDIVPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtCLASSNM.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+3, Row
			
			
		'ELSEIF Col = 6 Then
		'	IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN2") then exit Sub
		'
		'	vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVCD",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNM",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNM",Row))
		'	vntRet = gShowModalWindow("PDCMITEMDIVCLASSPOP.aspx",vntInParams , 413,435)
		'	
		'	IF isArray(vntRet) then
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntRet(0,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntRet(1,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSCD",Row, vntRet(2,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNM",Row, vntRet(3,0)
		'		mobjSCGLSpr.CellChanged .sprSht, Col,Row
		'	End IF
		'	.txtCLASSNM.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		'	.sprSht.Focus
		'	mobjSCGLSpr.ActiveCell .sprSht, Col, Row
				
		end if
	End with
End Sub

Sub sprSht_change(ByVal Col,ByVal Row)
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim intRtn
	
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF Col = 2 Then
		
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNM",.sprSht.ActiveRow)
			vntData = mobjPDCMGET.GetDIVCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntData(1,1)
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
			Else
				if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNM",Row) <> "" Then 
					intRtn = gYesNoMsgbox("��ϵ� �з��� �����ϴ�. �űԵ���Ͻðڽ��ϱ�?","���Ȯ��")
					if intRtn <> vbYes then
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End if 
				END IF 
			End If
			.txtCLASSNM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+3, Row
			
			
		'ELSEIF COL =5 THEN
		'	strCode = ""
		'	strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNM",.sprSht.ActiveRow)
		'	vntData = mobjPDCMGET.GetDIVCLASSCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding( .sprSht,"DIVCD",.sprSht.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNM",.sprSht.ActiveRow),"",strCodeName)
		'	If mlngRowCnt = 1 Then
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntData(0,1)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntData(1,1)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSCD",Row, vntData(2,1)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNM",Row, vntData(3,1)
		'		
		'		mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
		'	Else
		'		mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
		'	End If
		'	.txtCLASSNM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		'	.sprSht.Focus	
		'	mobjSCGLSpr.ActiveCell .sprSht, Col, Row
		END IF
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams

	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF Col = 2 Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"DIVNM",Row))
			vntRet = gShowModalWindow("PDCMITEMDIVPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtCLASSNM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+3, Row
			
		'ELSEIF Col = 5 Then
		'	
		'	vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVCD",Row),mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNM",Row),"",mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNM",Row))
		'	vntRet = gShowModalWindow("PDCMITEMDIVCLASSPOP.aspx",vntInParams , 413,435)
		'	
		'	IF isArray(vntRet) then
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVCD",Row, vntRet(0,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"DIVNM",Row, vntRet(1,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSCD",Row, vntRet(2,0)
		'		mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNM",Row, vntRet(3,0)
		'		mobjSCGLSpr.CellChanged .sprSht, Col,Row
		'	End IF
		'	.txtCLASSNM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		'	.sprSht.Focus	
		'	mobjSCGLSpr.ActiveCell .sprSht, Col, Row
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
	
	Dim strComboList
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
	gSetSheetDefaultColor()
	
	
	With frmThis
	'DIVCD,CLASSCD,DIVNM,CLASSNM
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 5, 0, 0,0,0
		mobjSCGLSpr.AddCellSpan  .sprSht, 2, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "DIVCD|DIVNM|BTN|CLASSCD|CLASSNM"
		mobjSCGLSpr.SetHeader .sprSht,		"��з��ڵ�|��з���|�ߺз��ڵ�|�ߺз���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","20      |39    |2|20        |39"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLASSNM", -1, -1, 255
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "DIVCD|DIVNM|BTN|CLASSCD|CLASSNM"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNM|CLASSNM",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVCD|CLASSCD",-1,-1,2,2,false
		
	End With
	
	InitPageData	
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

		vntData = mobjPDCMCODETR.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,.cmbDIV.value,.txtCLASSCD.value,.txtCLASSNM.value )

		if not gDoErrorRtn ("SelectRtn") then
			'mobjSCGLSpr.SpreadLayout .sprSht, 9, 0, 0,0,2
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,.sprSht.activeRow,1,7,true
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
   		end if
   	end with
End Sub
Sub CmbSetting
	with frmThis
	.cmbDIV.selectedIndex = 0
	End with
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"DIVCD|DIVNM|BTN|CLASSCD|CLASSNM")
	
		if .sprSht.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		intRtn = mobjPDCMCODETR.ProcessRtn_DTL(gstrConfigXml,vntData)
	
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
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNM",intCnt) = "" Then 
					gErrorMsgBox "��з����� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
			  if mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNM",intCnt) = "" Then 
					gErrorMsgBox "�ߺз����� Ȯ���Ͻʽÿ�","�Է¿���"
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
	Dim strDIVCD
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
			strDIVCD = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVCD",vntData(i))
			strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSCD",vntData(i))
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSCD",vntData(i)) <> ""  then
				intRtn2 = mobjPDCMCODETR.DeleteRtn_DTL(gstrConfigXml,strDIVCD,strITEMCODE,strERRMSG)
				If strERRMSG <>  "" Then
					gErrorMsgBox strERRMSG,"�����ȳ�"
					Exit Sub
				End If
				
			end if
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			end if
		next
		
		gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�"
		mobjSCGLSpr.DeselectBlock .sprSht
		mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		SelectRtn
	end with
End Sub



-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
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
												<TABLE cellSpacing="0" cellPadding="0" width="51" background="../../../images/back_p.gIF"
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
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton2"  cellSpacing="0" cellPadding="2" border="0"ALIGN="right">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgClose.gIF" border="0" name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<td>
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
								<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="left">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CmbSetting()" width="90">��з��ڵ�</TD>
											<TD class="SEARCHDATA" style="WIDTH: 368px"><SELECT id="cmbDIV" title="��з��ڵ�" style="WIDTH: 120px" name="cmbDIV">
												</SELECT></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLASSNM, txtCLASSCD)"
												width="90">�ߺз��ڵ�</TD>
											<TD class="SEARCHDATA" ><INPUT class="INPUT_L" id="txtCLASSNM" title="�����׸��" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="255" size="34" name="txtCLASSNM"> <IMG id="ImgCLASSCD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
													name="ImgCLASSCD"> <INPUT class="INPUT_L" id="txtCLASSCD" title="�����׸��ڵ�" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" size="5" name="txtCLASSCD"></TD>
											<TD class="SEARCHDATA" width="54" align="right"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</td>
				</tr>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 100%">
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="103" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���ۺз��ڵ����</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50"
										border="0">
										<TR>
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
						<!--�׽�Ʈ ��--></TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="WIDTH: 100%;POSITION: relative;HEIGHT: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31803">
								<PARAM NAME="_ExtentY" VALUE="12118">
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
		</form>
	</body>
</HTML>
