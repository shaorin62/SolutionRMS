<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOFEE.aspx.vb" Inherits="SC.SCCOFEE" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>Fee�ŷ������� ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOFEE.aspx
'��      �� : FEE �ŷ� �����ָ� ����AOR ��������� ��ȯ �����͸� �����Ѵ�.
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/10 By KimTH
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
Dim mobjSCCOFEE , mobjSCCOGET, mobjMDCOGET

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

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'���߰� ��ư Ŭ��
sub ImgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
End sub

'�����ư Ŭ�� 
Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End SUb

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		with frmThis
			mobjSCGLSpr.SetTextBinding .sprSht,"YEARMON",.sprSht.ActiveRow, .txtYEARMON.value
			mobjSCGLSpr.SetTextBinding .sprSht,"DEMANDDAY",.sprSht.ActiveRow, MID(gNowDate,1,4) & MID(gNowDate,6,2) & MID(gNowDate,9,2)
			mobjSCGLSpr.SetTextBinding .sprSht,"CONFIRMFLAG",.sprSht.ActiveRow, "0"
			.txtYEARMON.focus
			.sprSht.focus
		End with
	End If
End Sub

'���� ��ư Ŭ�� 
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

'--------------------------------
'------��ȸ �˾� ��ư �̺�Ʈ-----
'--------------------------------
'Fee�ŷ������� �˾� - ��ȸ���ǿ�
Sub ImgClient_onclick	
	CLIENTPOP
End Sub

Sub CLIENTPOP
	Dim vntRet
	Dim vntInParams
	Dim strMEDFLAG
	strMEDFLAG =""
	with frmThis
		strMEDFLAG = "A"
		vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME.value,strMEDFLAG)
		vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)
			
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = vntRet(0,0)		             ' Code�� ����
			.txtCLIENTNAME.value = vntRet(1,0)             ' �ڵ�� ǥ��
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
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value,.txtCLIENTNAME.value, "A")
			
			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
				Else
					Call CLIENTPOP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	with frmThis
		if mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",Row) = "" then
			IF mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = 1 THEN
				gErrorMsgBox "�����ָ� �������� �ʰ� ���� ���� �ϽǼ� �����ϴ�..","���� �ȳ�!" 
				Exit Sub
			END IF 
		end if 
	end with
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row	
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim strFDATE, strEDATE
	Dim strCLIENTCODE
	Dim vntData_Log
	Dim strCode
	Dim strCodeName
	Dim vntData
	
	With frmThis	
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CONFIRMFLAG") Then
			if mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",Row) = "" then
				IF mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = 1 THEN
					mobjSCGLSpr.SetTextBinding .sprSht,"CONFIRMFLAG",Row, ""
					
					gErrorMsgBox "�����ָ� �������� �ʰ� ���� ���� �ϽǼ� �����ϴ�..","���� �ȳ�!" 
					Exit Sub
				END IF 
			end if 
		END IF
		
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			strCode = "" : strCodeName = ""
			
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName, "A")
			If mlngRowCnt = 1 Then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
				'mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",Row, vntData(2,1)
				'mobjSCGLSpr.SetTextBinding .sprSht,"COMPANYNAME",Row, vntData(3,1)
				
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
			Else
				mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
			End If
			.txtYEARMON.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+5, Row
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
   	End with 
   	'���� Sprsht ���濡 ���� �÷��� ó��
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)	
				'mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",Row, vntRet(2,0)	
				'mobjSCGLSpr.SetTextBinding .sprSht,"COMPANYNAME",Row, vntRet(3,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtYEARMON.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
	End With
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	WITH frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'****************************************************************************************
'��Ʈ ��ư Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strRow
	Dim strGREATCODE
	
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array("","")
			vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)
				
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"COMPANYNAME",Row, vntRet(3,0)
				'mobjSCGLSpr.SetTextBinding .sprSht,"BUSINO",Row, vntRet(2,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtYEARMON.focus()
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+4,Row
			End IF
			CALL sprSht_Change (Col,Row)
		end if
		
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN_DEPT") Then
			vntInParams = array("","")
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
			CALL sprSht_Change (Col,Row)
		end if
		.sprSht.focus 
	End with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjSCCOFEE	= gCreateRemoteObject("cSCCO.ccSCCOFEE")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,  "YEARMON | SEQ | CLIENTCODE | BTN | CLIENTNAME | DEPT_CD | BTN_DEPT | DEPT_NAME | FDATE | EDATE | DEMANDDAY | MONTHAMT | SUSURATE | FEEAMT | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_PAP | MED_OUT | CONFIRMFLAG | VOCHNO"
		mobjSCGLSpr.SetHeader .sprSht,		  "���|����|�ŷ�ó�ڵ�|�ŷ�ó��|���μ��ڵ�|���μ���|����Ⱓ(����)|����Ⱓ(����)|û����|��ü���Ѿ�|�������������(%)|����������Ѿ�|������TV|������RD|������DMB|CATV|�μ�|����|���α���|��ǥ��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   8|   4|         8|2|    12|           8|2|      12|             0|             0|    12|        14|                8|            14|      10|      10|       10|  10|  10|  10|       4|       0"                
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN | BTN_DEPT"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CONFIRMFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " SUSURATE", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " MONTHAMT | FEEAMT | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_PAP | MED_OUT ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "FDATE | EDATE | DEMANDDAY"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME | DEPT_NAME ", -1, -1, 255
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON | CLIENTCODE | DEPT_CD | DEMANDDAY ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQ | CLIENTNAME | DEPT_NAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "SEQ | MONTHAMT | SUSURATE | FEEAMT | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_PAP | VOCHNO"
		mobjSCGLSpr.ColHidden .sprSht, "FDATE | EDATE | VOCHNO", true
		.sprSht.style.visibility = "visible"
    End With
		
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjSCCOFEE = Nothing
	set mobjSCCOGET = Nothing
	set mobjMDCOGET = Nothing
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
		.sprSht.maxrows = 0
		.txtYEARMON.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2) '���� �̰����� ��ó �ӽ÷� �׽�Ʈ�� ���� �Ͽ���
	End with
End Sub
	
'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn
	Dim vntData
   	Dim i, strCols
    Dim intCnt
    Dim strYEARMON
    Dim strCONFIRMFLAG
    Dim strCLIENTCODE
    Dim strCLIENTNAME
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON 		= .txtYEARMON.value
		strCONFIRMFLAG  = .cmbUSEFLAG.value
		strCLIENTCODE	= .txtCLIENTCODE.value	
		strCLIENTNAME	= .txtCLIENTNAME.value	
				
		vntData = mobjSCCOFEE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strCONFIRMFLAG,strCLIENTCODE,strCLIENTNAME)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				for i = 1 to .sprSht.MaxRows
			
					'��ǥ�� ������ ������ ���
					if mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",i) <> "" or mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "1" then 
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,i,1,-1,true
						
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HCCFFFF, &H000000,False '�����
					else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, i, i,&HFFFFFF, &H000000,False '���
					END IF 
					
					'���� ���� �����ᰡ ������ �������� 0
					if  mobjSCGLSpr.GetTextBinding(.sprSht,"FEEAMT",i) = 0  then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUSURATE",i, 0.0
					end if 
				next
   			Else
   			initpageData
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	'��������
   	Dim vntData
   	Dim intRtn
   	Dim intConfirm
   	Dim i
   	With frmThis
   		for i = 1 to .sprSht.MaxRows
   			if mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",i) <> "" then 
				gErrorMsgBox "��ǥ�� ������ ������ ���� ���� �ϽǼ� �����ϴ�..","����ȳ�"
				Exit Sub
			END IF 
   		next
   		
   		if DataValidation =false then exit sub	
   		intConfirm = gYesNoMsgbox("���� �Ͻðڽ��ϱ�?","����Ȯ��!")
		If intConfirm <> vbYes then exit Sub
		
   		On error resume next
   		vntData = mobjSCGLSpr.GetDataRows(.sprSht," YEARMON | SEQ | CLIENTCODE | BTN | CLIENTNAME | DEPT_CD | BTN_DEPT | DEPT_NAME | FDATE | EDATE | DEMANDDAY | MONTHAMT | SUSURATE | FEEAMT | MED_TV | MED_RD | MED_DMB | MED_CATV | MED_PAP | MED_OUT | CONFIRMFLAG | VOCHNO")
   		If  not IsArray(vntData) Then 
			gErrorMsgBox "����� ������ �����ϴ�.","����ȳ�"
			Exit Sub
		End If
				
   		intRtn = mobjSCCOFEE.ProcessRtn(gstrConfigXml,vntData)
   		
   		If intRtn = 0 Then
			gErrorMsgBox "����� ������ �����ϴ�.","����ȳ�!" 
		End If
   		
   		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox  "�ڷᰡ ����" & mePROC_DONE,"����ȳ�!"
			SelectRtn
   		End If
   	End With
End Sub

Function DataValidation ()
	DataValidation = false
   	Dim intCnt
   	
	On error resume next
	with frmThis
  		for intCnt = 1 to .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt) = "" OR mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",intCnt) = "" Then
   				gErrorMsgBox intCnt & " ���� Fee�ŷ�ó �Ǵ� ���μ� �� �ݵ�� �Է� �Ǿ�� �մϴ�.","����ȳ�!"
				Exit Function
   			End If
		next
	End with
	DataValidation = true
End Function

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
												<TABLE cellSpacing="0" cellPadding="0" width="115" background="../../../images/back_p.gIF"
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
											<td class="TITLE">Fee �ŷ������� ����</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 326px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End--></TD>
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
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="���ݰ�꼭��ȸ ������ �����մϴ�" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="80">�����</TD>
											<TD class="SEARCHDATA" width="88"><INPUT class="INPUT" id="txtYEARMON" title="��Ͽ�" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" title="���ݰ�꼭��ȸ ������ �����մϴ�" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbUSEFLAG, '')"
												width="80">���α���</TD>
											<TD class="SEARCHDATA" width="88"><SELECT id="cmbUSEFLAG" title="�����������" style="WIDTH: 96px" name="cmbUSEFLAG">
													<OPTION value="X" selected>��ü</OPTION>
													<OPTION value="1">����</OPTION>
													<OPTION value="2">�̽���</OPTION>
												</SELECT></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="80">Fee������</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="Fee�����ָ�" style="WIDTH: 224px; HEIGHT: 22px"
													maxLength="255" size="32" name="txtCLIENTNAME"> <IMG id="ImgClient" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgClient"> <INPUT class="INPUT" id="txtCLIENTCODE" title="Fee�������ڵ�" style="WIDTH: 88px; HEIGHT: 22px"
													maxLength="6" name="txtCLIENTCODE"></TD>
											<td class="SEARCHDATA" width="232"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery">
												<IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'"
													alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54" align="absMiddle" border="0"
													name="imgAddRow"> <IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
													align="absMiddle" border="0" name="imgSave"> <IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF"
													align="absMiddle" border="0" name="imgExcel">
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
						<OBJECT style="WIDTH: 100%; HEIGHT: 95%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="31829">
							<PARAM NAME="_ExtentY" VALUE="15213">
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
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
