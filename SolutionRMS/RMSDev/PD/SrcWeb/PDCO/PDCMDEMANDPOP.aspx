<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDEMANDPOP.aspx.vb" Inherits="PD.PDCMDEMANDPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>û����û �̸�����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMDEMANDPOP.aspx
'��      �� : û����û ȭ���� û����û�̸����� ��ư Ŭ���� �̸����� ȭ������ �����Ǹ�, û����û�� �Ͽ� PD_DIVAMT �� ���ԶǴ� ������Ʈ �Ѵ�.
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/06 By KimTH
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
		
Dim mlngRowCnt,mlngColCnt
Dim mobjPDCODEMAND
Dim mobjPDCMGET
Dim mobjSCCOGET
Dim mstrYEARMON1,mstrYEARMON2, mstrUSENO
Dim mstrCheck	
Dim mstrGBN
Dim mlngTempRowCnt,mlngTempColCnt
Dim mstrITEMCODESEQ


Dim mvntData

mstrCheck = True	

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	
	with frmThis
		window.returnvalue = "SAVETRUE"
	End with
	EndPage
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
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
Sub imgRowDel_onclick()

End Sub

Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCODEMAND = gCreateRemoteObject("cPDCO.ccPDCODEMANDLIST")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����

		'mstrPREESTNO,mstrITEMCODE,mlngIMESEQ
		for i = 0 to intNo
			select case i
				case 0 : mstrYEARMON1 = vntInParam(i)			'�ش��
				case 1 : mstrYEARMON2 = vntInParam(i)			'�ش��
				case 2 : mstrUSENO = vntInParam(i)				'�ش�����
			end select
		next
		'PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|AMT
	'**************************************************
	'***Sum Sheet ������
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 29, 0
	mobjSCGLSpr.SpreadDataField .sprSht,    "YEARMON|PREESTNO|JOBNAME|JOBNO|SEQ|CREDAY|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|DEMANDFLAGNAME|MEMO|TAXCODE|TAXCODENAME|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON"
	mobjSCGLSpr.SetHeader .sprSht,		    "��û��|������ȣ|���۰Ǹ�|JOBNO|SUBNO.|������|�������ڵ�|������|���ڵ�|����|�귣���ڵ�|�귣��|�����ݾ�|û���ݾ�|�ܾ�|û������|û������|�����|û�����|û�����|�����|�Ϸᱸ��|���α���|SORT|�׷���|�󼼽�����|���α���|�����̷�|���ο�û��"
	mobjSCGLSpr.SetColWidth .sprSht, "-1",  "     7|      10|15      |7    |6     |9     |0         |13    |0     |13  |0         |13    |11      |11      |11  |10      |10      |10      |10      |10      |6     |10      |10      |10  |10    |10        |10      |10      |10"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	'mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT|CHARGE", -1, -1, 0
	'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REQDAY|", -1, -1, 10
	'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,"YEARMON|PREESTNO|JOBNAME|JOBNO|SEQ|CREDAY|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV|DEMANDFLAGNAME|TAXCODENAME|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "PREESTNO|JOBNAME|CLIENTNAME|TIMNAME|SUBSEQNAME|MEMO|DEMANDFLAGNAME|TAXCODENAME",-1,-1,0,2,false ' ����
	mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON|JOBNO|SEQ|CREDAY|CLIENTCODE|TIMCODE|SUBSEQ|DEMANDFLAG|TAXCODE|USENO|ENDFLAG|CONFIRMFLAG|SORTGBN|RANKDIV",-1,-1,2,2,false '���
	'mobjSCGLSpr.ColHidden .sprSht, "ATTR", true 
	'CHK|PREESTNO|SEQ|ITEMCODESEQ|ITEMCODE|AMT
	'mobjSCGLSpr.ColHidden .sprSht, "PREESTNO", true

	pnlTab1.style.visibility = "visible" 
	.txtYEARMON1.value = mstrYEARMON1
	.txtYEARMON2.value = mstrYEARMON2
	.txtUSENO.value = mstrUSENO
	
	SelectRtn
	.txtEMPNAME.focus()
	End with
	 
End Sub

Sub InitpageData
	with frmThis
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub imgRowAdd_onclick ()
call sprSht_Keydown(meINS_ROW, 0)
End Sub
'================================================================
'UI
'================================================================
Sub txtDIVAMT_onfocus
	with frmThis
		.txtDIVAMT.value = Replace(.txtDIVAMT.value,",","")
	end with
End Sub
Sub txtDIVAMT_onblur
	with frmThis
		call gFormatNumber(.txtDIVAMT,0,true)
	end with
End Sub

Sub txtADJAMT_onfocus
	with frmThis
		.txtADJAMT.value = Replace(.txtADJAMT.value,",","")
	end with
End Sub
Sub txtADJAMT_onblur
	with frmThis
		call gFormatNumber(.txtADJAMT,0,true)
	end with
End Sub

Sub txtCHARGE_onfocus
	with frmThis
		.txtCHARGE.value = Replace(.txtCHARGE.value,",","")
	end with
End Sub
Sub txtCHARGE_onblur
	with frmThis
		call gFormatNumber(.txtCHARGE,0,true)
	end with
End Sub



'================================================================
'SpreadSheet Event
'================================================================
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
		'�ʵ� To ���ε� ����� ����
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
				strCOLUMN = "DIVAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
				strCOLUMN = "CHARGE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
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
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	Dim lngChCnt,IntChAMT,IntChAMTSUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtDIVAMT.value = 0
		else
			.txtDIVAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtDIVAMT,0,True)
		End If
		
		IntEXEAMTSUM = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtADJAMT.value = 0
		else
			.txtADJAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtADJAMT,0,True)
		End If
		
		IntChAMTSUM = 0
		For lngChCnt = 1 To .sprSht.MaxRows
			IntChAMT = 0	
			IntChAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHARGE", lngChCnt)
			IntChAMTSUM = IntChAMTSUM + IntChAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtCHARGE.value = 0
		else
			.txtCHARGE.value = IntChAMTSUM
			Call gFormatNumber(frmThis.txtCHARGE,0,True)
		End If
	End With
End Sub



Sub EndPage
	Set mobjPDCODEMAND = Nothing
	Set mobjPDCMGET = Nothing
	Set mobjSCCOGET = Nothing
	
	gEndPage
End Sub
'=============================================================
'Sheet Event
'=============================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	
End Sub


Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	
End Sub

'=============================================================
'��ȸ
'=============================================================

Sub SelectRtn
	Dim vntData
   	Dim i, strCols
    Dim strCHK
    Dim intCnt
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjPDCODEMAND.SelectRtn_PreView(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value,.txtYEARMON2.value,.txtUSENO.value)
		
		if not gDoErrorRtn ("SelectRtn_PreView") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				For intCnt = 1 To .sprSht.MaxRows 
					'JOB�� �÷� ����
					If mobjSCGLSpr.GetTextBinding(.sprSht,"MEMO",intCnt) <> "�����̿���" Then
						If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKDIV",intCnt) Mod 2 = "0" Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
					End If
				Next
   		
   			Else
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			
   		end if
   	window.setTimeout "AMT_SUM",1	
	.txtSELECTAMT.value = 0
   	end with
   	
   	
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
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../../../PD/SrcWeb/PDCO/PDCMEMPPOP_MANAGER.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()				' ��Ŀ�� �̵�
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtEMPNAME
			
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
			vntData = mobjPDCMGET.GetPDEMP_MANAGER(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			if not gDoErrorRtn ("GetPDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
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

'-----------------------------------------------------------------------------------------
' ���ο�û
'-----------------------------------------------------------------------------------------
Sub processRtn
	Dim vntData
	Dim intRtn
	Dim strSAVEGBN
	Dim intCnt,intCnt2,intCnt3,intMsgCnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg
	'SMS ����
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		
		strMasterData = gXMLGetBindingData (xmlBind)
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "û����û���� �����ϴ�.","û����û�ȳ�"
			Exit Sub
		End If
		
		'��Ʈ�� ����� �����͸� �����´�.
		If .txtEMPNO.value = "" Then
			gErrorMsgBox "���α��ڸ� ���� �Ͻʽÿ�.","û����û�ȳ�"
			Exit Sub
		End If
		
		
		'���α��� �� �׸��忡 ž��
		intMsgCnt = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"MANAGER",intCnt2,Trim(.txtEMPNO.value)
			'�׸����� ���۰Ǹ� �� �����´�
			If intCnt2 = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt2)
			End If
			intMsgCnt = intMsgCnt +1
		Next
	
	
		If intMsgCnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] ���ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ���ο�û���ֽ��ϴ�"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] ��" & intMsgCnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ��" & intMsgCnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			End If
		End If
		
		if DataValidation =false then exit sub 	

		intSaveRtn = gYesNoMsgbox("�ش絥���͸� û����û �Ͻðڽ��ϱ�?","û����û Ȯ��")
		IF intSaveRtn <> vbYes then 
			exit Sub
		Else
		
			'��ü���� �����;� �Ѵ�.
			For intCnt = 1 To .sprSht.MaxRows
				mobjSCGLSpr.CellChanged .sprSht, 1, intCnt	
			Next
			
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|JOBNO|SEQ|PREESTNO|JOBNAME|CLIENTCODE|TIMCODE|SUBSEQ|CREDAY|DIVAMT|ADJAMT|CHARGE|ENDFLAG|DEMANDFLAG|CONFIRMFLAG|MEMO|USENO|TAXCODE|OLDSEQ|MANAGER|CHARGEHISTORY|DATAYEARMON")
			
			intRtn = mobjPDCODEMAND.ProcessRtn_Demand(gstrConfigXml,vntData, .txtYEARMON1.value,.txtYEARMON2.value,.txtUSENO.value)
			If not gDoErrorRtn ("ProcessRtn_Demand") Then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
				
				'������ �����Ͽ����Ƿ� SMS �߼�
				'������ ����� ���� ��������
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				
				vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
				
				'�����»������
				strFromUserName		= vntData_info(0,2)
				strFromUserEmail	= vntData_info(1,2)
				strFromUserPhone	= vntData_info(2,2)
				
				'�޴»�� ����
				strToUserName		=  vntData_info(0,1)
				strToUserEmail		=  vntData_info(1,1)
				strToUserPhone		=  vntData_info(2,1)
			
				
				strAMT = .txtADJAMT.value 
				call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
				
				
				Window_OnUnload
			End If
		End If
		
	End with
End Sub
'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	
   	Dim intCnt
	'On error resume next
	with frmThis
		
   		for intCnt = 1 to .sprSht.MaxRows
   			'Sheet �ʼ� �Է»���
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht,"MANAGER",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ���� ���α��� �Է¿� ������ �ֽ��ϴ�" & vbcrlf & "��� ���� ���� �Ͻʽÿ�.","û����û�ȳ�"
				Exit Function
			End if
			
		next
   	
   	End with
   	
	DataValidation = true
End Function


		</script>
		<script language="javascript">
		//SMS �߼�
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
		</script>
		
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis"><br>
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table style="WIDTH: 100%; HEIGHT: 24px" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="138" background="../../../images/back_p.gIF"
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
								<td class="TITLE">û����û���� �̸�����</td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<TD>
						<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
							<TR>
								<td class="SEARCHDATA" style="WIDTH: 911px" width="911" colSpan="7">&nbsp;û����û�� <INPUT class="NOINPUTB" id="txtYEARMON1" title="û����û��" style="WIDTH: 96px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="10" size="10" name="txtYEARMON1">&nbsp;~&nbsp;<INPUT class="NOINPUTB" id="txtYEARMON2" title="û����û��" style="WIDTH: 96px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="10" size="10" name="txtYEARMON2">
									�����&nbsp; <INPUT class="NOINPUTB_R" id="txtUSENO" title="������" style="WIDTH: 112px; HEIGHT: 20px"
										accessKey=",NUM" readOnly type="text" maxLength="15" size="13" name="txtUSENO">&nbsp;�� 
									���� û�� ��û�Ͻ� �����Դϴ�.</td>
								<td align="right" ><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="ȭ���� �ݽ��ϴ�."
										src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
							</TR>
						</TABLE>
					</TD>
				</tr>
			</table>
			<BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">�� �� : <INPUT class="NOINPUTB_R" id="txtDIVAMT" title="�����ݾ��հ�" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtDIVAMT"> <INPUT class="NOINPUTB_R" id="txtADJAMT" title="û���ݾ��հ�" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtADJAMT">&nbsp;<INPUT class="NOINPUTB_R" id="txtCHARGE" title="�ܾ��հ�" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtCHARGE">&nbsp;<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="HEIGHT: 20px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<td style="FONT-WEIGHT: bold; FONT-SIZE: 12px" align="right" width="600"><span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">������:</span>
						&nbsp;<INPUT class="NOINPUTB_L" id="txtEMPNAME" title="���α���" style="WIDTH: 96px; HEIGHT: 20px"
							type="text" maxLength="100" size="10" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
							name="ImgEMPNO" title="���α��ڼ���"> <INPUT class="NOINPUTB" id="txtEMPNO" title="���α��ڻ��" style="WIDTH: 58px; HEIGHT: 20px"
							type="text" maxLength="100" size="4" name="txtEMPNO">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgDivDemandOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDivDemand.gIF'" height="20" alt="û����û�� �մϴ�.." src="../../../images/imgDivDemand.gif"
							align="absMiddle" border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF"
							width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</td>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--����-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="30506">
								<PARAM NAME="_ExtentY" VALUE="12435">
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
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
				</TR>
			</table>
		</form>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 0px;HEIGHT: 0px" name="frmSMS"></iframe> <!--DISPLAY: none; -->
	</body>
</HTML>
