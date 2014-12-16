<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_ESTLIST.aspx.vb" Inherits="PD.PDCMJOBMST_ESTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBMST_ESTDTL.aspx
'��      �� : JOBMST�� �ι�° �� - ��/�� �������� ���� �� ���� �Ѵ�. 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
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
		
option explicit

Dim mlngRowCnt, mlngColCnt		
Dim mobjPDCOPREESTLIST
Dim mobjPDCOGET
Dim mobjSCCOGET

Const meTab = 9

'=============================
' �̺�Ʈ ���ν��� 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'��ȸ
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
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

'����
Sub imgDelete_onclick()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF	
End Sub

'��������
Sub imgListcopy_onclick
	Dim vntData
	Dim i
	Dim strPREESTNO
	Dim strPREESTNAME
	Dim strJOBNO, strJOBNAME
	Dim intRtn
	Dim strNEWPREESTNO
	Dim intSaveRtn
	Dim intCnt
	Dim intEDITCODE
	Dim intCount
	
	strNEWPREESTNO = ""
	gFlowWait meWAIT_ON
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","��������ȳ�"
			Exit Sub
		end if
		
		'üũ�� �� �����Ͱ� �ִ��� ������ üũ�Ѵ�.
		intCount = 0
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strPREESTNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO", i)
				strPREESTNAME	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNAME", i)
				
				intCount = intCount + 1
			end if
		next
		
		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if intCount = 0 then
			gErrorMsgBox "������ �����͸� �����Ͻʽÿ�.","��������ȳ�"
			Exit Sub
		elseif intCount > 1 then
			gErrorMsgBox "�����ҵ����ʹ� ���ุ �����Ͻʽÿ�.", "��������ȳ�"
			Exit Sub
		end if
		
		intRtn = gYesNoMsgbox( strPREESTNAME & " �� ������ ���� �Ͻðڽ��ϱ�?","�������� Ȯ��")
		
		IF intRtn <> vbYes then exit Sub
		
		strJOBNO   = parent.document.forms("frmThis").txtJOBNO.value
		strJOBNAME = parent.document.forms("frmThis").txtPRIJOBNAME.value 
		
		intSaveRtn = mobjPDCOPREESTLIST.ProcessRtn_DataCopy(gstrConfigXml,strPREESTNO, strNEWPREESTNO, strJOBNO)
		
		If not gDoErrorRtn ("ProcessRtn_DataCopy") Then
			'��� �÷��� Ŭ����
			gOkMsgBox "����Ǿ����ϴ�.","��������ȳ�!"
			
			.txtFROM.value			= ""
			.txtTO.value			= ""
			.cmbJOBTYPE.value		= ""
			.txtCLIENTCODE1.value	= ""
			.txtCLIENTNAME1.value	= ""
			
			.txtJOBNO.value	  =  strJOBNO
			.txtJOBNAME.value =  strJOBNAME
			
			SelectRtn
			
			For intCnt = 1 To .sprSht.MaxRows 
				If strNEWPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",intCnt) Then
					intEDITCODE = intCnt 
					Exit For
				End If
			Next
			
			mobjSCGLSpr.ActiveCell .sprSht, 1,intEDITCODE
		End If
	end with
	gFlowWait meWAIT_OFF
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
     	end if
	End with
	SelectRtn
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
			
			if not gDoErrorRtn ("GetHIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
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

'-----------------------------------------------------------------------------------------
' ���ڰ��� COMMAND
'-----------------------------------------------------------------------------------------
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

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'���� ������List ��������
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value))
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub
			.txtJOBNO.value = trim(vntRet(0,0))
			.txtJOBNAME.value = trim(vntRet(1,0))
			SelectRtn
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
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO.value = trim(vntData(0,0))
					.txtJOBNAME.value = trim(vntData(1,0))
					SelectRtn
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' SpreadSheet ���� Command
'-----------------------------------------------------------------------------------------
Sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO, strPREESTNO
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			strJOBNO	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strPREESTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",.sprSht.ActiveRow)
			
			parent.document.forms("frmThis").txtPREESTNO.value = strPREESTNO
			
			If strJOBNO = parent.document.forms("frmThis").txtJOBNO.value Then 
				parent.document.forms("frmThis").txtSELECT.value = "T"
			Else
				parent.document.forms("frmThis").txtSELECT.value = "F"
			End If
			parent.jobMst_Call
			
			mobjSCGLSpr.ActiveCell .sprSht, strCol, strRow	
		End If
	End With
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCOLUMN
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	'Ű �����϶� ���ε�

	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				strCOLUMN = "SUMAMT"
			End If

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) Then
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
END SUB

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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")  Then
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

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ
'-----------------------------------------------------------------------------------------	
Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i
	
	set mobjPDCOPREESTLIST	= gCreateRemoteObject("cPDCO.ccPDCOPREESTLIST")
	set mobjPDCOGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 10, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONFIRMGBN | CONFIRMFLAG | JOBNO | JOBNAME | PREESTNAME | SUMAMT | MEMO | PREESTNO | ENDFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|��/��|������|JOBNO|JOB��|������|�����ݾ�|���|������ȣ|û������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|   10|    10|    9|   30|    30|      15|  30|      10|      10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUMAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONFIRMFLAG", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONFIRMGBN | JOBNO | JOBNAME | PREESTNAME | MEMO | PREESTNO | ENDFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "CONFIRMGBN | CONFIRMFLAG | JOBNO | JOBNAME | PREESTNAME | SUMAMT | MEMO | PREESTNO | ENDFLAG"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONFIRMGBN | JOBNO | PREESTNO | ENDFLAG",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0

		InitPageData	
		
		SelectRtn
	End With
End Sub

Sub InitPageData
	'�ʱ� ������ ����
	with frmThis
		'���ڰ��� ��ü��ȸ ����ڿ�û�� ���
		DateClean
		.txtFROM.value = ""
		
		.txtJOBNO.value	  =  parent.document.forms("frmThis").txtJOBNO.value 
		.txtJOBNAME.value =  parent.document.forms("frmThis").txtPRIJOBNAME.value 
		
		Call SEARCHCOMBO_TYPE()
	End with
End Sub

'�������ݱ�
Sub EndPage()
	set mobjPDCOPREESTLIST = Nothing
	set mobjPDCOGET = Nothing
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

'-----------------------------------------------------------------------------------------
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub SEARCHCOMBO_TYPE()'
	Dim vntJOBTYPE
  
   With frmThis   
	'On error resume next
	'Long Type�� ByRef ������ �ʱ�ȭ
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	
	vntJOBTYPE = mobjPDCOPREESTLIST.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt)  'JOB���� ȣ��
	
	if not gDoErrorRtn ("COMBO_TYPE") then 
		mobjSCGLSpr.TypeComboBox = True 
		gLoadComboBox .cmbJOBTYPE,  vntJOBTYPE, False
   	end if    				   		
   end with     
End Sub

'-----------------------------------------------------------------------------------------
' ��ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim intCnt
   	Dim strJOBNAME
	On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strFROM = REPLACE(.txtFROM.value, "-", "")
		strTO	= REPLACE(.txtTO.value, "-", "")
		
		strJOBNAME = REPLACE(.txtJOBNAME.value,"[","[[]")
		
		
		vntData = mobjPDCOPREESTLIST.SelectRtn_List(gstrConfigXml, mlngRowCnt, mlngColCnt, _
													strFROM, strTO, strJOBNAME, Trim(.txtJOBNO.value), _
													.cmbJOBTYPE.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value)
		If not gDoErrorRtn ("SelectRtn_List") then
			'��ȸ�� �����͸� ���ε�
			mobjSCGLSpr.SetClipBinding .sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht.MaxRows '��ȸ�� ������ ó������ ������ ���鼭
					'�������� ��� ���
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMGBN",intCnt) ="������" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HD3FED7, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
				Next
			ELSE
				.sprSht.MaxRows = 0	
			End If
			
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
		End If	
		
		window.setTimeout "AMT_SUM",1	
		.txtSELECTAMT.value = 0
	END WITH
End Sub

'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		
		If .sprSht.MaxRows > 0 Then
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		else
			.txtSUMAMT.value = 0
		End If
	End With
End Sub

'JOBMST ���� ȣ�� û�������� Est_Copy �� ������ �޾� ���⼭ ����ȸ�Ѵ�.  
Sub PreSelectData
	with frmThis
		.txtJOBNO.value = parent.document.forms("frmThis").txtJOBNOVIEW.value    
		.txtJOBNAME.value = parent.document.forms("frmThis").txtPRIJOBVIEW.value   
		SelectRtn
	End with
End Sub

'-----------------------------------------------------------------------------------------
' ����
'-----------------------------------------------------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCount, intRtn, i,intRtn2,lngCnt
	Dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim intChk
	Dim strJOBNO
	Dim intRntChFlag
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","��������ȳ�"
			Exit Sub
		end if
		
		'üũ�� �� �����Ͱ� �ִ��� ������ üũ�Ѵ�.
		intCount = 0
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				intCount = intCount + 1
			end if
		next
		
		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if intCount = 0 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, "�����ȳ�"
			Exit Sub
		end if
		
		for i = 1 to .sprSht.MaxRows
			strJOBNO = "" : strPREESTNO = ""
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strJOBNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO", i)
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO", i)	
				
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				
				vntData = mobjPDCOPREESTLIST.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strJOBNO, strPREESTNO) 
				
				If mlngRowCnt > 0  Then
					gOkMsgBox i & "���� ������ ���λ��� �Ǵ� û���� ����Ǿ����ϴ�. �����Ҽ� �����ϴ�","�����ȳ�!"
					Exit Sub
				End if
			end if
		Next
	
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		lngCnt =0
		intRtn2 = 0

		for i = .sprSht.MaxRows to 1 step -1
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)				
				intRtn2 = mobjPDCOPREESTLIST.DeleteRtn(gstrConfigXml, strPREESTNO)
				
				IF not gDoErrorRtn ("DeleteRtn") then
					lngCnt = lngCnt +1
					mobjSCGLSpr.DeleteRow .sprSht, i
   				End IF
			End If
		next
		
		If lngCnt <> 0 Then
			gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			If .sprSht.MaxRows = 0 Then
				strJOBNO	 = parent.document.forms("frmThis").txtJOBNO.value 
				
				parent.document.forms("frmThis").txtPREESTNO.value  = ""
				
				intRntChFlag = mobjPDCOPREESTLIST.FlagUpdateRtn(gstrConfigXml, strJOBNO)
			End If
		End If
		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		
		SelectRtn
		
		parent.jobMst_Tab2Search
		parent.jobMst_Tab5Search
	End with
	err.clear
End Sub

		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle1" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD0" align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="300" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="65" background="../../../images/back_p.gIF"
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
											<td class="TITLE">��������Ʈ</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start--></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<TABLE class="SEARCHDATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" align="left"
							border="0">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
									width="60">������</TD>
								<TD class="SEARCHDATA" width="214"><INPUT class="INPUT" id="txtFROM" title="�Ⱓ�˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
										accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"> <IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
										border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
									width="50">JOB��</TD>
								<TD class="SEARCHDATA" width="235"><INPUT class="INPUT_L" id="txtJOBNAME" title="���۰����� ��ȸ" style="WIDTH: 145px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="29" name="txtJOBNAME"> <IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgJOBNO">
									<INPUT class="INPUT" id="txtJOBNO" title="���۰����ڵ� ��ȸ" style="WIDTH: 60px; HEIGHT: 22px"
										type="text" maxLength="7" align="left" size="3" name="txtJOBNO"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbJOBTYPE,'')"
									width="45">����</TD>
								<TD class="SEARCHDATA" width="100"><select id="cmbJOBTYPE" style="WIDTH: 100px">
									</select></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
									width="50">������</TD>
								<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="��ȸ�뱤���ָ�" style="WIDTH: 140px; HEIGHT: 22px"
										type="text" maxLength="100" size="16" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
										name="ImgCLIENTCODE1"> <INPUT class="INPUT" id="txtCLIENTCODE1" title="��ȸ�뱤�����ڵ�" style="WIDTH: 57px; HEIGHT: 22px"
										type="text" maxLength="7" size="4" name="txtCLIENTCODE1">
								</TD>
								<TD class="SEARCHDATA" align="right" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
										align="right" border="0" name="imgQuery"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" id="spacebar" style="WIDTH: 100%; HEIGHT: 25px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20">
									<table style="WIDTH: 640px; HEIGHT: 20px" cellSpacing="0" cellPadding="0" width="640" border="0">
										<tr>
											<td class="TITLE">��������Ʈ �հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">&nbsp;&nbsp; 
												�����հ� : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgListcopy" onmouseover="JavaScript:this.src='../../../images/imglistcopyOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imglistcopy.gif'"
													height="20" alt="���������Ǻ��縦 �մϴ�." src="../../../images/imglistcopy.gIF" border="0"
													name="imgListcopy"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="13679">
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
				</tr>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
