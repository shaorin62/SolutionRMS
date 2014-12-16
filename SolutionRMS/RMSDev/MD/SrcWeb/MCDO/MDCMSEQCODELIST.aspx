<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMSEQCODELIST.aspx.vb" Inherits="MD.MDCMSEQCODELIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�귣���ڵ� ����</title>
		<meta content="True" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : �귣���ڵ� ��� ȭ��(MDCMSEQCODELIST)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMSEQCODELIST.aspx
'��      �� : �����ֿ� ���� ���� �귣�� �ڵ带 ��� �Ҽ� �ִ� ȭ��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/25 By Kim Tae Ho
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
Dim mobjMDCMCODETR,mobjMDCMGET
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mstrCheck

mstrCheck=True


Dim mstrInsert 
mstrInsert = False
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgClose_onclick
	EndPage
End Sub
Sub imgNew_Onclick
	'initpageData
	'SelectRtn
	frmThis.txtSEQNO.value = ""
	frmThis.txtSEQNAME.value = ""
	frmThis.txtDEPTCD.value = ""
	frmThis.txtDEPTNAME.value = ""
	frmThis.txtATTR02.value = ""
	frmThis.txtCUSTCODE.value = ""
	frmThis.txtCUSTNAME.value = "" 
	frmThis.txtCLIENTSUBCODE.value = ""
	frmThis.txtCLIENTSUBNAME.value = ""
	mstrInsert = True
	with frmThis
	.sprSht.MaxRows = 0
	End with
	call sprSht_Keydown(meINS_ROW, 0)
	'CLIENTCODE_NEWPOP
	
	
End Sub
Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	End With
	gFlowWait meWAIT_OFF
End Sub
Sub imgSave_Onclick
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgQuery_Onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub txtSEARCHCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		SelectRtn
	end if
End Sub
Sub sprSht_Keydown(KeyCode, Shift)
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
	Else
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW':
						'SetDefaultNewRow
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub
'-----------------------------------------------------------------------------------------
' �귣���ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtSEARCHSEQCODE.value), trim(.txtSEARCHSEQNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		if isArray(vntRet) then
			if .txtSEARCHSEQCODE.value = vntRet(0,0) and .txtSEARCHSEQNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			.txtSEARCHSEQCODE.value = trim(vntRet(1,0))		' �귣�� ǥ��
			.txtSEARCHSEQNAME.value = trim(vntRet(2,0))	' �귣��� ǥ��
			.txtSEARCHCUSTCODE.value = trim(vntRet(3,0))
			.txtSEARCHCUSTNAME.value = trim(vntRet(4,0))
			gSetChangeFlag .txtSEARCHSEQCODE
     	end if
	End with
	gSetChange
End Sub

Sub txtSEARCHSEQNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSEARCHSEQCODE.value),trim(.txtSEARCHSEQNAME.value),"",trim(.txtSEARCHCUSTNAME.value))
			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					.txtSEARCHSEQCODE.value = trim(vntData(1,0))
					.txtSEARCHSEQNAME.value = trim(vntData(2,0))
					.txtSEARCHCUSTCODE.value = trim(vntData(3,0))
					.txtSEARCHCUSTNAME.value = trim(vntData(4,0))
				Else
					Call SUBSEQCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' ������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCUSTCODE.value), trim(.txtCUSTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCUSTCODE.value = trim(vntRet(5,0))
			.txtCUSTNAME.value = trim(vntRet(6,0))
			txtCUSTCODE_onchange
			txtCUSTNAME_onchange
			txtCLIENTSUBCODE_onchange
			txtCLIENTSUBNAME_onchange
			'if .sprSht.ActiveRow >0 Then
			'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
			'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
			'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow, .txtCLIENTSUBCODE.value
			'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",.sprSht.ActiveRow, .txtCLIENTSUBNAME.value
			'	
			''	mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
			'	mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTSUBNAME",.sprSht.ActiveRow, .txtCLIENTSUBNAME.value
			'	
			'	mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			'end if
			
			.txtDEPTNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCUSTCODE.value),trim(.txtCUSTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCUSTCODE.value = trim(vntData(5,0))
					.txtCUSTNAME.value = trim(vntData(6,0))
						txtCUSTCODE_onchange
						txtCUSTNAME_onchange
						txtCLIENTSUBCODE_onchange
						txtCLIENTSUBNAME_onchange
					'if .sprSht.ActiveRow >0 Then
					'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
					'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
					'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",.sprSht.ActiveRow, .txtCLIENTSUBCODE.value
					'	mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",.sprSht.ActiveRow, .txtCLIENTSUBNAME.value
					'	
					'	mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
					'	mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTSUBNAME",.sprSht.ActiveRow, .txtCLIENTSUBNAME.value
					'	
					'	mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					'end if
					.txtDEPTNAME.focus()
					gSetChangeFlag .txtCLIENTSUBCODE
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
Sub txtCUSTCODE_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CUSTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCUSTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCUSTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CUSTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCUSTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCLIENTSUBCODE_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtCLIENTSUBNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTSUBNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTSUBNAME.value
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
Sub txtDEPTNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPTNAME",frmThis.sprSht.ActiveRow, frmThis.txtDEPTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtSEQNAME_onchange
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SEQNAME",frmThis.sprSht.ActiveRow, frmThis.txtSEQNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub txtATTR02_onchnage
	if frmThis.sprSht.ActiveRow >0  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ATTR02",frmThis.sprSht.ActiveRow, frmThis.txtATTR02.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub ImgCLIENTSUBApp_onclick
	Dim intCnt
	With frmThis
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",intCnt,.txtCLIENTSUBCODE.value 
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBNAME",intCnt,.txtCLIENTSUBNAME.value 
				
				mobjSCGLSpr.CellChanged .sprSht, 4,intCnt
			End If
		Next
	End With
End Sub



Sub ImgDEPTApp_onclick
	Dim intCnt
	With frmThis
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTCD",intCnt,.txtDEPTCD.value 
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPTNAME",intCnt,.txtDEPTNAME.value 
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,"1" 
			End If
		Next
	End With
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
Dim vntInParam
Dim intNo
	'����������ü ����	
	set mobjMDCMCODETR	= gCreateRemoteObject("cMDCO.ccMDCOCODETR")			'�����ڵ�
	set mobjMDCMGET     = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "208px"

	pnlTab1.style.left= "7px"
	mobjSCGLCtl.DoEventQueue
	
	
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor
    with frmThis
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		for i = 0 to intNo
			select case i
				case 0 : .txtSEARCHCUSTCODE.value = vntInParam(i)
				case 1 : .txtSEARCHCUSTNAME.value = vntInParam(i)		
			end select
		next
		'msgbox .txtSEARCHCUSTCODE.value
		.txtATTR01.style.visibility = "hidden"
		'.txtCLIENTCODE.style.visibility = "hidden"
        .txtACCCUSTCODE.style.visibility = "hidden"
		'ȭ���� �������� �����ϱ� ����(Tab�� ���� ó���� ǥ�õǴ� �͸� ��)
		'.sprSht.style.visibility = "hidden"
		
		'**************************************************
		'***ù��° Sheet ������
		'**************************************************
		
		'Sheet Į�� ����
	    gSetSheetColor mobjSCGLSpr, .sprSht
		'Sheet Layout ������
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0,0,0,2
		'Binding Field ����
	    mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|CUSTCODE  |CUSTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SEQNO|SEQNAME|ACCCUSTCODE|DEPTCD|DEPTNAME|ATTR02|ATTR01"
		'Header ������
		mobjSCGLSpr.SetHeader .sprSht,        "����|�������ڵ�|�����ָ�|������ڵ�|����θ�|�귣���ڵ�|�귣��� |ȸ���ڵ�|���μ��ڵ�|���μ���|���|���屸��",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1","4   |9         |22      |9         |11      |9         |20       |0       | 10         |15        |13  |0"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SEQNO|CUSTCODE|DEPTCD", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ACCCUSTCODE", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SEQNAME|CUSTNAME|DEPTNAME|ATTR02", -1, -1, 200
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellsLock2 .sprSht, true,"SEQNO|SEQNAME|CUSTCODE  |CUSTNAME|ACCCUSTCODE   |DEPTCD  |DEPTNAME|ATTR02|ATTR01|CLIENTSUBCODE|CLIENTSUBNAME|ATTR01"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQNAME|CUSTNAME|DEPTNAME|ATTR02",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ACCCUSTCODE",-1,-1,1,2,false
		mobjSCGLSpr.ColHidden .sprSht, "ATTR01|ACCCUSTCODE",TRUE
	End with

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	'InitPageData	
	
	SelectRtn
End Sub
'��ȸ
Sub SelectRtn
	Dim vntData
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjMDCMCODETR.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtSEARCHSEQCODE.value,.txtSEARCHSEQNAME.value,.txtSEARCHCUSTCODE.value,.txtSEARCHCUSTNAME.value)
		if not gDoErrorRtn ("SelectRtn") then
			'mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			if mlngRowCnt > 0 Then
			mobjSCGLSpr.SetClipBinding .sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True
			Else
			initpageData
			End If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			sprShtToFieldBinding 1,1
   		end if
   	end with
   	mstrInsert = False
End Sub
Sub SelectRtn_Dup
	Dim vntData
   	Dim i, strCols
   	Dim strCODE
	'On error resume next

End Sub
'����
'-----------------------------------------------------------------------------------------
' ������ ó�� 
'-----------------------------------------------------------------------------------------
Sub ProcessRtn ()
  	Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	with frmThis
	'On error resume next
  		'������ Validation
		if DataValidation =false then exit sub
		strCODE = .txtSEQNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|CUSTCODE|CUSTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SEQNO|SEQNAME|DEPTCD|DEPTNAME")
		if  not IsArray(vntData) then 
		gErrorMsgBox "����� " & meNO_DATA,"�������"
		exit sub
		End If
		'vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
      'if  not IsArray(vntData) then 
      'gErrorMsgBox "����� " & meNO_DATA,"�������"
      'exit sub
        'end if
		'ó�� ������ü ȣ��
		strMasterData = gXMLGetBindingData (xmlBind)
		
		if .txtSEQNO.value = "" then
			strSEQFlag = "new"
			intRtn = mobjMDCMCODETR.ProcessRtn(gstrConfigXml,strMasterData, strSEQFlag)
		else
			intRtn = mobjMDCMCODETR.ProcessRtnSheet(gstrConfigXml,vntData)
		end if
		

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			if strSEQFlag = "new" then
				gErrorMsgBox " �ڷᰡ �ű�����" & mePROC_DONE,"����ȳ�" 
			else
				gErrorMsgBox " �ڷᰡ" & intRtn & " �� ��������" & mePROC_DONE,"����ȳ�" 
			end if
			SelectRtn
  		end if
 	end with
End Sub

'-----------------------------------------------------------------------------------------
' ������ ó�� �� ���� ������ ����
'-----------------------------------------------------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
		
   	End with
	DataValidation = true
End Function

Sub EndPage()
	set mobjMDCMCODETR = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub


Sub CLIENTCODE_NEWPOP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtCUSTCODE.value, .txtCUSTNAME.value) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCUSTCODE.value = vntRet(0,0)  ' Code�� ����
			.txtCUSTNAME.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtACCCUSTCODE.value = vntRet(4,0) 'ȸ���ڵ� ǥ��
			.txtSEARCHCUSTNAME.value = vntRet(1,0)
							' ��Ŀ�� �̵�
			gSetChangeFlag .txtCUSTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
	SelectRtn
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgSEARCHCUST_onclick
	Call CLIENTCODESEARCH_POP()
End Sub
'���� ������List ��������
Sub CLIENTCODESEARCH_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtSEARCHCUSTCODE.value, .txtSEARCHCUSTNAME.value) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMBRANDCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtSEARCHCUSTCODE.value = vntRet(0,0) and .txtSEARCHCUSTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtSEARCHCUSTCODE.value = vntRet(0,0)  ' Code�� ����
			.txtSEARCHCUSTNAME.value = vntRet(1,0)  ' �ڵ�� ǥ��
			
     	end if
	End with
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtSEARCHCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtSEARCHCUSTCODE.value,.txtSEARCHCUSTNAME.value)
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtSEARCHCUSTCODE.value = vntData(0,0)
					.txtSEARCHCUSTNAME.value = vntData(1,0)
				Else
					Call CLIENTCODESEARCH_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub
'���� ������List ��������
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(.txtCUSTCODE.value, .txtCUSTNAME.value) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMBRANDCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCUSTCODE.value = vntRet(0,0) and .txtCUSTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCUSTCODE.value = vntRet(0,0)  ' Code�� ����
			.txtCUSTNAME.value = vntRet(1,0)  ' �ڵ�� ǥ��
			txtCUSTCODE_onchange
			txtCUSTNAME_onchange
			'.txtACCCUSTCODE.value = vntRet(4,0) 'ȸ���ڵ� ǥ��
							' ��Ŀ�� �̵�
			gSetChangeFlag .txtCUSTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCUSTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_HIGHCUST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCUSTCODE.value,.txtCUSTNAME.value)
			if not gDoErrorRtn ("txtCUSTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCUSTCODE.value = vntData(0,0)
					.txtCUSTNAME.value = vntData(1,0)
					txtCUSTCODE_onchange
					txtCUSTNAME_onchange
					'.txtACCCUSTCODE.value = vntRet(3,0)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' ����μ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------	
Sub ImgDEPTCODE_onclick
	Call JOBREQU_DEPTCD_POP()
End Sub
Sub JOBREQU_DEPTCD_POP
	Dim vntRet, vntInParams
	with frmThis
		'LOC,OC,MU,PU,CC Type,CC �ڵ�/��,optional(�����뿩��,���˻���,�߰���ȸ �ʵ�,Key Like����)
		vntInParams = array(trim(.txtDEPTNAME.value))
		'vntRet = gShowModalWindow("PDCMCC.aspx",vntInParams , 413,440)
		vntRet = gShowModalWindow("../MDCO/MDCMDEPTPOP.aspx",vntInParams , 413,440)
		if isArray(vntRet) then
		    .txtDEPTCD.value = vntRet(0,0)	'Code�� ����
			.txtDEPTNAME.value = vntRet(1,0)	'�ڵ�� ǥ��
			txtDEPTCD_onchange
			txtDEPTNAME_onchange
			'.txtATTR02.focus()
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
			vntData = mobjMDCMGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPTNAME.value))
			if not gDoErrorRtn ("txtJOBREQU_DEPTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
						.txtDEPTCD.value = vntData(0,0)
						.txtDEPTNAME.value = vntData(1,0)
						txtDEPTCD_onchange
						txtDEPTNAME_onchange
						'.txtATTR02.focus()
				Else
					Call JOBREQU_DEPTCD_POP()
				End If
   			end if
   		end with
	window.event.returnValue = false
	window.event.cancelBubble = true
	End If
End Sub
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ Ŭ���� 
'-----------------------------------------------------------------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	mstrInsert = False
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
	if Row > 0 and Col > 1 then		
			sprShtToFieldBinding Col,Row
	elseif Col = 1 and Row = 0 then
		mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
		if mstrCheck = True then 
			mstrCheck = False	
		elseif mstrCheck = False then 
			mstrCheck = True
		end if	
	end if 
	End With
End Sub  
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ���� Ŭ���� 
'-----------------------------------------------------------------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
'-----------------------------------------------------------------------------------------
' �������� ��Ʈ ����� üũ 
'-----------------------------------------------------------------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub
'-----------------------------------------------------------------------------------------
' ��Ʈ ���ε� 
'-----------------------------------------------------------------------------------------
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	dim vntData
	dim strSEQCODE
	with frmThis
	
		strSEQCODE=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQNO",Row)
		
		if strSEQCODE ="" Then EXIT Function
		
		vntData = mobjMDCMCODETR.SelectRtn_sprSht (gstrConfigXml, Row,Col ,strSEQCODE)
		
		IF not gDoErrorRtn ("SelectRtn_sprSht") then
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			.txtCLIENTSUBCODE.focus()
			.sprSht.focus()
		End IF
	
	END WITH
End Function
'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
gClearAllObject frmThis
with frmThis
.sprSht.MaxRows = 0
end With
gXMLNewBinding frmThis,xmlBind,"#xmlBind"		
With frmThis
.txtSEQNAME.focus()
End With
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="790" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="WIDTH: 790px; HEIGHT: 54px">
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
											<td class="TITLE">&nbsp;�귣��&nbsp;����</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 390px" vAlign="middle" align="right" height="28">
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
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 790px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 790px; HEIGHT: 15px" vAlign="top" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TBODY>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 105px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSEARCHCUSTCODE,txtSEARCHCUSTNAME)">������</TD>
												<TD class="SEARCHDATA" style="WIDTH: 257px"><INPUT class="INPUTB_L" id="txtSEARCHCUSTNAME" title="�����ָ�" style="WIDTH: 160px; HEIGHT: 21px"
														type="text" size="21" name="txtSEARCHCUSTNAME"><IMG id="ImgSEARCHCUST" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
														width="23" align="absMiddle" border="0" name="ImgSEARCHCUST"><INPUT class="INPUTB_L" id="txtSEARCHCUSTCODE" title="�������ڵ�" style="WIDTH: 72px; HEIGHT: 21px"
														type="text" size="5" name="txtSEARCHCUSTCODE"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSEARCHSEQCODE,txtSEARCHSEQNAME)">�귣��
												</TD>
												<TD class="SEARCHDATA" style="WIDTH: 297px"><INPUT class="INPUT_L" id="txtSEARCHSEQNAME" style="WIDTH: 186px; HEIGHT: 22px" tabIndex="0"
														type="text" size="25" name="txtSEARCHSEQNAME"><IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
														align="absMiddle" border="0" name="ImgSUBSEQCODE"><INPUT class="INPUT_L" id="txtSEARCHSEQCODE" title="�귣���" style="WIDTH: 72px; HEIGHT: 22px"
														tabIndex="0" type="text" size="6" name="txtSEARCHSEQCODE"><INPUT dataFld="ATTR01" id="txtATTR01" style="WIDTH: 5px; HEIGHT: 21px" dataSrc="#xmlBind"
														type="hidden" size="1" name="txtATTR01"><INPUT dataFld="ACCCUSTCODE" class="NOINPUT_L" id="txtACCCUSTCODE" title="������ ID" style="WIDTH: 8px; HEIGHT: 22px"
														dataSrc="#xmlBind" tabIndex="0" readOnly type="text" size="1" name="txtACCCUSTCODE"></TD>
												<td class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" border="0" name="imgQuery"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgClose.gIF"
														border="0" name="imgClose"></td>
											</TR>
										</TBODY>
									</TABLE>
									<table class="DATA" height="28" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 790px; HEIGHT: 25px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;�귣��&nbsp;���</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 390px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																src="../../../images/imgNew.gIF" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 790px; HEIGHT: 2px"></TD>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 790px; HEIGHT: 22px" vAlign="top" align="center">
									<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCUSTCODE,txtCUSTNAME)">������</TD>
											<TD class="DATA" width="300"><INPUT dataFld="CUSTNAME" class="INPUTB_L" id="txtCUSTNAME" title="�����ָ�" style="WIDTH: 157px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="text" size="28" name="txtCUSTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT dataFld="CUSTCODE" class="NOINPUT" id="txtCUSTCODE" title="�������ڵ�" style="WIDTH: 64px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" readOnly type="text" size="5" name="txtCUSTCODE"></TD>
											<TD class="LABEL" width="70" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSEQNO,'')">
												�귣��
											</TD>
											<TD class="DATA" width="350"><INPUT dataFld="SEQNAME" class="INPUT_L" id="txtSEQNAME" title="�ڵ��" style="WIDTH: 192px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" tabIndex="0" type="text" size="26" name="txtSEQNAME"><INPUT dataFld="SEQNO" class="NOINPUT" id="txtSEQNO" title="�ڵ�" style="WIDTH: 72px; HEIGHT: 22px"
													dataSrc="#xmlBind" tabIndex="0" type="text" maxLength="4" size="6" name="txtSEQNO" readOnly>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" title="����θ� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBCODE,txtCLIENTSUBNAME)">�����</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTSUBNAME" class="INPUTB_L" id="txtCLIENTSUBNAME" title="�μ���" style="WIDTH: 157px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="text" size="20" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTSUBCODE"><INPUT dataFld="CLIENTSUBCODE" class="NOINPUT" id="txtCLIENTSUBCODE" title="������ڵ�" style="WIDTH: 64px; HEIGHT: 21px"
													dataSrc="#xmlBind" readOnly type="text" size="5" name="txtCLIENTSUBCODE"><IMG id="ImgCLIENTSUBApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="����θ� �ϰ������մϴ�" src="../../../images/ImgApp.gif" width="54"
													align="absMiddle" border="0" name="ImgCLIENTSUBApp"></TD>
											<TD class="LABEL" title="���μ��� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEPTCD,txtDEPTNAME)">���μ�</TD>
											<TD class="DATA"><INPUT dataFld="DEPTNAME" class="INPUTB_L" id="txtDEPTNAME" title="�μ���" style="WIDTH: 192px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="text" size="26" name="txtDEPTNAME"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgDEPTCODE"><INPUT dataFld="DEPTCD" class="NOINPUT" id="txtDEPTCD" title="�μ��ڵ�" style="WIDTH: 64px; HEIGHT: 21px"
													accessKey=",M" dataSrc="#xmlBind" readOnly type="text" size="5" name="txtDEPTCD">
												<IMG id="ImgDEPTApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
													height="20" alt="���μ��� �ϰ� �����մϴ�" src="../../../images/ImgApp.gif" align="absMiddle"
													border="0" name="ImgDEPTApp">
											</TD>
										</TR>
										<TR>
											<TD class="LABEL" title="��� �����մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtATTR02,'')">��&nbsp;&nbsp; 
												��</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="ATTR02" class="INPUT_L" id="txtATTR02" title="�귣��" style="WIDTH: 714px; HEIGHT: 22px"
													dataSrc="#xmlBind" tabIndex="0" type="text" name="txtATTR02" size="113"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 790px"></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 790px; HEIGHT: 400px" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 400px"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 398px"
											width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="20876">
											<PARAM NAME="_ExtentY" VALUE="10530">
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
							<!--BodySplit Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 790px"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End-->
					</TD>
				</TR>
				<!--Top TR End-->
			</TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
