<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECEXCUTION.aspx.vb" Inherits="MD.MDCMELECEXCUTION" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ - ����� ������ ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : ��ü���� ������ ����
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMEXCUTION.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ������ ����
'�Ķ�  ���� : 
'Ư��  ���� : ��ü���� �μ�,����,���ͳ�,���� ������ ������ ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/20 By Kim Tae Ho
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECEXCUTION 
Dim mobjMDCMGET
Dim mstrCheck
mstrCheck = True
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
	'location.href ="http://"& meSERVERIP &"/a.html"
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�",""
		gFlowWait meWAIT_OFF
		exit Sub
	end if
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

sub imgDelRow_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn_Dtl
	gFlowWait meWAIT_OFF
end sub
Sub ImgApp_onclick()
	Dim intCnt
	Dim lngSUM
	Dim lngSUMSUM
	lngSUM = 0
	lngSUMSUM = 0
	with frmThis
	if .sprSht.MaxRows =0 Then
		gErrorMsgbox "��ȸ ���ǿ� �´� �����Ͱ� �����Ƿ� �����ϽǼ� �����ϴ�.","����ȳ�!"
		Exit Sub
	End If
	
	For intCnt =1  To .sprSht.MaxRows
	lngSUM = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
	lngSUMSUM = lngSUMSUM + lngSUM
	Next 
	
	If lngSUMSUM = 0 Then
		gErrorMsgbox "�����Ͻ� �����͸� �����Ͻʽÿ�.","����ȳ�!"
		Exit Sub
	End If
	
	If .txtSUSURATE.value <> "" AND .txtSUSURATE.value <> "0" Then
		gFlowWait meWAIT_ON
		Commition_batch
		gFlowWait meWAIT_OFF
	Else 
		gErrorMsgbox "�������� �� �� �����ž� �մϴ�.","����ȳ�!"
	End IF
	End With
End Sub

Sub ImgAppEx_onclick()
	Dim intCnt
	Dim lngSUM
	Dim lngSUMSUM
	lngSUM = 0
	lngSUMSUM = 0
	with frmThis
	if .sprSht.MaxRows =0 Then
		gErrorMsgbox "��ȸ ���ǿ� �´� �����Ͱ� �����Ƿ� �����ϽǼ� �����ϴ�.","����ȳ�!"
		Exit Sub
	End If
	
	For intCnt =1  To .sprSht.MaxRows
	lngSUM = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
	lngSUMSUM = lngSUMSUM + lngSUM
	Next 
	
	If lngSUMSUM = 0 Then
		gErrorMsgbox "�����Ͻ� �����͸� �����Ͻʽÿ�.","����ȳ�!"
		Exit Sub
	End If
	
	If .txtEXCLIENTCODE.value <> "" and  .txtEXCLIENTNAME.value <> "" Then
		gFlowWait meWAIT_ON
		ExClient_batch
		gFlowWait meWAIT_OFF
	Else 
		gErrorMsgbox "����� �ڵ�� ���� �Է��Ͻÿ�.","����ȳ�!"
	End IF
	End With
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
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				
				'mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			'.txtMEDNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	'GetBrandAndDept '������ �������� �������� ���μ��� �����´�.
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
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then
						'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						'mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					'.txtMEDNAME.focus()
					'GetBrandAndDept'������ �������� �������� ���μ��� �����´�.
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
' ����� �ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub EXCLIENTCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtEXCLIENTCODE.value), trim(.txtEXCLIENTNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtEXCLIENTCODE.value = vntRet(0,0) and .txtEXCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtEXCLIENTCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			if .sprSht.ActiveRow >0 Then
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
				'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				
				'mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
			end if
			'.txtMEDNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	'GetBrandAndDept '������ �������� �������� ���μ��� �����´�.
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEXCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(0,0))
					.txtEXCLIENTNAME.value = trim(vntData(1,0))
					if .sprSht.ActiveRow >0 Then
						'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",.sprSht.ActiveRow, .txtCLIENTCODE.value
						'mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						'mobjSCGLSpr.SetTextBinding .sprSht_EXCEL,"CLIENTNAME",.sprSht.ActiveRow, .txtCLIENTNAME.value
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
					end if
					'.txtMEDNAME.focus()
					'GetBrandAndDept'������ �������� �������� ���μ��� �����´�.
				Else
					Call EXCLIENTCODE_POP()
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
'���Ŭ���� �̺�Ʈ
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row > 0 and Col > 1 then		
			'sprShtToFieldBinding Col,Row			
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
'��ưŬ���� �̺�Ʈ
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""
		IF Col = 11 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		
		end if
		.txtYEARMON.focus
		.sprSht.focus 

	End with
	
End Sub
'��Ʈ ����� �̺�Ʈ
Sub sprSht_change(ByVal Col,ByVal Row)
	
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim strQTY,strPRICE,strAMT 
   	Dim intCnt,intCnt0,intCnt1
   	Dim lngSUSU
   	Dim lngSUSUAMT
   	Dim lngRATE
   	Dim lngMCSUSU
   	Dim intColor
   	intColor = ""
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		lngMCSUSU = 0
		IF  Col = 12 Then
		
			strCode		= ""'mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",frmThis.sprSht.ActiveRow)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
			
			vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

			if not gDoErrorRtn ("GetEXCUSTNO") then
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(1,0)			
					'mobjSCGLSpr.CellChanged .sprSht, frmThis.sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					.txtYEARMON.focus
					.sprSht.focus 
					'mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht, 9, Row
					.txtYEARMON.focus
					.sprSht.focus 
				End If
   			end if
   		
		end if
		If Col = 13 Then
		
			if  100 < mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HCCFFFF, &H000000,False
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,0
				Exit Sub
				
			Else
				'intColor = MOD(Row / 2)
				If Row Mod 2 = 0 Then
				
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HF4EDE3, &H000000,False
				
				Else
				
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HFFFFFF, &H000000,False
				End If
			
				'gSetSheetDefaultColor() 
				'gErrorMsgbox "�й���������� �������� ���� Ŭ�� �����ϴ�." & vbcrlf & "��������� ����� ���� Ȯ�� �Ͽ� �ֽʽÿ�.","���� �ȳ�����!"
			End if
			
			if (mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) <> "" Or mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) <> 0) Then 
			lngSUSUAMT = (mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) * mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) ) * 0.01
			lngSUSUAMT = gRound(lngSUSUAMT,0)
			'msgbox mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row)
			'msgbox lngSUSUAMT
			lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) - lngSUSUAMT
			mobjSCGLSpr.SetTextBinding .sprSht,"EXSUSU",Row,lngSUSUAMT
			mobjSCGLSpr.SetTextBinding .sprSht,"MCSUSU",Row,lngMCSUSU
				if mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) = 0.0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,0
				Else
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,1
				End If
			end if
		End if
		
		
		
		
		If Col = 14 Then
		
			if mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) < mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HCCFFFF, &H000000,False
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,0
				Exit Sub
				
			Else
				If Row Mod 2 = 0 Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HF4EDE3, &H000000,False
				
				Else
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, Row, Row,&HFFFFFF, &H000000,False
				End If
			End if
		
			if (mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) <> "" Or mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) <> 0) Then 
			lngRATE = (mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) / mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row)) * 100
			mobjSCGLSpr.SetTextBinding .sprSht,"EXSUSURATE",Row,lngRATE
			lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) - mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row)
			mobjSCGLSpr.SetTextBinding .sprSht,"MCSUSU",Row,lngMCSUSU
				if mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) = 0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,0
				Else
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,1
				End If
			end if
		
		End if
   	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
	AMT_SUM
End Sub	
'��Ʈ Ŭ�� Process
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
dim vntRet, vntInParams
	with frmThis
		IF Col = 9 Then			
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN1") then exit Sub
			Dim strGUBUN
			strGUBUN = ""
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row))
			
			vntRet = gShowModalWindow("../MDCO/MDCMEXCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		end if
		.txtYEARMON.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.sprSht.Focus
	end with
End Sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		'if Row > 0 and Col > 1 then		
		'	sprShtToFieldBinding Col,Row
		
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
		end if
	end with
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()

	'����������ü ����	
	
	set mobjMDCMELECEXCUTION			= gCreateRemoteObject("cMDET.ccMDETELEC_TRAN")
    set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 19, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 10, SPREAD_HEADER, 2, 1
		'CHK,YEARMON,MED_FLAG,CLIENTNAME,REAL_MED_NAME,PROGRAM_NAME,EXCLIENTCODE,BTN,EXCLIENTNAME,EXSUSURATE,EXSUSU,MCSUSU,AMT,SUSURATE,SUSU
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|YEARMON|SEQ|MED_FLAG|CLIENTCODE|CLIENTNAME|REAL_MED_CODE|REAL_MED_NAME|PROGRAM_NAME|EXCLIENTCODE|BTN|EXCLIENTNAME|EXSUSURATE|EXSUSU|MCSUSU|AMT|SUSURATE|SUSU|INPUT_MEDFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		"����|���|����|��ü����|�������ڵ�|�����ָ�|��ü���ڵ�|��ü���|���౸��|���۴����ڵ�|���۴�����|���۴���  ��������|���۴���   ������|MC������|����ݾ�|��������|������|��ü�����ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|0   |0   |4       |0         |15      |0         |16      |0       |8         |2|15          |8                 |10               |10      |11      |8       |10     |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "EXSUSURATE", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "EXSUSU|AMT|SUSURATE|SUSU|MCSUSU", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "EXCLIENTCODE|EXCLIENTNAME",-1,-1,100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "AMT|SUSURATE|SUSU"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "YEARMON|MED_FLAG|CLIENTNAME|REAL_MED_NAME|PROGRAM_NAME", -1, -1, 40
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON|SEQ|CLIENTCODE|REAL_MED_CODE|INPUT_MEDFLAG|PROGRAM_NAME", true
		.sprSht.style.visibility = "visible"
		
		
		gSetSheetColor mobjSCGLSpr, .sprShtSum
		mobjSCGLSpr.SpreadLayout .sprShtSum, 19, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprShtSum, "CHK|YEARMON|SEQ|MED_FLAG|CLIENTCODE|CLIENTNAME|REAL_MED_CODE|REAL_MED_NAME|PROGRAM_NAME|EXCLIENTCODE|BTN|EXCLIENTNAME|EXSUSURATE|EXSUSU|MCSUSU|AMT|SUSURATE|SUSU|INPUT_MEDFLAG"
		mobjSCGLSpr.SetText .sprShtSum, 4, 1, "�հ�"
	    mobjSCGLSpr.SetScrollBar .sprShtSum, 0
	    mobjSCGLSpr.SetBackColor .sprShtSum,"1|2|3|4",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "EXSUSU|AMT|SUSURATE|SUSU|MCSUSU", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprShtSum, "YEARMON|SEQ|CLIENTCODE|REAL_MED_CODE|INPUT_MEDFLAG|PROGRAM_NAME ", true
		
		mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "13"
	    mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub
'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub

'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
	End with
end sub
'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
'EXSUSU|  MCSUSU    |AMT|SUSU
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	Dim lngEXSUSU,lngEXSUSUSUM,lngMCSUSU, lngMCSUSUSUM
	With frmThis
		IntAMTSUM = 0
		IntPRICESUM = 0
		lngEXSUSUSUM = 0
		lngMCSUSUSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntPRICE = 0
			lngEXSUSU = 0
			lngMCSUSU = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU", lngCnt)
			lngEXSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"EXSUSU", lngCnt)
			lngMCSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"MCSUSU", lngCnt)
			
			
			lngEXSUSUSUM = lngEXSUSUSUM + lngEXSUSU
			lngMCSUSUSUM = lngMCSUSUSUM + lngMCSUSU
			IntAMTSUM = IntAMTSUM + IntAMT
			IntPRICESUM = IntPRICESUM + IntPRICE
		Next
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprShtSum,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprShtSum,"SUSU",1, IntPRICESUM	
			mobjSCGLSpr.SetTextBinding .sprShtSum,"EXSUSU",1, lngEXSUSUSUM
			mobjSCGLSpr.SetTextBinding .sprShtSum,"MCSUSU",1, lngMCSUSUSUM	
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		end if
	End With
End Sub

Sub EndPage()
	set mobjMDCMELECEXCUTION = Nothing
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
		.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		.txtYEARMON.focus
		
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	dim vntData
	Dim strYEARMON, strCUSTCODE, strGUBUN
	Dim strMED_FLAG,strGFLAGNAME,strPROGRAMNAME
   	Dim i, strCols
	'on error resume next
	
	with frmThis
	
		strYEARMON	= .txtYEARMON.value
		strCUSTCODE	= .txtCLIENTCODE.value
		strGUBUN = "F" '���౸��
		strMED_FLAG = .cmbMEDFLAG.value '��ü����
		'strGFLAGNAME =  "J"
		'strPROGRAMNAME = .txtPROGRAMNAME.value
		'If strGUBUN <> "C" Then
		'	msgbox "���� �μ��ü �� ���� �մϴ�."
		'	Exit Sub
		'End If
		'�ʱ�ȭ
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		'���,�������ڵ�,���౸��,��ü����/strYEARMON, strCUSTCODE,strGUBUN,strMED_FLAG
		vntData = mobjMDCMELECEXCUTION.EXLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCUSTCODE,strGUBUN,strMED_FLAG)
		
		IF not gDoErrorRtn ("EXLIST") then
			'��ȸ�� �����͸� ���ε�
			'call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			'�ʱ� ���·� ����
			If mlngRowCnt > 0 Then
			AMT_SUM
			Else
			.sprSht.MaxRows = 0
			End If
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			
			
		End IF
	end with
End Sub

Sub PreSearchFiledValue (strCUSTCODE, strCUSTNAME)
	frmThis.txtYEARMON.value = strCUSTCODE
	frmThis.txtCLIENTCODE.value = strCUSTNAME		
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
  		
  		''EXCLIENTCODE,EXCLIENTNAME,EXSUSURATE,EXSUSU
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",intCnt) = "" _
			 AND (mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",intCnt) = "" _
			 AND mobjSCGLSpr.GetTextBinding(.sprSht,"EXSUSURATE",intCnt) = 0.0 _
			 AND mobjSCGLSpr.GetTextBinding(.sprSht,"EXSUSU",intCnt) = 0 _
			 AND mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1)  Then 
					gErrorMsgBox intCnt & " ��° ���� �Է³��� �� Ȯ���Ͻʽÿ�" & vbcrlf & "�ʼ������� ������ڵ�,��,�����������,��������� �Դϴ�.","�Է¿���"
					Exit Function
			 End if
		next
   		
   		
   	End with
	DataValidation = true
End Function
'------------------------------------------
' ������ ���� Process
'------------------------------------------
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	with frmThis
   		'������ Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
			 lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,1,intCnt)
			 lngCHKSUM = lngCHKSUM + lngCHK
		Next
		If lngCHKSUM = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		
		if DataValidation =false then exit sub
	    '������ Validation End
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|SEQ|INPUT_MEDFLAG|CLIENTCODE|REAL_MED_CODE|EXCLIENTCODE|EXSUSURATE|EXSUSU|MCSUSU")
		
		intRtn = mobjMDCMELECEXCUTION.EXCUTION_ProcessRtn(gstrConfigXml,vntData)
	
		if not gDoErrorRtn ("EXCUTION_ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ ����" & mePROC_DONE
			'gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
   		end if
   		
   	end with
End Sub
'------------------------------------------
' ������ ���� Process
'------------------------------------------
Sub DeleteRtn_Dtl
	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	with frmThis
   		'������ Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","�����ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
			 lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,1,intCnt)
			 lngCHKSUM = lngCHKSUM + lngCHK
		Next
		If lngCHKSUM = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","�����ȳ�!"
			Exit Sub
		End If
		
		
		if DataValidation =false then exit sub
	    '������ Validation End
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|SEQ|INPUT_MEDFLAG|CLIENTCODE|REAL_MED_CODE")
		
		intRtn = mobjMDCMELECEXCUTION.EXCUTION_DeleteRtn(gstrConfigXml,vntData)
	
		if not gDoErrorRtn ("EXCUTION_DeleteRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
   		end if
   		
   	end with
End Sub

Sub Commition_batch
Dim intCnt
	with frmThis
		If Cint(.txtSUSURATE.value) > 99  Then
			gErrorMsgbox "���������� 100 ���Ͽ��� �մϴ�.","����ȳ�!"
			.txtSUSURATE.value = ""
			.txtSUSURATE.focus()
			Exit Sub
		end If
		
		for intCnt= 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXSUSURATE",intCnt, .txtSUSURATE.value	
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",intCnt, .txtEXCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",intCnt, .txtEXCLIENTNAME.value	
				sprSht_change 13,intCnt
			End If
		Next
	End With
	
End Sub


Sub ExClient_batch
	Dim intCnt
	with frmThis
		for intCnt= 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",intCnt, .txtEXCLIENTCODE.value
				mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",intCnt, .txtEXCLIENTNAME.value	
				sprSht_change 11,intCnt
			End If
		Next
	End With
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 1040px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
				width="684" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
									<td align="left">
										<TABLE cellSpacing="0" cellPadding="0" width="171" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���� ���۴���� ������ ����</td>
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
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 82px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="82">�� ��</TD>
											<TD class="SEARCHDATA" width="89" style="WIDTH: 89px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" width="90" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,txtCLIENTNAME)">������
											</TD>
											<TD class="SEARCHDATA" width="329" style="WIDTH: 329px"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 240px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="34" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE">
											</TD>
											<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" width="85">��ü����</TD>
											<TD class="SEARCHDATA" style="WIDTH: 276px; CURSOR: hand"><SELECT id="cmbMEDFLAG" title="��ü����" style="WIDTH: 104px" name="cmbMEDFLAG">
													<OPTION value="X" selected>��ü</OPTION>
													<OPTION value="TV">TV</OPTION>
													<OPTION value="RADIO">RADIO</OPTION>
													<OPTION value="DMB">DMB</OPTION>
												</SELECT></TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" width="54" border="0" name="imgQuery" align="right"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
														<td><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'"
																alt="�� �� ����" src="../../../images/imgDelRow.gif" border="0" name="imgDelRow"></td>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTCODE, txtEXCLIENTNAME)">���۴����</TD>
											<TD class="SEARCHDATA" style="WIDTH: 517px"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="�ڵ��" style="WIDTH: 184px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="25" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgEXCLIENTCODE"> <INPUT class="INPUT_L" id="txtEXCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtEXCLIENTCODE"><IMG id="ImgAppEx" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="���۴���� �� �����Ḧ �ϰ����� �մϴ�." src="../../../images/ImgApp.gif" width="54" align="absMiddle"
													border="0" name="ImgAppEx"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">��������</TD>
											<TD class="SEARCHDATA" style="WIDTH: 340px"><INPUT class="INPUT_R" id="txtSUSURATE" style="WIDTH: 104px; HEIGHT: 22px" accessKey="NUM"
													type="text" size="12" name="txtSUSURATE"><IMG id="ImgApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="���۴���� �� �����Ḧ �ϰ����� �մϴ�." src="../../../images/ImgApp.gif"
													width="54" align="absMiddle" border="0" name="ImgApp"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"></TD>
							</TR>
							<TR>
								<TD align="center">
									<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LISTFRAME" style="HEIGHT: 684px" height="101">
												<OBJECT id="sprSht" style="WIDTH: 1038px; HEIGHT: 660px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
													VIEWASTEXT>
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="27464">
													<PARAM NAME="_ExtentY" VALUE="17463">
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
												<OBJECT id="sprShtSum" style="WIDTH: 1038px; HEIGHT: 23px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
													VIEWASTEXT>
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="27464">
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
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
