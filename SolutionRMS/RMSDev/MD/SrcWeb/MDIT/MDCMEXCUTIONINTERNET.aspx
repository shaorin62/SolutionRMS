<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMEXCUTIONINTERNET.aspx.vb" Inherits="MD.MDCMEXCUTIONINTERNET" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ������ ����</title> 
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
Dim mobjMDCMINTERNETEXCUTION 
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

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgQuery_onclick
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�",""
		exit Sub
	end if
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
	
	If .txtEXCLIENTCODE.value <> "" and .txtEXCLIENTNAME.value <> ""  Then
		gFlowWait meWAIT_ON
		ExClient_batch
		gFlowWait meWAIT_OFF
	Else 
		gErrorMsgbox "����� �ڵ�� ���� �Է��Ͻÿ�.","����ȳ�!"
	End IF
	End With
End Sub



'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					SELECTRTN
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ����� �ڵ��˾� ��ư
'-----------------------------------------------------------------------------------------
Sub ImgEXCLIENTCODE_onclick
	Call EXCLIENTCODE_POP()
End Sub

Sub EXCLIENTCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtEXCLIENTCODE.value),trim(.txtEXCLIENTNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMEXEALLPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtEXCLIENTCODE.value = trim(vntRet(1,0))	'Code�� ����
			.txtEXCLIENTNAME.value = trim(vntRet(2,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtEXCLIENTCODE
		End If
	end With
End Sub


Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)

			vntData = mobjMDCOGET.Get_EXCLIENT_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEXCLIENTCODE.value,.txtEXCLIENTNAME.value,"")
		
			If not gDoErrorRtn ("Get_EXCLIENT_ALL") Then
				If mlngRowCnt = 1 Then
					.txtEXCLIENTCODE.value = trim(vntData(1,1))	'Code�� ����
					.txtEXCLIENTNAME.value = trim(vntData(2,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call EXCLIENTCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
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
		if Col = 1 then exit sub
		IF Col = 9 Then
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
   	Dim strQTY, strAMT 
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
		IF  Col = 10 Then
			strCode		= ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"EXCLIENTNAME",Row)
			
			IF strCodeName <> "" THEN
				vntData = mobjMDCMGET.GetEXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

				if not gDoErrorRtn ("GetEXCUSTNO") then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"EXCLIENTNAME",Row, vntData(1,0)			
						.txtYEARMON.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc .sprSht, 9, Row
						.txtYEARMON.focus
						.sprSht.focus 
					End If
   				end if
   			END IF
		end if
		
		If Col = 11 Then
			if  100 < mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) Then
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
			
			if (mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) <> "" Or mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) <> 0) Then 
				lngSUSUAMT = (mobjSCGLSpr.GetTextBinding( .sprSht,"COMMISSION",Row) * mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) ) * 0.01
				lngSUSUAMT = gRound(lngSUSUAMT,0)
				lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht,"COMMISSION",Row) - lngSUSUAMT
				mobjSCGLSpr.SetTextBinding .sprSht,"EXSUSU",Row,lngSUSUAMT
				mobjSCGLSpr.SetTextBinding .sprSht,"MCSUSU",Row,lngMCSUSU
				if mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSURATE",Row) = 0.0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,0
				Else
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row,1
				End If
			end if
		End if
		
		If Col = 12 Then
			if mobjSCGLSpr.GetTextBinding( .sprSht,"COMMISSION",Row) < mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) Then
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
			lngRATE = (mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row) / mobjSCGLSpr.GetTextBinding( .sprSht,"COMMISSION",Row)) * 100
			mobjSCGLSpr.SetTextBinding .sprSht,"EXSUSURATE",Row,lngRATE
			lngMCSUSU = mobjSCGLSpr.GetTextBinding( .sprSht,"COMMISSION",Row) - mobjSCGLSpr.GetTextBinding( .sprSht,"EXSUSU",Row)
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

'��ƮŬ���̺�Ʈ
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intcnt
			next
		end if
	end with
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
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") _
			or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"MCSUSU") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strCOLUMN = "COMMISSION"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") Then
				strCOLUMN = "EXSUSU"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"MCSUSU") Then
				strCOLUMN = "MCSUSU"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) _
					OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"MCSUSU"))   Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") _
				or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXSUSU") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"MCSUSU") Then
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
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjMDCMINTERNETEXCUTION = gCreateRemoteObject("cMDIT.ccMDITINTERNETREG")
    set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 19, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 8, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | MED_FLAG | CLIENTNAME | MEDNAME | MATTERNAME | EXCLIENTCODE | BTN | EXCLIENTNAME | EXSUSURATE | EXSUSU | MCSUSU | AMT | COMMI_RATE | COMMISSION | TBRDSTDATE | TBRDEDDATE | MEMO"
		mobjSCGLSpr.SetHeader .sprSht,		"����|���|��Ź��ȣ|��ü|������|��ü��|�����|������ �ڵ�|�������|�����  ��������|�����  ������|M&C������|�ݾ�|��������|������|��������|��������|���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|0   |4       |6   |10    |10    |10    |6          |2|10        |8               |8             |8        |8       |8       |8     |8       |8       |20"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "EXSUSURATE", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "EXSUSU | AMT | COMMI_RATE | COMMISSION | MCSUSU", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEMO", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "TBRDSTDATE | TBRDEDDATE | AMT | COMMI_RATE | COMMISSION"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "YEARMON | SEQ | MED_FLAG | CLIENTNAME | MEDNAME | MATTERNAME", -1, -1, 40
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON ", true
		.sprSht.style.visibility = "visible"
		
	
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
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

Sub EndPage()
	set mobjMDCMINTERNETEXCUTION = Nothing
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
		strPROGRAMNAME = .txtPROGRAMNAME.value

		'�ʱ�ȭ
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		vntData = mobjMDCMINTERNETEXCUTION.EXLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCUSTCODE,strPROGRAMNAME)
		
		
		IF not gDoErrorRtn ("EXLIST") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
			'mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			AMT_SUM
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
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTCODE",intCnt) = "" _
			 AND (mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",intCnt) = "" _
			 AND (mobjSCGLSpr.GetTextBinding(.sprSht,"EXSUSURATE",intCnt) <> 0.0 _
			 or mobjSCGLSpr.GetTextBinding(.sprSht,"EXSUSU",intCnt) <> 0) _
			 AND mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1)  Then 
					gErrorMsgBox intCnt & " ��° ���� �Է³��� �� Ȯ���Ͻʽÿ�" & vbcrlf & "����簡 ���� �� �����������/������� 0 �Դϴ�.","�Է¿���"
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | EXCLIENTCODE | EXSUSURATE | EXSUSU | MCSUSU | MEMO")
		
		intRtn = mobjMDCMINTERNETEXCUTION.EXCUTION_ProcessRtn(gstrConfigXml,vntData)
	
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ")
		
		intRtn = mobjMDCMINTERNETEXCUTION.EXCUTION_DeleteRtn(gstrConfigXml,vntData)
	
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
				sprSht_change 11,intCnt
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
			<TABLE id="tblForm" width="100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD id="Td1" align="left" width="400" height="20" runat="server">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="113" background="../../../images/back_p.gIF"
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
											<td class="TITLE">����� ������ ����&nbsp;</td>
										</tr>
									</table>
								</TD>
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
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="81">�� ��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 91px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" name="txtYEARMON" onchange="vbscript:Call gYearmonCheck(txtYEARMON)"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE,txtCLIENTNAME)"
												width="81">������
											</TD>
											<TD class="SEARCHDATA" width="266"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 176px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="20" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
												<INPUT class="INPUT_L" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPROGRAMNAME, '')"
												width="81">�����
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtPROGRAMNAME" title="�����" style="WIDTH: 288px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="42" name="txtPROGRAMNAME">
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" width="54" align="right" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="500" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">&nbsp;&nbsp;�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
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
											<TD class="BODYSPLIT" style="WIDTH: 100%"></TD>
										</TR>
									</TABLE>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtEXCLIENTCODE, txtEXCLIENTNAME)">���۴����</TD>
											<TD class="SEARCHDATA" style="WIDTH: 517px"><INPUT class="INPUT_L" id="txtEXCLIENTNAME" title="�ڵ��" style="WIDTH: 184px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="25" name="txtEXCLIENTNAME"> <IMG id="ImgEXCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgEXCLIENTCODE">
												<INPUT class="INPUT_L" id="txtEXCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtEXCLIENTCODE"> <IMG id="ImgAppEx" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="���۴���� �� �����Ḧ �ϰ����� �մϴ�." src="../../../images/ImgApp.gif" width="54" align="absMiddle"
													border="0" name="ImgAppEx"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 86px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">��������</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_R" id="txtSUSURATE" style="WIDTH: 104px; HEIGHT: 22px" accessKey="NUM"
													type="text" size="12" name="txtSUSURATE"> <IMG id="ImgApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'" height="20" alt="���۴���� �� �����Ḧ �ϰ����� �մϴ�." src="../../../images/ImgApp.gif"
													width="54" align="absMiddle" border="0" name="ImgApp"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblSheet" height="75%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="15425">
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
											<PARAM NAME="MaxCols" VALUE="44">
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
											<PARAM NAME="CellMEMOIndicator" VALUE="0">
											<PARAM NAME="ShowScrollTips" VALUE="0">
											<PARAM NAME="DataMember" VALUE="">
											<PARAM NAME="OLEDropMode" VALUE="0">
										</OBJECT>
									</DIV>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
