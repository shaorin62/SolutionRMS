<%@ Page CodeBehind="MDCMSENDCATVTRUTAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMSENDCATVTRUTAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���̺��� ����Ź ���ݰ�꼭 ����</title> 
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCOSENTTRUTAX , mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
CONST meTAB = 9
mstrGUBUN = ""
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
Sub imgQuery_onclick
	if frmThis.txtTAXYEARMON1.value = "" then
	    gErrorMsgBox "��� �Է��Ͻÿ�",""
		exit Sub
	end if
	If LEN(frmThis.txtTAXYEARMON1.value) <> 6 Then
		 gErrorMsgBox "����� 6�ڸ� �Դϴ�",""
		exit Sub
	End If
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub ImgSend_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "���ݰ�꼭 ������ �����Ͱ� �����ϴ�.","���۾ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdF.checked <> true then
		gErrorMsgBox "���ݰ�꼭������ �̿Ϸ�����϶� �����մϴ�..","���۾ȳ�!"
		Exit Sub
	end if
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgSendCancel_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "���ݰ�꼭 ��������� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "���ݰ�꼭 ������Ҵ� �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
		
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtTAXYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "CATV") 
		vntRet = gShowModalWindow("../MDCO/MDCMTAXCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(1,0) and .txtCLIENTNAME1.value = vntRet(2,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			if .txtTAXYEARMON1.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			End if
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetTAXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTAXYEARMON1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "CATV")
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					if .txtTAXYEARMON1.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					End if
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' ��ȸ�ʵ� ü���� �̺�Ʈ
'****************************************************************************************
'�Ϸ�üũ
Sub rdT_onclick
	SelectRtn
End Sub

'�̿Ϸ�üũ
Sub rdF_onclick
	SelectRtn
End Sub

Sub txtTAXYEARMON1_onkeydown
	'or window.event.keyCode = meTAB ���϶��� �ƴ� �����϶��� ��ȸ
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ	
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				If  right(mobjSCGLSpr.GetTextBinding( .sprSht,"STAT",intCnt),3) = "������" Then
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If		
				sprSht_Change 1, intcnt
			next
		end if
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim intCnt
   	
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EDITTYPECD") Then
			If mobjSCGLSpr.GetTextBinding( .sprSht,"EDITTYPECD",Row) = "" Then
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,Row,42,43,true
			Else
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,Row,42,43,true
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BUYNM") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"BUYNM",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"BUYEMAIL",Row, ""
			If strCode = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetSC_CUST_EMP(gstrConfigXml,mlngRowCnt,mlngColCnt, _ 
													 mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"BUYLDSCR",Row), _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",Row), _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"BUYNM",Row))		

				If not gDoErrorRtn ("GetSC_CUST_EMP") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"BUYNM",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"BUYEMAIL",Row, trim(vntData(3,1))
						
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"BUYNM"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF	
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPNM") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPNM",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"SUPPEMAIL",Row, ""
			If strCode = "" AND strCodeName <> "" Then	
				vntData = mobjMDCOGET.GetSC_CUST_EMP(gstrConfigXml,mlngRowCnt,mlngColCnt, _ 
													 mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_CODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPLDSCR",Row), _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row), _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPNM",Row))		

				If not gDoErrorRtn ("GetSC_CUST_EMP") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"SUPPNM",Row, trim(vntData(2,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"SUPPEMAIL",Row, trim(vntData(3,1))
						
						.txtCLIENTNAME1.focus
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPNM"), Row
						.txtCLIENTNAME1.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		END IF	
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BUYNM") Then
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"BUYLDSCR",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"BUYNM",.sprSht.ActiveRow))
								
			vntRet = gShowModalWindow("../MDCO/MDCMSENDEMAIL_CLIENT_POP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"BUYNM",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"BUYEMAIL",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPNM") Then
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPLDSCR",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPNM",.sprSht.ActiveRow))
			
								
			vntRet = gShowModalWindow("../MDCO/MDCMSENDEMAIL_CLIENT_POP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"SUPPNM",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUPPEMAIL",Row, vntRet(3,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If	
		
		'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.txtCLIENTNAME1.focus
		.sprSht_DTL.Focus
	End With
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	Dim strMEDFLAG
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If .rdT.checked = True Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row)) '<< �޾ƿ��°��
				gShowModalWindow "../MDCT/MDCMCATVTRUTAXDTL.aspx",vntInParams , 898,680
				'SelectRtn
			End IF
		end if	
	end with
end sub

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
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPAMT") OR _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VATAMT")  Then
			
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTAMT") Then
				strCOLUMN = "TOTAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPAMT") Then
				strCOLUMN = "SUPPAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VATAMT") Then
				strCOLUMN = "VATAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTAMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPAMT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VATAMT"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"TOTAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUPPAMT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VATAMT") Then
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

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim intRtn
	
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			vntInParams = array(.txtTAXYEARMON1.value, mobjSCGLSpr.GetTextBinding( .sprSht,"BUYLDSCR",.sprSht.ActiveRow), "A2", "CATV")
								
			vntRet = gShowModalWindow("../MDCO/MDCMFIRSTBILLNO_POP.aspx",vntInParams , 780,630)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"FIRSTBILLNO",Row, vntRet(2,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				.txtCLIENTCODE1.focus
				.sprSht.Focus
				mobjSCGLSpr.ActiveCell .sprSht, Col, Row
			End If
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNBUY") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTNBUY") then exit Sub
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"BUYLDSCR",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"TIMNAME",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"BUYNM",.sprSht.ActiveRow))
								
			vntRet = gShowModalWindow("../MDCO/MDCMSENDEMAIL_CLIENT_POP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				
				if .sprSht.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"BUYNM",Row, vntRet(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"BUYEMAIL",Row, vntRet(3,0)
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
				end if
				
				.txtCLIENTCODE1.focus
				.sprSht.Focus
				mobjSCGLSpr.ActiveCell .sprSht, Col, Row
			End If

		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUPP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTNSUPP") then exit Sub
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"REAL_MED_CODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPLDSCR",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",.sprSht.ActiveRow), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",.sprSht.ActiveRow), _
								mobjSCGLSpr.GetTextBinding( .sprSht,"SUPPNM",.sprSht.ActiveRow))
			
								
			vntRet = gShowModalWindow("../MDCO/MDCMSENDEMAIL_CLIENT_POP.aspx",vntInParams , 413,440)
			
			If isArray(vntRet) Then
				
				if .sprSht.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"SUPPNM",Row, vntRet(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"SUPPEMAIL",Row, vntRet(3,0)
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
				end if
				
				.txtCLIENTCODE1.focus
				.sprSht.Focus
				mobjSCGLSpr.ActiveCell .sprSht, Col, Row
			End If
		End If	
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
	set mobjMDCOSENTTRUTAX	= gCreateRemoteObject("cMDCO.ccMDCOSENDTRUTAX")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	With frmThis
		'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
		gInitComParams mobjSCGLCtl,"MC"
		
		mobjSCGLCtl.DoEventQueue
		
		gSetSheetDefaultColor() 
		
		'Sheet �⺻Color ����
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 60, 0, 1, 2
		mobjSCGLSpr.AddCellSpan  .sprSht, 42, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 44, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 47, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | STAT | TAXYEARMON | TAXNO | MEDFLAG | COMPANYCD | BILLNO | FISCALLYY | BILLFLAG | SUPPBSN | SUPPLDSCR | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYBSN | BUYLDSCR | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | REGDATE | TOTAMT | SUPPAMT | VATAMT | BILLRMRK | TITLE | REQFLAG | NORMFLAG | RECEIPTID | RECEIPTNM | PURTEAMCD | INSDATE | BILLSEQ | SUPPDATE | ITEMNM | SIZE | QTY | UNITPRC | ITEMRMRK | EDITTYPECD | FIRSTBILLNO | BTN | BUYNM | BTNBUY | BUYEMAIL | SUPPNM | BTNSUPP | SUPPEMAIL | SENDNTS_YN | ISTRUST_YN | TRUST_CUSCD | RMSNO | ERRCODE | CLIENTCODE | REAL_MED_CODE | TIMCODE | TIMNAME | MEDCODE | MEDNAME"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|����|���ݰ�꼭���|���ݰ�꼭��ȣ|��ü����|Company code|��ȣ|ȸ�迬��|����|�����ڵ�Ϲ�ȣ|�����ڻ�ȣ|�����ڴ�ǥ��|�������ּ�|�����ھ���|����������|���޹޴��ڵ�Ϲ�ȣ|���޹޴��ڻ�ȣ|���޹޴��ڴ�ǥ�ڸ�|���޹޴����ּ�|���޹޴��ھ���|���޹޴�������|������|�հ�ݾ�|���ް���|����|���|����|û������|����������|�����ID|����ڸ�|�μ��ڵ�|Create Date|ǰ��|�Ⱓ|ǰ�񳻿�|����|����|�ܰ�|ǰ����|��������|�����ݰ�꼭��ȣ|���޹޴��ڴ����|���޹޴���email|�����ڴ����|������email|����û���ۿ���|����Ź����|��Ź��ü����ڹ�ȣ|���Ϲ�ȣ|�����ڵ��ȣ|�������ڵ�|��ü���ڵ�|���ڵ�|����|��ü�ڵ�|��ü��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   7|             0|             0|       0|           0|  10|       0|  10|            10|	      17|           0|	       0|         0|         0|                10|            17|                 0|	         0|             0|             0|     8|      10|      10|  10|  10|  17|       6|         0|       8|      13|      10|          8|   0|   0|      18|   0|   0|   0|       0|      15|            12|2|            10|2|             15|        10|2|         15|             7|         0|                10|      21|          10|         0|         0|     0|   0|       0|    0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN | BTNBUY | BTNSUPP"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "SENDNTS_YN", -1, -1, "Y" & vbTab & "N" , 10, 40, False, False
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOTAMT | SUPPAMT | VATAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "STAT | COMPANYCD | BILLNO | FISCALLYY | SUPPBSN | SUPPLDSCR | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYBSN | BUYLDSCR | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | REGDATE | BILLRMRK | TITLE | REQFLAG | NORMFLAG | RECEIPTID | RECEIPTNM | PURTEAMCD | INSDATE | BILLSEQ | SUPPDATE | ITEMNM | SIZE | QTY | UNITPRC | ITEMRMRK | EDITTYPECD | FIRSTBILLNO | BUYNM | BUYEMAIL | SUPPNM | SUPPEMAIL | ISTRUST_YN | TRUST_CUSCD | RMSNO | ERRCODE", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "STAT | COMPANYCD | BILLNO | FISCALLYY | SUPPBSN | SUPPLDSCR | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYBSN | BUYLDSCR | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | REGDATE | TOTAMT | SUPPAMT | VATAMT | TITLE | NORMFLAG | RECEIPTID | RECEIPTNM | PURTEAMCD | INSDATE | BILLSEQ | SUPPDATE | SIZE | QTY | UNITPRC | ITEMRMRK | FIRSTBILLNO | BTN | ISTRUST_YN | TRUST_CUSCD | RMSNO | ERRCODE"
		mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON | TAXNO | MEDFLAG | COMPANYCD  | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | NORMFLAG | SUPPDATE | BILLSEQ | SIZE | QTY | UNITPRC | ISTRUST_YN  | CLIENTCODE | REAL_MED_CODE | TIMCODE | TIMNAME | MEDCODE | MEDNAME", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | STAT | BILLNO | SUPPBSN | BUYBSN | REGDATE | INSDATE | TRUST_CUSCD | RECEIPTID | PURTEAMCD | INSDATE | SUPPDATE | RMSNO",-1,-1,2,2,False
		
		.sprSht.style.visibility = "visible"
	
	End With
	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCOSENTTRUTAX = Nothing
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
		.txtTAXYEARMON1.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		Get_COMBO_VALUE				
		.txtCLIENTNAME1.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'-----------------------------------------------------------------------------------------
' �׸��� �޺��ڽ� ����
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData_REQFLAG, vntData_BILLFLAG, vntData_EDITTYPECD
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_REQFLAG = mobjMDCOSENTTRUTAX.Get_COMBOREQFLAG_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		vntData_BILLFLAG = mobjMDCOSENTTRUTAX.Get_COMBOBILLFLAG_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		vntData_EDITTYPECD = mobjMDCOSENTTRUTAX.Get_COMBOEDITTYPECD_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBOREQFLAG_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "REQFLAG",,,vntData_REQFLAG,,40 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "BILLFLAG",,,vntData_BILLFLAG,,100 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "EDITTYPECD",,,vntData_EDITTYPECD,,150 
			mobjSCGLSpr.TypeComboBox = True 
   		End If    
   	End With
End Sub


'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strTAXYEARMON, strCLIENTCODE
   	Dim i, strCols
   	Dim strMED_FLAG
   	Dim intCnt
   	Dim strRows
   	Dim intCnt2
   
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		If .txtTAXYEARMON1.value = "" Then
			gErrorMsgBox "����� �Է��Ͻʽÿ�","��ȸ�ȳ�!"
			Exit Sub
		End If	
		
		If Len(.txtTAXYEARMON1.value) <> 6 Then
			gErrorMsgBox "����� ������ �ƴմϴ�.","��ȸ�ȳ�!"
			Exit Sub
		End If
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		strTAXYEARMON	= .txtTAXYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strMED_FLAG		= "A2"
		
		IF .chkMAE.checked THEN
			if .rdT.checked then
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
				mobjSCGLSpr.ColHidden .sprSht, "BUYNM | BTNBUY | BUYEMAIL", true
				
				vntData = mobjMDCOSENTTRUTAX.Get_SENDED_TAX_NO(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strCLIENTCODE, strMED_FLAG, "CATV")
				If not gDoErrorRtn ("Get_SENDED_TAX_NO") then
					
					'��ȸ�� �����͸� ���ε�
					call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
					'�ʱ� ���·� ����
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					if .sprSht.MaxRows > 0 then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,-1,42,43,true
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,-1,41,41,true
						
						For intCnt = 1 To .sprSht.MaxRows
							mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
							mobjSCGLSpr.SetTextBinding .sprSht,"STAT",intCnt,"����ó����"
						Next
					end if
					AMT_SUM
					gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				END IF
			end if
		ELSE
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.ColHidden .sprSht, "BUYNM | BTNBUY | BUYEMAIL", false
			'���ݰ�꼭 �Ϸ���ȸ
			If .rdT.checked = True Then
				vntData = mobjMDCOSENTTRUTAX.Get_SEND_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strCLIENTCODE, strMED_FLAG, "CATV")
				If not gDoErrorRtn ("Get_SEND_TAX") then
					'��ȸ�� �����͸� ���ε�
					call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
					'�ʱ� ���·� ����
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					if .sprSht.MaxRows > 0 then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,-1,41,43,true
						
						intCnt2 = 1
						For intCnt = 1 To .sprSht.MaxRows
							If mobjSCGLSpr.GetTextBinding(.sprSht,"STAT",intCnt) = "����������" Then
								If intCnt2 = 1 Then
									strRows = intCnt
								Else
									strRows = strRows & "|" & intCnt
								End If
								intCnt2 = intCnt2 + 1
							End If
						Next
						mobjSCGLSpr.SetCellTypeStatic2 .sprSht, strRows, 1, 1, 0, 2,  TRUE
						
						for intcnt = 1 to .sprSht.MaxRows
							If  mobjSCGLSpr.GetTextBinding( .sprSht,"STAT",intCnt) = "����������" Then
								mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
							End If		
						next
					end if
					AMT_SUM
					gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
					mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				End If
			'�̿Ϸ� �ŷ����� ������ ��ȸ
			ElseIf .rdF.checked = True Then			
				vntData = mobjMDCOSENTTRUTAX.Get_SENDED_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strCLIENTCODE, strMED_FLAG, "CATV")
				If not gDoErrorRtn ("Get_SENDED_TAX") then
					'��ȸ�� �����͸� ���ε�
					call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
					'�ʱ� ���·� ����
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					if .sprSht.MaxRows > 0 then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,-1,42,43,true
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,-1,41,41,true
						
						intCnt2 = 1
						For intCnt = 1 To .sprSht.MaxRows
							If mobjSCGLSpr.GetTextBinding(.sprSht,"STAT",intCnt) = "������" Then
								If intCnt2 = 1 Then
									strRows = intCnt
								Else
									strRows = strRows & "|" & intCnt
								End If
								intCnt2 = intCnt2 + 1
							End If
						Next
						mobjSCGLSpr.SetCellTypeStatic2 .sprSht, strRows, 1, 1, 0, 2,  TRUE
						
						for intcnt = 1 to .sprSht.MaxRows
							If  mobjSCGLSpr.GetTextBinding( .sprSht,"STAT",intCnt) = "������" Then
								mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
							End If		
						next
					end if
					AMT_SUM
					gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
					mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				End If
			End If		
		END IF
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUPPAMT", lngCnt)
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

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn, intRtn2
   	Dim vntData, vntSelect
	Dim intCnt, intCnt2
	Dim chkcnt
	Dim strYEARMON
	Dim strSAVEYEARMON
	Dim strSAVESEQ
	Dim strSAVERMSNO
	Dim strTITLE, strBUYEMAIL, strBUYNM, strSUPPEMAIL, strSUPPNM
	
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .rdT.checked = True Then
			gErrorMsgBox "�̿Ϸ� ���¿��� ������ �����մϴ�.","����ȳ�!"
			Exit Sub
		End If
		
		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
   		
   		'üũ ���� ��� ���� �ȵǵ���
		chkcnt = 0
		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				
				strTITLE = "" :  strBUYEMAIL = "" : strBUYNM = "" : strSUPPEMAIL = "" : strSUPPNM = ""
				
				strTITLE = mobjSCGLSpr.GetTextBinding(.sprSht,"TITLE",intCnt)
				strBUYEMAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"BUYEMAIL",intCnt)
				strBUYNM = mobjSCGLSpr.GetTextBinding(.sprSht,"BUYNM",intCnt)
				strSUPPEMAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUPPEMAIL",intCnt)
				strSUPPNM = mobjSCGLSpr.GetTextBinding(.sprSht,"SUPPNM",intCnt)
				
				If strTITLE  = "" Then
					gErrorMsgBox "������ �ʼ� �Դϴ�.","����ȳ�!"
					Exit Sub
				End If
				If  strBUYEMAIL = "" Then
					gErrorMsgBox "���޹޴��� �̸����� �ʼ� �Դϴ�.","����ȳ�!"
					Exit Sub
				End If
				If  strBUYNM = "" Then
					gErrorMsgBox "���޹޴��ڴ� �ʼ� �Դϴ�.","����ȳ�!"
					Exit Sub
				End If
				If  strSUPPEMAIL = "" Then
					gErrorMsgBox "�������̸����� �ʼ� �Դϴ�.","����ȳ�!"
					Exit Sub
				End If
				If  strSUPPNM = "" Then
					gErrorMsgBox "�����ڴ� �ʼ� �Դϴ�.","����ȳ�!"
					Exit Sub
				End If
				chkcnt = chkcnt + 1
			END IF
		Next
		
		if chkcnt = 0 then
			gErrorMsgBox "���ݰ�꼭�� ������ �����͸� üũ �Ͻʽÿ�","���۾ȳ�!"
			exit sub
		end if
		
		intRtn2 = gYesNoMsgbox("���ݰ�꼭�� ���� �Ͻðڽ��ϱ�?","Ȯ��")
		IF intRtn2 <> vbYes then exit Sub
		
		strYEARMON = .txtTAXYEARMON1.value
		vntSelect = mobjMDCOSENTTRUTAX.SelectRtn_RMSNO(gstrConfigXml, strYEARMON)
		
		if  IsArray(vntSelect) then 
			strSAVEYEARMON = vntSelect(0,1)
			strSAVESEQ =vntSelect(1,1) 
			strSAVERMSNO =vntSelect(2,1)
		End If
		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
   		
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | STAT | TAXYEARMON | TAXNO | MEDFLAG | COMPANYCD | BILLNO | FISCALLYY | BILLFLAG | SUPPBSN | SUPPLDSCR | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYBSN | BUYLDSCR | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | REGDATE | TOTAMT | SUPPAMT | VATAMT | BILLRMRK | TITLE | REQFLAG | NORMFLAG | RECEIPTID | RECEIPTNM | PURTEAMCD | INSDATE | BILLSEQ | SUPPDATE | ITEMNM | SIZE | QTY | UNITPRC | ITEMRMRK | EDITTYPECD | FIRSTBILLNO | BTN | BUYNM | BTNBUY | BUYEMAIL | SUPPNM | BTNSUPP | SUPPEMAIL | SENDNTS_YN | ISTRUST_YN | TRUST_CUSCD | RMSNO | ERRCODE | CLIENTCODE | REAL_MED_CODE | TIMCODE | TIMNAME | MEDCODE | MEDNAME")
		
		'ó�� ������ü ȣ��
		intRtn = mobjMDCOSENTTRUTAX.ProcessRtn_SENDTAX(gstrConfigXml,vntData, "SEND", strSAVEYEARMON, strSAVESEQ, strSAVERMSNO)
		

		If not gDoErrorRtn ("ProcessRtn_SENDTAX") Then
			Call Excel_save (strSAVERMSNO)
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "���ݰ�꼭�� ���۵Ǿ����ϴ�.","���۾ȳ�!"
			'.rdT.checked = True
			selectRtn
   		End If
   	end with
End Sub

Sub DeleteRtn ()
	Dim intRtn, intRtn2
   	Dim vntData, vntSelect
	Dim intCnt, intCnt2
	Dim chkcnt
	Dim strYEARMON
	Dim strSAVEYEARMON
	Dim strSAVESEQ
	Dim strSAVERMSNO
	
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .rdF.checked = True Then
			gErrorMsgBox "�Ϸ� ���¿��� ������ �����մϴ�.","����ȳ�!"
			Exit Sub
		End If
		
		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
   		
   		'üũ ���� ��� ���� �ȵǵ���
		chkcnt = 0
		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				chkcnt = chkcnt + 1
			END IF
		Next
		
		if chkcnt = 0 then
			gErrorMsgBox "���ݰ�꼭�� ��������� �����͸� üũ �Ͻʽÿ�","���۾ȳ�!"
			exit sub
		end if
		
		intRtn2 = gYesNoMsgbox("���ݰ�꼭�� ������� �Ͻðڽ��ϱ�?","Ȯ��")
		IF intRtn2 <> vbYes then exit Sub
		
		strYEARMON = .txtTAXYEARMON1.value
		vntSelect = mobjMDCOSENTTRUTAX.SelectRtn_RMSNO(gstrConfigXml, strYEARMON)
		
		if  IsArray(vntSelect) then 
			strSAVEYEARMON = vntSelect(0,1)
			strSAVESEQ =vntSelect(1,1) 
			strSAVERMSNO =vntSelect(2,1)
		End If
		
		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
   		
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | STAT | TAXYEARMON | TAXNO | MEDFLAG | COMPANYCD | BILLNO | FISCALLYY | BILLFLAG | SUPPBSN | SUPPLDSCR | SUPPCEO | SUPPADDR | SUPPBUSICOND | SUPPBUSIITEM | BUYBSN | BUYLDSCR | BUYCEO | BUYADDR | BUYBUSICOND | BUYBUSIITEM | REGDATE | TOTAMT | SUPPAMT | VATAMT | BILLRMRK | TITLE | REQFLAG | NORMFLAG | RECEIPTID | RECEIPTNM | PURTEAMCD | INSDATE | BILLSEQ | SUPPDATE | ITEMNM | SIZE | QTY | UNITPRC | ITEMRMRK | EDITTYPECD | FIRSTBILLNO | BTN | BUYNM | BTNBUY | BUYEMAIL | SUPPNM | BTNSUPP | SUPPEMAIL | SENDNTS_YN | ISTRUST_YN | TRUST_CUSCD | RMSNO | ERRCODE | CLIENTCODE | REAL_MED_CODE | TIMCODE | TIMNAME | MEDCODE | MEDNAME")
		
		'ó�� ������ü ȣ��
		intRtn = mobjMDCOSENTTRUTAX.ProcessRtn_SENDTAX(gstrConfigXml,vntData, "SENDCANCEL", strSAVEYEARMON, strSAVESEQ, strSAVERMSNO)
		

		If not gDoErrorRtn ("ProcessRtn_SENDTAX") Then
			Call Excel_save (strSAVERMSNO)
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "���ݰ�꼭�� ���۵Ǿ����ϴ�.","���۾ȳ�!"
			'.rdT.checked = True
			selectRtn
   		End If
   	end with
End Sub

-->
		</script>
		<script language="javascript">
		
		//##########################################################################################################################################
		//������ SAM ������ �����ϱ� ���Ͽ� ���ϸ��� ������ file���� asp �������� �޷�����.
		//##########################################################################################################################################
		function Excel_save(strSAVERMSNO){
			ifrm.location.href = "../MDCO/MDCMSENETAXSUB.asp?temp_filename="+ strSAVERMSNO; 
		}
		
		//##########################################################################################################################################
		// ������ FTP ���� �������ο� ���� RFC ȣ���� �ϴ� �Լ� �̴�. FTP ������ �Ϸ�Ǹ� �Ϸ�޼����� �Բ� ���۵� ���Ϲ�ȣ�� 
		// frmSapVoch ���� ������ �� �̿��Ͽ� Submit �ϹǷν�[******************************************��1) ����] ���Ϲ�ȣ��,
		// server Control �� ��������, SubControl ���� ����� RFC ����� ������ ���� **********************************��2) vbscript �Լ��� �����Ѵ�. 
		//##########################################################################################################################################
		function RFC_Call(strMsg){
		var strConfirm;
		var strRmsNo;
		var array_data = strMsg.split("|");
			strConfirm = array_data[0];
			strRmsNo = array_data[1];
			if (strConfirm =="Put Successful!"){
				//���Ϲ�ȣ���� �� "200908" '���� RFC input ������ ���� 6�ڸ� �̱⶧���� 2009080001_T ���� ������ ������ ���Ƿ� ����!! ���� ���߿Ϸ�� ��ü
				//Set_IframeValue (strRmsNo);
			} else{
				alert("�������ۿ� ���� �Ͽ����ϴ�!");
			}
		}
		
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="144" background="../../../images/back_p.gIF"
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
											<td class="TITLE">����Ź ���ݰ�꼭 ���� ����</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 326px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End--></TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTAXYEARMON1, '')"
												width="50">���</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtTAXYEARMON1" title="�ŷ������" style="WIDTH: 89px; HEIGHT: 22px"
													accessKey="MON" type="text" maxLength="6" size="6" name="txtTAXYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="90">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 143px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="14" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
												<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHDATA" width="50">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL">����
											</TD>
											<TD class="SEARCHDATA" colspan="1">&nbsp;<INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN">
												�Ϸ�&nbsp;&nbsp; &nbsp; <INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" value="rdF" name="rdGBN" checked>&nbsp;�̿Ϸ�&nbsp;
											</TD>
											<TD class="SEARCHLABEL" colSpan="1">����ó����</TD>
											<TD class="SEARCHDATA" colspan="2">&nbsp;&nbsp; <INPUT id="chkMAE" title="����ó����" type="checkbox" name="chkMAE">&nbsp; 
												(����ó ������ �Ϸ���·θ� ��ȸ�����մϴ�.)
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="absmiddle" align="center"><TABLE class="DATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 20px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD colSpan="4" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
										</TR>
										<TR>
											<TD colSpan="4" height="4"></TD>
										</TR>
										<TR>
											<TD class="DATA_RIGHT" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="ImgSend" onmouseover="JavaScript:this.src='../../../images/ImgSendOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/ImgSend.gif'" height="20" alt="���ݰ�꼭����."
																src="../../../images/ImgSend.gif" align="absMiddle" border="0" name="ImgSend"></td>
														<TD><IMG id="imgSendCancel" onmouseover="JavaScript:this.src='../../../images/imgSendCancelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSendCancel.gif'"
																height="20" alt="���ݰ�꼭���� ���" src="../../../images/imgSendCancel.gIF" border="0" name="imgSendCancel"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31882">
										<PARAM NAME="_ExtentY" VALUE="14314">
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
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></form>
		<iframe id="ifrm" frameBorder="0" width="0" height="0"></iframe>
	</body>
</HTML>
