<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGEBASICLIST.aspx.vb" Inherits="PD.PDCMCHARGEBASICLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMCHARGEBASICLIST.aspx
'��      �� : JOBMST�� ����° �� - ���۸���Ʈ ���� ���ε�
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/24 By Ȳ����
'****************************************************************************************
-->
		<meta content="False" name="vs_snapToGrid">
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<script language="vbscript" id="clientEventHandlersVBS">
option explicit


Dim mlngRowCnt, mlngColCnt		
Dim mobjccPDDCCHARGEEXCOM
Dim mobjPDCOGET
Dim mcomecalender
Dim mstrCheck
Dim mstrGrid
Dim strPARENTJOBNO

CONST meTAB = 9
mcomecalender = FALSE
mstrCheck=True
mstrGrid = FALSE


'=============================
' �̺�Ʈ ���ν��� 
'=============================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN(byVal strmode)
	With frmThis
		If  strmode = "EXTENTION"  Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "60%"
			document.getElementById("tblSheet2").style.height = "30%"
		ELSEIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ELSEIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "60%"
		END IF
	End With
End Sub

Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'��ȸ��ư
Sub imgQuery_onclick
	mstrGrid = TRUE
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub		

'�űԹ�ư
Sub imgNEW_onclick ()
	mstrGrid = False
	Call sprSht_HDR_Keydown(meINS_ROW, 0)	
	Call sprSht_DTL_Keydown(meINS_ROW, 0)	
	
end Sub

'�����ư
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'������ư
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'������ư
Sub imgExcel_HDR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_DTL_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
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
	SelectRtn
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
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

'-----------------------------------------------------------------------------------------
' ����ó ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub imgOUTSCODE_onclick
	Call SEARCHOUT_POP()
End Sub

'���� ������List ��������
Sub SEARCHOUT_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	SelectRtn
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtOUTSNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCOGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHOUT_POP()
				End If
   			end if
   		end with
   		
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub



'-----------------------------------------------------------------------------------------
' �ʵ� ü����
'-----------------------------------------------------------------------------------------
Sub txtFROM_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	Dim strOLDYEARMON
	strdate = ""
	strFROM =""
	strFROM2 = ""

	With frmThis
		strdate=.txtFROM.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strFROM2 = strdate
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		
		.txtFROM.value = strFROM2
		DateClean strFROM
		txtTo_onchange
	End With

	gSetChange
End Sub

Sub txtTo_onchange
	SelectRtn
	gSetChange
End Sub

Sub txtOUTSNAME_onchange
	SelectRtn
	gSetChange
End Sub


Sub txtJOBNAME_onchange
	SelectRtn
	gSetChange
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
		SelectRtn
		gSetChange
	end with
End Sub



'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
'Ŭ��
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_HDR, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_HDR.MaxRows
				sprSht_HDR_Change 1, intcnt
			next
		elseif Row > 0 AND Col > 1 then
			if mstrGrid then SelectRtn_DTL Col, Row
		end if
	end with
End Sub


Sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		Else
			'strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"JOBNO",.sprSht_HDR.ActiveRow)
			'parent.jobMst_Call
			'mobjSCGLSpr.ActiveCell .sprSht_HDR, strCol, strRow	
		End If
	End With
End Sub


Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	Dim strRow, strCol
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub


Sub sprSht_HDR_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim strINPUTJOBNO
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If frmThis.txtJOBNO.value <> ""Then
		strINPUTJOBNO = frmThis.txtJOBNO.value 
		Else
		strINPUTJOBNO = parent.document.forms("frmThis").txtJOBNO.value
		End If
		frmThis.sprSht_HDR.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_HDR, cint(KeyCode), cint(Shift), -1, 1)
		
		'���⼭ �θ�â���� �޾ƿ� JOBNO�� �ִ´�.
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"JOBNO",frmThis.sprSht_HDR.ActiveRow,strINPUTJOBNO 
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"CONFIRMFLAG",frmThis.sprSht_HDR.ActiveRow, "��Ȯ��"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht_HDR,"CREDAY",frmThis.sprSht_HDR.ActiveRow, gNowDate
		
	End If
End Sub


Sub sprSht_DTL_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
	
		frmThis.sprSht_DTL.MaxRows = 0
		frmThis.sprSht_DTL.MaxRows = 100
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_DTL, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht_DTL, 1,1
		frmThis.sprSht_DTL.focus()
	End If
End Sub


Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
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
		If .sprSht_HDR.MaxRows >0 Then
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") Then
				If .sprSht_HDR.ActiveRow > 0 Then
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)
					
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
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,strCol,vntData_row(j))
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


Sub sprSht_DTL_Mouseup(KeyCode, Shift, X,Y)
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
		If mstrGrid Then
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"QTY") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		ELSE
			If .sprSht_DTL.MaxRows >0 Then
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"QTY") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE")  Then
					If .sprSht_DTL.ActiveRow > 0 Then
						vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
						vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)
						
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
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,strCol,vntData_row(j))
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
		END IF
		
	End With
End Sub

Sub sprSht_HDR_Keyup(KeyCode, Shift)
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
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub


'��Ʈ ��ưŬ��
Sub sprSht_HDR_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strCUSTCODE , strCUSTNAME

	with frmThis

		'����ó��
		IF Col = 8 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"BTN") then exit Sub
			strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow)
			strCUSTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow)
			
			vntInParams = array(trim(strCUSTCODE), trim(strCUSTNAME)) '<< �޾ƿ��°��
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		
			if isArray(vntRet) then
				if strCUSTCODE = vntRet(0,0) and strCUSTNAME = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit

				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow, trim(vntRet(0,0))   
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow, trim(vntRet(1,0))     
				
				mobjSCGLSpr.CellChanged .sprSht_HDR, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_HDR, Col+3, Row			
			end if
			gSetChange
     	END IF	
	End with
End Sub



Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
Dim vntData
   	Dim i, strCols , vntInParams
   	Dim strCode, strCodeName
   	Dim strCUSTCODE , strCUSTNAME
   	Dim intCnt
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
	
					
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"CUSTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjPDCOGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",trim(strCodeName))
				
				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",Row, trim(vntData(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",Row, trim(vntData(1,0))
						mobjSCGLSpr.CellChanged .sprSht_HDR, .sprSht_HDR.ActiveCol-1,frmThis.sprSht_HDR.ActiveRow
						
						.sprSht_HDR.focus
					Else
					
						strCUSTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow)
						strCUSTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow)
						
						vntInParams = array(trim(strCUSTCODE), trim(strCUSTNAME)) '<< �޾ƿ��°��
						vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
					
						if isArray(vntRet) then
							if strCUSTCODE = vntRet(0,0) and strCUSTNAME = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit

							mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTCODE",.sprSht_HDR.ActiveRow, trim(vntRet(0,0))   
							mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CUSTNAME",.sprSht_HDR.ActiveRow, trim(vntRet(1,0))     
						end if
						
					End If
					.sprSht_HDR.focus 
					mobjSCGLSpr.ActiveCell .sprSht_HDR, Col+4, Row
   				End If
   			End If
		End If
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row
End Sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ
'-----------------------------------------------------------------------------------------	
Sub InitPage()
	'����������ü ����	
	dim vntInParam
	dim intNo,i
	
	set mobjccPDDCCHARGEEXCOM	= gCreateRemoteObject("cPDCO.ccPDDCCHARGEEXCOM")
	set mobjPDCOGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	'�� ��ġ ���� �� �ʱ�ȭ
	'pnlTab1.style.position = "absolute"
	'pnlTab1.style.top = "260px"
	'pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	'JOBNO �޾ƿ��� �κ�==========================================================
	'vntInParam = window.dialogArguments
	'	intNo = ubound(vntInParam)
	'	'�⺻�� ����
	'	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'	
	'	for i = 0 to intNo
	''		select case i
	'			case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
	'			case 1 : frmThis.txtJOBNO.value = vntInParam(i)
	'		end select
	'	next
	'==============================================================================
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis

		gSetSheetColor mobjSCGLSpr, .sprSht_HDR
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 12, 0, 0, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_HDR,7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK|REVSEQ|JOBNO|OUTSCODE|OUTSNAME|CUSTCODE|CUSTNAME|BTN|AMT|CREDAY|CONFIRMFLAG|BIGO"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "����|����|JOBNO|�������ڵ�|��������|����ó�ڵ�|����ó��|�ݾ�|�ۼ���|Ȯ��|���"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1","  4|   4|   7|         9|      15|         9|      15|2|   9|    9|  9|  20"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_HDR,"..", "BTN"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_HDR, "CONFIRMFLAG", -1, -1, "Ȯ��" & vbTab & "��Ȯ��" , 10, 70, False, False
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "OUTSNAME | CUSTNAME | BIGO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "REVSEQ|JOBNO|OUTSCODE|CUSTCODE|AMT|CREDAY|"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "REVSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "OUTSNAME | CUSTNAME | BIGO",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "REVSEQ | JOBNO | OUTSCODE | CUSTCODE | CREDAY | CONFIRMFLAG",-1,-1,2,2,false
	
	    .sprSht_HDR.style.visibility  = "visible"
		
		
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 9, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ|ITEMNAME|STD|QTY|PRICE|AMT|BIGO"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "����|�������ڵ�|����������|�׸�|�԰�|����|�ܰ�|�ݾ�|���"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1","  5|         9|         9|  30|   12|   12|   12|   12|  30"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "QTY|PRICE|AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "ITEMNAME|STD|BIGO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "SEQ|OUTSCODE|REVSEQ"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ|OUTSCODE|REVSEQ",-1,-1,2,2,false
	
	    .sprSht_DTL.style.visibility  = "visible"

	InitPageData	

	'���ڰ��� ��ü��ȸ ����ڿ�û�� ���
	'msgbox parent.document.forms("frmThis").txtJOBNO.value 
	window.setTimeout "call time_data()",1000 
	SelectRtn
	End With
End Sub

Sub time_data
 with frmThis
	.txtJOBNO.value =  parent.document.forms("frmThis").txtJOBNO.value 
	.txtJOBNAME.value =  parent.document.forms("frmThis").txtJOBNAME.value 
	
 End with
End Sub
Sub EndPage()
	'set mobjccPDDCCHARGEEXCOM = Nothing
	'set mobjPDCOGET = Nothing
	gEndPage
End Sub

Sub InitPageData
	'�ʱ� ������ ����
	with frmThis
		
		'�ʱⰪ ����
		.txtFROM.value = gNowDate
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)	
		
		
		'.sprSht_HDR.focus
	End with
End Sub

Sub DateClean(strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	with frmThis
	
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		.txtTO.value = date2
	End With
End Sub


'-----------------------------------------------------------------------------------------
' ��ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn
	Dim vntData
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
	
		'���ݰ�꼭 �Ϸ���ȸ
		vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtJOBNO.value),Trim(.txtJOBNAME.value),TRIM(.txtOUTSCODE.value),TRIM(.txtOUTSNAME.value))
		
		If not gDoErrorRtn ("SelectRtn_HDR") then
			If mlngRowCnt >0 Then
				mobjSCGLSpr.SetClipBinding .sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True
				mobjSCGLSpr.SetFlag  frmThis.sprSht_HDR,meCLS_FLAG
					
				gWriteText lblstatus1, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE	
		
			ELSE
				.sprSht_HDR.MaxRows = 0
				gWriteText lblstatus1, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE	
				
			End If
		End If	
		
		sprSht_HDR_Click 2, 1
		AMT_SUM
			
	END WITH
End Sub

Sub SelectRtn_DTL (Col , Row)
	Dim vntData
	Dim strOUTSCODE,strREVSEQ
   	Dim i, strCols
   	Dim intCnt
   	Dim strRow
	
	'On error resume next
	
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strOUTSCODE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"OUTSCODE",Row)
		strREVSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"REVSEQ",Row)
		
		IF strOUTSCODE <> "" THEN
			vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(strOUTSCODE),Trim(strREVSEQ))
		end if
	
		If not gDoErrorRtn ("SelectRtn_DTL") then
			IF mlngRowCnt > 0 THEN
				mobjSCGLSpr.SetClipBinding .sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True
				mobjSCGLSpr.SetFlag  frmThis.sprSht_DTL,meCLS_FLAG
			ELSE 
				.sprSht_DTL.MaxRows = 0
				
			END IF
		End If	
		gWriteText lblstatus2, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
	
		
	END WITH
End Sub


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht_HDR.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht_HDR.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub



'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn()
	Dim intRtn ,Cnti , Cntj
  	dim vntData_BASICLIST_HDR , vntData_BASICLIST_DTL
	Dim strOLDSEQ , strRow
	Dim strSEQFlag
	Dim IntCnt ,i
	Dim lngchkCnt
	Dim strJOBNO
	
	with frmThis
 
  		
  		'�ʱ�ȭ
  		Cnti=0
  		Cntj=0
  		lngchkCnt = 0
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		'����� ����Ÿ�� �޴´�
		vntData_BASICLIST_HDR = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"REVSEQ|JOBNO|OUTSCODE|OUTSNAME|CUSTCODE|BTN|AMT|CREDAY|CONFIRMFLAG|BIGO")
		vntData_BASICLIST_DTL = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"SEQ|OUTSCODE|REVSEQ|ITEMNAME|STD|QTY|PRICE|AMT|BIGO")
		
		' validation ����
		'------------------------------------------------------------------------------------------------------------------------------------------
			
		'��� VALIDATION
		if  not IsArray(vntData_BASICLIST_HDR) then 
			IF not IsArray (vntData_BASICLIST_DTL )THEN
				gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
				exit sub
			END IF
		End If
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "������ ��������Ʈ�� ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
		FOR  i = 1 to .sprSht_HDR.maxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"OUTSNAME",i) = "" THEN
				gErrorMsgBox "���������� �ʼ� �Դϴ�.","����ȳ�"	
				EXIT SUB
			END IF
		Next
		
		'������ VALIDATION
		if  not IsArray(vntData_BASICLIST_DTL) and .sprSht_DTL.MaxRows = 0 Then
			gErrorMsgBox "�ּ� �ϳ��� �׸��� �Է��ؾ� �մϴ�. " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		
		'�θ�â���� ���� STRBJONO�� ��������
		strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"JOBNO",.sprSht_HDR.ActiveRow) 
		
		
		'��Ʈ1�� SEQ������ NEW���� UPDATE������ ���
		strRow = .sprSht_HDR.ActiveRow
		strOLDSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"REVSEQ",strRow)
		IF strOLDSEQ = "" THEN strOLDSEQ = 0
		
		
		'��Ʈ2���� SEQ������ NEW ,  UPDATE
		if strOLDSEQ = 0 then
			strSEQFlag = "new"
		else
			strSEQFlag = "update"
		end if
		
		'OLDSEQ (HDR�� SEQ)�� �� �����ư�����. ��������� �� �ϰ��� DTL������ �ϵ��� �Ǿ��ֱ� �����̴�.
		intRtn = mobjccPDDCCHARGEEXCOM.ProcessRtn(gstrConfigXml,vntData_BASICLIST_HDR,vntData_BASICLIST_DTL,strJOBNO,strOLDSEQ )
		
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
			
			if strSEQFlag = "new" then
				gErrorMsgBox " �ڷᰡ �ű����� " & mePROC_DONE ,"����ȳ�" 
			else
				gErrorMsgBox " �ڷᰡ �������� " & mePROC_DONE , "����ȳ�"
			end if
  		end if
		
		
		'������ ��ƮŬ���� �ǰ� �ϱ����ؼ����..
		mstrGrid = TRUE
  		SelectRtn
 	end with
End Sub



'�ڷ����
Sub DeleteRtn ()

	Dim vntData
	Dim intCnt, intRtn, i
	Dim lngchkCnt
	
	
	with frmThis
		
		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "������ ������ �����ϴ�.","�����ȳ�!"
			Exit Sub
		END IF
		
		For i = 1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | REVSEQ | OUTSCODE ")
		intRtn = mobjccPDDCCHARGEEXCOM.DeleteRtn(gstrConfigXml,vntData)
		
		
		IF not gDoErrorRtn ("DeleteRtn") then
			'���õ� �ڷḦ ������ ���� ����
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "���ְ����� �����Ǿ����ϴ�.","�����ȳ�!"
			if .sprSht_HDR.MaxRows > 0 then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, 1,1
				mstrGrid = true
				SelectRtn_DTL 1,1
			else
				mstrGrid = FALSE
				SelectRtn
			end if
   		End IF
		
	End with
	err.clear
End Sub


		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="95%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<tr>
					<TD id="Td2" align="left" width="100%" height="20" runat="server">
						<TABLE id="tblTitle1" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="85" background="../../../images/back_p.gIF"
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
											<td class="TITLE">���ְ��� ��Ȳ&nbsp;</td>
										</tr>
									</table>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</tr>
				<TR>
					<TD vAlign="top">
						<TABLE class="SEARCHDATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" align="left"
							border="0">
							<TR>
								<TD class="SEARCHLABEL" style="CURSOR: hand" width="60"><FONT face="����">������</FONT></TD>
								<TD class="SEARCHDATA" width="224"><INPUT class="INPUT" id="txtFROM" title="�Ⱓ�˻�(FROM)" style="WIDTH: 88px; HEIGHT: 22px"
										accessKey="DATE" type="text" maxLength="10" size="9" name="txtFROM"> <IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"										border="0" name="imgCalEndarFROM1">~<INPUT class="INPUT" id="txtTO" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
										type="text" maxLength="10" size="7" name="txtTO"> <IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
										 align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)"
									width="60">����ó</TD>
								<TD class="SEARCHDATA" width="263"><INPUT class="INPUT_L" id="txtOUTSNAME" title="����ó" style="WIDTH: 170px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="37" name="txtOUTSNAME"> <IMG id="imgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
										border="0" name="imgOUTSCODE"> <INPUT class="INPUT" id="txtOUTSCODE" title="����ó" style="WIDTH: 65px; HEIGHT: 22px" accessKey=",M"
										type="text" maxLength="6" size="9" name="txtOUTSCODE"></TD>
								<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)"
									width="60">JOB��</TD>
								<TD class="SEARCHDATA" width="263"><INPUT class="INPUT_L" id="txtJOBNAME" title="JOBNO" style="WIDTH: 170px; HEIGHT: 22px"
										type="text" maxLength="100" align="left" size="23" name="txtJOBNAME"> <IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"  align="absMiddle"
										border="0" name="ImgJOBNO"> <INPUT class="INPUT" id="txtJOBNO" title="JOBNO" style="WIDTH: 65px; HEIGHT: 22px" type="text"
										maxLength="7" align="left" size="5" name="txtJOBNO"></TD>
								<TD class="SEARCHDATA2" align="right" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
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
						<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
							<TR>
								<TD class="TITLE" style="WIDTH: 100%; HEIGHT: 8px" vAlign="absmiddle"></TD>
							</TR>
							<TR>
								<TD class="TITLE" width="210" vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_normal.gif'
											align='absMiddle' border='0' name='btn_normal'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')">
										<IMG id='btn_multi' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_multi.gif'
											align='absMiddle' border='0' name='btn_multi'></span>&nbsp; <span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')">
										<IMG id='btn_hide' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_hide.gif'
											align='absMiddle' border='0' name='btn_hide'></span>
								</TD>
							</TR>
						</table>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" align="left" width="400" height="20">
									<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td class="TITLE" vAlign="absmiddle">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
												<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgNEW" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" height="20" alt="�ڷḦ �߰��մϴ�."
													src="../../../images/imgNew.gIF" border="0" name="imgNEW"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											<TD><IMG id="imgExcel_HDR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel_HDR"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<!--���̺��� �������°��� �����ش�-->
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR id="tblBody1">
					<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_HDR" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="42545">
								<PARAM NAME="_ExtentY" VALUE="3334">
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
					<TD class="BOTTOMSPLIT" id="lblStatus1" style="WIDTH: 1040px"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="64" background="../../../images/back_p.gIF"
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
												<td class="TITLE">���۸���Ʈ&nbsp;</td>
											</tr>
										</table>
									</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgExcel_DTL" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel_DTL"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="left">
						<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="42545">
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
					<TD class="BOTTOMSPLIT" id="lblStatus2" style="WIDTH: 1040px"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
