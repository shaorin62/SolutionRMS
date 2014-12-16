<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMEXELIST.aspx.vb" Inherits="PD.PDCMEXELIST" %>
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
'HISTORY    :1) 2009/09/18 By KimTH
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
'=============================
' �̺�Ʈ ���ν��� 
'=============================
option explicit
Const meTAB = 9
Dim mlngRowCnt, mlngColCnt
Dim mobjPDCMEXE, mobjPDCMGET
'����üũ��
Dim mstrCheck
Dim mALLCHECK

'�������� ���������� true   �ƴϰ� exe_hdr �� �ִٸ�  �ʱⰪ�� false
Dim strACTUALFLAG
'����� ���泻�� ����    �⺻ false ���� true
Dim mstrHEADERFLAG 
Dim mstrPROCESS

Dim strJOBNO 
Dim strPREESTNO

mALLCHECK = TRUE
mstrCheck=TRUE
strACTUALFLAG = FALSE
mstrPROCESS = False
mstrHEADERFLAG = false
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

Sub imgSetting_onclick ()
	with frmThis
'		If .txtENDDAY.value <> "" Then
'			gErrorMsgbox "�̹� Ȯ���� ���ֺ� �Դϴ�.","ó���ȳ�!"
'			Exit Sub
'		End If
	End with
	
	gFlowWait meWAIT_ON
	UpdateRtn_ENDDAY
	gFlowWait meWAIT_OFF
End Sub

Sub imgConfirmCancel_onclick ()	
	with frmThis
'		If .txtENDDAY.value = "" Then
'			gErrorMsgbox "�̹� Ȯ����� ���ֺ� �Դϴ�.","ó���ȳ�!"
'			Exit Sub
'		End If
	End with	
	
	gFlowWait meWAIT_ON
	DeleteRtn_ENDDAY
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowAdd_onclick ()
	CALL sprSht_Keydown(meINS_ROW, 0)	
	mstrPROCESS = False
end Sub

Sub imgRowDel_onclick
	with frmThis
'		If .txtENDDAY.value <> "" Then
'			gErrorMsgbox "Ȯ���� ���ֺ�� ���� �Ҽ������ϴ�.","ó���ȳ�!"
'			Exit Sub
'		End If
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
'	with frmThis
'		If .txtENDDAY.value <> "" Then
'			gErrorMsgbox "Ȯ���� ó���� Ȯ������� �����մϴ�.","ó���ȳ�!"
'			Exit Sub
'		End If
'	End with
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick
	with frmThis
'		If .txtENDDAY.value <> "" Then
'			gErrorMsgbox "Ȯ���� ���ֺ�� ���� �Ҽ������ϴ�.","ó���ȳ�!"
'			Exit Sub
'		End If
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn_ALL
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

Sub ImgAccInput_onclick()
	Dim vntInParams
	Dim vntRet
	Dim vntData
	Dim strGUBN
	Dim intRtn
	
	with frmThis
		If .txtJOBNO.value = "" Then
			gErrorMsgbox "���۹�ȣ ��ȸ�� �Է� ���� �մϴ�.","ó���ȳ�!"
			Exit Sub
		End If
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		If .txtENDDAY.value <> "" Then
			strGUBN = "END"
		Else
			strGUBN = ""
		End If
	
		'exe_hdr�� ��ȸ ���ʿ���� strACTUALFLAG�� �Ǵ�   ( strACTUALFLAG�� ó�� ���۽� ���������� ������������ ��ȸ�Ѵ�.)
		If .txtJOBNOINS.value = ""  then
			intRtn = gYesNoMsgbox("������ �����ؾ� ����� �Է��Ҽ��ֽ��ϴ�. �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
			if intRtn <> vbYes then exit sub
			
			IF .sprSht.MaxRows <> 0 THEN
				ProcessRtn
			END IF 
		end if
			
		vntInParams = array(strJOBNO,strGUBN)
		vntRet = gShowModalWindow("PDCMACCLISTPOP.aspx",vntInParams , 550,540)
		
		SelectRtn
		mstrHEADERFLAG = true
		Payment_changevalue
		DelProc
	End with
End Sub

Sub imgPrint_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim intRtn
	Dim i, j, intCount
	Dim strJOBNO
	Dim strUSERID
	Dim vntDataTemp
	
		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if frmThis.sprSht.MaxRows = 0 then
			gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.","�μ����"
			Exit Sub
		end if
		
		gFlowWait meWAIT_ON
		with frmThis
		
			'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			intRtn = mobjPDCMEXE.DeleteRtn_TEMP(gstrConfigXml)
		
			ModuleDir = "PD"
			ReportName = "PDCMEXEAMT.rpt"
			
		
			strJOBNO = parent.document.forms("frmThis").txtJOBNO.value 
			strUSERID = ""
			vntDataTemp = mobjPDCMEXE.ProcessRtn_TEMP(gstrConfigXml,strJOBNO, 1, strUSERID)
	
			Params = strUSERID 
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	
End Sub	


'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMEXE.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub


'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------

Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub

Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		CALL gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub

Sub txtDEMANDAMT_onfocus
	with frmThis
		.txtDEMANDAMT.value = Replace(.txtDEMANDAMT.value,",","")
	end with
End Sub
Sub txtDEMANDAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub

Sub txtESTAMT_onfocus
	with frmThis
		.txtESTAMT.value = Replace(.txtESTAMT.value,",","")
	end with
End Sub
Sub txtESTAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtESTAMT,0,true)
	end with
End Sub

Sub txtPAYMENT_onfocus
	with frmThis
		.txtPAYMENT.value = Replace(.txtPAYMENT.value,",","")
	end with
End Sub

Sub txtPAYMENT_onblur
	with frmThis
		CALL gFormatNumber(.txtPAYMENT,0,true)
	end with
End Sub

Sub txtINCOM_onfocus
	with frmThis
		.txtINCOM.value = Replace(.txtINCOM.value,",","")
	end with
End Sub
Sub txtINCOM_onblur
	with frmThis
		CALL gFormatNumber(.txtINCOM,0,true)
	end with
End Sub

Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		CALL gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub

Sub txtACCAMT_onfocus
	with frmThis
		.txtACCAMT.value = Replace(.txtACCAMT.value,",","")
	end with
End Sub
Sub txtACCAMT_onblur
	with frmThis
		CALL gFormatNumber(.txtACCAMT,0,true)
	end with
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
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

sub sprSht_DblClick (ByVal Col, ByVal Row)
Dim vntInParams
Dim vntRet
Dim strCONTRACTNO
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If Col = 17 AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row) <> "" Then
				strCONTRACTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row)	
				vntInParams = array(strCONTRACTNO)
				vntRet = gShowModalWindow("PDCMCONTRACTPOP.aspx",vntInParams , 1060,900)
			End If
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

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT"))   Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next
				
			.txtSELECTAMT.value = strSUM
			CALL gFormatNumber(.txtSELECTAMT,0,True)
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")  or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
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
		CALL gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

'����ó ��ưŬ��
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams 
	Dim strBUSINO
	with frmThis
	
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then 
	
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				
				strBUSINO =  vntRet(2,0)
				if left(strBUSINO,3) = "000" then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 0
				else
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 1
				end if 
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			.txtCLIENTNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		END IF
	End with
End Sub

'���������� �׸��� ���ҽ� ��� �Լ��� �¿���� �Ҷ� ���
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		
		IF Col = 7 Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)

			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtACCAMT.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		If mstrPROCESS = True Then
			frmThis.sprSht.MaxRows = 0
		End If
		frmThis.txtSUMAMT.value = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)

		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"QTY",frmThis.sprSht.ActiveRow,1
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"VATCODE",frmThis.sprSht.ActiveRow,"�ڵ弱��"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"INCOMCODE",frmThis.sprSht.ActiveRow,"����ҵ�(3,3%)"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"REGDATE",frmThis.sprSht.ActiveRow,gNowDate
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMTFLAG",frmThis.sprSht.ActiveRow,"1"
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		frmThis.txtCLIENTNAME.focus
		frmThis.sprSht.focus
	End if
End Sub


Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, strCols
	Dim strJOBNO, strJOBNOName
	Dim strBUSINO
	Dim lngQTY, lngPRICE

	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strJOBNO = ""
		strJOBNOName = ""
		IF Col =  mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then
			strJOBNO = ""
			strJOBNOName = mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",.sprSht.ActiveRow)
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strJOBNOName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntData(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntData(1,0)
				strBUSINO =  vntData(2,0)
				if left(strBUSINO,3) = "000" then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 0
				else
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 1
				end if 
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
			Else
				mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
			End If
			.txtCLIENTNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		
		ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") then
			lngQTY = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",Row))
			lngPRICE = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngQTY*lngPRICE
		ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") then
			lngQTY = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",Row))
			lngPRICE = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngQTY*lngPRICE
		ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") then
			Payment_changevalue
		END IF
	end with
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


'���������� �׸��� ���ҽ� ��� �Լ��� �¿���� �Ҷ� ���
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	Dim strBUSINO
	With frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				
				strBUSINO =  vntRet(2,0)
				if left(strBUSINO,3) = "000" then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 0
				else
					mobjSCGLSpr.SetTextBinding .sprSht,"AMTFLAG",Row, 1
				end if 
						
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtCLIENTNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		
		end if
	End With
End Sub


'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i
	Dim strComboList
	Dim strComboList2
	Dim strMSG
	
	'����������ü ����	
	set mobjPDCMEXE	= gCreateRemoteObject("cPDCO.ccPDCOEXE")
	set mobjPDCMGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis
		strComboList =  "�ڵ弱��" & vbTab & "���ݰ�꼭(10%)" & vbTab & "���ݰ�꼭�Ұ���" & vbTab & "���ݰ�꼭������" & vbTab & "��꼭" & vbTab & "INVOICE" & vbTab & "����ҵ�(3,3%)" & vbTab & "��Ÿ�ҵ�(22%)" & vbTab & "��Ÿ�ҵ�(�ʿ���80%)" & vbTab & "�������(���Ѽ���)" & vbTab & "�������" & vbTab & "��Ÿ"
		strComboList2 =  "������"
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0, 13
		mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,   "CHK | JOBNO | PREESTNO | SORTSEQ | OUTSCODE | BTN | OUTSNAME | STD | QTY | PRICE | AMT | ADJAMT | CONTRACTNO | VATCODE | INCOMCODE | AMTFLAG | REGDATE | ADDFLAG | SEQ | ADJDAY | VOCHNO"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|���۹�ȣ|������ȣ|����|����ó�ڵ�|����ó|�����׸�|����|�ܰ�|�ݾ�|���޾�|��༭��ȣ|�����ڵ�|�ҵ汸���ڵ�|�ϵ���|�������|���Ա���|��ȣ|��ǥ��|��ǥ��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "  4|        0|       0|   0|         6|2|  18|      10|   4|   9|   9|     9|         9|      10|           0|     6|	    8|       0|   0|     8|       9"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | AMTFLAG "
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SORTSEQ | QTY | PRICE | AMT | ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "OUTSNAME | STD | VATCODE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY | REGDATE", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "OUTSCODE | SORTSEQ | ADJDAY | CONTRACTNO | PREESTNO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PREESTNO | SORTSEQ | OUTSCODE | CONTRACTNO | VOCHNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "JOBNO | QTY | PRICE | AMT | ADDFLAG | SEQ | INCOMCODE", true
		
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,14,14,-1,-1,strComboList
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,15,15,-1,-1,strComboList2
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0
		
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_PREEST
		mobjSCGLSpr.SpreadLayout .sprSht_PREEST, 2, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_PREEST,   "ITEMNAME| AMT"
		mobjSCGLSpr.SetHeader .sprSht_PREEST,		   "�����׸�|�ݾ�"
		mobjSCGLSpr.SetColWidth .sprSht_PREEST, "-1", "       20|   9"
		mobjSCGLSpr.SetRowHeight .sprSht_PREEST, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_PREEST, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_PREEST, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_PREEST, true, "ITEMNAME | AMT"
		mobjSCGLSpr.SetCellAlign2 .sprSht_PREEST, "ITEMNAME",-1,-1,0,2,false
		
	    .sprSht_PREEST.style.visibility  = "visible"
		.sprSht_PREEST.MaxRows = 0
		
		InitPageData
		'�θ�â�� ������ ��������  (�������������)
		
		.txtJOBNO.value = parent.document.forms("frmThis").txtJOBNO.value 
		strJOBNO = parent.document.forms("frmThis").txtJOBNO.value 
		
		.txtPREESTNO.value = parent.document.forms("frmThis").txtPREESTNO.value 
		strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value 
		
		SelectRtn
	End With
End Sub

Sub EndPage()
	'set mobjPDCMEXE = Nothing
	'set mobjPDCMGET = Nothing
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	TestControl_hidding

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub TestControl_hidding
	With frmThis
		'.txtPREESTNO.style.visibility = "hidden" 
		'.txtJOBNO.style.visibility = "hidden" 
		'.txtENDDAY.style.visibility = "hidden"
	End With
End Sub
'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	with frmThis
		if strJOBNO = "" Or Len(strJOBNO) <> 7 Then
			gErrorMsgBox "���۹�ȣ��Ȯ���Ͻʽÿ�.","��ȸ�ȳ�!"
			Exit Sub
		End if
	
		'�������������� ������� preest_dtl��ȸ
		CALL SelectRtn_Preest ()
		
		'JOBNO�� ���굥��Ÿ�� �����´�. ������FALSE
		IF SelectRtn_Head Then 
			CALL SelectRtn_Detail ()
		else
			CALL SelectRtn_Actual_Head ()
			CALL SelectRtn_Actual_Detail ()	
		END IF
		
		txtSUSUAMT_onblur
		txtCOMMITION_onblur
		txtDEMANDAMT_onblur
		txtPAYMENT_onblur
		txtINCOM_onblur
		txtNONCOMMITION_onblur
		txtACCAMT_onblur
		txtESTAMT_onblur
		AMT_SUM
		mstrHEADERFLAG = false
	End with
End Sub

Function SelectRtn_Preest
	Dim vntData
	
	With frmThis 
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		vntData = mobjPDCMEXE.SelectRtn_Preest(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
		IF not gDoErrorRtn ("SelectRtn_Preest") then
			'��ȸ�� �����͸� ���ε�
			CALL mobjSCGLSpr.SetClipBinding (frmThis.sprSht_PREEST,vntData,1,1,mlngColCnt,mlngRowCnt,true)
			'�ʱ� ���·� ����
			IF mlngRowCnt > 0 THEN
				gWriteText lblStatus_PREEST, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht_PREEST.MaxRows = 0
			END IF
			mobjSCGLSpr.SetFlag  .sprSht_PREEST,meCLS_FLAG
		End IF
		
	End With
End Function

Function SelectRtn_Head
	Dim vntData
	SelectRtn_Head = false
	'on error resume next
	'�ʱ�ȭ
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMEXE.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	IF not gDoErrorRtn ("SelectRtn_HDR") then
		IF mlngRowCnt <=0 then
			'gErrorMsgBox "Ȯ���������� " & meNO_DATA ,""
			SelectRtn_Head = FALSE
			strACTUALFLAG = TRUE
			gClearAllObject frmThis
		else
			'��ȸ�� �����͸� ���ε�
			SelectRtn_Head = True
			CALL gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
		End IF
	End IF
End Function

'���� ���̺� ��ȸ
Function SelectRtn_Detail
	dim vntData
	Dim strRows
	Dim intCnt
	Dim lngRowCnt
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMEXE.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
	IF not gDoErrorRtn ("SelectRtn_DTL") then
		'��ȸ�� �����͸� ���ε�
		CALL mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		
		lngRowCnt = mlngRowCnt
		SelectRtn_Detail = True
		
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt = 1 To .sprSht.MaxRows
					'.txtENDDAY.value <> "" or
					If  mobjSCGLSpr.GetTextBinding(.sprSht, "CONTRACTNO",intCnt) <> "" or mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt) <> "" then 'Ư������ �ش� ������ �⺻���� ����
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,2,-1,true
						if mobjSCGLSpr.GetTextBinding(.sprSht, "CONTRACTNO",intCnt) = "" then
							mobjSCGLSpr.SetCellsLock2 .sprSht, FALSE, "AMTFLAG"
						end if 
					ELSE
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False '�̰� ���
						mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,intCnt,-1,-1,true
						mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SORTSEQ|ADJDAY|CONTRACTNO"
					END IF
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",intCnt,"�ڵ弱��"
					sprSht_Change 18,intCnt
					
					End If
					
				Next
				gWriteText lblStatus, lngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function


Function SelectRtn_Actual_Head
	Dim vntData
	Dim vntData_empty
	'on error resume next
	
	'�ʱ�ȭ
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'
	vntData	= mobjPDCMEXE.SelectRtn_Actual_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
	IF not gDoErrorRtn ("SelectRtn_Actual_HDR") then
		IF mlngRowCnt > 0 then
			'��ȸ�� �����͸� ���ε�
			CALL gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			
			'���ε��� �Ŀ��� ������ jobno�� preestno �� �ٽ� ���������� �־��ش�.
			'strJOBNO	= frmThis.txtJOBNO.value
			'strPREESTNO = frmThis.txtPREESTNO.value
		Else
		'gClearAllObject frmThis
			vntData_empty =  mobjPDCMEXE.SelectRtn_Actual_HDR_EMPTY(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
			CALL gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData_empty)
		End IF
	End IF
End Function

'���� ���̺� ��ȸ
Function SelectRtn_Actual_Detail
	
	Dim vntData
	Dim intCnt
	Dim strRows
	Dim intCnt2
	Dim lngRowCnt
	'on error resume next	
	'�ʱ�ȭ
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMEXE.SelectRtn_Actual_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)

	IF not gDoErrorRtn ("SelectRtn_Actual_DTL") then
		'��ȸ�� �����͸� ���ε�
		CALL mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'�ʱ� ���·� ����
		lngRowCnt = mlngRowCnt
		
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt = 1 To .sprSht.MaxRows
					'.txtENDDAY.value <> ""  or
					If  mobjSCGLSpr.GetTextBinding(.sprSht, "CONTRACTNO",intCnt) <> "" or mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt) <> "" then 'Ư������ �ش� ������ �⺻���� ����
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HCCFFFF, &H000000,False
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,-1,-1,true
					ELSE
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False '�̰� ���
						mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,intCnt,-1,-1,true
						mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SORTSEQ|ADJDAY|CONTRACTNO"
					END IF
					
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",intCnt,"�ڵ弱��"
					sprSht_Change 18,intCnt
					
					End If
					
				Next
				gWriteText lblStatus, lngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			CALL gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'***********************************************
' Ȯ�� ENDDAY���ó�¥�� ������Ʈ
'***********************************************
Sub UpdateRtn_ENDDAY ()
	Dim intRtn
	
	with frmThis
		
		intRtn = mobjPDCMEXE.UpdateRtn_ENDDAY(gstrConfigXml,strJOBNO)
		
		if not gDoErrorRtn ("UpdateRtn_ENDDAY") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "�ڷᰡ Ȯ���Ǿ����ϴ�.","����ȳ�" 
			SelectRtn
  		end if
	End with
End Sub

'***********************************************
' Ȯ�����  ENDDAY  '' ������Ʈ
'***********************************************
Sub DeleteRtn_ENDDAY ()
	Dim intRtn
	
	with frmThis
		intRtn = mobjPDCMEXE.DeleteRtn_ENDDAY(gstrConfigXml,strJOBNO)
		
		if not gDoErrorRtn ("DeleteRtn_ENDDAY") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "�ڷᰡ Ȯ����� �Ǿ����ϴ�.","����ȳ�" 
			SelectRtn
  		end if
	End with
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
    Dim intRtn , intCnt,intCnt2
  	dim vntData , lngRow
	Dim strMasterData
	Dim intCHK
	Dim intConRtn
	Dim strDataCHK
	with frmThis
	
	'On error resume next
		if strJOBNO = "" Then
			gErrorMsgBox "��ȸ�� ���۰�����ȣ�� �����ϴ�.","����ȳ�!"
			Exit Sub
		End If
		
		intCHK = 0
		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) = "" then 
				intCHK = intCHK + 1
			End If
		next
		If intCHK <> 0  Then
			intConRtn = gYesNoMsgbox("����ó���� ���� �ڷ�� �����˴ϴ�.�ڵ����� �ϰ� �����Ͻðڽ��ϱ�?","����Ȯ��")
			If intConRtn <> vbYes Then exit Sub
			for intCnt2 = .sprSht.MaxRows to 1 step -1
				if mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt2) = "" then 
					mobjSCGLSpr.DeleteRow .sprSht,intCnt2
					
				End If
			Next
		End If
		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, " OUTSCODE | OUTSNAME | STD | ADJAMT  ", False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ����ó/�׸��/���ݾ� ��  �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		End If
		
  		'������ Validation
		if DataValidation = false then exit sub
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | JOBNO | PREESTNO | SORTSEQ | OUTSCODE | BTN | OUTSNAME | STD | QTY | PRICE | AMT | ADJAMT | CONTRACTNO | VATCODE | INCOMCODE | AMTFLAG | REGDATE | ADDFLAG | SEQ | ADJDAY | VOCHNO")
		strMasterData = gXMLGetBindingData (xmlBind)
		
		if  not IsArray(vntData)  Then 'XML ������ �� ����Ȱ��� �ִٸ� 'AND mstrHEADERFLAG = FALSE
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		'ó�� ������ü ȣ��
		intRtn = mobjPDCMEXE.ProcessRtn(gstrConfigXml,strMasterData,vntData,strJOBNO,strPREESTNO)
				
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ" & intRtn & " �� ����" & mePROC_DONE,"����ȳ�" 
			SelectRtn
			parent.jobMst_Tab6Search
			parent.jobMst_Tab4Search
  		end if
  		Payment_changevalue
 	end with
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
   		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) <> "" AND mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "�ڵ弱��" Then 
				gErrorMsgBox intCnt & " ��° ���� �����ڵ� �� Ȯ���Ͻʽÿ�","�Է¿���"
				Exit Function
			End if
			
		next
   	End with
	DataValidation = true
End Function


'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
'�ڷ����
Sub DeleteRtn ()
	Dim intRtn, i , intCnt
	Dim lngchkCnt
	Dim strSEQFLAG 
	Dim dblSEQ  'JOBNO �� ��������
	
	strSEQFLAG= false
	
	with frmThis
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",i) <> "" Then
					gErrorMsgbox "�����̵� ������ �����Ҽ� �����ϴ�.","ó���ȳ�!"
					Exit Sub
				end if
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
					gErrorMsgbox "��༭�� �ִ� ������ �����Ҽ� �����ϴ�.","ó���ȳ�!"
					Exit Sub
				end if
				lngchkCnt = lngchkCnt +1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				dblSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",i))
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					'JOBNO�����������
					intRtn = mobjPDCMEXE.DeleteRtn(gstrConfigXml,strJOBNO,dblSEQ)
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		
		'�� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			Payment_changevalue
			DelProc
		End If
	End with
	err.clear
End Sub

'��ü ����
Sub DeleteRtn_ALL
	Dim intRtn , intCnt, i
	Dim vntData
	
	with frmThis

		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		'�������� ���嵵���ߴµ� �����ҽ� validation
		vntData = mobjPDCMEXE.SelectRtn_ACCEXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)
		IF not gDoErrorRtn ("SelectRtn_Detail") then
			If mlngRowCnt = 0 Then
				gErrorMsgBox "���� �� �����Ͱ� �����ϴ�.","��ü�����ȳ�!"
				Exit Sub	
			End If
		End If

		'������,��༭ validation
		For intCnt =1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJDAY",frmThis.sprSht.ActiveRow) <> "" then
				gErrorMsgBox "�������� �����ϴ� ���� ��ü���� �ɼ� �����ϴ�.","��ü�����ȳ�!"
				Exit Sub
			End If
			if mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CONTRACTNO",frmThis.sprSht.ActiveRow) <> "" then
				gErrorMsgBox "��༭�� �ִ� ������ �����Ҽ� �����ϴ�","��ü�����ȳ�!"
				Exit Sub
			End If
		Next
		
		intRtn = gYesNoMsgbox("�ڷḦ ��ü �����Ͻðڽ��ϱ�?" & vbcrlf & "��ü�ڷᰡ �����˴ϴ�.","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		intRtn = mobjPDCMEXE.DeleteRtn_ALL(gstrConfigXml,strJOBNO)
		if not gDoErrorRtn ("DeleteRtn_ALL") then
			gOkMsgbox "��ü������ �Ϸ�Ǿ����ϴ�.","�����ȳ�"
			SelectRtn
		End If
	End with 
End Sub

Sub Payment_changevalue
	Dim lngDEMANDAMT
	Dim lngPAYMENTAMT
	Dim lngACCAMT , lngAMT , lngAMTSUM
	Dim lngRATE
	Dim intCnt
	
	with frmThis
		lngACCAMT = Replace(.txtACCAMT.value,",","")
		
		For intCnt = 1 To .sprSht.MaxRows
			lngAMT = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"ADJAMT",intCnt))
			lngAMTSUM = lngAMTSUM + lngAMT
		Next
		
		'���ֺ� ���޾� + �����
		.txtPAYMENT.value = lngAMTSUM +lngACCAMT
		
		'û���ݾ�
		lngDEMANDAMT = Replace(.txtDEMANDAMT.value,",","")
		'���ֺ�
		lngPAYMENTAMT = Replace(.txtPAYMENT.value,",","")
		
		'������  û���� - ���ֺ�
		If lngDEMANDAMT = 0 Then lngRATE = 0
		.txtINCOM.value = lngDEMANDAMT-lngPAYMENTAMT
		If lngDEMANDAMT = 0 Then
			.txtRATE.value = 0
		Else
			
			.txtRATE.value = gRound(((lngDEMANDAMT-lngPAYMENTAMT)/lngDEMANDAMT)*100,2)
		End If
		txtPAYMENT_onblur
		txtINCOM_onblur
		 
	End with
End Sub

Sub DelProc
	Dim intHDR
	Dim strMasterData
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		intHDR = mobjPDCMEXE.ProcessRtn_DELHDR(gstrConfigXml,strMasterData)
		if not gDoErrorRtn ("ProcessRtn_DELHDR") then
			SelectRtn
		End If
	End with
End Sub
		</script>
	</HEAD>
	<body class="base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="54" background="../../../images/back_p.gIF"
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
											<td class="TITLE">�������&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="20">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 246px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton2" style="WIDTH: 110px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<td><INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 20px" dataSrc="#xmlBind" size="1" name="txtJOBNO"
													type="hidden"><INPUT dataFld="JOBNOINS" id="txtJOBNOINS" style="WIDTH: 20px" dataSrc="#xmlBind" size="1"
													name="txtJOBNOINS" type="hidden"><INPUT dataFld="PREESTNO" id="txtPREESTNO" style="WIDTH: 20px" dataSrc="#xmlBind" size="1"
													name="txtPREESTNO" type="hidden"><INPUT dataFld="ENDDAY" id="txtENDDAY" style="WIDTH: 20px" dataSrc="#xmlBind" size="1"
													name="txtENDDAY" type="hidden"><IMG id="ImgAccInput" onmouseover="JavaScript:this.src='../../../images/ImgAccInputOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAccInput.gIF'" height="20" alt="���������" src="../../../images/ImgAccInput.gIF"
													align="absMiddle" border="0" name="ImgAccInput"></td>
											<!--<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
													height="20" alt="�ڷḦ Ȯ���մϴ�." src="../../../images/imgSetting.gIF" border="0" name="imgSetting"></TD>
											<TD><IMG id="imgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgConfirmCancelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirmCancel.gIF'"
													height="20" alt="�ڷḦ Ȯ������մϴ�." src="../../../images/imgConfirmCancel.gIF" border="0"
													name="imgConfirmCancel"></TD>-->
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></TD>
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
				<TR>
					<TD vAlign="top" width="100%">
						<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1"
							cellPadding="0" align="right" border="0">
							<TR>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">������Ʈ��</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="PROJECTNM" class="NOINPUTB_R" id="txtPROJECTNM" title="������Ʈ��" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPROJECTNM"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">������</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="CLIENTNAME" class="NOINPUTB_R" id="txtCLIENTNAME" title="������" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTNAME"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">�����ݾ�</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="ESTAMT" class="NOINPUTB_R" id="txtESTAMT" title="�����ݾ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtESTAMT"></TD>
								<TD class="SEARCHLABEL" style="WIDTH: 106px">Noncommition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="����������ұݾ�"
										style="WIDTH: 152px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtNONCOMMITION"></TD>
							</TR>
							<TR>
								<TD class="SEARCHLABEL">JOB��</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="JOBNAME" class="NOINPUTB_R" id="txtJOBNAME" title="JOB��" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBNAME"></TD>
								<TD class="SEARCHLABEL">��</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="TIMNAME" class="NOINPUTB_R" id="txtTIMNAME" title="����" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtTIMNAME"></TD>
								<TD class="SEARCHLABEL">û���ݾ�</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDAMT" class="NOINPUTB_R" id="txtDEMANDAMT" title="û���ݾ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDAMT"></TD>
								<TD class="SEARCHLABEL">Commition</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="���������ұݾ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCOMMITION"></TD>
							</TR>
							<tr>
								<TD class="SEARCHLABEL">��ü�ι�</TD>
								<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="JOBGUBN" class="NOINPUTB_R" id="txtJOBGUBN" title="��ü�ι�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="txtJOBGUBN"></TD>
								<TD class="SEARCHLABEL">�귣��</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUTB_R" id="txtSUBSEQNAME" title="�귣��" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUBSEQNAME"></TD>
								<TD class="SEARCHLABEL">���ֺ�</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="PAYMENT" class="NOINPUTB_R" id="txtPAYMENT" title="���ֺ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPAYMENT"></TD>
								<TD class="SEARCHLABEL">������</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="�������հ�ݾ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUSUAMT"></TD>
							</tr>
							<tr>
								<TD class="SEARCHLABEL">��ü�з�</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUTB_R" id="txtCREPART" title="��ü�з�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="6" name="txtCREPART"></TD>
								<TD class="SEARCHLABEL">û����</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="DEMANDDAY" class="NOINPUTB_R" id="txtDEMANDDAY" title="û����" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDDAY"></TD>
								<TD class="SEARCHLABEL">�����</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ACCAMT" class="NOINPUTB_R" id="txtACCAMT" title="��� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtACCAMT"></TD>
								<TD class="SEARCHLABEL">��������</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="SUSURATE" class="NOINPUTB_R" id="txtSUSURATE" title="��������" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtSUSURATE">&nbsp;(%)</TD>
							</tr>
							<TR>
								<TD class="SEARCHLABEL">����</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="ENDFLAG" class="NOINPUTB_R" id="cmbENDFLAG" title="����" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="8" name="cmbENDFLAG"></TD>
								<TD class="SEARCHLABEL">�����</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="CLOSEDAY" class="NOINPUTB_R" id="txtClOSEDAY" title="�����" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtClOSEDAY"></TD>
								<TD class="SEARCHLABEL">������</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="INCOM" class="NOINPUTB_R" id="txtINCOM" title="������" style="WIDTH: 152px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="20" name="txtINCOM"></TD>
								<TD class="SEARCHLABEL">������</TD>
								<TD class="SEARCHDATA"><INPUT dataFld="RATE" class="NOINPUTB_R" id="txtRATE" title="������" style="WIDTH: 128px; HEIGHT: 22px"
										dataSrc="#xmlBind" readOnly type="text" size="16" name="txtRATE">&nbsp;(%)</TD>
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
								<TD align="left" width="80" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="54" background="../../../images/back_p.gIF"
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
											<td class="TITLE">��������&nbsp;</td>
										</tr>
									</table>
								</TD>
								<td class="TITLE">
									�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
									<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
								<TD style="WIDTH: 100%" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton" style="WIDTH: 110px; HEIGHT: 24px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'"
													height="20" alt="�ڷ��Է��� ���� �����߰��մϴ�." src="../../../images/imgRowAdd.gIF" border="0"
													name="imgRowAdd"></TD>
											<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'"
													height="20" alt="������ ���������մϴ�." src="../../../images/imgRowDel.gIF" border="0" name="imgRowDel"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
											<!--<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
											-->
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="1075" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<TABLE height="98%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
							<TR>
								<td style="WIDTH: 290px; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 290px; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_PREEST" height="100%" width="290" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="7673">
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
								</td>
								<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="24183">
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
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus_PREEST" style="WIDTH: 1040px"></TD>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
						</TABLE>
					</TD>
				</tr>
			</TABLE>
		</FORM>
	</body>
</HTML>
