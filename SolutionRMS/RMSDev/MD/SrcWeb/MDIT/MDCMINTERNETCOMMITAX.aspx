<%@ Page CodeBehind="MDCMINTERNETCOMMITAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMINTERNETCOMMITAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ͳݱ��� ������ ���ݰ�꼭 ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ó�� ǥ�� ����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/17 By Kim Tae Yub
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
Dim mobjMDITINTERNETCOMMITAX , mobjMDCOGET
Dim mstrCheck
Dim mstrGUBUN
CONST meTAB = 9
mstrGUBUN = ""
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
	if frmThis.txtTRANSYEARMON1.value = "" then
	    gErrorMsgBox "��� �Է��Ͻÿ�",""
		exit Sub
	end if
	If LEN(frmThis.txtTRANSYEARMON1.value) <> 6 Then
		 gErrorMsgBox "����� 6�ڸ� �Դϴ�",""
		exit Sub
	End If
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgPrint_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "�μ�� �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim VATFLAG
	Dim FLAG
	Dim strUSERID
	
	IF frmThis.sprSht.MaxRows = 0 then
		gFlowWait meWAIT_ON
		with frmThis		
			ModuleDir = "MD"
			ReportName = "TAXNO_BLACK.rpt"
						
			IF .cmbFLAG.value = "receipt" THEN
				FLAG = "Y"
			ELSE
				FLAG = "N"
			END IF
						
			Params = FLAG
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
		end with
		gFlowWait meWAIT_OFF
	else
	
		'üũ�� �� �����Ͱ� �ִ��� ������ üũ�Ѵ�.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
				intCount = 1
			end if
		next
		
		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if intCount = 0 then
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�. �μ��� �����͸� üũ�Ͻÿ�",""
			Exit Sub
		end if
		
		gFlowWait meWAIT_ON
		with frmThis
			'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			'md_trans_temp���� ����
			intRtn = mobjMDITINTERNETCOMMITAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp���� ��
			
			ModuleDir = "MD"
			ReportName = "TAX_NEW.rpt"
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
					strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
					
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"VAT",i) = 0 THEN
						VATFLAG = "N"
					ELSE
						VATFLAG = "Y"
					END IF
					IF .cmbFLAG.value = "receipt" THEN
						FLAG = "Y"
					ELSE
						FLAG = "N"
					END IF
					strUSERID = ""
					
					vntDataTemp = mobjMDITINTERNETCOMMITAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			Params = "MD_TAXINTERNET_TEMP" & ":" & strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
			
		end with
		gFlowWait meWAIT_OFF
	END IF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDITINTERNETCOMMITAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub ImgTaxCre_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "���ݰ�꼭 ������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdF.checked <> true then
		gErrorMsgBox "���ݰ�꼭������ �̿Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	For i = 1 To frmThis.sprSht.MaxRows
		IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = 1 THEN
			chkcnt = chkcnt + 1
		END IF
	next
	if chkcnt = 0 then
		gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","����ȳ�!"
		exit sub
	end if
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick ()
	Dim i
	Dim chkcnt
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "������ �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	For i = 1 To frmThis.sprSht.MaxRows
		IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = 1 THEN
			chkcnt = chkcnt + 1
		END IF
	next
	if chkcnt = 0 then
		gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
		exit sub
	end if
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub btnCOMMISSION_onclick ()
	Dim intCnt
	Dim intRtn
	Dim strDEMANDDAY
	Dim strTAXYEARMON
	With frmThis
		If .rdT.checked = True OR .rdA.checked = TRUE Then
			gErrorMsgBox "û���� ������ �̿Ϸ���� ���� ����˴ϴ�.","ó���ȳ�!"
			Exit Sub
		End If
		
		if .txtDEMANDDAY.value = "" then
			gErrorMsgBox "������ û������ �Է��Ͻÿ�.","ó���ȳ�!"
			Exit Sub
		end if
		
		strDEMANDDAY = .txtDEMANDDAY.value
		strTAXYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		intRtn = gYesNoMsgbox("���õ� �׸��� û������ ���� �Ͻðڽ��ϱ�?","���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
			
		For intCnt = 1 To .sprSht.MaxRows
			If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
				mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
			End If
		Next
		
	End With
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtTRANSYEARMON1.value, .txtREAL_MED_CODE1.value, .txtREAL_MED_NAME1.value, "INTERNET") 
		vntRet = gShowModalWindow("../MDCO/MDCMTAXREAL_MEDPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtREAL_MED_CODE1.value = vntRet(1,0) and .txtREAL_MED_NAME1.value = vntRet(2,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTRANSYEARMON1.value = vntRet(0,0)		  ' Code�� ����
			.txtREAL_MED_CODE1.value = vntRet(1,0)		  ' Code�� ����
			.txtREAL_MED_NAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			if .txtTRANSYEARMON1.value <> "" then
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			End if
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetTAXREAL_MED_NO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON1.value,.txtREAL_MED_CODE1.value,.txtREAL_MED_NAME1.value, "INTERNET")
			if not gDoErrorRtn ("GetTAXREAL_MED_NO") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE1.value = vntData(1,1)
					.txtREAL_MED_NAME1.value = vntData(2,1)
					if .txtTRANSYEARMON1.value <> "" then
						gFlowWait meWAIT_ON
						SelectRtn
						gFlowWait meWAIT_OFF
					End if
				Else
					Call REAL_MED_CODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' �޷�
'-----------------------------------------------------------------------------------------
Sub imgDEMANDDAY_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgDEMANDDAY,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

Sub txtTRANSYEARMON1_onblur
	With frmThis
	If .txtTRANSYEARMON1.value <> "" AND Len(.txtTRANSYEARMON1.value) = 6 Then DateClean
	End With
End Sub

Sub imgFrom_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgTO_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTO_onchange
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), mobjSCGLSpr.CnvtDataField(.sprSht,"CHK"), , , "", , , , , mstrCheck
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

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else	
			If .rdT.checked = True Then
			
				If mobjSCGLSpr.GetTextBinding(.sprSht,"MERGEFLAG",Row) = "1" Then
					gErrorMsgBox "����û����꼭�� �����Ȱ� �Դϴ�.",""
					EXIT SUB
				End If 
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row)) '<< �޾ƿ��°��
				gShowModalWindow "../MDIT/MDCMINTERNETCOMMITAXDTL.aspx",vntInParams , 898,680
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
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
		If mstrGUBUN = "TAX" Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
					strCOLUMN = "SUMAMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_AMT") Then
					strCOLUMN = "REAL_AMT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_AMT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) Then
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
		else
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
					strCOLUMN = "SUMAMT"
				ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
					strCOLUMN = "COMMISSION"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION")) Then
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
		end if
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
		If mstrGUBUN = "TAX" Then
			If .sprSht.MaxRows >0 Then
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"REAL_AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or _
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
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
		ELSE
			If .sprSht.MaxRows >0 Then
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMISSION") Then
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
		END IF
		
	End With
End Sub

Sub cmbGUBUN_onchange
	with frmThis
		If .cmbGUBUN.value = "taxdiv" Then
			pnlFLAG.style.display = "none"
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			pnlFLAG.style.display = "inline"
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
		End If
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
	set mobjMDITINTERNETCOMMITAX= gCreateRemoteObject("cMDIT.ccMDITINTERNETCOMMITAX")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDITINTERNETCOMMITAX = Nothing
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
		.txtTRANSYEARMON1.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		DateClean
		.sprSht.MaxRows = 0
		
		CALL Grid_Setting ("TRANS")
		
		.txtREAL_MED_NAME1.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = MID(frmThis.txtTRANSYEARMON1.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
		.txtDEMANDDAY.value = date2
	End With
End Sub

Sub Grid_Setting (strGUBUN)
	With frmThis
		'Sheet �⺻Color ����
		.sprSht.MaxRows = 0
		.sprSht.style.visibility = "hidden"
		Call Grid_init()
		gSetSheetDefaultColor() 
		If strGUBUN = "TAX" Then
			'Sheet �⺻Color ����
			 gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 23, 0, 1, 2
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXYEARMON | TAXNO | TAXMANAGE | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | REAL_AMT | COMMI_RATE | AMT | VAT | SUMAMT | DEMANDDAY | SUMM | PRINTDAY | DEPT_NAME | REAL_MEDOWNER | REAL_MEDADDR1 | REAL_MEDADDR2 | VOCHNO | MERGEFLAG | SUPPNM | SUPPEMAIL"
			mobjSCGLSpr.SetHeader .sprSht,		 "����|���ݰ�꼭���|���ݰ�꼭��ȣ|������ȣ|��ü��|��ü�����ڹ�ȣ|��ü��|��ް�|��������|������|�ΰ���|�հ�ݾ�|û����|����|������|���μ�|��ü���ǥ��|�ּ�1|�ּ�2|��ǥ��ȣ|����û������|��ü������|��ü�������̸���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|             0|             0|      10|    15|	           13|    12|    12|	   7|    13|    10|      13|     9|	 20|     9| 	  9|           0|    0|    0|      10|           0|           0|                 0 "
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | PRINTDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "REAL_AMT | AMT | VAT | SUMAMT ", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE ", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXYEARMON | TAXNO | TAXMANAGE | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | SUMM | DEPT_NAME | VOCHNO | SUPPNM | SUPPEMAIL", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXYEARMON | TAXNO | TAXMANAGE | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | REAL_AMT | COMMI_RATE | AMT | VAT | SUMAMT | DEMANDDAY | SUMM | PRINTDAY | DEPT_NAME | VOCHNO"
			mobjSCGLSpr.ColHidden .sprSht, "REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | MERGEFLAG | SUPPNM | SUPPEMAIL", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE | REAL_MED_BISNO | VOCHNO",-1,-1,2,2,False
			mstrGUBUN = "TAX"
		Else
			'Sheet �⺻Color ����
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 39, 0, 1, 2
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | TIMNAME | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1 | CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1 | REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | MEDFLAG | RANKTRANS"
			mobjSCGLSpr.SetHeader .sprSht,		 "����|�ŷ���ȣ|û�����|������|�����ֻ���ڹ�ȣ|��|��ü��|��ü�����ڹ�ȣ|��ü��|�ݾ�|��������|������|�ΰ���|�հ�ݾ�|����|������|���μ�|�����ִ�ǥ�ڸ�|�������ּ�1|�������ּ�2|��ü���ǥ�ڸ�|��ü���ּ�1|��ü���ּ�2|��ǥ��ȣ|���ݰ�꼭���|���ݰ�꼭��ȣ|�ŷ��������|�ŷ�������ȣ|�ŷ���������|��Ź���|��Ź����|�������ڵ�|������AC�ڵ�|��ü���ڵ�|��ü��AC�ڵ�|��ü�ڵ�|�μ��ڵ�|��ü����|����"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|      10|       8|    15|	          13|15|    15|              13|     9|  10|       5|	  10|    10|      11|  20|     9|       8|             0|          0|          0|             0|          0|          0|       0|             0|             0|             0|             0|             0|       0|       0|         0|           0|         0|           0|       0|       0|       0|  10"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY | PRINTDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT | SUMAMT | COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXMANAGE | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | SUMM | DEPT_NAME | CLIENTOWNER | CLIENTADDR1 | CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1 | REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | MEDFLAG | RANKTRANS | TIMNAME", -1, -1, 100 
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1 | CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1 | REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | MEDFLAG | RANKTRANS | TIMNAME"
			mobjSCGLSpr.ColHidden .sprSht, "CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1| REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | MEDFLAG ", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE | CLIENTBISNO | REAL_MED_BISNO | VOCHNO",-1,-1,2,2,False
			mstrGUBUN = "TRANS"
		End If
		
		.sprSht.style.visibility = "visible"
		
	End With
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, ""
		mobjSCGLSpr.SetHeader .sprSht,		 ""
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
	End With
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'��üüũ
Sub rdA_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdMED_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub rdREAL_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
   	Dim i, strCols
   	Dim strGUBUN
   	Dim strGROUPGBN
   	Dim strMED_FLAG
   	Dim strVOCH_TYPE_TEMP
   
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		If .txtTRANSYEARMON1.value = "" Then
			gErrorMsgBox "����� �Է��Ͻʽÿ�","��ȸ�ȳ�!"
			Exit Sub
		End If	
		If Len(.txtTRANSYEARMON1.value) <> 6 Then
			gErrorMsgBox "����� ������ �ƴմϴ�.","��ȸ�ȳ�!"
			Exit Sub
		End If
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF .rdT.checked = TRUE THEN
			CALL Grid_Setting ("TAX")
		ELSE
			CALL Grid_Setting ("TRANS")
		END IF
		
		strYEARMON		= .txtTRANSYEARMON1.value
		strREAL_MED_CODE= .txtREAL_MED_CODE1.value
		strFROM			=  MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO			=  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		strGUBUN = .cmbGUBUN.value
		
		IF strGUBUN = "taxgroup" then
			IF .rdMED.checked = TRUE THEN
				strGROUPGBN = "MED"
			ELSEIF .rdREAL.checked = TRUE THEN
				strGROUPGBN = "REAL"
			END IF 
		else
			strGROUPGBN = ""
		end if
		
		'���ݰ�꼭 �Ϸ���ȸ
		If .rdT.checked = True Then
			vntData = mobjMDITINTERNETCOMMITAX.Get_INTERNET_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strREAL_MED_CODE, strFROM, strTO)
			If not gDoErrorRtn ("Get_PRINT_TAX") then
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				Layout_change
				AMT_SUM
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				if .sprSht.MaxRows = 0 then
					.imgDelete.style.display = "none"
				else
					.imgDelete.style.display = "inline"
				end if
			End If
		'�̿Ϸ� �ŷ����� ������ ��ȸ
		ElseIf .rdF.checked = True Then			
			vntData = mobjMDITINTERNETCOMMITAX.Get_INTERNET_TAXBUILD(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strREAL_MED_CODE, strFROM, strTO, strGUBUN, strGROUPGBN)
			If not gDoErrorRtn ("Get_PRINT_TAXBUILD") then
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				'Layout_change
				AMT_SUM
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				if .sprSht.MaxRows = 0 then
					.ImgTaxCre.style.display = "none"
				else
					.ImgTaxCre.style.display = "inline"
				end if
			End If
		ElseIf .rdA.checked = True Then
			vntData = mobjMDITINTERNETCOMMITAX.Get_INTERNET_TAXALL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strREAL_MED_CODE, strFROM, strTO, strGUBUN, strGROUPGBN)
			If not gDoErrorRtn ("Get_PRINT_TAXALL") then
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'Layout_change
				AMT_SUM
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
				.ImgTaxCre.style.display = "none"
				.imgDelete.style.display = "none"
			End If
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
		For intCnt = 1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MERGEFLAG",intCnt) = "1" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HAAAAAA, &H000000,False
				mobjSCGLSpr.SetCellsLock2 .sprSht,TRUE,"CHK",-1,-1,false
			ELSE
				mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,"CHK",-1,-1,false
			End If 
		Next 
	End With
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		If .sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		ELSE
			IF mstrGUBUN = "TAX" THEN
				For lngCnt = 1 To .sprSht.MaxRows
					IntAMT = 0
					IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
					IntAMTSUM = IntAMTSUM + IntAMT
				Next
			ELSE
				For lngCnt = 1 To .sprSht.MaxRows
					IntAMT = 0
					IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION", lngCnt)
					IntAMTSUM = IntAMTSUM + IntAMT
				Next
			END IF
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
    Dim intRtn2
   	Dim vntData, vntData1
	Dim strMasterData
	Dim strTAXYEARMON
	Dim intTAXNO
	Dim strTAXSET
	Dim strSUMM
	Dim intCnt
	Dim strDEMANDDAY,strPRINTDAY
	Dim chkcnt
	Dim intCnt2
	Dim intColFlag
	Dim intMaxCnt
	Dim bsdiv
	Dim strVALIDATION
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
   		
		intRtn2 = gYesNoMsgbox("û������ Ȯ���ϼ̽��ϱ�?","Ȯ��")
		IF intRtn2 <> vbYes then exit Sub
		
		'üũ ���� ��� ���� �ȵǵ���
		chkcnt = 0
		For intCnt = 1 To .sprSht.MaxRows
			strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
			strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
			If strDEMANDDAY  = "" Then
				gErrorMsgBox "û������ �ʼ� �Դϴ�.","����ȳ�!"
				Exit Sub
			End If
			If  strPRINTDAY = "" Then
				gErrorMsgBox "û������ �ʼ� �Դϴ�.","����ȳ�!"
				Exit Sub
			End If
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		Next
	
		if chkcnt = 0 then
			gErrorMsgBox "���ݰ�꼭�� ������ �����͸� üũ �Ͻʽÿ�","����ȳ�!"
			exit sub
		end if
		'�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
   		
		'if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1 | CLIENTADDR2 | REAL_MEDOWNER | REAL_MEDADDR1 | REAL_MEDADDR2 | VOCHNO | TAXYEARMON | TAXNO | TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON | TRUST_SEQ | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPT_CD | MEDFLAG | RANKTRANS")
		
		'������ �����͸� ���� �´�.
		'ó�� ������ü ȣ��
		intTAXNO = 0
		If .cmbGUBUN.value = "taxdiv" Then
		intRtn = mobjMDITINTERNETCOMMITAX.ProcessRtn_Div(gstrConfigXml,vntData, intTAXNO)
		Else
			If Not TaxGroup(strVALIDATION) Then 
				gErrorMsgBox strVALIDATION & vbCrlf & "������ [������,��ü��] [û����,�ۼ���,����,����] �� ���� �Ͽ��� �մϴ�.","����ȳ�!"
				Exit Sub
			Else
				'�ִ밪
				intColFlag = 0
				For intMaxCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intMaxCnt) = 1 Then
						bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intMaxCnt))
						IF intColFlag < bsdiv THEN
							intColFlag = bsdiv
						END IF
					End IF
				Next
				'�ƽ����� �߰��Ͽ� ������
				intRtn = mobjMDITINTERNETCOMMITAX.ProcessRtn_Group(gstrConfigXml,vntData, intTAXNO,intColFlag)
			End IF
		End If

		If not gDoErrorRtn ("ProcessRtn") Then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "���ݰ�꼭�� �����Ǿ����ϴ�.","����ȳ�!"
			.rdT.checked = True
			gFlowWait meWAIT_ON
			SelectRtn
			gFlowWait meWAIT_OFF
   		End If
   	end with
End Sub

Function TaxGroup(ByRef strVALIDATION)
	Dim intCnt
	Dim strREAL_MED_CODE 'û���� ����� ��Ϲ�ȣ
	Dim strDEMANDDAY
	Dim strPRINTDAY
	Dim strSUMM
	Dim strStartRank
	
	TaxGroup = False
	with frmThis
		strStartRank = "0"
		strREAL_MED_CODE = ""
		strDEMANDDAY = ""
		strPRINTDAY = ""
		strSUMM = ""
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				
				If strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Then
					If strREAL_MED_CODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt) Then
						Exit Function
					End If 
					If strDEMANDDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt) Then
						strVALIDATION = "û����Ȯ�� �ŷ�������ȣ " & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If 
					If strPRINTDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt) Then
						strVALIDATION = "������Ȯ�� �ŷ�������ȣ" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If 
					If strSUMM <> mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt) Then
						strVALIDATION = "����Ȯ�� �ŷ�������ȣ" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If
				End If
				
				strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt)
				strREAL_MED_CODE = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt)
				strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
				strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
				strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt)
			End If
		Next
	End With
	TaxGroup = True
End Function

Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strDESCRIPTION
	with frmThis
		strDESCRIPTION = ""
		For intCnt2 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"MERGEFLAG",intCnt2) = "1" THEN
					if NOT VOCHNO_CHECKED_MERGE (mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2), mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2)) then
						gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "����û�� ���ݰ�꼭�� ����� ���� ������ ���� �ʽ��ϴ�.","�����ȳ�!"
						Exit Sub
					END IF
				ELSE
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt2) <> "" THEN
						gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "��ǥ�� �����ϴ� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
						Exit Sub
					ELSE
						if NOT VOCHNO_CHECKED (mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2), mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2)) then
							gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "��ǥó�� �������� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
							Exit Sub
						END IF
					End If
				END IF
			END IF
		Next
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
			
				intRtn = mobjMDITINTERNETCOMMITAX.DeleteRtn_CommiTax(gstrConfigXml,strTAXYEARMON, strTAXNO)
				IF not gDoErrorRtn ("DeleteRtn_TruTax") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"�����ȳ�!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn_CommiTax") then
			gWriteText lblstatus, intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	End with
	err.clear	
End Sub

Function VOCHNO_CHECKED (ByRef strTAXYEARMON, ByRef strTAXNO)
	Dim vntData
	Dim intCnt
	Dim strCOUNT
	'on error resume next

	'�ʱ�ȭ
	VOCHNO_CHECKED = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCOGET.COMMIVOCHNO_CHECKED(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON,strTAXNO)
	
	IF mlngRowCnt >0 THEN
		VOCHNO_CHECKED = false
	ELSE
		VOCHNO_CHECKED = TRUE	
	End IF
End Function

Function VOCHNO_CHECKED_MERGE (ByRef strTAXYEARMON, ByRef strTAXNO)
	Dim vntData
	Dim intCnt
	Dim strCOUNT
	'on error resume next

	'�ʱ�ȭ
	VOCHNO_CHECKED_MERGE = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCOGET.COMMIVOCHNO_CHECKED_MERGE(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON,strTAXNO, "O")
	
	IF mlngRowCnt >0 THEN
		VOCHNO_CHECKED_MERGE = false
	ELSE
		VOCHNO_CHECKED_MERGE = TRUE	
	End IF
End Function

-->
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
													<TABLE cellSpacing="0" cellPadding="0" width="142" background="../../../images/back_p.gIF"
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
											<td class="TITLE">������ ���ݰ�꼭 ����</td>
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
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, '')"
												width="50">���</TD>
											<TD class="SEARCHDATA" width="100"><INPUT class="INPUT" id="txtTRANSYEARMON1" title="�ŷ������" style="WIDTH: 98px; HEIGHT: 22px"
													accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
												width="50">��ü��
											</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="�ڵ��" style="WIDTH: 223px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="30" name="txtREAL_MED_NAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE1"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtREAL_MED_CODE1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
												width="50">û����
											</TD>
											<TD class="DATA"><INPUT class="INPUT" id="txtFROM" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="2" name="txtFROM">&nbsp;<IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													 align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="6" name="txtTO">&nbsp;<IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"
													 align="absMiddle" border="0" name="imgTo">
											</TD>
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
											<TD class="SEARCHDATA" colspan='6'><INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN">
												&nbsp;�Ϸ�&nbsp; &nbsp; <INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" CHECKED value="rdF" name="rdGBN">
												&nbsp;�̿Ϸ�&nbsp;&nbsp;<INPUT id="rdA" title="��ü ������ȸ" type="radio" value="rdA" name="rdGBN">&nbsp;��ü</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">�������հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="ImgTaxCre" onmouseover="JavaScript:this.src='../../../images/ImgTaxCreOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTaxCre.gif'"
																height="20" alt="���õǾ��� ��Ŀ� ���� ���ݰ�꼭�� �ۼ��մϴ�." src="../../../images/ImgTaxCre.gif"
																align="absMiddle" border="0" name="ImgTaxCre"></td>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
																name="imgDelete"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
																name="imgPrint"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 67px">����/�ջ�</TD>
											<TD class="DATA" style="WIDTH: 300px"><SELECT id="cmbGUBUN" title="��ü����" style="WIDTH: 96px" name="cmbGUBUN">
													<OPTION value="taxdiv" selected>���ҹ���</OPTION>
													<OPTION value="taxgroup">�ջ����</OPTION>
												</SELECT>
												<DIV id="pnlFLAG" style="DISPLAY: none; WIDTH: 170px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout">&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="rdMED" title="��ü���ջ�" type="radio" CHECKED value="MED" name="rdGROUP">&nbsp;��ü��&nbsp;&nbsp;&nbsp; 
													&nbsp; <INPUT id="rdREAL" title="��ü���ջ�" type="radio" value="REAL" name="rdGROUP">&nbsp;��ü��</DIV>
											</TD>
											<TD class="LABEL" style="WIDTH: 73px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY, '')">û���� 
												����</TD>
											<TD class="DATA" style="WIDTH: 294px"><INPUT class="INPUT" id="txtDEMANDDAY" title="û������" style="WIDTH: 112px; HEIGHT: 22px"
													accessKey="date" type="text" maxLength="10" size="13" name="txtDEMANDDAY">&nbsp;<IMG id="imgDEMANDDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF"  align="absMiddle"
													border="0" name="imgDEMANDDAY">&nbsp;<IMG id="btnCOMMISSION" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" height="20" alt="�ش� �����Ϸ� �������� ���� ���׸��� Setting �մϴ�"
													src="../../../images/imgApp.gIF" width="54" align="absMiddle" border="0" name="btnCOMMISSION">
											</TD>
											<TD class="LABEL" style="WIDTH: 83px">û��/��������</TD>
											<td class="DATA">
												<SELECT id="cmbFLAG" title="����/û������" style="WIDTH: 120px" name="cmbFLAG">
													<OPTION value="receipt" selected>û��</OPTION>
													<OPTION value="demand">����</OPTION>
												</SELECT>
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31882">
										<PARAM NAME="_ExtentY" VALUE="13838">
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
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></form>
	</body>
</HTML>
