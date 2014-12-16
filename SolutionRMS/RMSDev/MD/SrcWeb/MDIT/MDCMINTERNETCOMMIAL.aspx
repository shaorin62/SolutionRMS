<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNETCOMMIAL.aspx.vb" Inherits="MD.MDCMINTERNETCOMMIAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ�����</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ����� ��� ȭ��(MDCMPRINTTRANS1.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ����Ź�ŷ����� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/28 By Kim Tae Yub
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
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDITINTERNETCOMMI, mobjMDCOGET
Dim mstrCheck, mstrCheck1
Dim mstrGrid
CONST meTAB = 9
mstrCheck=True
mstrCheck1=True
mstrGrid = FALSE
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
	IF frmThis.txtYEARMON1.value = "" and frmThis.txtREAL_MED_CODE1.value = "" then
		gErrorMsgBox "��ȸ������ �Է��Ͻÿ�.","��ȸ�ȳ�"
		Exit Sub
	end if
	mstrGrid = FALSE
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�ʱ�ȭ��ư
Sub imgCho_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
Sub ImgSave_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgALLPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData, vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMINTERNETCOMMI_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = ""
		vntDataTemp = mobjMDITINTERNETCOMMI.ProcessRtn_TEMP_ALL(gstrConfigXml, vntData, strUSERID)

		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		'window.setTimeout "call printSetTimeout_All()", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout_All()
	Dim intRtn
	with frmThis
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
	end with
end sub


'���� ���
Sub imgSelectPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim vntData, vntDataTemp
	Dim strcnt
	Dim intRtn
	Dim strUSERID
	
	strcnt = 0
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	for j = 1 to frmthis.sprSht_HDR.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmthis.sprSht_HDR,"CHK",j) = "1" Then
			strcnt = strcnt + 1
		end if
	next 
	
	if strcnt = 0 then
		gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�μ�ȳ�"
		EXIT SUB
	end if 
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMINTERNETCOMMI_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = "1" Then
				mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
			END IF 
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = ""
		vntDataTemp = mobjMDITINTERNETCOMMI.ProcessRtn_TEMP_Select(gstrConfigXml, vntData, strUSERID)

		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout_Select()", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout_Select()
	Dim intRtn
	with frmThis
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
	end with
end sub


Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strCNT
	Dim vntData
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht_HDR.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if

	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMINTERNETCOMMI_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjMDITINTERNETCOMMI.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout('" & strTRANSYEARMON & "', '" & strTRANSNO & "')", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout(strTRANSYEARMON, strTRANSNO)
	Dim intRtn, intRtn2
	with frmThis
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn_temp(gstrConfigXml)
		'intRtn2 = mobjMDITINTERNETCOMMI.DeleteRtnUpdate_PRINTSEQ(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
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
	
	strDATE = MID(frmThis.txtYEARMON1.value,1,4) & "-" & MID(frmThis.txtYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgREAL_MED_CODE1_onclick
	Call REAL_MED_CODE1_POP ()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(.txtYEARMON1.value, .txtREAL_MED_CODE1.value, .txtREAL_MED_NAME1.value, "commi", "INTERNET") 
		vntRet = gShowModalWindow("../MDCO/MDCMCOMMIREALMEDPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtREAL_MED_CODE1.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			IF vntRet(3,0) = "�Ϸ�" THEN
				.txtYEARMON1.value = vntRet(0,0)
				.txtREAL_MED_CODE1.value = vntRet(4,0)		  ' Code�� ����
				.txtREAL_MED_NAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			ELSE
				.txtYEARMON1.value = vntRet(0,0)
				.txtREAL_MED_CODE1.value = vntRet(1,0)		  ' Code�� ����
				.txtREAL_MED_NAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			END IF
			selectRtn
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtREAL_MED_CODE1.value,.txtREAL_MED_NAME1.value,"","commi", "INTERNET")
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtREAL_MED_CODE1.value = vntData(1,1)
					.txtREAL_MED_NAME1.value = vntData(2,1)
					selectRtn
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
' û������ ����
'-----------------------------------------------------------------------------------------
Sub txtYEARMON1_onblur
	With frmThis
		If .txtYEARMON1.value <> "" AND Len(.txtYEARMON1.value) = 6 Then DateClean
	End With
End Sub
'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------
Sub imgCalDemandday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û�����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
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
			mstrGrid = TRUE
			CALL Grid_Setting ()
			SelectRtn_DTL Col, Row
			'mstrGrid = false
		end if
	end with
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		IF mstrGrid = FALSE THEN
			if Row = 0 and Col = 1 then
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_DTL, 1, 1, , , "", , , , , mstrCheck1
				if mstrCheck1 = True then 
					mstrCheck1 = False
				elseif mstrCheck1 = False then 
					mstrCheck1 = True
				end if
				for intcnt = 1 to .sprSht_DTL.MaxRows
					sprSht_DTL_Change 1, intcnt
				next
			end if
		END IF
	end with
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
		mstrGrid = true
		Grid_Setting
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
			.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") Then
				strCOLUMN = "VAT"
			ELSEIF .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
				strCOLUMN = "SUMAMTVAT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_HDR,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT")) OR _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT")) OR _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT")) Then
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_HDR,vntData_col(i),vntData_row(j))
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
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") OR _
				.sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"SUMAMTVAT") Then
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

sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		end if
		
		if Row > 0 and Col >1 then
			if mstrGrid then
				'���ݰ�꼭�� ������� �󼼳������� �ΰ��� ������ �Ұ��� �ϴ�.
				IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TAXNO", Row) <> "" THEN 
					gErrorMsgBox "���ݰ�꼭�� ������ �������� ��������� Ȯ���Ͻ� �� �����ϴ�."," �󼼳��� Ȯ�� �ȳ�!!"		
				ELSE
					vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSNO", Row)) '<< �޾ƿ��°��
					gShowModalWindow "MDCMINTERNETCOMMIDTL.aspx",vntInParams , 898,680
					
					imgQuery_onclick
				END IF 
			else
				gErrorMsgBox "�ŷ������� ���� �ϼž� �ŷ������� �� ������ Ȯ�� �Ͻ� �� �ֽ��ϴ�.!"," �󼼳��� Ȯ�� �ȳ�!!"		
			end if 
		end if 
	end with
end sub

Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub

Sub sprSht_DTL_Keyup(KeyCode, Shift)
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
		If mstrGrid Then
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
				.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
					strCOLUMN = "SUMAMTVAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT")) Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
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
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
				.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
					strCOLUMN = "SUMAMTVAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") Then
					strCOLUMN = "COMMISSION"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION")) Then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht_DTL,vntData_col(i),vntData_row(j))
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") Then
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"SUMAMTVAT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"COMMISSION") Then
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

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim i
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strSEQ
	Dim strPRINT_SEQ
	Dim intRtn
	
	with frmThis
		If mstrGrid Then
			if Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRINT_SEQ") then
				for i=1 to .sprSht_DTL.MaxRows
					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) <> "" then
						if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) = 0 then
							mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
						else
							
							if Row <> i then
								if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row) = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) then
									gErrorMsgBox "��¼����� �ߺ��ԷµǾ����ϴ�.",""
									mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
									.txtCLIENTNAME1.focus() 
									.sprSht_DTL.focus()
									EXIT SUB
								end if
							end if
						end if
					end if
				next
				
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSYEARMON",Row)
				strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSNO",Row)
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",Row)
				strPRINT_SEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row)
				
				intRtn = mobjMDITINTERNETCOMMI.UPDATE_PRINTSEQ(gstrConfigXml,strTRANSYEARMON,strTRANSNO, strSEQ, strPRINT_SEQ)
			end if
		END IF
	end with
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	'����������ü ����	
	set mobjMDITINTERNETCOMMI	= gCreateRemoteObject("cMDIT.ccMDITINTERNETCOMMI")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'�ŷ����� ��� �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 14, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CONFIRMFLAG | REAL_MED_NAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | MEMO | CNT | MED_FLAG"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "����|��꼭|��ü��|��ü����|�������Ѿ�|�ΰ�����|�հ�ݾ�|û����|������|�ŷ����|��ȣ|���|�����|��ü�����ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "  4|    6|     15|       8|        12|      11|      12|     9|     9|       8|   5|  14|      10|          0"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY | PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TRANSNO | AMT | VAT | SUMAMTVAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMFLAG | REAL_MED_NAME | MED_FLAGNAME | TRANSYEARMON | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMFLAG | REAL_MED_NAME | MED_FLAGNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | MEMO | MED_FLAG"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "MED_FLAG", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "CONFIRMFLAG | TRANSYEARMON | MED_FLAGNAME" ,-1,-1,2,2,false

		.sprSht_HDR.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDITINTERNETCOMMI = Nothing
	set mobjMDCOGET = Nothing
	gEndPage
End Sub

Sub Grid_Setting ()
	With frmThis
		mobjSCGLCtl.DoEventQueue
		If mstrGrid Then
			'Sheet �⺻Color ����
			gSetSheetDefaultColor() 
			'******************************************************************
			''�ŷ����� ������
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 20, 0, 0, 2
			mobjSCGLSpr.SpreadDataField .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | TRUST_YEARMON"
			mobjSCGLSpr.SetHeader .sprSht_DTL,		"�ŷ������|�ŷ�����ȣ|����|��Ź����|������|��ü��|��ü��|��޾�|��������|������|�ΰ���|��|��ü����|���|�μ���|û����|������|��꼭���|��꼭��ȣ|��Ź���" 
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "        0|	         0|	  0|       4|	 15|    13|    15|    10|       4|    10|    10|11|       4|  10|    10|     9|     9|         9|         5|      0"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "DEMANDDAY | PRINTDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "TRUST_SEQ | AMT | COMMISSION | VAT | SUMAMTVAT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | MED_FLAGNAME | MEMO | DEPT_NAME | TAXYEARMON | TAXNO | TRUST_YEARMON ", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, false, "TRANSYEARMON | TRANSNO | SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | TRUST_YEARMON" 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "TRANSYEARMON | TRANSNO | SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | TRUST_YEARMON" 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | TRUST_SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MED_FLAGNAME | MEMO | DEPT_NAME | DEMANDDAY | PRINTDAY | TAXYEARMON | TAXNO | TRUST_YEARMON", FALSE
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | TRUST_YEARMON  ", TRUE
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "TRUST_SEQ | MED_FLAGNAME",-1,-1,2,2,false
		Else
			'Sheet �⺻Color ����
			gSetSheetDefaultColor() 
			'******************************************************************
			'û�೻�� �׸���
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 26, 0, 1, 0,0
			mobjSCGLSpr.SpreadDataField .sprSht_DTL,   "CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK "
			mobjSCGLSpr.SetHeader .sprSht_DTL,		   "����|���|����|��ü�����ڵ�|����|��ü��|��ü�����ڹ�ȣ|������|�����ֻ���ڹ�ȣ|��ü��|�귣���|�����|��޾�|��������|������|�ΰ���|��|���|��ü���ڵ�|�������ڵ�|��ü�ڵ�|�μ��ڵ�|���۴�����ڵ�|��ǥ����|�ΰ�������|��������"
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "   4|	0|   4|           0|   4|    15|              15|    15|              15|    12|      10|    10|    11|       4|    11|    10|12|  13|         0|         0|       0|       0|             0|       0|         0|      0"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "SEQ | AMT | VAT | SUMAMTVAT | COMMISSION", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "COMMI_RATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "YEARMON | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK ", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, false, "CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK " 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | AMT | COMMI_RATE | COMMISSION | SUMAMTVAT | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK  "
			mobjSCGLSpr.ColHidden .sprSht_DTL, "CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK ", FALSE
			mobjSCGLSpr.ColHidden .sprSht_DTL, "YEARMON | MED_FLAG | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK", true
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "MED_FLAGNAME",-1,-1,2,2,false
		End If
		
		.sprSht_DTL.style.visibility = "visible"
	End With
End Sub
'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		DateClean

		.txtPRINTDAY.value  = gNowDate
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		
		CALL Grid_Setting ()
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strREAL_MED_CODE, strREAL_MED_NAME
	chkcnt = 0
	
	with frmThis
		if mstrGrid then exit sub
		
		intColFlag = 0
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'�׷��ִ밪 ����
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if

		'�����÷��� ����

		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | YEARMON | SEQ | MED_FLAG | MED_FLAGNAME | REAL_MED_NAME | REAL_MED_BISNO | CLIENTNAME | CLIENTBISNO | MEDNAME | SUBSEQNAME | MATTERNAME | AMT | COMMI_RATE | COMMISSION | VAT | SUMAMTVAT | MEMO | REAL_MED_CODE | CLIENTCODE | MEDCODE | DEPT_CD | EXCLIENTCODE | VOCH_TYPE | COMMI_TAX_FLAG | TRANSRANK ")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strREAL_MED_CODE= .txtREAL_MED_CODE1.value
		strREAL_MED_NAME= .txtREAL_MED_NAME1.value
		
		intRtn = mobjMDITINTERNETCOMMI.ProcessRtn(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON,intColFlag)
   		
   		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtREAL_MED_CODE1.value = strREAL_MED_CODE
				.txtREAL_MED_NAME1.value = strREAL_MED_NAME
				selectRtn
			Else
				initpagedata
			End If
			DateClean
   		end if

   	end with
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	'On error resume next
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "�������� �ʼ� �Է� ���� �Դϴ�.",""
			Exit Function
		End If
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON, strDEMANDYEARMON
	Dim strREAL_MED_CODE
   	Dim strMED_FLAG
   	Dim i, strCols
    
	'On error resume next
	with frmThis
	
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value
		strREAL_MED_CODE= .txtREAL_MED_CODE1.value
		
		CALL Grid_Setting()
		
		vntData = mobjMDITINTERNETCOMMI.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strREAL_MED_CODE)
													
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   		
   		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
   		vntData2 = mobjMDITINTERNETCOMMI.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												strYEARMON, strREAL_MED_CODE)
		
		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData2,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   		End If
   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)
				
		vntData = mobjMDITINTERNETCOMMI.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strTRANSYEARMON, strTRANSNO)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatusDTR, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			'mstrGrid = TRUE
   		End If
   	end with
End Sub
'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To .sprSht_DTL.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"COMMISSION", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub
'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	Dim strPRINTDAY
   	Dim strMED_FLAG
   	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
   	Dim strCLIENTCODE, strCLIENTNAME
   	Dim lngchkCnt
   	
	with frmThis
		strDESCRIPTION = ""

		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "������ ������ �����ϴ�.","�����ȳ�!"
			Exit Sub
		END IF
		
		For i = 1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
				strTRANSYEARMON = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TRANSNO",i)
				
				vntData = mobjMDITINTERNETCOMMI.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO) 
				If mlngRowCnt > 0 Then
					gErrorMsgBox i & "���� �ŷ������� ���ݰ�꼭�� �߻��� �󼼳����� �����մϴ�.","�����ȳ�!"
					Exit Sub
				End If
				lngchkCnt = lngchkCnt + 1
			End If
		Next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
				
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0
		mobjSCGLSpr.SetFlag  .sprSht_HDR, meINS_TRANS
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO ")
		
		intRtn = mobjMDITINTERNETCOMMI.DeleteRtn(gstrConfigXml,vntData)

		IF not gDoErrorRtn ("DeleteRtn") then
			'���õ� �ڷḦ ������ ���� ����
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "�ŷ������� �����Ǿ����ϴ�.","�����ȳ�!"
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

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="260" background="../../../images/back_p.gIF"
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
											<td class="TITLE">û����� - ������ŷ����� ����/��ȸ/����</td>
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
						<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="1"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" height="93%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
										border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1,'')"
												width="60">û�����</TD>
											<TD class="SEARCHDATA" width="130"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_NAME1, txtREAL_MED_CODE1)"
												width="60">��ü��</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="�ڵ��" style="WIDTH: 203px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="27" name="txtREAL_MED_NAME1">
												<IMG id="ImgREAL_MED_CODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"
													src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgREAL_MED_CODE1">
												<INPUT class="INPUT_L" id="txtREAL_MED_CODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtREAL_MED_CODE1"></TD>
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
									</TABLE>
								</TD>
							<tr>
								<td>
									<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
										<TR>
											<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
										</TR>
									</table>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
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
												<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="ȭ���� �ʱ�ȭ �մϴ�."
																src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="���õ� �ŷ������� �����մϴ�." src="../../../images/imgDelete.gIF" border="0"
																name="imgDelete"></TD>
														<TD><IMG id="imgALLPrint" onmouseover="JavaScript:this.src='../../../images/imgALLPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgALLPrint.gif'"
																height="20" alt="��ȸ�� �ŷ������� ��ü�μ��մϴ�.." src="../../../images/imgALLPrint.gIF" border="0"
																name="imgALLPrint"></TD>
														<TD><IMG id="imgSelectPrint" onmouseover="JavaScript:this.src='../../../images/imgSelectPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSelectPrint.gif'"
																height="20" alt="��ȸ�� �ŷ������� �����μ��մϴ�.." src="../../../images/imgSelectPrint.gIF" border="0"
																name="imgSelectPrint"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="4339">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD class="TITLE" align="left" width="400" height="22" vAlign="absmiddle"></TD>
											<TD class="TITLE" vAlign="absmiddle" align="left" width="400" height="22">û������ : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 100px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="32" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">&nbsp;&nbsp;&nbsp;&nbsp; 
												�������� : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 94px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalPrintday"></TD>
											<TD vAlign="middle" align="right" height="22">
												<!--Common Button Start-->
												<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="ImgSave" onmouseover="JavaScript:this.src='../../../images/ImgTransCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTransCre.gIF'"
																height="20" alt="��ü�纰�� �ŷ������� �����մϴ�.." src="../../../images/ImgTransCre.gIF" border="0"
																name="ImgSave"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="���� �ŷ������� ����մϴ�.." src="../../../images/imgPrint.gIF" border="0"
																name="imgPrint"></TD>
														<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213" VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="8625">
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
								<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
