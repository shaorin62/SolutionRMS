<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMTRANS.aspx.vb" Inherits="PD.PDCMTRANS" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ�������</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ������� ��� ȭ��(PDCMTRANS.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMTRANS.aspx
'��      �� : �ŷ������� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/09/30 By Ȳ����
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331"
			VIEWASTEXT>
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMTRANS, mobjSCCMGET
Dim mstrCheck, mstrCheck1
Dim mstrGrid

CONST meTAB = 9
mstrCheck=True
mstrCheck1=True
mstrGrid = FALSE

'=========================================================================================
' �̺�Ʈ ���ν���
'=========================================================================================
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

'-----------------------------------
' ���� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	IF frmThis.txtYEARMON1.value = "" and frmThis.txtCLIENTCODE1.value = "" then
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

'������ư
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
'�����ֺ�������ư
Sub ImgCustSave_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_CUST
	gFlowWait meWAIT_OFF
End Sub

'tim��������ư
Sub ImgTimSave_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_TIM
	gFlowWait meWAIT_OFF
End Sub

'HDR����
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub

'DTL����
Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub


Sub ImgConfirmRequest_onclick()
	Dim intcnt
	Dim lngchkCnt
	Dim vntData
	Dim vntRet
	
	gFlowWait meWAIT_ON
	with frmThis
		IF .sprSht_HDR.MaxRows = 0 THEN
			gErrorMsgBox "û����û�� ������ �����ϴ�.","�����ȳ�!"
			Exit Sub
		END IF
		
		for intcnt =1 to .sprSht_HDR.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intcnt) = 1 Then
				vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | CLIENTCODE | CLIENTNAME | AMT | SUMAMTVAT")
				lngchkCnt = lngchkCnt + 1
			end if
		next
		
		IF lngchkCnt = 0 Then
			gErrorMsgBox "���õ� ������ �����ϴ�.","�����ȳ�!"
			EXIT SUB
		END IF
		
		'����� ����Ÿ�� �����´�
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOMMSPOP.aspx", vntData , 413,435)
		
	end with
	gFlowWait meWAIT_OFF
End Sub




'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout_CONFIRM()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub


'��ü�μ�
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
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺��� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjPDCMTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "PD"
		ReportName = "PDCMTRANS.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = ""
		vntDataTemp = mobjPDCMTRANS.ProcessRtn_TEMP_ALL(gstrConfigXml, vntData, strUSERID)

		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout_All()", 10000
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout_All()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

'�μ�
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
	
	if mstrGrid = FALSE THEN
		gErrorMsgBox "�̻��� ������ �μ��Ҽ� �����ϴ�.",""	
		Exit Sub
	END IF
	
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺��� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjPDCMTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "PD"
		ReportName = "PDCMTRANS.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjPDCMTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
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
		intRtn = mobjPDCMTRANS.DeleteRtn_temp(gstrConfigXml)
		intRtn2 = mobjPDCMTRANS.DeleteRtnUpdate_PRINTSEQ(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

'�ݱ�
Sub imgClose_onclick ()
	Window_OnUnload
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
			SelectRtn
     	end if
	End with
	SelectRtn
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value) , "A")
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
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
' ���ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgTIMCODE1_onclick
	Call TIMCODE1_POP()
End Sub

'���� ������List ��������
Sub TIMCODE1_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array( trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value)) 
							
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,465)
		'TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE
		if isArray(vntRet) then
			.txtTIMCODE1.value = trim(vntRet(0,0))
			.txtTIMNAME1.value = trim(vntRet(1,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			.txtCLIENTNAME1.value = trim(vntRet(5,0))
			SelectRtn
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtTIMNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCMGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												  trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
												  trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value))
			
			if not gDoErrorRtn ("GetTRANSTIMCODE") then
				If mlngRowCnt = 1 Then
					.txtTIMCODE1.value = trim(vntData(0,1))
					.txtTIMNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(5,1))
					SelectRtn
				Else
					Call TIMCODE1_POP()
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

Sub cmbJOBGUBN1_onchange
	SelectRtn
End Sub

Sub txtYEARMON1_onkeydown
	'or window.event.keyCode = meTAB ���϶��� �ƴ� �����϶��� ��ȸ
	If window.event.keyCode = meEnter Then
		SELECTRTN
		frmThis.txtCLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
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
			CALL Grid_Setting ()
			SelectRtn_DTL Col, Row
			
			mstrGrid = TRUE
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
		End IF
	end with
End Sub  


sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		end if
	end with
end sub



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
		Else
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
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") OR _
				.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE") Then
					strCOLUMN = "SUMAMTVAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE")) Then
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
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") OR _
			   .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE")  Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE") Then
					strCOLUMN = "SUMAMTVAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE")) Then
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE") Then
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") OR _
					.sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"BALANCE")  Then
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


Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_HDR, Col, Row  
End Sub

Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	Dim intAMT
	Dim intBALANCE , intOLDBALANCE
	Dim intADJAMT
	Dim intCalCul
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strSEQ
	Dim strPRINT_SEQ
	Dim intRtn, i
	
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
				
				intRtn = mobjPDCMTRANS.UPDATE_PRINTSEQ(gstrConfigXml,strTRANSYEARMON,strTRANSNO, strSEQ, strPRINT_SEQ)
			end if
		else
			if Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"ADJAMT") THEN	
				intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",Row)
   				intBALANCE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"BALANCE",Row)
   				intOLDBALANCE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"BALANCE",Row)
   				intADJAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ADJAMT",Row)
   				intCalCul = intAMT - intADJAMT
   				
   				If intADJAMT > intAMT Then	
   					gErrorMsgBox "����ݾ��� �ܾ� ���� Ŭ�� �����ϴ�.","�Է¾ȳ�!"
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"ADJAMT",Row,0
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"BALANCE",Row, intOLDBALANCE
   					mobjSCGLSpr.ActiveCell .sprSht_DTL,Col,Row
	   				
					.sprSht_DTL.focus 
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht_DTL, 1, 1,Row ,Row , "", , , , , False
					Exit Sub
				else
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"BALANCE",Row, intCalCul
   				end if
   				
   				If intADJAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CHK",Row, "0"
   				End If
   			end if
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
	dim vntInParam
	dim intNo,i
	'����������ü ����	
	
	
	set mobjPDCMTRANS	= gCreateRemoteObject("cPDCO.ccPDCOTRANS")
	set mobjSCCMGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'�ŷ������� ��� �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 16, 0, 4, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CREFLAG | CONFIRMGBN | CONFIRMFLAG | CLIENTCODE | CLIENTNAME|  AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | CONFIRM_USER | MEMO | CNT"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "����|����|����|��꼭|�������ڵ�|������|���ް���|�ΰ�����|�հ�ݾ�|û����|������|�ŷ����|��ȣ|������|���|�����"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "  4|   4|   4|     6|         0|    20|      12|      12|      12|     9|     9|       9|   5|    10|  21|       4"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "DEMANDDAY | PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TRANSNO | AMT | VAT | SUMAMTVAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | TRANSYEARMON | CONFIRM_USER | MEMO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMGBN | CONFIRMFLAG | CLIENTNAME | AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TRANSYEARMON | TRANSNO | CONFIRM_USER | MEMO | CLIENTCODE"
		mobjSCGLSpr.ColHidden .sprSht_HDR, "CLIENTCODE|CNT|CONFIRMGBN|CREFLAG", TRUE '
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "" ,-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "CONFIRMGBN | CONFIRMFLAG | TRANSYEARMON | TRANSNO | CONFIRM_USER | CNT" ,-1,-1,2,2,false

		.sprSht_HDR.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjPDCMTRANS = Nothing
	set mobjSCCMGET = Nothing
	gEndPage
End Sub

Sub Grid_Setting ()
	With frmThis
		mobjSCGLCtl.DoEventQueue
		If mstrGrid Then	
			'Sheet �⺻Color ����
			gSetSheetDefaultColor() 
			'******************************************************************
			''�ŷ������� ������
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 19, 0, 0, 2
			mobjSCGLSpr.SpreadDataField .sprSht_DTL, "TRANSYEARMON | TRANSNO |  PRINT_SEQ | SEQ | CLIENTNAME | TIMNAME | SUBSEQNAME | JOBNAME | JOBNO | JOBSEQ | AMT | ADJAMT | BALANCE | VAT | MEMO | TAXYEARMON | TAXNO|TAXCODE|TAXNAME "
			mobjSCGLSpr.SetHeader .sprSht_DTL,		"�ŷ��������|�ŷ�������ȣ|����|����|������|��|�귣��|JOB��|JOBNO|SEQ|û���ݾ�|�����|����|�ΰ���|���|��꼭���|��꼭��ȣ|�ؽ��ڵ�|���ݰ�꼭����" 
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "        0|	         0|	  4|   4|    20|20|    20|   23|   7|   4|      10|    10|  10|    10|	15|         9|        10|      10|            12"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "TAXYEARMON", -1, -1, 10
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "", -1, -1, 2
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "PRINT_SEQ |  AMT | ADJAMT | BALANCE | VAT ", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | CLIENTNAME | TIMNAME | SUBSEQNAME | MEMO | TAXYEARMON ", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, false, "TRANSYEARMON | TRANSNO | SEQ | PRINT_SEQ | CLIENTNAME | TIMNAME | SUBSEQNAME | AMT | ADJAMT | BALANCE | VAT | MEMO | TAXYEARMON | TAXNO|TAXCODE|TAXNAME " 
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "TRANSYEARMON | TRANSNO | SEQ | CLIENTNAME |  TIMNAME | SUBSEQNAME | AMT | ADJAMT | BALANCE | VAT | MEMO | TAXYEARMON | TAXNO|TAXCODE|TAXNAME " 
			mobjSCGLSpr.ColHidden .sprSht_DTL, "TAXCODE", true
			mobjSCGLSpr.ColHidden .sprSht_DTL, "SEQ | PRINT_SEQ |  CLIENTNAME | TIMNAME | SUBSEQNAME | JOBNAME | JOBNO | JOBSEQ | AMT | ADJAMT | BALANCE | MEMO | TAXYEARMON | TAXNO  |TAXNAME", FALSE
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, " SEQ | JOBNO | TAXYEARMON | TAXNO ",-1,-1,2,2,false
		Else
			
			'Sheet �⺻Color ����
			gSetSheetDefaultColor()		
			'******************************************************************
			'DIVAMT �׸���
			'******************************************************************
			gSetSheetColor mobjSCGLSpr, .sprSht_DTL
			mobjSCGLSpr.SpreadLayout .sprSht_DTL, 26, 0, 0, 2
			mobjSCGLSpr.SpreadDataField .sprSht_DTL,   "CHK | PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | ADJAMT | BALANCE | VAT |  DEMANDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | ENDFLAGNAME | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME | DEMANDFLAG"
			mobjSCGLSpr.SetHeader .sprSht_DTL,		   "����|������ȣ|�������ڵ�|�����ָ�|���ڵ�|����|�귣���ڵ�|�귣���|JOB��|JOBNO|SEQ|û���ݾ�|�����|����|�ΰ���|û������|���|TRANSTIMRANK|TRANSCUSTRANK|ENDFLAGNAME|INCJOBNOSEQ|INCJOBNO|���μ�|�ؽ��ڵ�|���ݰ�꼭����|û�������ڵ�"
			mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "   4|       0|         0|      20|     0|  20|         0|      20|   23|    9|  4|      12|    12|  12|    12|      10|  15|           0|            0|          0|          0|       0|       0|10      |12            |0" 
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "AMT | ADJAMT | BALANCE | VAT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | ENDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | DEMANDFLAG | INCJOBNOSEQ | INCJOBNO | DEPT_CD", -1, -1, 255
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,false,"CHK | PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | ADJAMT | BALANCE | ENDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | DEMANDFLAG | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME"
			mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,"PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | BALANCE | ENDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | ENDFLAGNAME | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME"
			mobjSCGLSpr.ColHidden .sprSht_DTL, "PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | ADJAMT | BALANCE | DEMANDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | ENDFLAGNAME | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME | DEMANDFLAG", true
			mobjSCGLSpr.ColHidden .sprSht_DTL, "CLIENTNAME | TIMNAME | SUBSEQNAME | JOBNO | JOBNAME | SEQ | AMT | ADJAMT | BALANCE | MEMO | TAXNAME | DEMANDFLAGNAME",False
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CLIENTNAME | TIMNAME | SUBSEQNAME | JOBNAME | DEMANDFLAGNAME | MEMO | TAXCODE",-1,-1,0,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ | CLIENTCODE | TIMCODE | SUBSEQ | JOBNO | ENDFLAGNAME",-1,-1,2,2,false
			
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

		.txtPRINTDAY.value  = gNowDate2
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		
		Call COMBO_TYPE()
		.cmbJOBGUBN1.selectedIndex = -1
		
		CALL Grid_Setting ()
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
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
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	Dim vntJOBGUBN  
	
    With frmThis   

		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntJOBGUBN = mobjPDCMTRANS.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  '���۱���

		if not gDoErrorRtn ("COMBO_TYPE") then 

			'mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "ENDFLAG",,,vntENDFLAG,,65 
			
			mobjSCGLSpr.TypeComboBox = True 
			 gLoadComboBox .cmbJOBGUBN1, vntJOBGUBN, False
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
' �ŷ������� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON, strDEMANDYEARMON
	Dim strCLIENTCODE, strCLIENTNAME
	Dim strTIMCODE, strTIMNAME 
   	Dim strJOBGUBN
   	Dim i, strCols
   	Dim strCLIENTSUBCODE, strCLIENTSUBNAME
   	Dim strVOCH_TYPE
   	Dim test
    
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
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strJOBGUBN		= .cmbJOBGUBN1.value
		
		CALL Grid_Setting()
		

		vntData = mobjPDCMTRANS.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strJOBGUBN)
													
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
		
   		vntData2 = mobjPDCMTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													strYEARMON, strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strJOBGUBN)
		
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
				
		vntData = mobjPDCMTRANS.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, _
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
   			mstrGrid = FALSE
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
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT", lngCnt)
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
' ������ ó��
'****************************************************************************************
Sub ProcessRtn_CUST ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt, intCnt2
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME , strJOBGUBN
	chkcnt = 0
	
	with frmThis
		if mstrGrid then exit sub
		
		if .txtDEMANDDAY.value = "" then
			gErrorMsgBox "û�����ڴ� �ʼ� �Դϴ�..","����ȳ�!"
			Exit Sub
		end if
		
		'�׷� �ִ밪 ����
		intColFlag = 0
		For intCnt = 1 To .sprSht_DTL.MaxRows
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSCUSTRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
		For intCnt2 = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt2) = 1 THEN
				chkcnt = chkcnt + 1
				IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ADJAMT",intCnt2) = ""  Or  mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ADJAMT",intCnt2) = 0 THEN
					gErrorMsgBox "������ �ŷ��������� ����ݾ��� �ݵ�� �Է� �Ͽ��� �մϴ�.","����ȳ�!"
					Exit Sub
				End If
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ��������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if

		'�����÷��� ����

		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
		
   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
													   
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | ADJAMT | BALANCE | VAT |  DEMANDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | ENDFLAGNAME | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME | DEMANDFLAG")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strJOBGUBN		= .cmbJOBGUBN1.value
		
		
		intTRANSNO = 0
		If .cmbGUBUN.value = "taxgroup" Then
			intRtn = mobjPDCMTRANS.ProcessRtn_CUST(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON, intTRANSNO, intColFlag)
  		
   			if not gDoErrorRtn ("ProcessRtn_CUST") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
				InitPageData
				gOkMsgBox "�ŷ��������� �����Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					.txtTIMCODE1.value = strTIMCODE
					.txtTIMNAME1.value = strTIMNAME
					.cmbJOBGUBN1.value = strJOBGUBN
					selectRtn
				Else
					initpagedata
				End If
   			end if
   		Else
   			intRtn = mobjPDCMTRANS.ProcessRtn_Div(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON,intTRANSNO)
   			if not gDoErrorRtn ("ProcessRtn_Div") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
				InitPageData
				gOkMsgBox "�ŷ��������� �����Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					.txtTIMCODE1.value = strTIMCODE
					.txtTIMNAME1.value = strTIMNAME
					.cmbJOBGUBN1.value = strJOBGUBN
					selectRtn
				Else
					initpagedata
				End If
   			end if
   		End IF

   	end with
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn_TIM ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt, intCnt2
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME, strJOBGUBN
	chkcnt = 0
	
	with frmThis
		if mstrGrid then exit sub
		
		if .txtDEMANDDAY.value = "" then
			gErrorMsgBox "û�����ڴ� �ʼ� �Դϴ�..","����ȳ�!"
			Exit Sub
		end if
		
		intColFlag = 0
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'�׷��ִ밪 ����
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSTIMRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		next
		
		For intCnt2 = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt2) = 1 THEN
				chkcnt = chkcnt + 1
				IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ADJAMT",intCnt2) = ""  Or  mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ADJAMT",intCnt2) = 0 THEN
					gErrorMsgBox "������ �ŷ��������� ����ݾ��� �ݵ�� �Է� �Ͽ��� �մϴ�.","����ȳ�!"
					Exit Sub
				End If
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ��������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if

		'�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | PREESTNO | CLIENTCODE | CLIENTNAME | TIMCODE | TIMNAME | SUBSEQ | SUBSEQNAME | JOBNAME | JOBNO | SEQ | AMT | ADJAMT | BALANCE | VAT |  DEMANDFLAGNAME | MEMO | TRANSTIMRANK | TRANSCUSTRANK | ENDFLAGNAME | INCJOBNOSEQ | INCJOBNO | DEPT_CD | TAXCODE | TAXNAME | DEMANDFLAG")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strJOBGUBN		= .cmbJOBGUBN1.value
		
		
		intTRANSNO = 0
		If .cmbGUBUN.value = "taxgroup" Then
			intRtn = mobjPDCMTRANS.ProcessRtn_TIM(gstrConfigXml,strMasterData,vntData,strTRANSYEARMON, intTRANSNO, intColFlag)
  		
   			if not gDoErrorRtn ("ProcessRtn_TIM") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
				InitPageData
				gOkMsgBox "�ŷ��������� �����Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					.txtTIMCODE1.value = strTIMCODE
					.txtTIMNAME1.value = strTIMNAME
					.cmbJOBGUBN1.value = strJOBGUBN
					selectRtn
				Else
					initpagedata
				End If
   			end if
   		Else
   			intRtn = mobjPDCMTRANS.ProcessRtn_Div(gstrConfigXml,strMasterData,vntData, strTRANSYEARMON, intTRANSNO)
   			if not gDoErrorRtn ("ProcessRtn_Div") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
				InitPageData
				gOkMsgBox "�ŷ��������� �����Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtYEARMON1.value = strTRANSYEARMON
					.txtCLIENTCODE1.value = strCLIENTCODE
					.txtCLIENTNAME1.value = strCLIENTNAME
					.txtTIMCODE1.value = strTIMCODE
					.txtTIMNAME1.value = strTIMNAME
					.cmbJOBGUBN1.value = strJOBGUBN
					selectRtn
				Else
					initpagedata
				End If
   			end if
   		End IF
   	end with
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
   	Dim strJOBGUBN
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
				
				vntData = mobjPDCMTRANS.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO) 
				If mlngRowCnt > 0 Then
					gErrorMsgBox i & "���� �ŷ��������� ���ݰ�꼭�� �߻��� �󼼳����� �����մϴ�.","�����ȳ�!"
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
		
		intRtn = mobjPDCMTRANS.DeleteRtn(gstrConfigXml,vntData)

		IF not gDoErrorRtn ("DeleteRtn") then
			'���õ� �ڷḦ ������ ����	
			for i = .sprSht_HDR.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   				End If
			Next
			
			gErrorMsgBox "�ŷ��������� �����Ǿ����ϴ�.","�����ȳ�!"
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
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="96" background="../../../images/back_p.gIF"
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
												<td class="TITLE" valign="top">���� �ŷ�������</td>
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
												<TD class="SEARCHDATA" width="100"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">������</TD>
												<TD class="SEARCHDATA" width="250"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 173px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="27" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
													<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1,txtTIMCODE1)"
													width="60">��</TD>
												<TD class="SEARCHDATA" width="250"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 173px; HEIGHT: 22px" type="text"
														maxLength="100" size="20" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" size="6" name="txtTIMCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbJOBGUBN1,'')"
													width="60">���۱���</TD>
												<TD class="SEARCHDATA"><SELECT dataFld="cmbJOBGUBN1" id="cmbJOBGUBN1" title="���۱���" style="WIDTH: 98px" name="cmbJOBGUBN1"></SELECT></TD>
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
															<TD><IMG id="ImgConfirmRequest" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'"
																	height="20" alt="���õ� �ŷ��������� ���ο�û�մϴ�." src="../../../images/ImgConfirmRequest.gIF"
																	border="0" name="ImgConfirmRequest"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="���õ� �ŷ��������� �����մϴ�." src="../../../images/imgDelete.gIF" border="0"
																	name="imgDelete"></TD>
															<TD><IMG id="imgALLPrint" onmouseover="JavaScript:this.src='../../../images/imgALLPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgALLPrint.gif'"
																	height="20" alt="��ȸ�� �ŷ��������� ��ü�μ��մϴ�.." src="../../../images/imgALLPrint.gIF" border="0"
																	name="imgALLPrint"></TD>
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
								<TR id="tblBody1">
									<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="30083">
												<PARAM NAME="_ExtentY" VALUE="2540">
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
												<TD class="TITLE" vAlign="absmiddle" align="left" width="250" height="22"></TD>
												<TD class="TITLE" vAlign="absmiddle" align="left" width="650" height="22">�ۼ����� :
													<SELECT id="cmbGUBUN" title="�׷챸��" style="WIDTH: 90px" name="cmbGUBUN">
														<OPTION value="taxgroup" selected>�ջ��ۼ�</OPTION>
														<OPTION value="nongroup">�����ۼ�</OPTION>
													</SELECT>
													û������ : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 100px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="32" name="txtDEMANDDAY">
													<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">
													�������� : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 94px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtPRINTDAY">
													<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalPrintday"></TD>
												<TD vAlign="middle" align="right" height="22">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgTimSave" onmouseover="JavaScript:this.src='../../../images/ImgTimSaveOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTimSave.gif'"
																	alt="������ �ŷ��������� �����մϴ�." src="../../../images/ImgTimSave.gif" border="0" name="ImgTimSave"></TD>
															<TD><IMG id="ImgCustSave" onmouseover="JavaScript:this.src='../../../images/ImgCustSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgCustSave.gIF'"
																	height="20" alt="�����ֺ��� �ŷ��������� �����մϴ�.." src="../../../images/ImgCustSave.gIF" border="0"
																	name="ImgCustSave"></TD>
															<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="���� �ŷ��������� ����մϴ�.." src="../../../images/imgPrint.gIF" border="0"
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
								<TR id="tblBody2">
									<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 60%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
												VIEWASTEXT DESIGNTIMEDRAGDROP="213">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="30083">
												<PARAM NAME="_ExtentY" VALUE="5159">
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
					<!--Top TR End--></TBODY></TABLE>
			</TR></TBODY></TABLE></FORM>
	</body>
</HTML>