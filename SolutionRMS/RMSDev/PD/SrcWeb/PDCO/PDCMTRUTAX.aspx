<%@ Page CodeBehind="PDCMTRUTAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="PD.PDCMTRUTAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ݰ�꼭 ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMTRUTAX.aspx
'��      �� : ���ݰ�꼭 ����/��ȸ/����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/09 Ȳ����
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
Dim mobjPDCMTRUTAX , mobjSCCMGET , mobjPDCMGET
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
			ModuleDir = "PD"
			ReportName = "TRANSTAXNO_BLACK.rpt"
						
			Params = ""
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
			intRtn = mobjPDCMTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp���� ��
			
			ModuleDir = "PD"
			'������/���޹޴��� �������� ���忡 �ٺ����ְų� ���޹޴��� �����븸 �����ִ� ��
			'IF .chkPRINT.value THEN
			ReportName = "TRANSTAX_BLACK_NEW.rpt"
			'ELSE
			'	ReportName = "TRANSTAX_BLACKONE_NEW.rpt"
			'END IF
			
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
					
					vntDataTemp = mobjPDCMTRUTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			
			Params = strUSERID & ":" & "PD_TAX_TEMP"
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	

Sub imgConfirmPrint_onclick ()
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
		gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�. �μ��� �����͸� üũ�Ͻÿ�",""
		Exit Sub
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
			intRtn = mobjPDCMTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp���� ��
			
			ModuleDir = "PD"
			ReportName = "PDCMTRANS_CONFIRM_NEW.rpt"
			
			
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
					
					vntDataTemp = mobjPDCMTRUTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			
			Params = strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			'window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMTRUTAX.DeleteRtn_TEMP(gstrConfigXml)
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
		
		If .cmbGUBUN.value = "taxdiv" Then
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
					mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
					mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
					'mobjSCGLSpr.setTextBinding .sprSht,"SUMM",intCnt,"���ۺ� - (" & mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",intCnt) & ")"
				End If
			Next
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
					mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
					mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
					'mobjSCGLSpr.setTextBinding .sprSht,"SUMM",intCnt,"���ۺ�"
				End If
			Next
		End If
	End With
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
' ���ڵ��˾� ��ư[�Է¿�]
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
		if isArray(vntRet) then
			.txtTIMCODE1.value = trim(vntRet(0,0))
			.txtTIMNAME1.value = trim(vntRet(1,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			.txtCLIENTNAME1.value = trim(vntRet(5,0))
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
				Else
					Call TIMCODE1_POP()
				End If
   			end if
   		end with
   		SelectRtn
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub


Sub cmbJOBGUBN1_onChange ()
	if frmThis.txtFROM.value <> "" then
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE0_onClick ()
	if frmThis.txtFROM.value <> "" then
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE1_onClick ()
	if frmThis.txtFROM.value <> "" then
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

Sub chkVOCH_TYPE2_onClick ()
	if frmThis.txtFROM.value <> "" then
		gFlowWait meWAIT_ON
		SelectRtn
		gFlowWait meWAIT_OFF
	End if
End Sub

'-----------------------------------------------------------------------------------------
' �޷�
'-----------------------------------------------------------------------------------------
Sub imgDEMANDDAY_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgDEMANDDAY,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
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

'û����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTO_onchange
	gSetChange
End Sub

Sub cmbGUBUN_onchange
	with frmThis
		If .cmbGUBUN.value = "taxdiv" Then
			selectRtn
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			selectRtn
		End If
	End with
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
	Dim vntRet
	Dim vntInParams
	Dim strMEDFLAG
	DIM strTAXYEARMON
	DIM strTAXNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		else
			strTAXYEARMON =  mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",Row)
			strTAXNO =  mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",Row)
			
			IF .rdT.checked THEN
			
				If mobjSCGLSpr.GetTextBinding(.sprSht,"MERGEFLAG",Row) = "1" Then
					gErrorMsgBox "����û����꼭�� �����Ȱ� �Դϴ�.",""
					EXIT SUB
				End If 
				
				vntInParams = array(strTAXYEARMON, strTAXNO) '<< �޾ƿ��°��
				vntRet = gShowModalWindow("PDCMTRUTAXDTL.aspx",vntInParams , 813,545)
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
				
				
			END IF
		
			if isArray(vntRet) then
     		end if
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
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
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
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
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")  Then
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
				
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR _ 
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
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR _
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT") Then
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
					.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMT")  Then
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

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjPDCMTRUTAX	= gCreateRemoteObject("cPDCO.ccPDCOTRUTAX")
	set mobjSCCMGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	set mobjPDCMGET		= gCreateRemoteObject("cPDCO.ccPDCOGET")
	

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjPDCMTRUTAX = Nothing
	set mobjSCCMGET = Nothing
	set mobjPDCMGET = Nothing
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
		DateClean 
		
		.sprSht.MaxRows = 0
		CALL COMBO_TYPE()
		.cmbJOBGUBN1.selectedIndex = -1
		
		CALL Grid_Setting ("TRANS")
		
		.txtCLIENTNAME1.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
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

		vntJOBGUBN = mobjPDCMTRUTAX.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"JOBGUBN")  '���۱���

		if not gDoErrorRtn ("COMBO_TYPE") then 
			mobjSCGLSpr.TypeComboBox = True 
			gLoadComboBox .cmbJOBGUBN1, vntJOBGUBN, False
   		end if    	
   		
   	end with     	
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	
	strDATE = Mid(gNowDate2,1,4) & "-" & Mid(gNowDate2,6,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub


Sub Grid_Setting (strGUBUN)
	With frmThis
		'Sheet �⺻Color ����
		.sprSht.MaxRows = 0
		.sprSht.style.visibility = "hidden"
		Call Grid_init()
		gSetSheetDefaultColor() 
		
		'�Ϸ��϶�
		If strGUBUN = "TAX" Then
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 24, 0, 1, 2
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBUSINO | TIMNAME | SUBSEQNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | VOCHNO | TAXYEARMON | TAXNO | JOBGUBN|TAXCODE|TAXNAME|JOBPARTNAME | MERGEFLAG"
			mobjSCGLSpr.SetHeader .sprSht,		  "����|������ȣ|û�����|������|�����ֻ���ڹ�ȣ|��|�귣��|�ݾ�|�ΰ���|�հ�ݾ�|����|������|���μ�|�����ִ�ǥ�ڸ�|�������ּ�1|�������ּ�2|��ǥ��ȣ|���ݰ�꼭���|���ݰ�꼭��ȣ|���۱���|���ݰ�꼭�ڵ�|���ݰ�꼭����|���ۺз�|����û������"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "  5|      11|       8|    15|	           13|13|	  13|  10|    10|	   11|  20|     9| 	     8|             0|          0|          0|      10|             0|             0|     0 |0             |12            |10      |           0"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TAXMANAGE | CLIENTNAME | CLIENTBUSINO | TIMNAME | SUBSEQNAME | SUMM | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | VOCHNO | TAXYEARMON | TAXNO | JOBGUBN", -1, -1, 100
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXMANAGE | DEMANDDAY | CLIENTNAME | CLIENTBUSINO | TIMNAME | SUBSEQNAME | AMT| VAT | SUMAMT | SUMM | PRINTDAY | DEPT_NAME | CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | VOCHNO | TAXYEARMON | TAXNO | JOBGUBN|JOBPARTNAME "
			mobjSCGLSpr.ColHidden .sprSht, "CLIENTOWNER | CLIENTADDR1| CLIENTADDR2 | VOCHNO | JOBGUBN|TAXCODE|TAXYEARMON | TAXNO | MERGEFLAG", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE | CLIENTBUSINO | VOCHNO",-1,-1,2,2,False
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXNAME|JOBPARTNAME",-1,-1,0,2,false
			mstrGUBUN = "TAX"
			
		'�Ϸᰡ �ƴҶ�
		Else
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 38, 0, 10, 2
			mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|TAXMANAGE|TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|SEQ|JOBNOSEQ|JOBNO|JOBNAME|SUMM|DEMANDDAY|AMT|VAT|SUMAMT|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DEPTCD|DEPTNAME|PRINTDAY|ACCODE|REALBUSINO|CLIENTBUSINO|CLIENTOWNER|CLIENTADDR1|CLIENTADDR2|VOCHNO|RANKTRANS|INCJOBNO|JOBGUBN|TAXCODE|TAXNAME|JOBPART|JOBPARTNAME"
			mobjSCGLSpr.SetHeader .sprSht,		  "����|��꼭��ȣ|���|��ȣ|���|��ȣ|����|JOB����|JOBNO|JOBNAME|����|û����|�ݾ�|�ΰ�����|�հ�ݾ�|�������ڵ�|�����ָ�|���ڵ�|����|�귣���ڵ�|�귣���|�μ��ڵ�|�μ���|������|ȸ���ڵ�|����ڹ�ȣ|�����ֻ���ڹ�ȣ|��ǥ�ڸ�|�ּ�1|�ּ�2|��ǥ��ȣ|�ջ����|����û����|���۱���|���ݰ�꼭�ڵ�|���ݰ�꼭����|���ۺз�|���ۺз���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1","   5|        11|   5|   4|   5|   4|   4|      0|   12|      0|  19|     8|   9|       9|       9|         0|      25|     0|  25|         0|      25|       0|    18|     8|       0|        12|               0|      0|    0|    0|      10|       0|         0|       9 |0             |12            |0       |0"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TRANSNO|SEQ|JOBNOSEQ|AMT|VAT|SUMAMT", -1, -1, 0
			mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|SEQ|JOBNOSEQ|JOBNO|JOBNAME|DEMANDDAY|CLIENTNAME|TIMNAME|SUBSEQNAME|AMT|SUMAMT|DEPTNAME|PRINTDAY|CLIENTCODE|TIMCODE|ACCODE|REALBUSINO|CLIENTBUSINO|CLIENTOWNER|CLIENTADDR1|CLIENTADDR2|DEPTCD|VOCHNO|RANKTRANS|INCJOBNO|JOBGUBN|TAXCODE|TAXNAME|JOBPART|JOBPARTNAME"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "SUMM", -1, -1, 255
			mobjSCGLSpr.SetCellAlign2 .sprSht, "TAXMANAGE|TRANSYEARMON|TRANSNO|SEQ|TAXNO|REALBUSINO|CLIENTBUSINO|CLIENTOWNER|TAXYEARMON|JOBNO|JOBGUBN",-1,-1,2,2,false
			mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|TIMNAME|SUMM|DEPTNAME|SUBSEQNAME|CLIENTADDR1|CLIENTADDR2|TAXNAME|JOBPARTNAME",-1,-1,0,2,false
			mobjSCGLSpr.ColHidden .sprSht, "TAXMANAGE|TAXYEARMON|TAXNO|JOBNAME|CLIENTCODE|ACCODE|TIMCODE|REALBUSINO|CLIENTBUSINO|CLIENTOWNER|CLIENTADDR1|CLIENTADDR2|DEPTCD|RANKTRANS|INCJOBNO|SUBSEQ|JOBGUBN|TAXCODE|JOBGUBN|JOBPART|JOBPARTNAME", true
			mstrGUBUN = "TRANS"
		End If
		
		'Get_COMBO_VALUE
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
	
	SelectRtn
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	SelectRtn
End Sub
'��üüũ
Sub rdA_onclick
	SelectRtn
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCLIENTCODE
	Dim strTIMCODE
	Dim strFROM,strTO 
   	Dim i, strCols
   	Dim strGUBUN
   	Dim strMED_FLAG
   	Dim strVOCH_TYPE_TEMP
   
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
				
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IF .rdT.checked = TRUE THEN
			CALL Grid_Setting ("TAX")
		ELSE
			CALL Grid_Setting ("TRANS")
		END IF
		
		strFROM			=  MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO			=  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strTIMCODE		= .txtTIMCODE1.value
		strGUBUN		= .cmbGUBUN.value
		strMED_FLAG		= .cmbJOBGUBN1.value
		
		strVOCH_TYPE_TEMP = ""
		
		
		'���ݰ�꼭 �Ϸ���ȸ
		If .rdT.checked = True Then
			vntData = mobjPDCMTRUTAX.Get_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM, strTO,  strCLIENTCODE, strMED_FLAG)
			If not gDoErrorRtn ("Get_TAX") then
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
			vntData = mobjPDCMTRUTAX.Get_TAXBUILD(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM, strTO, strCLIENTCODE,strTIMCODE,  strGUBUN, strMED_FLAG)
			If not gDoErrorRtn ("Get_TAXBUILD") then
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
			
			vntData = mobjPDCMTRUTAX.Get_TAXALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM, strTO, strCLIENTCODE,strTIMCODE,  strGUBUN, strMED_FLAG)
			If not gDoErrorRtn ("Get_TAXALL") then
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

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
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

Sub Layout_change ()
	Dim intCnt
	with frmThis
		For intCnt = 1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"MERGEFLAG",intCnt) = "1" Then
				mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HAAAAAA, &H000000,False 
				'mobjSCGLSpr.SetCellsLock2 .sprSht,TRUE,"CHK",-1,-1,false
			else
				'mobjSCGLSpr.SetCellsLock2 .sprSht,FALSE,"CHK",-1,-1,false
			End If 
		Next 
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TAXMANAGE|TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|SEQ|JOBNOSEQ|JOBNO|JOBNAME|SUMM|DEMANDDAY|AMT|VAT|SUMAMT|CLIENTCODE|CLIENTNAME|TIMCODE|TIMNAME|SUBSEQ|SUBSEQNAME|DEPTCD|DEPTNAME|PRINTDAY|ACCODE|REALBUSINO|CLIENTBUSINO|CLIENTOWNER|CLIENTADDR1|CLIENTADDR2|VOCHNO|RANKTRANS|INCJOBNO|JOBGUBN|TAXCODE|JOBGUBN")
		
		'������ �����͸� ���� �´�.
		'ó�� ������ü ȣ��
		intTAXNO = 0
		If .cmbGUBUN.value = "taxdiv" Then
		intRtn = mobjPDCMTRUTAX.ProcessRtn_Div(gstrConfigXml,vntData, intTAXNO)
		Else
			If Not TaxGroup(strVALIDATION) Then 
				gErrorMsgBox strVALIDATION & vbCrlf & "������ [������] [û����,�ۼ���,����,����] �� ���� �Ͽ��� �մϴ�.","����ȳ�!"
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
				intRtn = mobjPDCMTRUTAX.ProcessRtn_Group(gstrConfigXml,vntData, intTAXNO,intColFlag)
			End IF
		End If

		If not gDoErrorRtn ("ProcessRtn") Then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "���ݰ�꼭�� �����Ǿ����ϴ�.","����ȳ�!"
			.rdT.checked = True
			selectRtn
   		End If
   	end with
End Sub

Function TaxGroup(ByRef strVALIDATION)
	Dim intCnt
	Dim strCLIENTCODE '������ ����� ��Ϲ�ȣ
	Dim strDEMANDDAY
	Dim strPRINTDAY
	Dim strSUMM
	Dim strStartRank
	Dim strVOCH_TYPE
	
	TaxGroup = False
	with frmThis
		strStartRank = "0"
		strCLIENTCODE = ""
		strDEMANDDAY = ""
		strPRINTDAY = ""
		strSUMM = ""
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				
				If strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Then
					If strCLIENTCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBUSINO",intCnt) Then
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
					'If strSUMM <> mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt) Then
					'	strVALIDATION = "����Ȯ�� �ŷ�������ȣ" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
					'	Exit Function
					'End If
					
				End If
				
				strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBUSINO",intCnt)
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
			
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
			
				intRtn = mobjPDCMTRUTAX.DeleteRtn_TruTax(gstrConfigXml,strTAXYEARMON, strTAXNO)
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
		
		IF not gDoErrorRtn ("DeleteRtn_TruTax") then
			gWriteText lblstatus, intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub

'��ǥ��ȣ üũ
Function VOCHNO_CHECKED (ByRef strTAXYEARMON, ByRef strTAXNO)
	Dim vntData
	Dim intCnt
	Dim strCOUNT
	'on error resume next

	'�ʱ�ȭ
	VOCHNO_CHECKED = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMGET.VOCHNO_CHECKED(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON,strTAXNO) 
	
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
	
	vntData = mobjPDCMGET.COMMIVOCHNO_CHECKED_MERGE(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON,strTAXNO, "P")
	
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
											<td class="TITLE">���ݰ�꼭 ����</td>
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
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM,'')"
												width="50">�����</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtFROM" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="2" name="txtFROM">&nbsp;<IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
													border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="10" size="6" name="txtTO">&nbsp;<IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle"
													border="0" name="imgTo"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="50">������
											</TD>
											<TD class="SEARCHDATA" width="220"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 143px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="14" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE1">
												<INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)"
												width="50">��
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="����" style="WIDTH: 143px; HEIGHT: 22px" type="text"
													maxLength="100" size="14" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgTIMCODE1"> <INPUT class="INPUT_L" id="txtTIMCODE1" title="���ڵ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
													maxLength="6" size="6" name="txtTIMCODE1">
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
											<TD class="SEARCHDATA"><INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN">
												&nbsp;�Ϸ�&nbsp; <INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" CHECKED value="rdF" name="rdGBN">
												&nbsp;�̿Ϸ�&nbsp;&nbsp;<INPUT id="rdA" title="��ü ������ȸ" type="radio" value="rdA" name="rdGBN">&nbsp;��ü</TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(cmbJOBGUBN1, '')"
												width="50">���۱���</TD>
											<TD class="SEARCHDATA" width="90" colSpan="6"><SELECT dataFld="cmbJOBGUBN1" id="cmbJOBGUBN1" title="���۱���" style="WIDTH: 98px" name="cmbJOBGUBN1"></SELECT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="absmiddle" align="center">
									<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 20px" cellSpacing="1"
										cellPadding="0" align="left" border="0">
										<TR>
											<TD height="20" colspan="4">
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
											<TD height="4" colspan="4"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 67px">û��������</TD>
											<TD class="SEARCHDATA" style="WIDTH: 350px"><INPUT class="INPUT" id="txtDEMANDDAY" title="û������" style="WIDTH: 120px; HEIGHT: 22px"
													accessKey="date" type="text" maxLength="10" size="14" name="txtDEMANDDAY"> <IMG id="imgDEMANDDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0"
													name="imgDEMANDDAY">&nbsp;<IMG id="btnCOMMISSION" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" height="20" alt="�ش� �����Ϸ� �������� ���� ���׸��� Setting �մϴ�"
													src="../../../images/imgApp.gIF" width="54" align="absMiddle" border="0" name="btnCOMMISSION">
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 250px"><SELECT id="cmbGUBUN" title="��ü����" style="WIDTH: 80px" name="cmbGUBUN">
													<OPTION value="taxdiv" selected>���ҹ���</OPTION>
													<OPTION value="taxgroup">�ջ����</OPTION>
												</SELECT>&nbsp;<SELECT id="chkPRINT" title="��¹�����" style="WIDTH: 80px" name="chkPRINT">
													<OPTION value="1" selected>���ڿ�</OPTION>
													<OPTION value="0">���޹޴��ڿ�</OPTION>
												</SELECT>&nbsp;<SELECT id="cmbFLAG" title="����/û������" style="WIDTH: 80px" name="cmbFLAG">
													<OPTION value="receipt" selected>û��</OPTION>
													<OPTION value="demand">����</OPTION>
												</SELECT></TD>
											<TD class="DATA_RIGHT" vAlign="middle" align="right" height="20">
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
														<TD><IMG id="imgConfirmPrint" onmouseover="JavaScript:this.src='../../../images/imgConfirmPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirmPrint.gif'"
																height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgConfirmPrint.gIF"  border="0"
																name="imgConfirmPrint"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
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
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 99%">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31882">
										<PARAM NAME="_ExtentY" VALUE="14235">
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
	</body>
</HTML>
