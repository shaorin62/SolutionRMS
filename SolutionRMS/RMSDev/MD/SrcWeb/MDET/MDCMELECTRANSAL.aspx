<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANSAL.aspx.vb" Inherits="MD.MDCMELECTRANSAL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ŷ����� ����</title>
		<meta content="False" name="vs_snapToGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ����� ��� ȭ��(MDCMCATVTRANSAL.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMCATVTRANSAL.aspx
'��      �� : ����Ź�ŷ����� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/11/21 By HWANG DUCK SU
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDETELECTRANS, mobjMDETELEC
Dim mobjMDCMGET
Dim mstrCheck, mstrCheck1, mstrGrid

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
			document.getElementById("tblSheet1").style.height = "70%"
			document.getElementById("tblSheet2").style.height = "30%"
		ELSEIf strmode = "HIDDEN" Then
			document.getElementById("tblBody1").style.display = "none"
			document.getElementById("tblSheet2").style.height = "100%"
		ELSEIF strmode = "STANDARD" Then
			document.getElementById("tblBody1").style.display = "inline"
			document.getElementById("tblSheet1").style.height = "30%"
			document.getElementById("tblSheet2").style.height = "70%"
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
' ��� ��ư Ŭ�� �̺�Ʈ
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

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
	
Sub ImgCRE_onclick
	If frmThis.sprSht_DTL.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON

	IF frmThis.chkSPONSOR.checked = True then
		ProcessRtn_ALL
	Else
		ProcessRtn
	End if
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
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
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
     				 'MDCMCATVTRANS_NEW.rpt
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		For i = 1 to .sprSht_HDR.MaxRows
			mobjSCGLSpr.CellChanged .sprSht_HDR, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"TRANSYEARMON | TRANSNO | CNT")
		
		strUSERID = "" 
		vntDataTemp = mobjMDETELECTRANS.ProcessRtn_TEMP_ALL(gstrConfigXml, vntData, strUSERID)

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
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
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
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",.sprSht_HDR.ActiveRow)
		strCNT			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CNT",.sprSht_HDR.ActiveRow)
		
		strUSERID = ""
		vntData = mobjMDETELECTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
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
		intRtn = mobjMDETELECTRANS.DeleteRtn_temp(gstrConfigXml)
		'intRtn2 = mobjMDETELECTRANS.DeleteRtnUpdate_PRINTSEQ(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP ()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(.txtYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "trans", "ELEC") 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			IF vntRet(3,0) = "�Ϸ�" THEN
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(4,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			ELSE
				.txtYEARMON1.value = vntRet(0,0)
				.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			END IF
			selectRtn
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag ���� �˸�
		end if
	End with
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
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON1.value, .txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"","trans", "ELEC")
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = vntData(0,1)
					.txtCLIENTCODE1.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
					selectRtn
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub	

'-----------------------------------------------------------------------------------------
' ���˾� ��ư[��ȸ��]
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
		vntInParams = array(trim(.txtYEARMON1.value), trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
							trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "trans", "ELEC") 
							
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSTIMPOP.aspx",vntInParams , 413,465)
		'TRANSYEARMON | TIMNAME | CLIENTNAME | GBN | CLIENTCODE | TIMCODE
		if isArray(vntRet) then
			.txtYEARMON1.value = trim(vntRet(0,0))
			.txtTIMCODE1.value = trim(vntRet(5,0))
			.txtTIMNAME1.value = trim(vntRet(1,0))
			.txtCLIENTCODE1.value = trim(vntRet(4,0))
			.txtCLIENTNAME1.value = trim(vntRet(2,0))
			selectrtn
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
			vntData = mobjMDCMGET.GetTRANSTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, _
												  trim(.txtYEARMON1.value),trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value), _
												  trim(.txtTIMCODE1.value), trim(.txtTIMNAME1.value), "", "trans", "ELEC")
			
			if not gDoErrorRtn ("GetTRANSTIMCODE") then
				If mlngRowCnt = 1 Then
					.txtYEARMON1.value = trim(vntData(0,1))
					.txtTIMCODE1.value = trim(vntData(5,1))
					.txtTIMNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(4,1))
					.txtCLIENTNAME1.value = trim(vntData(2,1))
					selectrtn
				Else
					Call TIMCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if	
End Sub


'�귣��
Sub ImgSUBSEQ1_onclick
	Call SUBSEQCODE1_POP()
End Sub

Sub SUBSEQCODE1_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ1.value), trim(.txtSUBSEQNAME1.value), trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		If isArray(vntRet) Then
			If .txtSUBSEQ1.value = vntRet(0,0) and .txtSUBSEQNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ1.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME1.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE1.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtTIMNAME1.value = trim(vntRet(5,0))	' �����ָ� ǥ��
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ1.value),trim(.txtSUBSEQNAME1.value),  _
												trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ1.value = trim(vntData(0,1))
					.txtSUBSEQNAME1.value = trim(vntData(1,1))
					.txtCLIENTCODE1.value = trim(vntData(2,1))		' ������ ǥ��
					.txtCLIENTNAME1.value = trim(vntData(3,1))	' ������
					.txtTIMCODE1.value = trim(vntData(4,1))	' ������
					.txtTIMNAME1.value = trim(vntData(5,1))	' ������
				Else
					Call SUBSEQCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' �����ڵ��˾� ��ư
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgMATTER1_onclick
	Call MATTERCODE1_POP()
	
End Sub

Sub MATTERCODE1_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		
		vntInParams = array(trim(.txtCLIENTNAME1.value), trim(.txtTIMNAME1.value), trim(.txtSUBSEQNAME1.value),"", _
							trim(.txtMATTERNAME1.value), "", "A") '<< �޾ƿ��°��
							
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP_ALL.aspx",vntInParams , 780,630)
		if isArray(vntRet) then
			if .txtMATTERCODE1.value = vntRet(0,0) and .txtMATTERNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			.txtMATTERCODE1.value = trim(vntRet(0,0))	' �����ڵ� ǥ��
			.txtMATTERNAME1.value = trim(vntRet(1,0))	' ����� ǥ��
			.txtCLIENTCODE1.value = trim(vntRet(2,0))	' �������ڵ� ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			.txtTIMCODE1.value = trim(vntRet(4,0))		' ���ڵ� ǥ��
			.txtTIMNAME1.value = trim(vntRet(5,0))		' ���� ǥ��
			.txtSUBSEQ1.value = trim(vntRet(6,0))		' �귣�� ǥ��
			.txtSUBSEQNAME1.value = trim(vntRet(7,0))	' �귣��� ǥ��
			'.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' ���ۻ��ڵ� ǥ��
			'.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' ���ۻ��ڵ� ǥ��
			'.txtDEPT_CD.value = trim(vntRet(10,0))		' �μ��ڵ� ǥ��
			'.txtDEPT_NAME.value = trim(vntRet(11,0))	' �μ��� ǥ��
			'.txtCLIENTSUBCODE.value = trim(vntRet(12,0))	' ������ڵ� ǥ��
			'.txtCLIENTSUBNAME.value = trim(vntRet(13,0))	' ����θ� ǥ��
			'.txtGREATCODE.value = trim(vntRet(14,0))	' ����ó�ڵ� ǥ��
			'.txtGREATNAME.value = trim(vntRet(15,0))	' ����ó�� ǥ��
			
     	end if
	End with
	'gSetChange
	
End Sub

Sub txtMATTERNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME1.value),trim(.txtTIMNAME1.value), trim(.txtSUBSEQNAME1.value),"", _
											trim(.txtMATTERNAME1.value), "", "A")
			if not gDoErrorRtn ("GetMATTER") then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE1.value = trim(vntData(0,1))		' �귣�� ǥ��
					.txtMATTERNAME1.value = trim(vntData(1,1))	' �귣��� ǥ�� 2,3,6,7
					.txtCLIENTCODE1.value = trim(vntData(2,1))
					.txtCLIENTNAME1.value = trim(vntData(3,1))
					.txtTIMCODE1.value	  = trim(vntData(4,1))	' ���ڵ� ǥ��
					.txtTIMNAME1.value	  = trim(vntData(5,1))	' ���� ǥ��
					.txtSUBSEQ1.value      = trim(vntData(6,1))
					.txtSUBSEQNAME1.value  = trim(vntData(7,1))
				Else
					Call MATTERCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub




'****************************************************************************************
'��ȸ�ʵ� onclick onkeydown 
'****************************************************************************************

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
' ������ �޷�
'****************************************************************************************
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

'-----------------------------------------------------------------------------------------
' û������ ����
'-----------------------------------------------------------------------------------------
Sub txtYEARMON1_onblur
	With frmThis
		If .txtYEARMON1.value <> "" AND Len(.txtYEARMON1.value) = 6 Then DateClean
	End With
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

Sub sprSht_CLIENT_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 AND Col > 1 then
			mstrGrid = TRUE
			Call SelectRtn_HDR (Col, Row)
			Call selectRtn_ELECTRIC_MEDIUM ()
			mstrGrid = false
		end if
	end with
End Sub


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
			SelectRtn_DTL Col, Row
			mstrGrid = false
		end if
	end with
End Sub

Sub sprSht_DTL_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		IF mstrGrid = false THEN
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


sub sprSht_CLIENT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_CLIENT, ""
		end if
	end with
end sub

sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

Sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End Sub

Sub sprSht_CLIENT_Keyup(KeyCode, Shift)
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
		Call SelectRtn_HDR (frmThis.sprSht_CLIENT.ActiveCol,frmThis.sprSht_CLIENT.ActiveRow)
		Call SelectRtn_ELECTRIC_MEDIUM ()
	End If
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
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") _
			   or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE") Then
					strCOLUMN = "PRICE"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")) OR _ 
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRICE")) Then
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
			If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") OR .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")  THEN

				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""
				
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") Then
					strCOLUMN = "AMT"
				ELSEIF .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") Then
					strCOLUMN = "VAT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht_DTL,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT")) OR _
												(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT"))   Then
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
			If .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or .sprSht_HDR.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"VAT") Then
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT")   Then
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
				If .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"AMT") or _
			       .sprSht_DTL.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"VAT") THEN
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
	Dim i
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strSEQ
	Dim strPRINT_SEQ
	Dim intRtn

	with frmThis
		If mstrGrid Then
			'���� �׸��尡 ���ݰ�꼭 �󼼳����϶� ��¼����� ���Ҷ� �߻�
			'if Col = mobjSCGLSpr.CnvtDataField(.sprSht_DTL,"PRINT_SEQ") then
			'	for i=1 to .sprSht_DTL.MaxRows
			'		if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) <> "" then
			'			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) = 0 then
			'				mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
			'			else
			'				
			'				if Row <> i then
			'					'�Է��Ѽ��ڿ� ���� ��¼����� ���ִٸ� ����
			'					if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row) = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",i) then
			'						gErrorMsgBox "��¼����� �ߺ��ԷµǾ����ϴ�.",""
			'						mobjSCGLSpr.SetTextBinding .sprSht_DTL,"PRINT_SEQ",Row, ""
			'						.txtCLIENTNAME1.focus() 
			'						.sprSht_DTL.focus()
			'						EXIT SUB
			'					end if
			'				end if
			'			end if
			'		end if
			'	next
				
			'	strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSYEARMON",Row)
			'	strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSNO",Row)
			'	strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"SEQ",Row)
			'	strPRINT_SEQ = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"PRINT_SEQ",Row)
				
			'	intRtn = mobjMDETELECTRANS.UPDATE_PRINTSEQ(gstrConfigXml,strTRANSYEARMON,strTRANSNO, strSEQ, strPRINT_SEQ)
			'end if
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
	set mobjMDETELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")
	set mobjMDETELEC		= gCreateRemoteObject("cMDET.ccMDETELECT_MEDIUM")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    With frmThis
		
		'******************************************************************
		'�ŷ����� ������ �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_CLIENT	
		mobjSCGLSpr.SpreadLayout .sprSht_CLIENT, 4, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht_CLIENT, "TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.SetHeader .sprSht_CLIENT,		  "�ŷ��������|�ŷ�������ȣ|�������ڵ�|�����ָ�"
		mobjSCGLSpr.SetColWidth .sprSht_CLIENT, "-1", "             0|             0|         0|      20"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CLIENT, "0", "15"
		mobjSCGLSpr.SetCellsLock2 .sprSht_CLIENT, true, "CLIENTCODE | CLIENTNAME"
		mobjSCGLSpr.ColHidden .sprSht_CLIENT, "TRANSYEARMON | TRANSNO | CLIENTCODE", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_CLIENT, "CLIENTNAME" ,-1,-1,0,2,false
		.sprSht_CLIENT.style.visibility = "visible"
		
		'******************************************************************
		'�ŷ����� ��� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 12, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | TRANSYEARMON | TRANSNO | SEQ | VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | VAT | SUMAMTVAT | CNT"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		  "����|�ŷ��������|�ŷ�������ȣ|����|����|����|��ü���ڵ�|��ü��|����ݾ�|�ΰ���|��|�����"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", "   4|            0|             0|   0|   6|   6|         0|    30|      20|    20|15|      10"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT | VAT | SUMAMTVAT |  CNT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE | REAL_MED_NAME | AMT | VAT | SUMAMTVAT "
		mobjSCGLSpr.ColHidden .sprSht_HDR, "REAL_MED_CODE", TRUE
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "REAL_MED_NAME" ,-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, "TRANSYEARMON | TRANSNO | SEQ | VOCH_TYPE | CONFIRMFLAGNAME | REAL_MED_CODE" ,-1,-1,2,2,false
		.sprSht_HDR.style.visibility = "visible"
		
		
		'******************************************************************
		'' û�೻�� ��ȸ  ��   �ŷ����� ������ 
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 37, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		  "����|�ŷ��������|�ŷ�������ȣ|����|�������ڵ�|�����ָ�|������ڵ�|����θ�|��ü�ڵ�|��ü��|��ü���ڵ�|��ü��|�귣���ڵ�|�귣���|���ڵ�|����|�����ڵ�|�����|���α׷���|����|��ۿ���|�μ��ڵ�|û������|��������|���ް���|ȸ��|������|�ΰ���|��ü�����ڵ�|��ü����|���ݰ�꼭���|���ݰ�꼭��ȣ|��Ź����|TRANSCUSTRANK | TRANSSPONRANK | ��ǥ���� | ���ۻ��ڵ�" 
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "  4|             0|	            0|  0|         0|	    0|          0|       0|       0|    15|         0|    15|         0|      15|     0|  15|       0|    15|        10|   6|      10|       0|       0|       8|       9|   5|       9|     9|            0|       9|             0|             0|       0|             0 |            5|         0|          0"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "PRICE | AMT| VAT | CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht_DTL, "MEDNAME|REAL_MED_NAME|PROGRAM|ADLOCALFLAG|WEEKDAY|MED_FLAG|SUBSEQNAME|TIMNAME|MATTERNAME", -1, -1, 50
		mobjSCGLSpr.ColHidden .sprSht_DTL, "TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | REAL_MED_CODE | DEPT_CD | DEMANDDAY | TAXYEARMON | TAXNO | SUBSEQ | TIMCODE | MATTERCODE | INPUT_MEDFLAG | TRANSCUSTRANK | VOCH_TYPE | EXCLIENTCODE", true	
		.sprSht_DTL.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDETELEC = Nothing
	set mobjMDETELECTRANS = Nothing
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
		.txtYEARMON1.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		DateClean

		.txtPRINTDAY.value  = gNowDate
		.sprSht_CLIENT.MaxRows = 0
		.sprSht_HDR.MaxRows = 0	
		.sprSht_DTL.MaxRows = 0
		'.chkSPONSOR.checked = TRUE
		.cmbREAL_MED_CODE1.selectedIndex = 0
		.cmbMEDGUBUN1.selectedIndex = 0

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
' �׸��� �޺��ڽ� ����
'-----------------------------------------------------------------------------------------
Sub Get_COMBO_VALUE ()
	Dim vntData, vntData_VOCH, vntData_DUTY
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData_VOCH = mobjMDETELEC.Get_COMBOVOCH_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt)
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "VOCH_TYPE",,,vntData_VOCH,,60 
			mobjSCGLSpr.TypeComboBox = True 
			
   		End If    
   	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData_CLIENT, vntData_HDR , vntData_DTL
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE, strCLIENTNAME 
   	Dim strTIMCODE, strTIMNAME
	Dim strSUBSEQ, strSUBSEQNAME
	Dim strMATTERCODE , strMATTERNAME
	Dim strREAL_MED_CODE
	Dim strMEDGUBN
    
	'On error resume next
	with frmThis
	
		If .txtYEARMON1.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		'Sheet�ʱ�ȭ
		.sprSht_CLIENT.MaxRows = 0
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
		strSUBSEQ		= .txtSUBSEQ1.value
		strSUBSEQNAME	= .txtSUBSEQNAME1.value
		strMATTERCODE	= .txtMATTERCODE1.value
		strMATTERNAME	= .txtMATTERNAME1.value
		strREAL_MED_CODE= .cmbREAL_MED_CODE1.value
		strMEDGUBN		= .cmbMEDGUBUN1.value
		
		vntData_CLIENT = mobjMDETELECTRANS.SelectRtn_CLIENT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strCLIENTNAME, _
							   			strTIMCODE, strTIMNAME, strSUBSEQ, strSUBSEQNAME, strMATTERCODE, strMATTERNAME, strREAL_MED_CODE, strMEDGUBN)
		
		If not gDoErrorRtn ("SelectRtn_CLIENT") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_CLIENT,vntData_CLIENT,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				Call SelectRtn_HDR (1, 1)
   				Call SelectRtn_ELECTRIC_MEDIUM ()
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   		
   		

   	end with
End Sub

Sub SelectRtn_HDR (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strCLIENTCODE
   	Dim i, strCols
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"TRANSYEARMON",Row)
		strCLIENTCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",Row)
				
		vntData = mobjMDETELECTRANS.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strCLIENTCODE)
																							
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   			mstrGrid = TRUE
   		End If
   	end with
End Sub


Sub SelectRtn_ELECTRIC_MEDIUM ()
	Dim vntData , vntDataCHK
	Dim strTRANSYEARMON
   	Dim i, strCols
   	Dim strYEARMON
   	Dim strCLIENTCODE, strCLIENTNAME 
   	Dim strTIMCODE, strTIMNAME
	Dim strSUBSEQ, strSUBSEQNAME
	Dim strMATTERCODE , strMATTERNAME
	Dim strREAL_MED_CODE
	Dim strMEDGUBN
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		strSUBSEQ		= .txtSUBSEQ1.value
		strSUBSEQNAME	= .txtSUBSEQNAME1.value
		strMATTERCODE	= .txtMATTERCODE1.value
		strMATTERNAME	= .txtMATTERNAME1.value
		strREAL_MED_CODE= .cmbREAL_MED_CODE1.value
		strMEDGUBN		= .cmbMEDGUBUN1.value
		
		If strCLIENTCODE = "" Then
   			strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht_CLIENT,"CLIENTCODE",.sprSht_CLIENT.ActiveRow)
   		End If
   		
   		'û�� Ȯ�� ���� ���� ������ �����Ѵ�. [�귣��][���μ�]
   		vntDataCHK = mobjMDETELECTRANS.SelectRtn_DTLCHK(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE)
		If not gDoErrorRtn ("SelectRtn_DTLCHK") Then
			IF mlngRowCnt > 0 THEN
			
			gErrorMsgBox "û�� ��������� ���� ���� �����Ͱ� �ֽ��ϴ� ." & vbcrlf & "   û�� ������ ���� ���� �����ʹ� �ŷ��������� ������ �ʽ��ϴ�.","�ŷ����� �ȳ�"
			END IF
		END IF    		
		
		vntData = mobjMDETELECTRANS.SelectRtn_ELECTRIC_MEDIUM(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE, strCLIENTNAME, _
							   								  strTIMCODE, strTIMNAME, strSUBSEQ, strSUBSEQNAME, strMATTERCODE, strMATTERNAME, _
							   								  strREAL_MED_CODE, strMEDGUBN)
																							
		If not gDoErrorRtn ("SelectRtn_ELECTRIC_MEDIUM") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatusDTL, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatusDTL, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = false
   		End If
   	end with
End Sub

Sub SelectRtn_DTL (Col, Row)
	Dim vntData
	Dim strTRANSYEARMON, strTRANSNO
   	Dim i, strCols, intcnt
    
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		intcnt = 1

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSYEARMON",Row)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TRANSNO",Row)
		
		vntData = mobjMDETELECTRANS.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO)
																							
		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
				for intcnt = 1 to .sprSht_DTL.MaxRows  
					mobjSCGLSpr.SetCellTypeStatic .sprSht_DTL, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht_DTL,"CHK",intcnt," "
				next
				
   				gWriteText lblStatusDTL, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			else
   				gWriteText lblStatusDTL, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_DTL.MaxRows = 0
   			End If
   			AMT_SUM
   			mstrGrid = TRUE
   		End If
   	end with
End Sub



'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM, IntVAT, IntVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		
		For lngCnt = 1 To .sprSht_DTL.MaxRows
			IntAMT = 0
			IntVAT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT", lngCnt)
			IntVAT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"VAT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
			IntVATSUM = IntVATSUM + IntVAT
		Next
		If .sprSht_DTL.MaxRows = 0 Then
			.txtSUMAMT.value = 0
			.txtSUMVAT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			.txtSUMVAT.value = IntVATSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
			Call gFormatNumber(frmThis.txtSUMVAT,0,True)
		End If
	End With
End Sub


'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	Dim strVOCH_TYPE
	
	chkcnt = 0
	intColFlag = 0
	strVOCH_TYPE = 0
	
	with frmThis
		
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'�׷��ִ밪 ����
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSSPONRANK",intCnt))
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE")
	
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		
		intRtn = mobjMDETELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag,strVOCH_TYPE)
   		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				.txtTIMCODE1.value = strTIMCODE
				.txtTIMNAME1.value = strTIMNAME
				selectRtn
			Else
				initpagedata
			End If
			DateClean
   		end if
   	end with
End Sub

Sub ProcessRtn_ALL ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE, strCLIENTNAME, strTIMCODE, strTIMNAME
	Dim strVOCH_TYPE
	
	chkcnt = 0
	strVOCH_TYPE=1
	with frmThis
		intColFlag = 0
		
		For intCnt = 1 To .sprSht_DTL.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
			'�׷��ִ밪 ����
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"TRANSCUSTRANK",intCnt))
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"CHK | TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | CLIENTSUBCODE | CLIENTSUBNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | SUBSEQ | SUBSEQNAME | TIMCODE | TIMNAME | MATTERCODE | MATTERNAME | PROGRAM | ADLOCALFLAG | WEEKDAY | DEPT_CD | DEMANDDAY | PRINTDAY | PRICE | CNT | AMT | VAT | INPUT_MEDFLAG | MED_FLAG | TAXYEARMON | TAXNO | TRANSCUSTRANK | TRANSSPONRANK | VOCH_TYPE | EXCLIENTCODE")
	
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strTIMCODE		= .txtTIMCODE1.value
		strTIMNAME		= .txtTIMNAME1.value
		
		msgbox "ProcessRtn_ALL"
		intRtn = mobjMDETELECTRANS.ProcessRtn_ALL(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag,strVOCH_TYPE)
   		if not gDoErrorRtn ("ProcessRtn_ALL") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meCLS_FLAG
			InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			
			If intRtn <> 0  Then
				.txtYEARMON1.value = strTRANSYEARMON
				.txtCLIENTCODE1.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				.txtTIMCODE1.value = strTIMCODE
				.txtTIMNAME1.value = strTIMNAME
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
				
				'DB�� ������ Ȯ�����ؾ� �������� ���������� �ߺ��������� ����ġ���� ������ ���´�.(����)
				'����� TRANSYEARMON�� TRANSNO�� DTL�� �󼼳���(�ŷ����� �߻�����)�� �����ϴ��� ��ȸ
				vntData = mobjMDETELECTRANS.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON, strTRANSNO) 
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

		'�׸����� �÷��׸� INSERT�� �ٲ۴�.  �ʱⰪ�� UPDATE...
		mobjSCGLSpr.SetFlag  .sprSht_HDR, meINS_TRANS
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | TRANSYEARMON | TRANSNO ")
		
		intRtn = mobjMDETELECTRANS.DeleteRtn(gstrConfigXml,vntData)

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
				<TBODY>
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
													<TABLE cellSpacing="0" cellPadding="0" width="95" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�ŷ����� ����</td>
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
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
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
												<TD class="SEARCHDATA" width="215"><INPUT class="INPUT" id="txtYEARMON1" title="�����ȸ" style="WIDTH: 98px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="7" name="txtYEARMON1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">������</TD>
												<TD class="SEARCHDATA" width="215"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="17" name="txtCLIENTNAME1"> <IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgCLIENTCODE1"> <INPUT class="INPUT_L" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
													width="60">��ü��</TD>
												<TD class="SEARCHDATA"><SELECT id="cmbREAL_MED_CODE1" title="��ü��" name="cmbREAL_MED_CODE1">
														<OPTION value="" selected>��ü</OPTION>
														<OPTION value="B00107">�ѱ���۱�����纻��</OPTION>
														<OPTION value="B00111">�ѱ���۱������λ�����</OPTION>
														<OPTION value="B00109">�ѱ���۱������뱸����</OPTION>
														<OPTION value="B00110">�ѱ���۱�������������</OPTION>
														<OPTION value="B00108">�ѱ���۱�����籤������</OPTION>
														<OPTION value="B00112">�ѱ���۱��������������</OPTION>
														<OPTION value="B01092">�̵��ũ������Ʈ</OPTION>
													</SELECT><SELECT id="cmbMEDGUBUN1" title="��ü����" style="WIDTH: 56px" name="cmbMEDGUBUN1">
														<OPTION value="" selected>��ü</OPTION>
														<OPTION value="01">TV</OPTION>
														<OPTION value="02">RADIO</OPTION>
														<OPTION value="10">DMB</OPTION>
													</SELECT>&nbsp;&nbsp;&nbsp;&nbsp; �������� <INPUT id="chkSPONSOR" title="����" type="checkbox" name="chkSPONSOR"></TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
														align="right" border="0" name="imgQuery">
												</td>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME1,txtSUBSEQ1)">�귣��</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME1" title="�귣���ڵ��" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="16" name="txtSUBSEQNAME1"> <IMG id="ImgSUBSEQ1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgSUBSEQ1"> <INPUT class="INPUT" id="txtSUBSEQ1" title="�귣�����ȸ" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" align="left" size="5" name="txtSUBSEQ1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTIMNAME1, txtTIMCODE1)">��</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtTIMNAME1" title="������ڵ��" style="WIDTH: 138px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="11" name="txtTIMNAME1"> <IMG id="ImgTIMCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgTIMCODE1"> <INPUT class="INPUT" id="txtTIMCODE1" title="����θ�" style="WIDTH: 53px; HEIGHT: 22px" type="text"
														maxLength="6" align="left" size="5" name="txtTIMCODE1"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME1, txtMATTERCODE1)">�����</TD>
												<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtMATTERNAME1" title="�����ڵ�" style="WIDTH: 162px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="21" name="txtMATTERNAME1"> <IMG id="ImgMATTER1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'"src="../../../images/imgPopup.gIF" align="absMiddle"
														border="0" name="ImgMATTER1"> <INPUT class="INPUT" id="txtMATTERCODE1" title="�����ڵ��" style="WIDTH: 53px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtMATTERCODE1"></TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD class="TITLE" style="HEIGHT: 8px" vAlign="absmiddle"></TD>
											</TR>
											<TR>
												<TD class="TITLE" width="210"  vAlign="middle"><span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('STANDARD')"><IMG id='btn_normal' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_normal.gif' align='absMiddle' border='0' name='btn_normal'></span>&nbsp;
																<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('EXTENTION')"><IMG id='btn_multi' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_multi.gif' align='absMiddle' border='0' name='btn_multi'></span>&nbsp;
																<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ('HIDDEN')"><IMG id='btn_hide' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/btn_hide.gif' align='absMiddle' border='0' name='btn_hide'></span>
															</TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="297" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle" width="292">�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
																<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
																	readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
															</td>
														</tr>
													</table>
												</TD>
												<td class="TITLE" style="WIDTH: 340px" vAlign="absmiddle">û���� : <INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 85px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="8" name="txtDEMANDDAY">&nbsp;<IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgCalDemandday">&nbsp;&nbsp; 
													������ : <INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 85px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="8" name="txtPRINTDAY">&nbsp;<IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="16" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0"
														name="imgCalPrintday"></td>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" alt="ȭ���� �ʱ�ȭ �մϴ�."
																	src="../../../images/imgCho.gif" border="0" name="imgCho"></TD>
															<TD><IMG id="ImgAllCustSave" onmouseover="JavaScript:this.src='../../../images/ImgAllCustSaveOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAllCustSave.gIF'"
																	height="20" alt="�����ֺ��� �ŷ������� �ϰ������մϴ�.." src="../../../images/ImgAllCustSave.gIF"
																	border="0" name="ImgAllCustSave"></TD>
															<TD><IMG id="imgDeleteALL" onmouseover="JavaScript:this.src='../../../images/imgDeleteALLOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDeleteALL.gif'"
																	height="20" alt="���õ� �ŷ������� ��ü�����մϴ�." src="../../../images/imgDeleteALL.gIF" border="0"
																	name="imgDeleteALL"></TD>
															<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																	height="20" alt="���õ� �ŷ������� �����մϴ�." src="../../../images/imgDelete.gIF" border="0"
																	name="imgDelete"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" id="TD1" style="WIDTH: 100%; HEIGHT: 3px" runat="server"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR id="tblBody1">
									<TD id="tblSheet1" style="WIDTH: 100%; HEIGHT: 30%" vAlign="top" align="center">
										<table height="100%" cellSpacing="1" cellPadding="0" width="100%" align="left" border="0">
											<tr>
												<td width="20%">
													<DIV id="pnlTab_1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
														ms_positioning="GridLayout">
														<OBJECT id="sprSht_CLIENT" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
															VIEWASTEXT DESIGNTIMEDRAGDROP="213">
															<PARAM NAME="_Version" VALUE="393216">
															<PARAM NAME="_ExtentX" VALUE="6324">
															<PARAM NAME="_ExtentY" VALUE="3519">
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
												<td align="left" width="80%">
													<DIV id="pnlTab1_2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
														align="left" ms_positioning="GridLayout">
														<OBJECT id="sprSht_HDR" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
															<PARAM NAME="_Version" VALUE="393216">
															<PARAM NAME="_ExtentX" VALUE="25400">
															<PARAM NAME="_ExtentY" VALUE="3519">
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
											</tr>
										</table>
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
												<TD align="left" width="297" height="20">
													<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td class="TITLE" vAlign="absmiddle" width="292">�ΰ����հ� : <INPUT class="NOINPUTB_R" id="txtSUMVAT" title="�ΰ����հ�" style="WIDTH: 120px; HEIGHT: 22px"
																	accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMVAT">
															</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="22">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="ImgCRE" onmouseover="JavaScript:this.src='../../../images/ImgCREOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/ImgCRE.gif'" alt="�ŷ������� �����մϴ�."
																	src="../../../images/ImgCRE.gif" border="0" name="ImgCRE"></TD>
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
								<TR id="tblBody2">
									<TD id="tblSheet2" style="WIDTH: 100%; HEIGHT: 70%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht_DTL" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="8573">
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
									<TD class="BOTTOMSPLIT" id="lblStatusDTL" style="WIDTH: 100%"></TD>
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
