<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVMEDIUMSEARCH.aspx.vb" Inherits="MD.MDCMCATVMEDIUMSEARCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��ü����</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : PROJECT ��� ȭ��(PDCMPONO)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPONO.aspx
'��      �� : ������Ʈ ��� �� ����
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/10/27 By Tae Ho Kim
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
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMMEDIUMLIST
Dim mobjMDCOGET
Dim mcomecalender, mcomecalender2
Dim mstrHIDDEN
CONST meTAB = 9
mcomecalender = FALSE
mcomecalender2 = FALSE
mstrHIDDEN = 0

'=============================
' �̺�Ʈ ���ν��� 
'=============================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------

'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("tblBody").style.display = "inline"
		Else
			document.getElementById("tblBody").style.display = "none"
		End If
		
		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
	End With
End Sub

Sub imgQuery_onclick
	gFlowWait meWAIT_ON
		SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'��� �μ��ư Ŭ���� �̺�Ʈ
Sub imgPrint99999999_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j,k
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
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
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "PD"
		ReportName = "PDCMTRANS.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjPD_TRANS.Get_TRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_TRANS_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					
					datacnt = strcntsum
					strUSERID = ""
					vntDataTemp = mobjPD_TRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
					
				End IF
			END IF
		next
		
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
		intRtn = mobjPD_TRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
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
		gSetChange
	end with
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndar,"txtFOM_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgCalEndarREQ_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		mcomecalender2 = true
		gShowPopupCalEndar frmThis.txtTO,frmThis.imgCalEndar,"txtTO_onchange()"
		mcomecalender2 = false
		gSetChange
	end with
End Sub

'-----------------------------------------------------------------------------------------
' ��ü���ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgMEDCODE_onclick
	Call MEDCODE_POP()
End Sub

'���� ������List ��������
Sub MEDCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array("", "",trim(.txtMEDCODE.value), trim(.txtMEDNAME.value))
	    
	    vntRet = gShowModalWindow("../MDCO/MDCMMEDPOP.aspx",vntInParams , 413,435)
	    
		If isArray(vntRet) Then
			If .txtMEDCODE.value = vntRet(0,0) and .txtMEDNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtMEDCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtMEDNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			'.txtREAL_MED_CODE.value = trim(vntRet(3,0))       ' �ڵ�� ǥ��
			'.txtREAL_MED_NAME.value = trim(vntRet(4,0))       ' �ڵ�� ǥ��
			
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtMEDNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetMEDCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"","", trim(.txtMEDCODE.value),trim(.txtMEDNAME.value))
			
			If not gDoErrorRtn ("GetMEDCODE") Then
				If mlngRowCnt = 1 Then
					.txtMEDCODE.value = trim(vntData(0,1))	    ' Code�� ����
					.txtMEDNAME.value = trim(vntData(1,1))       ' �ڵ�� ǥ��
					'.txtREAL_MED_CODE.value = trim(vntData(3,1))
					'.txtREAL_MED_NAME.value = trim(vntData(4,1))
				Else
					Call MEDCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' ������˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgMATTERCODE_onclick
	Call MATTERCODE_POP()
End Sub

Sub MATTERCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTNAME.value),"" , trim(.txtSUBSEQNAME.value),"", _
							trim(.txtMATTERNAME.value), "" , "B") '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMMATTERPOP.aspx",vntInParams , 780,630)
		
		If isArray(vntRet) Then
			If .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtMATTERCODE.value = trim(vntRet(0,0))	' �����ڵ� ǥ��
			.txtMATTERNAME.value = trim(vntRet(1,0))	' ����� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' �������ڵ� ǥ��
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			'.txtTIMCODE.value = trim(vntRet(4,0))		' ���ڵ� ǥ��
			'.txtTIMNAME.value = trim(vntRet(5,0))		' ���� ǥ��
			.txtSUBSEQ.value = trim(vntRet(6,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(7,0))	' �귣��� ǥ��
			'.txtEXCLIENTCODE.value = trim(vntRet(8,0))	' ���ۻ��ڵ� ǥ��
			'.txtEXCLIENTNAME.value = trim(vntRet(9,0))	' ���ۻ��ڵ� ǥ��
			'.txtDEPT_CD.value = trim(vntRet(10,0))		' �μ��ڵ� ǥ��
			'.txtDEPT_NAME.value = trim(vntRet(11,0))	' �μ��� ǥ��
			
     	End If
	End With
	gSetChange
End Sub

Sub txtMATTERNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
                              
			vntData = mobjMDCOGET.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
											trim(.txtCLIENTNAME.value),"", trim(.txtSUBSEQNAME.value), "" , _
											trim(.txtMATTERNAME.value), "" , "B")
											
			If not gDoErrorRtn ("GetMATTER") Then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntRet(0,1))	' �����ڵ� ǥ��
					.txtMATTERNAME.value = trim(vntRet(1,1))	' ����� ǥ��
					.txtCLIENTCODE.value = trim(vntRet(2,1))	' �������ڵ� ǥ��
					.txtCLIENTNAME.value = trim(vntRet(3,1))	' �����ָ� ǥ��
					.txtTIMCODE.value	 = trim(vntRet(4,1))	' ���ڵ� ǥ��
					.txtTIMNAME.value	 = trim(vntRet(5,1))	' ���� ǥ��
					.txtSUBSEQ.value	 = trim(vntRet(6,1))	' �귣�� ǥ��
					.txtSUBSEQNAME.value = trim(vntRet(7,1))	' �귣��� ǥ��
					'.txtEXCLIENTCODE.value = trim(vntRet(8,1))	' ���ۻ��ڵ� ǥ��
					'.txtDEPT_CD.value	 = trim(vntRet(10,1))	' �μ��ڵ� ǥ��
					'.txtDEPT_NAME.value	 = trim(vntRet(11,1))	' �μ��� ǥ��
				
				Else
					Call MATTERCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
' �귣���ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgSUBSEQCODE_onclick
	Call SUBSEQCODE_POP()
End Sub

Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value), trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,430)
		
		If isArray(vntRet) Then
			If .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
			.txtSUBSEQ.value = trim(vntRet(0,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(1,0))	' �귣��� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(2,0))	' ������ ǥ��
			.txtCLIENTNAME.value = trim(vntRet(3,0))	' �����ָ� ǥ��
			'.txtTIMCODE.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			'.txtTIMNAME.value = trim(vntRet(5,0))	' �����ָ� ǥ��
     	End If
	End With
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
												trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),  _
												trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			If not gDoErrorRtn ("Get_BrandInfo") Then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(0,1))
					.txtSUBSEQNAME.value = trim(vntData(1,1))
					.txtCLIENTCODE.value = trim(vntData(2,1))	' ������ ǥ��
					.txtCLIENTNAME.value = trim(vntData(3,1))	' ������
					'.txtTIMCODE.value = trim(vntData(4,1))	' �����
					'.txtTIMNAME.value = trim(vntData(5,1))	' ����
				Else
					Call SUBSEQCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'-----------------------------------------------------------------------------------------
'CIC/����� �˾�  ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

Sub CLIENTSUBCODE_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCLIENTSUBPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtCLIENTSUBCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(3,0))	'Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(4,0))	'�ڵ�� ǥ��
			
			gSetChangeFlag .txtCLIENTCODE
		End If
	end With
End Sub

Sub txtCLIENTSUBNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetCLIENTSUBCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value))
			
			If not gDoErrorRtn ("GetCLIENTSUBCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))	'Code�� ����
					.txtCLIENTNAME.value = trim(vntData(1,0))	'�ڵ�� ǥ��
					.txtCLIENTSUBCODE.value = trim(vntData(3,0))	'Code�� ����
					.txtCLIENTSUBNAME.value = trim(vntData(4,0))	'�ڵ�� ǥ��
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
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
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
Sub cmbVAT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtFROM.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************


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
		
		
		IF .txtFROM.value <> "" THEN
			.txtTO.value = strFROM
			DateClean strFROM
		END IF
	
	End With
	gSetChange
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

'****************************************************************************************
' Amt ������ ������ �հ踦 �ؽ�Ʈ�ڽ��� �ѷ��ش�
'****************************************************************************************
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")  Then
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


'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ݾ�
Sub txtFROMAMOUNT_onblur
	with frmThis
		'COMMI_RATE_Cal
		call gFormatNumber(.txtFROMAMOUNT,0,true)
	end with
End Sub

'�ݾ�
Sub txtTOAMOUNT_onblur
	with frmThis
		'COMMI_RATE_Cal
		call gFormatNumber(.txtTOAMOUNT,0,true)
	end with
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ���ֱ� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------

Sub txtFROMAMOUNT_onfocus
	with frmThis
		.txtFROMAMOUNT.value = Replace(.txtFROMAMOUNT.value,",","")
	end with
End Sub

Sub txtTOAMOUNT_onfocus
	with frmThis
		.txtTOAMOUNT.value = Replace(.txtTOAMOUNT.value,",","")
	end with
End Sub



'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
Sub cmbMED_onchange
	with frmThis
		If .cmbMED.selectedIndex = 0 Then
		End If
		'�μ�
		If .cmbMED.selectedIndex = 1 Then
		End IF
	End With
end Sub



'****************************************************************************************
'****************************************************************************************
'=============================
' UI���� ���ν��� 
'=============================
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	set mobjMDCMMEDIUMLIST = gCreateRemoteObject("cMDCO.ccMDCOMEDIUMLIST")
	set mobjMDCOGET		= gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis
	   	gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 20, 0, 3, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, " CHK|MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|MEDCODE|MEDNAME|MATTERCODE|MATTERNAME|SUBSEQ|SUBSEQNAME|TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|AMT|COMMI_RATE|VAT"
		mobjSCGLSpr.SetHeader .sprSht,        "����|��ü����|û����|û�౸��|û������|û����|��ü�ڵ�|��ü��|�����ڵ�|�����|�귡���ڵ�|�귣��|���ڵ�|CIC/��|�������ڵ�|������|����ݾ�|��������|VAT"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|       8|     8|       8|        8|     8|       0|    15|       0|    15|         0|    15|     0|    15|         0|    14|      14|      8|  6|" 
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME|MATTERNAME|SUBSEQNAME|TIMNAME|CLIENTNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|VAT",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "MEDFLAG|PUB_DATE|VOCH_TYPE|DEMANDFLAG|DEMANDDAY|MEDCODE|MEDNAME|MATTERCODE|MATTERNAME|SUBSEQ|SUBSEQNAME|TIMCODE|TIMNAME|CLIENTCODE|CLIENTNAME|AMT|COMMI_RATE|VAT"
		mobjSCGLSpr.ColHidden .sprSht, "MEDCODE | MATTERCODE | SUBSEQ | TIMCODE | CLIENTCODE", True
		.sprSht.style.visibility = "visible"
	End With
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	gEndPage
End Sub


'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'�ʱ� ������ ����
	with frmThis
		.sprSht.MaxRows = 0
		.txtFROM.value = gNowDate
		.txtTO.value  = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)	
		
		'��¥���� - ���۴��� ��������
		DateClean .txtTO.value
		
		.txtFROM.focus
	End with
End Sub

'��¥ ��ȸ���� ����
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE

	strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
	With frmThis
		.txtTO.value = date2
	End With
End Sub

'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
   	Dim i, strCols
   	Dim strFROM
   	Dim strTO
   	Dim strVOCH_TYPE
   	Dim strVOCH_TYPE2
   	Dim strMEDCODE
   	Dim strMATTERCODE
   	Dim strSUBSEQ
   	Dim	strCLIENTSUBCODE
   	Dim strCLIENTCODE
   	Dim strFROMAMOUNT
   	Dim strTOAMOUNT
   	Dim strCOMMI_RATE
   	Dim strVAT
   	
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM				= .txtFROM.value
		strTO				= .txtTO.value
		strVOCH_TYPE		= .cmbVOCH_TYPE.value
		strVOCH_TYPE2		= .cmbVOCH_TYPE2.value
		strMEDCODE			= .txtMEDCODE.value
		strMATTERCODE		= .txtMATTERCODE.value
		strSUBSEQ			= .txtSUBSEQ.value
		strCLIENTSUBCODE	= .txtCLIENTSUBCODE.value
		strCLIENTCODE		= .txtCLIENTCODE.value
		strFROMAMOUNT		= .txtFROMAMOUNT.value
		strTOAMOUNT			= .txtTOAMOUNT.value
		strCOMMI_RATE		= .cmbCOMMI_RATE.value
		strVAT				= .cmbVAT.value

		vntData = mobjMDCMMEDIUMLIST.SelectRtn_CATV(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strVOCH_TYPE,strVOCH_TYPE2,strMEDCODE,strMATTERCODE,strSUBSEQ,strCLIENTSUBCODE,strCLIENTCODE,strFROMAMOUNT,strTOAMOUNT,strCOMMI_RATE,strVAT)
		
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
			.sprSht.MaxRows = 0	
			End If
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
		
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
	AMT_SUM
End Sub



'****************************************************************************************
'****************************************************************************************
'****************************************************************************************
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
				<TR>
					<TD class="TOPSPLIT" style="WIDTH: 1040px" colSpan="2"></TD>
				</TR>
				<!--TopSplit End-->
				<!--Input Start-->
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="430" colSpan="2" height="28">
									<table cellSpacing="0" cellPadding="0" width="800" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;û�����-���γ�����ȸ <span id="spnHIDDEN" style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()">
													(�����)</span>
											</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" colSpan="2" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 600px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<td><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" align="right" border="0"
													name="imgQuery"></td>
											<td><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"></td>
											<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE style="WIDTH: 100%; HEIGHT: 90%" cellSpacing="0" cellPadding="0" align="left" border="0">
							<TR>
								<TD style="HEIGHT: 4px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD id="tblBody" style="WIDTH: 280px; HEIGHT: 91%" vAlign="top">
									<table class="DATA" id="tblKey2" style="WIDTH: 272px; HEIGHT: 302px" cellSpacing="1" cellPadding="0"
										width="272" align="left" border="0">
										<tr>
											<td class="TITLE" width="272" colSpan="2">��ü�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 202px; HEIGHT: 22px"
													accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT"></td>
										</tr>
										<tr>
											<td class="TITLE" colSpan="2">�����հ� : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 202px; HEIGHT: 22px"
													readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
											</td>
										</tr>
										<tr>
											<TD class="GROUP" colSpan="2">��ȸ����</TD>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">��ü����</TD>
											<td class="DATA" width="184"><SELECT id="cmbMED" title="��ü����" style="WIDTH: 111px" name="cmbMED">
													<OPTION value="" selected>CATV</OPTION>
												</SELECT></td>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">�Ⱓ</TD>
											<td class="DATA"><INPUT class="INPUT" id="txtFROM" title="�Ƿ��� �˻�(FROM)" style="WIDTH: 76px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarFROM1">~<INPUT class="INPUT" id="txtTO" title="�Ƿ��� �˻�(TO)" style="WIDTH: 76px; HEIGHT: 22px" accessKey="DATE"
													type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">û�౸��</TD>
											<td class="DATA"><SELECT id="cmbVOCH_TYPE" title="û�౸��" style="WIDTH: 111px" name="cmbVOCH_TYPE">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="0">����Ź</OPTION>
													<OPTION value="1">����</OPTION>
													<OPTION value="2">�Ϲ�</OPTION>
												</SELECT></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px" width="88">û������</TD>
											<td class="DATA"><SELECT id="cmbVOCH_TYPE2" title="û������" style="WIDTH: 111px" name="cmbVOCH_TYPE2">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="0">����Ź</OPTION>
													<OPTION value="2">�Ϲ�</OPTION>
												</SELECT></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMEDNAME, txtMEDCODE)"
												width="88">��ü��</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtMEDNAME" title="��ü��" style="WIDTH: 125px; HEIGHT: 22px" type="text"
													maxLength="100" size="12" name="txtMEDNAME"><IMG id="ImgMEDCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF"
													width="22" align="absMiddle" border="0" name="ImgMEDCODE"><INPUT class="INPUT_L" id="txtMEDCODE" title="��ü���ڵ�" style="WIDTH: 59px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="6" size="4" name="txtMEDCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtMATTERNAME, txtMATTERCODE)"
												width="88">�����</TD>
											<td class="DATA"><INPUT dataFld="MATTERNAME" class="INPUT_L" id="txtMATTERNAME" title="�����" style="WIDTH: 125px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="500" size="30" name="txtMATTERNAME"><IMG id="ImgMATTERCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="22" align="absMiddle" border="0"
													name="ImgMATTERCODE"><INPUT dataFld="MATTERCODE" class="INPUT_L" id="txtMATTERCODE" title="�����ڵ�" style="WIDTH: 59px; HEIGHT: 22px"
													accessKey=",M" dataSrc="#xmlBind" type="text" maxLength="10" size="4" name="txtMATTERCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 88px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME, txtSUBSEQ)"
												width="88">�귣��</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="�귣���" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" size="12" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgSUBSEQCODE"><INPUT class="INPUT_L" id="txtSUBSEQ" title="�������ڵ�" style="WIDTH: 59px; HEIGHT: 22px" type="text"
													maxLength="9" name="txtSUBSEQ"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)"
												width="83">CIC/��</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="������θ�" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" size="26" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT_L" id="txtCLIENTSUBCODE" title="������ڵ�" style="WIDTH: 59px; HEIGHT: 22px"
													type="text" maxLength="9" name="txtCLIENTSUBCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, txtCLIENTCODE)"
												width="83">������</TD>
											<td class="DATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 125px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="16" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 59px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtCLIENTCODE"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROMAMOUNT, txtTOAMOUNT)"
												width="83">����ݾ�</TD>
											<td class="DATA"><INPUT class="INPUT_R" id="txtFROMAMOUNT" title="�ݾ�" style="WIDTH: 99px; HEIGHT: 22px"
													accessKey=",M" type="text" maxLength="13" size="20" name="txtFROMAMOUNT">~<INPUT class="INPUT_R" id="txtTOAMOUNT" title="�ݾ�" style="WIDTH: 99px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="13" size="9" name="txtTOAMOUNT"></td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px" width="83">��������</TD>
											<td class="DATA"><SELECT id="cmbCOMMI_RATE" title="��������" style="WIDTH: 80px" name="cmbCOMMI_RATE">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="1">1</OPTION>
													<OPTION value="2">2</OPTION>
													<OPTION value="3">3</OPTION>
													<OPTION value="4">4</OPTION>
													<OPTION value="5">5</OPTION>
													<OPTION value="6">6</OPTION>
													<OPTION value="7">7</OPTION>
													<OPTION value="8">8</OPTION>
													<OPTION value="9">9</OPTION>
													<OPTION value="10">10</OPTION>
													<OPTION value="15">15</OPTION>
													<OPTION value="20">20</OPTION>
													<OPTION value="25">25</OPTION>
													<OPTION value="30">30</OPTION>
													<OPTION value="35">35</OPTION>
													<OPTION value="40">40</OPTION>
													<OPTION value="35">45</OPTION>
													<OPTION value="40">50</OPTION>
													<OPTION value="100">50�ʰ�</OPTION>
												</SELECT>
												(%)</td>
										</tr>
										<tr>
											<TD class="LABEL" style="WIDTH: 83px" width="83">VAT</TD>
											<td class="DATA"><SELECT id="cmbVAT" title="VAT" style="WIDTH: 111px" name="cmbVAT">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="1">����</OPTION>
													<OPTION value="01">�鼼</OPTION>
													<OPTION value="02">����</OPTION>
												</SELECT></td>
										</tr>
									</table>
								</TD>
								<td style="WIDTH: 100%; HEIGHT: 100%" vAlign="top">
									<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" ms_positioning="GridLayout">
										<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 100%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="32438">
											<PARAM NAME="_ExtentY" VALUE="20929">
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
									</DIV>
								</td>
							</TR>
							<TR>
								<TD><!--��������! ���������� ������ �Ʒ�TD�� COLSPAN �߰�--></TD>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
