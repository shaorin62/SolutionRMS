<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANSDIV.aspx.vb" Inherits="MD.MDCMELECTRANSDIV" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ �ŷ����� ���һ��� �� ����</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMELECSPONTRANS, mobjMDCMGET, mobjMDCMELECTRANSLIST
Dim mobjMDCMCODETR	
Dim mstrCheck
Dim mALLCHECK
Dim mbuttonchk
Dim mobjMDCMELECTRANS
mALLCHECK = TRUE
mstrCheck=True
mbuttonchk = true
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

Sub imgNew_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgAllSet_onclick()
	'Dim intRtn
	'intRtn = gYesNoMsgbox("��ü �ŷ������� ���� �Ͻðڽ��ϱ�?","����ȳ�!")
	'IF intRtn <> vbYes then exit Sub
	'gFlowWait meWAIT_ON
	'If ProcessRtn_All = True Then
	'ProcessAllRtn
	'End If
	'gFlowWait meWAIT_OFF ''�̻� ���� ��ư ���� 
	'���� 20081031
	
	ProcessRtn_BatchProc
End Sub
Sub ProcessRtn_BatchProc
	dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		
			
		vntInParams = ""
		vntRet = gShowModalWindow("MDCMELECTRANS.aspx",vntInParams , 1062,900)
		SelectRtn
	End with
	'gSetChange
End Sub


Function ProcessRtn_All ()

	ProcessRtn_All = False
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
   	Dim intRtn
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
	Dim vntData2
	
	with frmThis
	If .txtTRANSYEARMON.value = "" Then
		gErrorMsgBox "����� �Է��Ͻʽÿ�.","����ȳ�!"
		Exit Function
	End If
	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtTRANSYEARMON.value
		
		vntPreData = mobjMDCMELECTRANS.SelectRtn_PreCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
			if not gDoErrorRtn ("SelectRtn_PreCnt") then
				lngCnt = vntPreData(0,0)
				If lngCnt < 100 Then
					lngCnt = 1
				Else
					lngCnt = int(lngCnt/100)
					lngCnt = lngCnt+1
				End If
			End if
		
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		vntDataConfirm = mobjMDCMELECTRANS.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		
		If IngCOMMITRowCnt = 0 Then
			gErrorMsgBox strYEARMON & "���� ����ó������ �ʾҽ��ϴ�.",""
			EXIT Function	
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprShtAll, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprShtAll,meCLS_FLAG
   			end if
   			strST = strST + 100
   			strED = strED + 100
   		Next
   		for intDelCnt = .sprShtAll.MaxRows to 1 step -1				
			If mobjSCGLSpr.GetTextBinding(.sprShtAll,"YEARMON",intDelCnt) = "" Then
				mobjSCGLSpr.DeleteRow .sprShtAll,intDelCnt
			End If		
		next
   		'��ȸ �Ϸ� ����
   		End with
   ProcessRtn_All = True
End Function

Sub ProcessAllRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
	with frmThis
		strDESCRIPTION = ""
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtDEMANDDAY.value = "" Then
			msgbox "û������ �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If
		
		If .txtPRINTDAY.value = "" Then
			msgbox "�������� �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If

		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprShtAll,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'�׷� �ִ밪 ����
		intColFlag = 0
		For intCnt = 1 To .sprShtAll.MaxRows
		'�ִ밪
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprShtAll,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
   		'������ Validation
   		If .sprShtAll.MaxRows = 0 Then
   			msgbox "������ ������ �����ϴ�."
   			Exit Sub
   		End If
		'if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprShtAll,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = .txtTRANSYEARMON.value
		
		intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
   		end if
   	end with
End Sub


Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
		
	gFlowWait meWAIT_ON
	with frmThis
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		vntData = mobjMDCMELECTRANSLIST.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_ELETRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt
			strUSERID = ""
			vntDataTemp = mobjMDCMELECTRANSLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
		End IF
	
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
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub



Sub imgClose_onclick ()
	Window_OnUnload
End Sub


Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
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
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP ()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		strSPONSOR = ""
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELEC", strSPONSOR) 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			IF vntRet(3,0) = "�Ϸ�" THEN
				.txtTRANSYEARMON.value = vntRet(0,0)
				.txtTRANSNO.value = vntRet(1,0)
				.txtCLIENTCODE.value = vntRet(4,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			ELSE
				.txtTRANSYEARMON.value = vntRet(0,0)
				.txtTRANSNO.value = ""
				.txtCLIENTCODE.value = vntRet(1,0)		  ' Code�� ����
				.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			END IF
			gSetChangeFlag .txtCLIENTCODE             ' gSetChangeFlag objectID	 Flag ���� �˸�
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strSPONSOR
   		
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strSPONSOR = ""
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtCLIENTCODE.value,.txtCLIENTNAME1.value,"","trans", "ELEC", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,1)
					.txtTRANSNO.value = ""
					.txtCLIENTCODE.value = vntData(1,1)
					.txtCLIENTNAME1.value = vntData(2,1)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub	

'-----------------------------------------------------------------------------------------
' �ŷ�ó��ȣ�˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------

'�̹�����ư Ŭ����
Sub ImgTRU_onclick
	Call TRU_POP()
End Sub

Sub txtTRANSNO_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strTRANSYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strTRANSYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELEC", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)  ' Code�� ����
					.txtTRANSNO.value = vntData(1,0)  ' �ڵ�� ǥ��
					.txtCLIENTCODE.value = vntData(2,0)  ' �ڵ�� ǥ��
					.txtCLIENTNAME1.value = vntData(3,0)  ' �ڵ�� ǥ��
					'Call SelectRtn ()
				Else
					Call TRU_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRU_POP
	dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
			strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
			
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "ELEC") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = vntRet(2,0)  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = vntRet(3,0)  ' �ڵ�� ǥ��
			'Call SelectRtn ()
		end if
	End with
	gSetChange
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


'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ܰ�
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

'�ݾ�
Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

'������
Sub txtSUMAMTVAT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMTVAT,0,true)
	end with
End Sub

Sub txtTRANSYEARMON_onblur
	With frmThis
		if .txtTRANSNO.value ="" then
			If .txtTRANSYEARMON.value <> "" AND Len(.txtTRANSYEARMON.value) = 6 Then DateClean
		end if
	End With
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************


sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	dim amt1, vat1, sumamtvat
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
	if col = 15 then
		amt1 = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"AMT",Row) 
		vat1 = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"VAT",Row) 
		sumamtvat = amt1 + vat1
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMTVAT",Row, sumamtvat
		AMT_SUM1
		
	end if	
End Sub

sub sprSht_ButtonClicked(Col,Row,ButtonDown)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	with frmThis
	IF mbuttonchk THEN
		if ButtonDown then
			lngAMT = 0
			lngSUMAMT = 0
			lngTOT = 0
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) 
			lngSUMAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"REAL_MED_NAME",1)
			lngTOT = lngAMT + lngSUMAMT
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			mobjSCGLSpr.DeselectBlock .sprSht
		else
			lngAMT = 0
			lngSUMAMT = 0
			lngTOT = 0
			lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) 
			lngSUMAMT = mobjSCGLSpr.GetTextBinding(.sprSht_SUM,"REAL_MED_NAME",1)
			lngTOT = lngSUMAMT - lngAMT
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			mobjSCGLSpr.DeselectBlock .sprSht
		end if
	END IF
	end with
end sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT
	
	With frmThis
	if Col = 1 and Row = 0 then
		mbuttonchk = false
		mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
		if mstrCheck = True then 
			mstrCheck = False
			lngTOT = 0
			For i=1 to .sprSht.MaxRows
				lngAMT = 0
				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",i) 
				lngTOT = lngTOT + lngAMT
				mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, lngTOT
			Next
		elseif mstrCheck = False then 
			mstrCheck = True
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, 0
		end if
	elseif Col = 1 and Row > 0 then
		mbuttonchk = true
	end if 
	End With
End Sub  

'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM	
	End with
end sub

'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_SUM, NewTop, NewLeft
End Sub


'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt
	Dim IntAMT, IntVAT, IntSUMAMTVAT, IntAMTSUM, IntVATSUM, IntSUMAMTVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMAMTVATSUM = 0
		'����Ź �׸��� �հ�׸��� ���ֱ�
		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntVAT = 0
				IntSUMAMTVAT = 0
				
				IntAMT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntVAT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
				IntSUMAMTVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt)
				
				IntAMTSUM		= IntAMTSUM + IntAMT
				IntVATSUM		= IntVATSUM + IntVAT
				IntSUMAMTVATSUM	= IntSUMAMTVATSUM + IntSUMAMTVAT
			Next
		END IF
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, IntSUMAMTVATSUM
		end if
	End With
End Sub

'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub AMT_SUM1
	Dim lngCnt
	Dim IntAMT, IntVAT, IntSUMAMTVAT, IntAMTSUM, IntVATSUM, IntSUMAMTVATSUM
	With frmThis
		IntVATSUM = 0
		IntSUMAMTVATSUM = 0
		'����Ź �׸��� �հ�׸��� ���ֱ�
		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntVAT = 0
				IntSUMAMTVAT = 0
				
				IntVAT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
				IntSUMAMTVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt)
				
				IntVATSUM		= IntVATSUM + IntVAT
				IntSUMAMTVATSUM	= IntSUMAMTVATSUM + IntSUMAMTVAT
			Next
		END IF
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, IntSUMAMTVATSUM
		end if
	End With
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
	set mobjMELECSPONTRANS	= gCreateRemoteObject("cMDET.ccMDETELECNOSPON")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECTRANSLIST = gCreateRemoteObject("cMDET.ccMDETELECTRANSLIST")
	set mobjMDCMCODETR	= gCreateRemoteObject("cMDCO.ccMDCOCODETR")
	set mobjMDCMELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "240px"
	pnlTab1.style.left= "7px"
	
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "240px"
	pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 32, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,   "CHK |YEARMON| SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		
		mobjSCGLSpr.SetHeader .sprSht,		   "����|���|����|������|��ü��|�׷챸��|��ü��|�����|��ü�����ڵ�|��ü����|���α׷�|����|����|����ݾ�|�ΰ���|��|COMMISSION|DEPTCD|�ܰ�|Ƚ��|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|�귣���ڵ�|�귣���|������ڵ�|����θ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|   0|   0|    20|     0|       0|    14|    28|           0|       8|       0|   0|   0|      12|    12|13|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |0    |0         |16      |0         |15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " CLIENTNAME|REAL_MED_NAME|GFLAGNAME|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME|ATTR02", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|CLIENTNAME|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		'�հ� ǥ�� �׸��� �⺻ȭ�� ����
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 32, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "CHK|YEARMON | SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetText .sprSht_SUM, 1, 1, "�հ�"
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"1|1",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|VAT|SUMAMTVAT|REAL_MED_NAME", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_SUM, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
	    
		'******************************************************************
		'�ŷ����� ��ȸ �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 28, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht1, "TRANSYEARMON | TRANSNO | SEQ |CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|PROGRAM |ADLOCALFLAG|WEEKDAY|DEPT_CD|DEMANDDAY|PRINTDAY| PRICE| CNT|AMT| VAT| MED_FLAG|ATTR02|TAXYEARMON|TAXNO|TRUST_SEQ|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME"
		mobjSCGLSpr.SetHeader .sprSht1,		  "TRANSYEARMON|TRANSNO|SEQ|CLIENTCODE|CLIENTNAME|MEDCODE|��ü��|REAL_MED_CODE|��ü��|���α׷���|����|��ۿ���|DEPT_CD|DEMANDDAY|��������|���ް���|ȸ��|������|�ΰ���|��ü����|�����|���ݰ�꼭���|���ݰ�꼭��ȣ|��Ź����|�귣���ڵ�|�귣���|������ڵ�|����θ�"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "         0|	  0|  0|         0|	        0|      0|    10|            0|    10|        10|   6|       8|      0|        0|       8|       9|   5|       9|     9|       9|    15|0             |0             |0       |0         |12      |0         |12"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "PRICE | AMT| VAT|CNT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "MEDNAME|REAL_MED_NAME|PROGRAM|ADLOCALFLAG|WEEKDAY|MED_FLAG|ATTR02|BRANDNAME|CLIENTSUBNAME", -1, -1, 50
		mobjSCGLSpr.ColHidden .sprSht1, "TRANSYEARMON|TRANSNO | SEQ|CLIENTCODE|CLIENTNAME|MEDCODE|REAL_MED_CODE|DEPT_CD|DEMANDDAY|TAXYEARMON|TAXNO|TRUST_SEQ|SUBSEQ|CLIENTSUBCODE", true
		'******************************************************************
		'��ü �ŷ����� ������ Hidden �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprShtAll
		mobjSCGLSpr.SpreadLayout .sprShtAll, 30, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprShtAll,   "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetHeader .sprShtAll,		   "YEARMON|SEQ|������|MEDNAME| ��ü��|INPUT_MEDFLAG|��ü����|PROGRAM|ADLOCALFLAG|WEEKDAY|����ݾ�|�ΰ���|��|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|�����|GFLAG|SUBSEQ|�귣���|CLIENTSUBCODE|����θ�"
		mobjSCGLSpr.SetColWidth .sprShtAll, "-1", "	  0|  0|    20|      0|     20|            0|       8|      0|          0|      0|      10|    10|10|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |19    |0    |0     |13      |0            |12"
		mobjSCGLSpr.SetRowHeight .sprShtAll, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprShtAll, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprShtAll, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprShtAll, " CLIENTNAME|REAL_MED_NAME|ATTR02|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprShtAll, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprShtAll, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
    End With    
	pnlTab1.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
	'			case 1 : .txtCLIENTCODE.value = vntInParam(i)
	'			case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'��ȸ�߰��ʵ�
	'			case 3 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
	'			case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
	'			case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
	'		end select
	'	next
	.txtTRANSYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	end with
	'SelectRtn
End Sub

Sub EndPage()
	set mobjMELECSPONTRANS = Nothing
	set mobjMDCMGET = Nothing
	set mobjMDCMCODETR = Nothing
	Set mobjMDCMELECTRANS = Nothing
	
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean
		.txtPRINTDAY.value  = gNowDate
		.sprSht.MaxRows = 0	
		.sprSht1.MaxRows = 0	
		
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"REAL_MED_NAME",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"PRICE",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, ""
		mobjSCGLSpr.SetTextBinding .sprSht_SUM,"CNT",1, ""
	
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
	Clean_Hdr	
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
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim intRtnCf
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strPROGNAME
	Dim strTRANSNO
	
	chkcnt = 0
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "�������� �ʼ� �Է� ���� �Դϴ�.",""
			Exit Sub
		End If

		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ������� ������ �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if

		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'�׷� �ִ밪 ����
		intColFlag = 0
		For intCnt = 1 To .sprSht.MaxRows
		'�ִ밪
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
   		'������ Validation
   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		'��Ʈ�� ����� �����͸� �����´�.
		intRtnCf = gYesNoMsgbox("û������ Ȯ���ϼ̽��ϱ�?","����ȳ�!")
		IF intRtnCf <> vbYes then exit Sub
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON | SEQ | CLIENTNAME | MEDNAME | GFLAGNAME |REAL_MED_NAME  |ATTR02| INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMAMTVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = MID(.txtDEMANDDAY.value,1,4) &  MID(.txtDEMANDDAY.value,6,2) 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strREAL_MED_CODE	= .cmbREAL_MED_CODE.value
		'strPROGNAME		= .txtPROGNAME.value
		strTRANSNO		= 0
		
		
		intRtn = mobjMELECSPONTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		'if not gDoErrorRtn ("ProcessRtn") then
		'	'��� �÷��� Ŭ����
		'	mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
		'	gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
		'	InitPageData
		'	PreSearchFiledValue strTRANSYEARMON, strTRANSNO, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strPROGNAME
   		'end if
   		
   		''''''''''''''''''''''''''
   		'intRtn = mobjMDCMINTERNETTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)
   		
   		if not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			If intRtn <> 0  Then
				.txtTRANSYEARMON.value = strTRANSYEARMON
				.txtTRANSNO.value = intTRANSNO
				.txtCLIENTCODE.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
				selectRtn
			Else
				initpagedata
			End If
   		end if
   		
   		
   		'''''''''''''''''''''''''
   	end with
End Sub


'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntData1
	Dim strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME, strTRANSNO
	Dim strPRINTDAY
	Dim strREAL_MED_CODE, strREAL_MED_NAME, strPROGNAME
   	Dim i, strCols
	Dim strMEDGUBUN
	'On error resume next
	with frmThis

		If .txtTRANSYEARMON.value = "" Then
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		strTRANSNO = ""
		strTRANSNO = .txtTRANSNO.value
		IF strTRANSNO = "" THEN
			IF  .txtCLIENTCODE.value = "" THEN
				gErrorMsgBox "��ȸ�� �������ڵ�� �ݵ�� �־�� �մϴ�.",""
				Exit SUb
			END IF
			IF  .cmbREAL_MED_CODE.value = "" THEN
				gErrorMsgBox "��ȸ�� ��ü���ڵ�� �ݵ�� �־�� �մϴ�.",""
				Exit SUb
			END IF
		END IF
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.sprSht1.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strTRANSYEARMON	= .txtTRANSYEARMON.value
		strTRANSNO		= .txtTRANSNO.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strREAL_MED_CODE	= .cmbREAL_MED_CODE.value
		strMEDGUBUN = .cmbMEDGUBUN.value
		'strPROGNAME = .txtPROGNAME.value
		
		
		Dim strSUBSEQ
		Dim strSUBSEQNAME
		Dim strCLIENTSUBCODE
		Dim strCLIENTSUBNAME
		Dim strMATTERCODE
		Dim strMATTERNAME
		
		strSUBSEQ = .txtSUBSEQ.value 
		strSUBSEQNAME = .txtSUBSEQNAME.value 
		strCLIENTSUBCODE = .txtCLIENTSUBCODE.value 
		strCLIENTSUBNAME = .txtCLIENTSUBNAME.value 
		strMATTERCODE = .txtMATTERCODE.value 
		strMATTERNAME = .txtMATTERNAME.value 
		
		'strSUBSEQ,strSUBSEQNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strMATTERCODE,strMATTERNAME
		'msgbox strTRANSYEARMON & "��ȣ" & strTRANSNO & "�������ڵ�" & strCLIENTCODE & "��ü���ڵ�" & strREAL_MED_CODE
		'exit sub
		IF strTRANSNO <> "" THEN
			'InitPageData
			IF not SelectRtn_HDR (strTRANSYEARMON, strTRANSNO, strCLIENTCODE) Then Exit Sub
			
			pnlTab1.style.visibility = "HIDDEN"
			pnlTab2.style.visibility = "visible"
						
			.txtDEMANDDAY.readOnly = "TRUE"
			.txtDEMANDDAY.className = "NOINPUT"	
			.imgCalDemandday.disabled = True		

			'��Ʈ ��ȸ
			Call SelectRtn_DTL (strTRANSYEARMON, strTRANSNO, strCLIENTCODE)
			
			PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
			.txtCLIENTSUBNAME.value = "" 
			.txtCLIENTSUBCODE.value = ""
			.txtMATTERNAME.value = ""
			.txtMATTERCODE.value = ""
			.txtSUBSEQNAME.value = ""
			.txtSUBSEQ.value = ""		
		ELSE
			'InitPageData
			'�̻��� ����
			vntData = mobjMELECSPONTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, strCLIENTCODE, strREAL_MED_CODE, strSUBSEQ,strSUBSEQNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strMATTERCODE,strMATTERNAME, strMEDGUBUN)
			
			pnlTab1.style.visibility = "visible"
			pnlTab2.style.visibility = "HIDDEN"
			
			.txtDEMANDDAY.readOnly = "FALSE"
			.txtDEMANDDAY.className = "INPUT"
			.imgCalDemandday.disabled = FALSE
			
			if not gDoErrorRtn ("SelectRtn") then
				if mlngRowCnt > 0 then
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   					PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
   				else
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   					InitPageData
   					PreSearchFiledValue strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE
   				end if
   				Clean_Hdr
   			end if
		END IF
		DateClean
		AMT_SUM
   	end with
End Sub
Sub Clean_Hdr
with frmThis
.txtCLIENTNAME.value = ""
.txtAMT.value = ""
.txtVAT.value = ""
.txtSUMAMTVAT.value = ""
End with
End Sub
Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strCLIENTCODE )
	dim vntData
	on error resume next

	'�ʱ�ȭ
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMELECSPONTRANS.Get_ELECTRANS_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ �ŷ�����ȣ�� ���Ͽ�" & meNO_DATA, ""
			Clean_Hdr
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			gWriteText "", "������ �ŷ�����ȣ�� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			
			SelectRtn_HDR = True
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByVal strYEARMON,ByVal strTRANSNO, ByVal strCLIENTCODE)
	Dim vntData
	on error resume next

	'�ʱ�ȭ
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMELECSPONTRANS.Get_ELECTRANS_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_LIST") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG

		SelectRtn_DTL = True
		
		gWriteText "", "������ �ŷ�����ȣ���� �󼼳����� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME,strREAL_MED_CODE)
' strTRANSYEARMON, strTRANSNO,strCLIENTCODE, strCLIENTNAME, strMATTERCODE,strMATTERNAME,strSUBSEQ,strSUBSEQNAME
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtTRANSNO.value = strTRANSNO
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
	frmThis.txtMATTERCODE.value =strMATTERCODE
	frmThis.txtMATTERNAME.value = strMATTERNAME
	frmThis.txtSUBSEQ.value = strSUBSEQ
	frmThis.txtSUBSEQNAME.value = strSUBSEQNAME
	frmThis.cmbREAL_MED_CODE.value = strREAL_MED_CODE
	
	
	'frmThis.txtPROGNAME.value = strPROGNAME
	'frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
	
	
End Sub

'****************************************************************************************
' �ŷ����� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	with frmThis
		IF .sprSht1.MaxRows = 0 THEN
			gErrorMsgBox "������ ���� �󼼳����� �����ϴ�.","�����ȳ�!"
			Exit Sub
		END IF
		
		For intCnt2 = 1 To .sprSht1.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(.sprSht1,"TAXNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "���ݰ�꼭��ȣ�� �����ϴ� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
				Exit Sub
			End If
		Next
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		intCnt = 0
		
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		'mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|PROGRAM |ADLOCALFLAG|WEEKDAY|DEPT_CD|DEMANDDAY|PRINTDAY| PRICE| CNT|AMT| VAT| MED_FLAG|ATTR02|TAXYEARMON|TAXNO|TRUST_SEQ")
		
		'���õ� �ڷḦ ������ ���� ����
		strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
	
		intRtn = mobjMELECSPONTRANS.DeleteRtn(gstrConfigXml,vntData, strTRANSYEARMON, strTRANSNO)

		IF not gDoErrorRtn ("DeleteRtn") then
			for i = .sprSht1.MaxRows to 1 step -1
				mobjSCGLSpr.DeleteRow .sprSht1,i
			next
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", strTRANSYEARMON & "-" & strTRANSNO & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht1
		initpagedata
		.txtDEMANDDAY.readOnly = false
		.txtDEMANDDAY.className = "INPUT"
		'SelectRtn
	End with
	err.clear	
End Sub
'-----------------------------------------------------------------------------------------
' ������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTSUBCODE_POP

	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME1.value = trim(vntRet(6,0))
			.txtMATTERNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTSUBNAME_onkeydown

	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value))
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME1.value = trim(vntData(6,0))
					
					.txtMATTERNAME.focus()
					gSetChangeFlag .txtCLIENTSUBCODE
				Else
					Call CLIENTSUBCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' �귣���ڵ��˾� ��ư
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgSUBSEQ_onclick
	Call SUBSEQCODE_POP()
End Sub


Sub SUBSEQCODE_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTSEQPOP.aspx",vntInParams , 520,440)
		if isArray(vntRet) then
			if .txtSUBSEQ.value = vntRet(0,0) and .txtSUBSEQNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			.txtSUBSEQ.value = trim(vntRet(1,0))		' �귣�� ǥ��
			.txtSUBSEQNAME.value = trim(vntRet(2,0))	' �귣��� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(3,0))	' �������ڵ� ǥ��
			.txtCLIENTNAME1.value = trim(vntRet(4,0))	' �����ָ� ǥ��
			.txtCLIENTSUBCODE.value = trim(vntRet(7,0))	' ������ڵ� ǥ��
			.txtCLIENTSUBNAME.value = trim(vntRet(8,0))	' ����θ� ǥ��
			
			.txtMATTERNAME.focus()					' ��Ŀ�� �̵�
			gSetChangeFlag .txtSUBSEQ		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

Sub txtSUBSEQNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetDEPT_CDBYCUSTSEQList(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtSUBSEQ.value),trim(.txtSUBSEQNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value))
			if not gDoErrorRtn ("GetDEPT_CDBYCUSTSEQList") then
				If mlngRowCnt = 1 Then
					.txtSUBSEQ.value = trim(vntData(1,0))
					.txtSUBSEQNAME.value = trim(vntData(2,0))
					.txtCLIENTCODE.value = trim(vntRet(3,0))		' �������ڵ� ǥ��
					.txtCLIENTNAME1.value = trim(vntRet(4,0))	' �����ָ� ǥ��
					.txtCLIENTSUBCODE.value = trim(vntRet(7,0))		' ������ڵ� ǥ��
					.txtCLIENTSUBNAME.value = trim(vntRet(8,0))	' ����θ� ǥ��
					.txtMATTERNAME.focus()
				Else
					Call SUBSEQCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'-----------------------------------------------------------------------------------------
' �����ڵ��˾� ��ư
'-----------------------------------------------------------------------------------------
'������ ��������������
Sub ImgMATTER_onclick
	Call MATTERCODE_POP()
	
End Sub

Sub MATTERCODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		
		vntInParams = array(trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("MDCMMATTERPOP.aspx",vntInParams , 783,473)
		if isArray(vntRet) then
			if .txtMATTERCODE.value = vntRet(0,0) and .txtMATTERNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			
			.txtMATTERCODE.value = trim(vntRet(0,0))	
			.txtMATTERNAME.value = trim(vntRet(1,0))	
			
			.txtCLIENTSUBCODE.value = trim(vntRet(4,0))
			.txtCLIENTSUBNAME.value =  trim(vntRet(5,0))
			.txtSUBSEQ.value = trim(vntRet(6,0))
			.txtSUBSEQNAME.value = trim(vntRet(7,0))
			
			'gSetChangeFlag .txtMATTERCODE
			
			.txtCLIENTCODE.value = trim(vntRet(2,0))
			.txtCLIENTNAME1.value = trim(vntRet(3,0))
     	end if
	End with
	'gSetChange
	
End Sub

Sub txtMATTERNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			'vntData = mobjMDCMCODETR.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtMATTERCODE.value),trim(.txtPROGNAME.value))
			vntData = mobjMDCMCODETR.GetMATTER(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtMATTERCODE.value), trim(.txtMATTERNAME.value),trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value), trim(.txtSUBSEQ.value), trim(.txtSUBSEQNAME.value),"", "")
			if not gDoErrorRtn ("GetMATTER") then
				If mlngRowCnt = 1 Then
					.txtMATTERCODE.value = trim(vntData(0,1))		' �귣�� ǥ��
					.txtMATTERNAME.value = trim(vntData(1,1))	' �귣��� ǥ�� 2,3,6,7
					.txtCLIENTCODE.value = trim(vntData(2,1))
					.txtCLIENTNAME1.value = trim(vntData(3,1))
					.txtCLIENTSUBCODE.value = trim(vntData(4,1))
					.txtCLIENTSUBNAME.value =  trim(vntData(5,1))
					.txtSUBSEQ.value = trim(vntData(6,1))
					.txtSUBSEQNAME.value = trim(vntData(7,1))
				Else
					Call MATTERCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'��ȣ�� Ŭ�����Ѵ�.
Sub CleanField (objField1, objField2, objField3)
	if isobject(objField1) then objField1.value = ""
	if isobject(objField2) then objField2.value = ""
	if isobject(objField3) then objField3.value = ""
	'InitPageData
End Sub
-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
				border="0">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;����Ź �ŷ����� ����</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 342px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End-->
									<!--Common Button Start--></TD>
							</TR>
							<!--Top Define Table End-->
							<!--Input Define Table End--></TABLE>
						<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)"
												width="90">��&nbsp;&nbsp;��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 347px" width="347"><INPUT class="INPUT" id="txtTRANSYEARMON" title="�ŷ������" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey="NUM" type="text" maxLength="6" size="6" name="txtTRANSYEARMON"><IMG id="ImgTRU" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTRU"><INPUT class="INPUT" id="txtTRANSNO" title="�ŷ�����ȣ" style="WIDTH: 72px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="6" name="txtTRANSNO"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="90">��ü��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 427px"><SELECT id="cmbREAL_MED_CODE" title="��ü��" name="cmbREAL_MED_CODE">
													<OPTION value="B00140" selected>�ѱ���۱�����纻��</OPTION>
													<OPTION value="B00144">�ѱ���۱������λ�����</OPTION>
													<OPTION value="B00142">�ѱ���۱������뱸����</OPTION>
													<OPTION value="B00143">�ѱ���۱�������������</OPTION>
													<OPTION value="B00141">�ѱ���۱�����籤������</OPTION>
													<OPTION value="B00145">�ѱ���۱��������������</OPTION>
												</SELECT>
												<SELECT id="cmbMEDGUBUN" title="��ü����" style="WIDTH: 80px" name="cmbMEDGUBUN">
													<OPTION value="" selected>��ü</OPTION>
													<OPTION value="TV">TV</OPTION>
													<OPTION value="RD">RADIO</OPTION>
													<OPTION value="DMB">DMB</OPTION>
												</SELECT>
											</TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call CleanField(txtCLIENTNAME1, txtCLIENTCODE, txtTRANSNO) ">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="�������ڵ���ȸ" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBNAME, txtCLIENTSUBCODE)">�����
											</TD>
											<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="������ڵ��" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTSUBCODE"><INPUT class="INPUT" id="txtCLIENTSUBCODE" title="����θ�" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtCLIENTSUBCODE"></TD>
										</tr>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUBSEQNAME, txtSUBSEQ)">�귣��&nbsp;
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUBSEQNAME" title="�귣���ڵ��" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtSUBSEQNAME"><IMG id="ImgSUBSEQ" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgSUBSEQ"><INPUT class="INPUT" id="txtSUBSEQ" title="�귣�����ȸ" style="HEIGHT: 22px" type="text" maxLength="6"
													align="left" size="5" name="txtSUBSEQ"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtMATTERNAME, txtMATTERCODE)">�����
											</TD>
											<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtMATTERNAME" title="�����ڵ�" style="HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="30" name="txtMATTERNAME"><IMG id="ImgMATTER" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgMATTER"><INPUT class="INPUT" id="txtMATTERCODE" title="�����ڵ��" style="HEIGHT: 22px" type="text"
													maxLength="6" align="left" size="5" name="txtMATTERCODE"></TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"><FONT face="����"></FONT></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;�ŷ����� ����</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="imgAllSet" onmouseover="JavaScript:this.src='../../../images/imgAllSetOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAllSet.gIF'"
																height="20" alt="�ش��� �� �ŷ������� ��ü �����մϴ�." src="../../../images/imgAllSet.gIF" border="0"
																name="imgAllSet"></td>
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgTransCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTransCre.gIF'"
																height="20" alt="���õ� �ŷ����� ������ �����մϴ�." src="../../../images/imgTransCre.gIF" border="0"
																name="imgSave"></td>
														<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></td>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
									<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="right" border="0">
										<TR>
											<TD class="LABEL" width="90">������</TD>
											<TD class="DATA" width="256"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="37" name="txtCLIENTNAME">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"
												width="90">��������</TD>
											<TD class="DATA" width="257"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="���μ���" style="WIDTH: 213px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="30" name="txtPRINTDAY"><IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
													name="imgCalPrintday">
											</TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')"
												width="90">û������</TD>
											<TD class="DATA" width="257"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 210px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="100" size="29" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></TD>
										</TR>
										<TR>
											<TD class="LABEL">���ް���</TD>
											<TD class="DATA"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="����ݾ�" style="WIDTH: 257px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="37" name="txtAMT">
											</TD>
											<TD class="LABEL">�ΰ�����</TD>
											<TD class="DATA"><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="�ΰ���" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="37" name="txtVAT"></TD>
											<TD class="LABEL">��</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMTVAT" class="NOINPUT_R" id="txtSUMAMTVAT" title="��" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="37" name="txtSUMAMTVAT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				<!--BodySplit Start-->
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
				</TR>
				<!--BodySplit End-->
				<!--List Start-->
				<TR>
					<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 604px" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht" style="WIDTH: 1040px; HEIGHT: 580px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
								<PARAM NAME="_ExtentY" VALUE="15346">
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
							<OBJECT id="sprSht_SUM" style="WIDTH: 1038px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27464">
								<PARAM NAME="_ExtentY" VALUE="635">
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
						<DIV id="pnlTab2" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
							<OBJECT id="sprSht1" style="WIDTH: 1040px; HEIGHT: 604px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
								<PARAM NAME="_ExtentY" VALUE="15981">
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
						<DIV id="pnlTab3" style="POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout"><!--VISIBILITY: hidden;-->
							<OBJECT id="sprShtAll" style="WIDTH: 1040px; HEIGHT: 0px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="27517">
								<PARAM NAME="_ExtentY" VALUE="0">
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
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
				</TR>
				<!--Bottom Split End--></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TABLE> 
			<!--Main End--></FORM>
		</TR></TABLE>
	</body>
</HTML>
