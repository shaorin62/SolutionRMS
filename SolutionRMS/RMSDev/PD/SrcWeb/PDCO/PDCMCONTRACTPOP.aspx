<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACTPOP.aspx.vb" Inherits="PD.PDCMCONTRACTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��༭ ��� �� Ȯ��</title>
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
Dim mobjPDCMCONTRACT, mobjPDCMGET
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
mALLCHECK = TRUE
mstrCheck=True
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgFind_onclick()
Dim vntRet
	vntRet = gShowModalWindow("PDCMCHARGELISTPOP.aspx","" , 1060,730)
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
Dim intCnt
Dim lngCnt
Dim lngSumCnt
	with frmThis
	
	lngCnt = 0
	lngSumCnt = 0
	For intCnt = 1 To .sprSht.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1"  Then
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CONFIRMFLAG", intCnt) = "Y" Then
				gErrorMsgBox intCnt & " ���� �˼� Ȯ�� �����Դϴ�.Ȯ�γ����� �����Ͽ��ֽʽÿ�.","�����ȳ�"
				Exit Sub
			End If
			lngCnt = 1
			lngSumCnt = lngSumCnt + lngCnt
		End if
	Next
	If lngSumCnt = 0 Then
		gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�����ȳ�"
		Exit Sub
	End If
	End with
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
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
	Dim strUSERID
	Dim intCnt2
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht1.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	For intCnt2 = 1 To frmThis.sprSht1.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXNO",intCnt2) <> "" THEN
			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "���ݰ�꼭��ȣ�� �����ϴ� ������ ������� �� �����ϴ�.","�μ�ȳ�!"
			Exit Sub
		End If
	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�.
		'md_trans_temp���� ����
		intRtn = mobjPDCMCONTRACT.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		
		vntData = mobjPDCMCONTRACT.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
	
		strcntsum = 0
		IF not gDoErrorRtn ("Get_CATVTRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt
			
			for i=1 to 3
				strUSERID = ""
				vntDataTemp = mobjPDCMCONTRACT.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
				
		window.setTimeout "printSetTimeout", 10000
	
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMCATVTRANS.DeleteRtn_temp(gstrConfigXml)
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
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFrom.value = date1
		.txtTo.value = date2
	End With
End Sub

Sub STEDClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtSTDATE.value = date1
		.txtEDDATE.value = date2
	End With
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'�˻����� ������
Sub imgFrom_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtFrom,.imgFrom,"txtFrom_onchange()"
		gSetChange
	end with
End Sub

Sub txtFrom_onchange
	gSetChange
End Sub

'�˻����� ������
Sub imgTo_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtTo,.imgTo,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

Sub imgSTDATE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtSTDATE,.imgSTDATE,"txtSTDATE_onchange()"
		gSetChange
	end with
End Sub

Sub txtSTDATE_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value  = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtCONTRACTDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value  = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgEDDATE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtEDDATE,.imgEDDATE,"txtEDDATE_onchange()"
		gSetChange
	end with
End Sub

Sub txtEDDATE_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgDELIVERYDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtDELIVERYDAY,.imgDELIVERYDAY,"txtDELIVERYDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtDELIVERYDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVERYDAY",frmThis.sprSht.ActiveRow, frmThis.txtDELIVERYDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub imgTESTDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtTESTDAY,.imgTESTDAY,"txtTESTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtTESTDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTDAY",frmThis.sprSht.ActiveRow, frmThis.txtTESTDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'txtLOCALAREA,txtAMT,txtTESTMENT,txtCOMENT

Sub txtLOCALAREA_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"LOCALAREA",frmThis.sprSht.ActiveRow, frmThis.txtLOCALAREA.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtAMT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtAMT_onfocus
	with frmThis
		.txtAMT.value = Replace(.txtAMT.value,",","")
	end with
End Sub
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

Sub txtTESTMENT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTMENT",frmThis.sprSht.ActiveRow, frmThis.txtTESTMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtCOMENT_Onchange

	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMENT",frmThis.sprSht.ActiveRow, frmThis.txtCOMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtPAYMENTGBN_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T" Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PAYMENTGBN",frmThis.sprSht.ActiveRow, frmThis.txtPAYMENTGBN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
Sub cmbENDGBN_onchange
'txtCONTRACTNO,cmbTEST
	with frmThis
		If .cmbENDGBN.value = "T" Then
			.txtCONTRACTNO.style.visibility = "visible"
			.cmbTEST.style.visibility = "visible"
			.txtJOBNO.style.visibility = "hidden"
			.txtJOBNAME.style.visibility = "hidden"
			.ImgJOBNO.style.visibility = "hidden"
		Elseif  .cmbENDGBN.value = "F" Then
			.txtCONTRACTNO.style.visibility = "hidden"
			.cmbTEST.style.visibility = "hidden"
			.txtJOBNO.style.visibility = "visible"
			.txtJOBNAME.style.visibility = "visible"
			.ImgJOBNO.style.visibility = "visible"
		Elseif  .cmbENDGBN.value = "" Then
			.txtCONTRACTNO.style.visibility = "visible"
			.cmbTEST.style.visibility = "hidden"
			.txtJOBNO.style.visibility = "visible"
			.txtJOBNAME.style.visibility = "visible"
			.ImgJOBNO.style.visibility = "visible"
		End If
	End with
	SelectRtn
End Sub
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i


	'����������ü ����	
	set mobjPDCMCONTRACT	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "380px"
	pnlTab1.style.left= "7px"
	
	

	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    Input_Layout
	pnlTab1.style.visibility = "visible"
	frmThis.txtCONTRACTNO.style.visibility = "hidden"
	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	'�̰��� �Ķ���� �ޱ�
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	with frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtCONTRACTNO.value = vntInParam(i)	'CC Code or Name
			end select
		next
	.cmbENDGBN.selectedIndex = 0 
	cmbENDGBN_onchange
	
	
	Call gCleanField("txtFrom","txtTo")
	End with
	 
	SelectRtn
	
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

Sub Input_Layout
	gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONTRACTNO|OUTSCODE|OUTSNAME|JOBNO|JOBNAME|ADJAMT|JOBGUBN|CREPART|RANKTRANS|SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|��༭��ȣ|����ó�ڵ�|����ó|JOBNO|JOB��|�ݾ�|���ۺι�|���ۺз�|��ũ|����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 6  |10        |0         |30    |12   |30   |14  |15      |15      |0   |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " OUTSCODE|OUTSNAME|JOBNO|JOBNAME|RANKTRANS|JOBGUBN|CREPART|CONTRACTNO", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"OUTSCODE|OUTSNAME|JOBNO|JOBNAME|JOBGUBN|CREPART|ADJAMT"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE|RANKTRANS|SEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO|OUTSCODE|JOBGUBN|CREPART|CHK",-1,-1,2,2,false
		'mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|ITEMNAME",-1,-1,0,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSNAME|CONTRACTNO"
		.txtOUTSCODE.style.visibility = "hidden"
	    		
    End With    
End Sub

Sub Select_Layout
	Dim strComboList
	gSetSheetDefaultColor() 
	With frmThis
		strComboList =  "��༭ ��Ȯ��" & vbTab & "��༭ Ȯ��"
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONTRACTNO|CONTRACTNAME|CONTRACTDAY|LOCALAREA|STDATE|EDDATE|AMT|DELIVERYDAY|TESTDAY|PAYMENTGBN|TESTMENT|COMENT|OUTSCODE|CONFIRMFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		"����|��༭��ȣ|����|�����|��ǰ���|�뿪������|�뿪������|���ݾ�|��ǰ��|�˼���|������޹��|�˼����|Ư�����|����ó�ڵ�|��༭Ȯ��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "6|10        |18    |8     |13      |10        |10        |12      |9     |9     |9           |9       |10     |0         |13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "LOCALAREA|PAYMENTGBN|TESTMENT|COMENT|CONFIRMFLAG|CONTRACTNO", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "STDATE|EDDATE|DELIVERYDAY|TESTDAY|CONTRACTDAY"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DELIVERYDAY|TESTDAY|CONTRACTDAY"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE", true
		mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO|TESTDAY|PAYMENTGBN", false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK",-1,-1,2,2,false
	    mobjSCGLSpr.SetCellAlign2 .sprSht, "CONTRACTNAME",-1,-1,0,2,false
		'mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSNAME"
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,15,15,-1,-1,strComboList
		mobjSCGLSpr.CellGroupingEach .sprSht,"CONTRACTNAME|LOCALAREA",,,,0
		
	    		
    End With    
End Sub
'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		DateClean
		STEDClean
		.txtDELIVERYDAY.value = gNowDate
		.txtTESTDAY.value = gNowDate
		.txtCONTRACTDAY.value = gNowDate
		.txtLOCALAREA.value = "�������� �÷���(��) ����"
		'.txtCONTRACTNO.style.visibility = "hidden"
		.cmbTEST.style.visibility = "hidden"
		.txtTESTMENT.value  = ""
		.txtCOMENT.value  = ""
		.txtPAYMENTGBN.value = ""
		.txtAMT.value  = 0
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub
'****************************************************************************************
' �̺�Ʈ ó��
'****************************************************************************************
Sub sprSht_Change(ByVal Col, ByVal Row)
	
	Dim intCnt
	Dim lngAMT
	Dim lngSUMAMT
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
	if Col = 1 Then
		lngAMT = 0
		lngSUMAMT = 0
		
		For intCnt = 1 To frmThis.sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1" And frmThis.cmbENDGBN.value = "F" Then
				lngAMT = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJAMT", intCnt)		
				lngSUMAMT = lngSUMAMT + lngAMT
			End if
		Next
		frmThis.txtAMT.value = lngSUMAMT
		txtAMT_onblur
	End if
End Sub
Sub sprSht_Click(ByVal Col, ByVal Row)
	
	dim intcnt
	with frmThis
		if .cmbENDGBN.value = "" then
			exit Sub
		End if
		If Row = 0 and Col = 1  then 
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
				
			next
			For intCnt = 1 To .sprSht.MaxRows
				If  .cmbENDGBN.value = "" Then
					'����ƽ
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				End If			
			Next
		Elseif Row > 0 and Col > 0 then
			If .cmbENDGBN.value  = "T" Then
			sprShtToFieldBinding Col,Row
			End IF
		end if
		If .cmbENDGBN.value = "F" then
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
		End If
	end with
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
			.txtLOCALAREA.value = mobjSCGLSpr.GetTextBinding(.sprSht,"LOCALAREA",Row)
			.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
			.txtEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			.txtDELIVERYDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYDAY",Row)
			.txtTESTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTDAY",Row)
			.txtPAYMENTGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PAYMENTGBN",Row)
			.txtTESTMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTMENT",Row)
			.txtCOMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMENT",Row)
			.txtOUTSCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",Row)
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
			.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
		If .txtAMT.value <> "" Then
			txtAMT_onblur
		End If
	End with
End Function
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub
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
		vntInParams = array(trim(.txtOUTSCODE1.value), trim(.txtOUTSNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtOUTSCODE1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
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
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE1.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE1.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
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
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
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
' ��������ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strGBN
	Dim strOUTSCODE
	Dim strOUTSNAME
	Dim strFROM
	Dim strTO
	Dim strJOBNO
	Dim strJOBNAME
	Dim vntData
	Dim intCnt
	Dim strCONFIRM
	Dim strCONTRACTNO
	'On error resume next
	with frmThis
		.sprSht.MaxRows = 0
		strGBN = .cmbENDGBN.value 
		strOUTSCODE = TRIM(.txtOUTSCODE1.value)
		strOUTSNAME =  TRIM(.txtOUTSNAME.value)
		strJOBNO = TRIM(.txtJOBNO.value)
		strJOBNAME =  TRIM(.txtJOBNAME.value)
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strCONTRACTNO = .txtCONTRACTNO.value 
		
		If Len(strCONTRACTNO) = 10 Then
			strCONTRACTNO = MID(strCONTRACTNO,1,7) & "-" & MID(strCONTRACTNO,8,3)
		End if
		
		strCONFIRM = .cmbTEST.value
		
		
		
		IF strGBN = "F" THEN  '�̿Ϸ���ȸ
		 Call Input_Layout()
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strJOBNO,strJOBNAME)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE		
   					If mlngRowCnt > 0 Then
   					
   						For intCnt = 1 To .sprSht.MaxRows
								If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
								End If
						Next	
						initpageData
   					Else
   						.sprSht.MaxRows = 0
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO", true
   					.imgDelete.disabled = true
   					.imgSave.disabled = false
   			end if
		ELSEIF strGBN = "T" THEN  '�Ϸ���ȸ
			Call Input_Layout()
			Call Select_Layout()
		
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn_EXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strCONFIRM,strCONTRACTNO)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE			
   					If mlngRowCnt > 0 Then	
   						sprShtToFieldBinding 1,1
   					Else
   						.sprSht.MaxRows = 0
   						
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNAME", false	
   					.imgDelete.disabled = false
   					.imgSave.disabled = false
   			end if
   		ELSEIF strGBN = "" THEN  '��ü��ȸ
   			Call Input_Layout()
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn_ALL(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,strOUTSCODE,strOUTSNAME,strJOBNO,strJOBNAME,strCONTRACTNO)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
					mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE		
   					If mlngRowCnt > 0 Then
   					
   						For intCnt = 1 To .sprSht.MaxRows
								If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
								End If
						Next	
						initpageData
   					Else
   						.sprSht.MaxRows = 0
   					End If
   					mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO", false	
   					.imgDelete.disabled = true
   					.imgSave.disabled = true
   			end if
		END IF
   	end with
End Sub
'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strGUBN
	Dim strOUTSCODE
	Dim strCONTRACTNAME
	Dim strSAVEFLAG
	Dim strCOMENT 
	Dim intCnt2
	Dim lngCNTSUM
	Dim lngCNT
	
	strMasterData = gXMLGetBindingData (xmlBind)
		with frmThis
		
		If .cmbENDGBN.value  = "F" Then
			strSAVEFLAG = "F"
		Elseif .cmbENDGBN.value = "T" Then
			strSAVEFLAG = "T"
		End If
		
		'txtCONTRACTNAME,txtCONTRACTDAY
		
		If .sprSht.MaxRows = 0 Then
				gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
				Exit Sub
		End IF
		lngCNTSUM = 0
		lngCNT = 0
	
		For intCnt2 = 1 To .sprSht.MaxRows
			lngCNT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2)
			lngCNTSUM = lngCNTSUM + lngCNT
		Next
		If lngCNTSUM = 0 Then
			gErrorMsgBox "���õǾ��� �ڷᰡ �����ϴ�.","����ȳ�"
			Exit Sub
		End if
		
		If .cmbENDGBN.value ="F" Then
			if DataValidation =false then exit sub
		End If
		If strSAVEFLAG = "F" Then
			If .txtCONTRACTNAME.value = "" Then
				gErrorMsgBox "������ �־��ֽʽÿ�.","����ȳ�"
				Exit Sub
			End If
			If .txtCONTRACTDAY.value = "" Then
				gErrorMsgBox "������� �־��ֽʽÿ�.","����ȳ�"
				Exit Sub
			End If
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|JOBNO|SEQ")
		Elseif strSAVEFLAG = "T" then
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|CONTRACTNO|CONTRACTNAME|LOCALAREA|STDATE|EDDATE|AMT|DELIVERYDAY|TESTDAY|PAYMENTGBN|TESTMENT|COMENT|OUTSCODE|CONFIRMFLAG|CONTRACTDAY")
		End If
		
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
				gErrorMsgBox "���õ� " & meNO_DATA,"����ȳ�"
				exit Sub
			Else
				gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
				exit sub
			End If
		End If
		strGUBN = ""
		If strSAVEFLAG = "F" then
			For intCnt = 1 to .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
					strGUBN = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt)
					strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
					'strCONTRACTNAME =  mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",intCnt)
				End If
				If strGUBN <> "" AND strOUTSCODE <> "" AND strCONTRACTNAME <> "" Then
					Exit For
				End If
			Next
			strCONTRACTNAME = .txtCONTRACTNAME.value 
			
			strGUBN = MID(strGUBN,1,1)
			
			If strGUBN = "" Then
				gErrorMsgBox "���õǾ���JOB��ȣ�� �����ϴ�.","����ȳ�"
				Exit Sub
			End If
		End If
		strCOMENT = .txtCOMENT.value 
	
		intRtn = mobjPDCMCONTRACT.ProcessRtn(gstrConfigXml,strMasterData,vntData,strGUBN,strOUTSCODE,strCONTRACTNAME,strSAVEFLAG,strCOMENT )
			if not gDoErrorRtn ("ProcessRtn") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
				SelectRtn
			End If
		End with
End Sub

Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim strOUTSCODE
   	Dim lngCnt
   	Dim strSTDINT
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		strSTDINT = ""
   		for intCnt = 1 To .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
   				strSTDINT = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
   				
   				If strSTDINT <> ""  Then
   					Exit For
   				End If
   			End if
   		Next
  
   		for intCnt = 1 to .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
				if strSTDINT <> mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) Then
					gErrorMsgBox intCnt & " ��° ���� ����ó�� Ȯ���Ͻʽÿ�." & vbcrlf & "���Ͽ���ó �ϰ�쿡�� ������ �����մϴ�.","�Է¿���"
					Exit Function
				End If
			End If
		next
   	
   	End with
	DataValidation = true
End Function
'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
	Dim strRow
	Dim strCONTRACTNO
	with frmThis
	
		
		
		'���õ� �ڷḦ ������ ���� ����
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		for i = .sprSht.MaxRows to 1 step -1
		
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "Y" Then
				gErrorMsgBox "Ȯ�������� �����ϽǼ� ������, �󼼳������� Ȯ���� ����� �����Ͻʽÿ�.","�����ȳ�"
				Exit Sub
			End if
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
				
					strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					intRtn = mobjPDCMCONTRACT.DeleteRtn(gstrConfigXml,strCONTRACTNO)
				End IF
				
   			End If
   			IF not gDoErrorRtn ("DeleteRtn") then
					mobjSCGLSpr.DeleteRow .sprSht,i
					
   			End IF
		next
		gWriteText lblstatus, "�ڷᰡ " & intRtn & " �� �����Ǿ����ϴ�."
		'���� ���� ����
		'mobjSCGLSpr.DeselectBlock .sprSht
		'strRow = .sprSht.ActiveRow
		SelectRtn
		'mobjSCGLSpr.ActiveCell .sprSht, 1, strRow
		'Call sprSht_Click(1,strRow)
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
			<TABLE id="tblForm" style="WIDTH: 100%" HEIGHT="100%" cellSpacing="0" cellPadding="0" border="0">
				<!--Top TR Start-->
				<TBODY>
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
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;������</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<TABLE id="tblButton1"  cellSpacing="0" cellPadding="2"
											 border="0"ALIGN="right">
											<TR>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
														height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgClose.gIF" border="0" name="imgClose"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call DateClean()"
										width="83">�Ⱓ</TD>
									<TD class="SEARCHDATA" style="WIDTH: 249px"><INPUT class="INPUT" id="txtFrom" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtFrom"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtTo"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
											width="23" align="absMiddle" border="0" name="imgTo">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand">�Ϸᱸ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 135px; CURSOR: hand"><SELECT id="cmbENDGBN" style="WIDTH: 128px" name="cmbENDGBN">
											<OPTION value="">��ü</OPTION>
											<OPTION value="F" selected>�̿Ϸ�</OPTION>
											<OPTION value="T">�Ϸ�</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE1)">����ó</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtOUTSNAME" title="����ó�� ��ȸ" style="WIDTH: 224px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtOUTSNAME"><IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtOUTSCODE1" title="����ó�ڵ���ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtOUTSCODE1"></TD>
									<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
											src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 83px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')"
										width="83">��༭��ȣ</TD>
									<TD class="SEARCHDATA" style="WIDTH: 249px"><INPUT class="INPUT_L" id="txtCONTRACTNO" title="JOB�� ��ȸ" style="WIDTH: 240px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand">��༭Ȯ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 135px; CURSOR: hand"><SELECT id="cmbTEST" style="WIDTH: 128px" name="cmbTEST">
											<OPTION value="" selected>��ü</OPTION>
											<OPTION value="��༭ ��Ȯ��">��༭ ��Ȯ��</OPTION>
											<OPTION value="��༭ Ȯ��">��༭ Ȯ��</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)">JOB��</TD>
									<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtJOBNAME" title="JOB�� ��ȸ" style="WIDTH: 224px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
											border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtJOBNO" title="JOBNO ��ȸ" style="WIDTH: 65px; HEIGHT: 22px" type="text"
											maxLength="7" align="left" size="3" name="txtJOBNO"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left"  height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;��༭ ��� �� Ȯ��</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD>
												<!--		
												<td><IMG id="imgTestOK" onmouseover="JavaScript:this.src='../../../images/imgTestOKOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTestOK.gIF'"
														height="20" alt="�˼��� Ȯ��ó�� �մϴ�." src="../../../images/imgTestOK.gIF" border="0" name="imgTestOK"></td>
												<td><IMG id="imgTestCancel" onmouseover="JavaScript:this.src='../../../images/imgTestCancelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgTestCancel.gIF'"
														height="20" alt="�˼��� ����մϴ�." src="../../../images/imgTestCancel.gIF" border="0" name="imgTestCancel"></td>-->
												<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
														height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
														name="imgPrint"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							
							
							<!---->
							<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
							
							
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="LEFT" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')"
													width="85"><FONT face="����">����</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px" width="251"></FONT><INPUT dataFld="CONTRACTNAME" id="txtCONTRACTNAME" style="WIDTH: 240px; HEIGHT: 21px" accessKey="M"
														dataSrc="#xmlBind" type="text" size="33" name="txtCONTRACTNAME" title="����"></TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTDAY,'')"
													width="89"><FONT face="����">�����</FONT></TD>
												<TD class="DATA" width="257"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="�����" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="M,DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" alt="ImgCONTRACTDAY" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="ImgCONTRACTDAY"></TD>
												<TD class="LABEL" width="90" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtLOCALAREA,'')"><FONT face="����">��ǰ���</FONT></TD>
												<TD class="DATA" width="257"><FONT face="����"><INPUT dataFld="LOCALAREA" class="INPUT_L" id="txtLOCALAREA" title="��ǰ���" style="WIDTH: 251px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="36" name="txtLOCALAREA"></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDELIVERYDAY, '')"><FONT face="����">��ǰ��</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px"><FONT face="����"></FONT><INPUT dataFld="DELIVERYDAY" class="INPUT" id="txtDELIVERYDAY" title="��ǰ��" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDELIVERYDAY"><IMG id="imgDELIVERYDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgDELIVERYDAY">&nbsp;
													<INPUT dataFld="OUTSCODE" id="txtOUTSCODE" title="����ó�ڵ�_����" style="WIDTH: 121px; HEIGHT: 21px"
														dataSrc="#xmlBind" type="text" size="14" name="txtOUTSCODE">
												</TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSTDATE, txtEDDATE)"><FONT face="����"><FONT face="����">�뿪�Ⱓ</FONT></FONT></TD>
												<TD class="DATA"></FONT><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="�뿪�Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtSTDATE"><IMG id="imgSTDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgSTDATE">&nbsp;~ <INPUT dataFld="EDDATE" class="INPUT" id="txtEDDATE" title="�뿪�Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtEDDATE"><IMG id="imgEDDATE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgEDDATE"></FONT></TD>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAMT, '')"><FONT face="����"><FONT face="����">���ݾ�</FONT></FONT></TD>
												<TD class="DATA"></FONT><FONT face="����"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="���ݾ�" style="WIDTH: 251px; HEIGHT: 22px"
															accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtAMT"></FONT></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTESTMENT, '')"><FONT face="����">�˼����</FONT></TD>
												<TD class="DATA" style="WIDTH: 251px"><FONT face="����"><INPUT dataFld="TESTMENT" class="INPUT_L" id="txtTESTMENT" title="�˼����" style="WIDTH: 240px; HEIGHT: 22px"
															dataSrc="#xmlBind" type="text" maxLength="255" size="35" name="txtTESTMENT"></FONT>
												</TD>
												<TD class="LABEL" style="WIDTH: 89px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTESTDAY, '')"><FONT face="����"><FONT face="����">�˼���</FONT></FONT></TD>
												<TD class="DATA"><INPUT dataFld="TESTDAY" class="INPUT" id="txtTESTDAY" title="�˼���" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtTESTDAY"><IMG id="imgTESTDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
														name="imgTESTDAY"></TD>
												<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPAYMENTGBN,'')"><FONT face="����">������޹��</FONT></TD>
												<TD class="DATA"><INPUT dataFld="PAYMENTGBN" class="INPUT_L" id="txtPAYMENTGBN" title="������޹��" style="WIDTH: 251px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="255" size="37" name="txtPAYMENTGBN"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 85px; CURSOR: hand; HEIGHT: 130px" onclick="vbscript:Call gCleanField(txtCOMENT,'')"><FONT face="����">Ư�����</FONT></TD>
												<TD class="DATA" colSpan="5"><TEXTAREA dataFld="COMENT" id="txtCOMENT" style="WIDTH: 952px" dataSrc="#xmlBind" name="txtCOMENT"
														rows="8" wrap="hard" cols="116"></TEXTAREA></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27464">
									<PARAM NAME="_ExtentY" VALUE="11642">
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
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
