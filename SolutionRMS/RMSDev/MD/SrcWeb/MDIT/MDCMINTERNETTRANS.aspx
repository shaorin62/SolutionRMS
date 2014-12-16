<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMINTERNETTRANS.aspx.vb" Inherits="MD.MDCMINTERNETTRANS" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���ͳ� �ŷ�����ǥ ���� �� ����</title>
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
Dim mobjMDCMINTERNETTRANS, mobjMDCMGET
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

'-----------------------------------
' ���� ��ư Ŭ�� �̺�Ʈ
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
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺��� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCMINTERNETTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMINTERNETTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		
		vntData = mobjMDCMINTERNETTRANS.Get_INTERNETTRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_INTERNETTRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum 
			for i=1 to 1
				strUSERID = ""
				vntDataTemp = mobjMDCMINTERNETTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
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
		intRtn = mobjMDCMINTERNETTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'��� �μ��ư Ŭ���� �̺�Ʈ
Sub imgFind_onclick
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		IF .txtTRANSYEARMON.value = "" THEN
			gErrorMsgBox "��ȸ ����� �Է��Ͻÿ�","��ȸ�ȳ�"
			Exit Sub
		END IF
		
		IF Len(.txtTRANSYEARMON.value) <> 6 THEN
			gErrorMsgBox "��ȸ ����� 6�ڸ� �Է��Դϴ�.","��ȸ�ȳ�"
			Exit Sub
		END IF
		
		strTRANSYEARMON = .txtTRANSYEARMON.value
		vntInParams = array(strTRANSYEARMON, "INTERNET") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMGROUPTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = vntRet(0,0)  ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(1,0)  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
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
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		strSPONSOR = ""
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "INTERNET", strSPONSOR) 
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = vntRet(1,0)		  ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			IF vntRet(3,0) = "�Ϸ�" THEN
				window.event.keyCode = meEnter
				txtTRANSNO_onkeydown
			ELSE
				.txtTRANSNO.value = ""
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
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtTRANSNO.value,.txtCLIENTNAME1.value,"ALL","trans", "INTERNET", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(1,0)
					.txtCLIENTNAME1.value = vntData(2,0)
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
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "INTERNET", "0")
			if not gDoErrorRtn ("GetTRANSNO") Then
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
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
		strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
		
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "INTERNET") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) Then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = vntRet(2,0)  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = vntRet(3,0)  ' �ڵ�� ǥ��
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
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 Then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
			if mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
	IF Col = 15 Then
		Dim strSUM
		strSUM = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"AMOUNT",Row) + mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"VAT",Row)
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SUMAMOUNT",Row, strSUM
	END IF
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

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjMDCMINTERNETTRANS	= gCreateRemoteObject("cMDIT.ccMDITINTERNETTRANS")
	set mobjMDCMGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "150px"
	pnlTab1.style.left= "7px"
	
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "150px"
	pnlTab2.style.left= "7px"

	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'�ŷ������� ���� �׸���
		'******************************************************************
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,   " CHK | YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | PROGRAM_NAME | DEPT_CD | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | NOTE | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRU_TAX_FLAG | MED_FLAG | TRANSRANK | SPONSOR | ATTR01"
		
		mobjSCGLSpr.SetHeader .sprSht,		   " ����|YEARMON|����|������|��ü��| ��ü��|�����|DEPT_CD|������|������|���ް���|����|��|NOTE|CLIENTCODE|MEDCODE|REAL_MED_CODE|TRU_TAX_FLAG |MED_FLAG|TRANSRANK|SPONSOR|ATTR01"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 	4|      0|   4|    12|    12|     0|     12|      0|     8|     8|      10|  10|10|   0|	        0|	    0|		      0|	        0|  0| 0|0|0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDSTDATE|TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "TBRDSTDATE|TBRDEDDATE|AMT|VAT|SUMAMTVAT" 
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "PROGRAM_NAME|SEQ |CLIENTNAME|MEDNAME  ", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht, " YEARMON | REAL_MED_NAME|DEPT_CD|NOTE|CLIENTCODE | MEDCODE | REAL_MED_CODE | TRU_TAX_FLAG |MED_FLAG| TRANSRANK | SPONSOR | ATTR01 ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLIENTNAME|MEDNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQ",-1,-1,2,2,false
		
		'�հ� ǥ�� �׸��� �⺻ȭ�� ����
		gSetSheetColor mobjSCGLSpr, .sprSht_SUM
		mobjSCGLSpr.SpreadLayout .sprSht_SUM, 22, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_SUM, "CHK | YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | PROGRAM_NAME | DEPT_CD | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | NOTE | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRU_TAX_FLAG | MED_FLAG | TRANSRANK | SPONSOR | ATTR01"
		mobjSCGLSpr.SetText .sprSht_SUM, 4, 1, "��      ��"
	    mobjSCGLSpr.SetScrollBar .sprSht_SUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_SUM,"1|3|4",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUM, "AMT|VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_SUM, "YEARMON | REAL_MED_NAME|DEPT_CD|NOTE|CLIENTCODE | MEDCODE | REAL_MED_CODE | TRU_TAX_FLAG |MED_FLAG| TRANSRANK | SPONSOR | ATTR01 ", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_SUM, "-1", "13"
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_SUM
		
		'******************************************************************
		'�ŷ������� ��ȸ �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 25, 0, 0, 2
		mobjSCGLSpr.SpreadDataField .sprSht1, "TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME |TRUST_SEQ| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|DEPT_CD|DEMANDDAY|PRINTDAY|PROGRAM |TBRDSTDATE|TBRDEDDATE| AMT| TRU_TAX_FLAG|VAT|SUMAMTVAT|MEMO|MED_FLAG|SPONSOR | TAXYEARMON | TAXNO"
		mobjSCGLSpr.SetHeader .sprSht1,		"TRANSYEARMON|TRANSNO|SEQ|CLIENTCODE|������|����|MEDCODE|��ü��|REAL_MED_CODE|û����|DEPT_CD|DEMANDDAY|PRINTDAY|�����|������|������| ���ް���|TRU_TAX_FLAG|�ΰ�����|��|MEMO|MED_FLAG|SPONSOR| TAXYEARMON | TAXNO"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "        0|	    0|	0|	       0|	  0|   4|      0|    18|	        0|     0|	   0|		 0|       0|    16|     9|     9|       12|           0|      10|13|   0|       0|      0|           0|  0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "TBRDSTDATE|TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT|VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht1, true, "TBRDSTDATE|TBRDEDDATE|AMT|VAT|SUMAMTVAT" 
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "MEDNAME|TRUST_SEQ|PROGRAM", -1, -1, 20
		mobjSCGLSpr.ColHidden .sprSht1, "TRANSYEARMON|TRANSNO|SEQ|CLIENTCODE|CLIENTNAME|REAL_MED_NAME|MEDCODE|REAL_MED_CODE|DEPT_CD|DEMANDDAY|PRINTDAY|TRU_TAX_FLAG|MEMO|MED_FLAG|SPONSOR| TAXYEARMON | TAXNO ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "MEDNAME|PROGRAM",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "TRUST_SEQ",-1,-1,2,2,false
		
		
    End With    
	pnlTab1.style.visibility = "visible"
	

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
				case 1 : .txtCLIENTCODE.value = vntInParam(i)
				case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'��ȸ�߰��ʵ�
				case 3 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	SelectRtn	
End Sub

Sub EndPage()
	set mobjMDCMINTERNETTRANS = Nothing
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
		.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean
		'.txtDEMANDDAY.value = gNowDate
		.txtPRINTDAY.value  = gNowDate
		.sprSht.MaxRows = 0	
		.sprSht1.MaxRows = 0
		
		.txtDEMANDDAY.readOnly = "FALSE"
		.txtDEMANDDAY.className = "INPUT"
		.imgCalDemandday.disabled = FALSE

	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim chkcnt
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	chkcnt = 0
	strCLIENTCODE = ""
	strCLIENTNAME = ""
	
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtPRINTDAY.value = "" Then
			gErrorMsgBox "�������� �ʼ� �Է� ���� �Դϴ�.",""
			Exit Sub
		End If
		
		For intCnt = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "�ŷ��������� ������ �����͸� üũ �Ͻʽÿ�",""
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
			IF intColFlag < bsdiv Then
				intColFlag = bsdiv
			END IF
		Next

		
   		'�հ�ݾ� ��� �� VAT ���
   		'Call SUMAMT_SUM ()
   		'������ Validation
   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | PROGRAM_NAME | DEPT_CD | TBRDSTDATE | TBRDEDDATE | AMT | VAT | SUMAMTVAT | NOTE | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRU_TAX_FLAG | MED_FLAG | TRANSRANK | SPONSOR | ATTR01")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = .txtTRANSYEARMON.value
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME1.value
		
		intRtn = mobjMDCMINTERNETTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)
   		
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
'		intColSum = 0
' 		for intCnt = 1 to .sprSht.MaxRows
'			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1  Then 
'				intColSum = intColSum + 1
'			End if
'		next
'		If intColSum = 0 Then exit Function
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
	Dim vntData, vntData1
	Dim strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME, strTRANSNO
	Dim strPRINTDAY
	Dim strSPONSOR
   	Dim i, strCols
   
	'On error resume next
	with frmThis
		If .txtTRANSYEARMON.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If
		
		strTRANSNO = ""
		strTRANSNO = .txtTRANSNO.value
		IF strTRANSNO = "" Then
			IF  .txtCLIENTCODE.value = "" Then
				gErrorMsgBox "��ȸ�� �������ڵ�� �ݵ�� �־�� �մϴ�.",""
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
		
		IF strTRANSNO <> "" Then
			InitPageData
			IF not SelectRtn_HDR (strTRANSYEARMON, strTRANSNO, strCLIENTCODE) Then Exit Sub
			
			pnlTab1.style.visibility = "HIDDEN"
			pnlTab2.style.visibility = "visible"
			
			.txtDEMANDDAY.readOnly = "TRUE"
			.txtDEMANDDAY.className = "NOINPUT"
			.imgCalDemandday.disabled = True

			'��Ʈ ��ȸ
			if SelectRtn_DTL (strTRANSYEARMON, strTRANSNO, strCLIENTCODE) then
				.txtTRANSYEARMON.value = strTRANSYEARMON
				.txtTRANSNO.value = strTRANSNO
				.txtCLIENTCODE.value = strCLIENTCODE
				.txtCLIENTNAME1.value = strCLIENTNAME
			end if
		ELSE
			InitPageData
			
			vntData = mobjMDCMINTERNETTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, strCLIENTCODE)
			
			pnlTab1.style.visibility = "visible"
			pnlTab2.style.visibility = "HIDDEN"
			
			.txtDEMANDDAY.readOnly = "FALSE"
			.txtDEMANDDAY.className = "INPUT"
			.imgCalDemandday.disabled = FALSE
			
			if not gDoErrorRtn ("SelectRtn") Then
				if mlngRowCnt > 0 Then
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   					
   					PreSearchFiledValue strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				else
   					'InitPageData
   					PreSearchFiledValue strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME
   					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				end if
   				DateClean
   				AMT_SUM '�հ�׸��� ǥ��
   			end if
		END IF
   	end with
End Sub

Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strCLIENTCODE)
	Dim vntData
	on error resume next

	'�ʱ�ȭ
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMGET.Get_INTERNETTRANS_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_INTERNETTRANS_HDR") Then
		IF mlngRowCnt<=0 Then
			gErrorMsgBox "������ �ŷ�������ȣ�� ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			gWriteText "", "������ �ŷ�������ȣ�� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			SelectRtn_HDR = True
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByRef strYEARMON,ByRef strTRANSNO, ByRef strCLIENTCODE)
	Dim vntData
	Dim intCnt
	on error resume next

	'�ʱ�ȭ
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMGET.Get_INTERNETTRANS_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCLIENTCODE)
	
	IF not gDoErrorRtn ("Get_INTERNETTRANS_LIST") Then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
		
		SelectRtn_DTL = True
		gWriteText "", "������ �ŷ�������ȣ���� �󼼳����� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME)
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
End Sub


'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMOUNT, IntAMTSUM, IntVAT, IntVATSUM, IntSUMAMOUNT, IntSUMAMOUNTSUM
	'AMT|VAT|SUMAMTVAT
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMAMOUNTSUM = 0
		IF .sprSht.MaxRows > 0 THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMOUNT = 0
				IntVAT = 0
				IntSUMAMOUNT = 0
				
				IntAMOUNT	 = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt) '�ݾ�
				IntVAT		 = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)	'�ΰ���
				IntSUMAMOUNT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMAMTVAT", lngCnt) '�ݾ׺ΰ����հ�
				
				IntAMTSUM		= IntAMTSUM + IntAMOUNT
				IntVATSUM		= IntVATSUM + IntVAT
				IntSUMAMOUNTSUM = IntSUMAMOUNTSUM + IntSUMAMOUNT
			Next
		end if
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, IntVATSUM		
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, IntSUMAMOUNTSUM
			'mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		ELSE
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"AMT",1, 0
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"VAT",1, 0		
			mobjSCGLSpr.SetTextBinding .sprSht_SUM,"SUMAMTVAT",1, 0
		end if
	End With
End Sub

'****************************************************************************************
' �ŷ������� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strDESCRIPTION
	with frmThis
		strDESCRIPTION = ""
		
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
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"TRANSYEARMON | TRANSNO | SEQ | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|DEPT_CD|DEMANDDAY|PRINTDAY|PROGRAM |TBRDSTDATE|TBRDEDDATE| AMT| TRU_TAX_FLAG|VAT|SUMAMTVAT|TRUST_SEQ|MEMO|MED_FLAG|SPONSOR | TAXYEARMON | TAXNO")
		
		'���õ� �ڷḦ ������ ���� ����
		strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
	
		intRtn = mobjMDCMINTERNETTRANS.DeleteRtn(gstrConfigXml,vntData, strTRANSYEARMON, strTRANSNO)

		IF not gDoErrorRtn ("DeleteRtn") then
			If strDESCRIPTION <> "" Then
				gErrorMsgBox strDESCRIPTION,"�����ȳ�!"
				Exit Sub
			End If
			for i = .sprSht1.MaxRows to 1 step -1
				mobjSCGLSpr.DeleteRow .sprSht1,i
			next
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", strTRANSYEARMON & "-" & strTRANSNO & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ������ ����
		mobjSCGLSpr.DeselectBlock .sprSht1
		
		'SelectRtn
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
			<TABLE id="tblForm" style="WIDTH: 793px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TBODY>
									<TR>
										<TD style="WIDTH: 300px" align="left" width="427" height="28">
											<table cellSpacing="0" cellPadding="0" width="100%" border="0">
												<tr>
													<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
													<td align="left" height="4"><FONT face="����"></FONT></td>
												</tr>
												<tr>
													<td class="TITLE">���ͳ�&nbsp;����Ź�ŷ����� ���� �� ����</td>
												</tr>
											</table>
										</TD>
										<TD style="WIDTH: 495px" vAlign="middle" align="right" height="28">
											<!--Wait Button Start-->
											<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
												cellSpacing="1" cellPadding="1" width="75%" border="0">
												<TR>
													<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
															border="0" name="imgWaiting">
													</TD>
												</TR>
											</TABLE>
											<!--Wait Button End-->
											<!--Common Button Start-->
											<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
												width="203" border="0">
												<TBODY>
													<TR>
														<TD></TD>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
																name="imgQuery"></TD>
														<TD><IMG id="imgNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" height="20" alt="�ű��ڷḦ �ۼ��մϴ�."
																src="../../../images/imgNew.gIF" width="54" border="0" name="imgNew"></TD>
														<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></TD>
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
																name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�."
																src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgDelete.gIF" width="54"
																border="0" name="imgDelete"></TD>
														<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
																height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
																name="imgClose"></TD>
													</TR>
												</TBODY>
											</TABLE>
											<!--Common Button End--></TD>
									</TR>
									<!--Top Define Table End-->
									<!--Input Define Table End--></TBODY></TABLE>
							<TABLE id="tblBody" style="WIDTH: 792px" cellSpacing="0" cellPadding="0" width="792" border="0"> <!--TopSplit Start->
								
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="����">
											<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
												<TR>
													<TD class="SEARCHLABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)"
														width="91">�� ��</TD>
													<TD class="SEARCHDATA" width="312"><INPUT class="INPUT" id="txtTRANSYEARMON" title="�ŷ��������" style="WIDTH: 72px; HEIGHT: 22px"
															accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON"><IMG id="ImgTRU" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
															onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTRU"><INPUT class="INPUT" id="txtTRANSNO" title="�ŷ�������ȣ" style="WIDTH: 72px; HEIGHT: 22px" type="text"
															maxLength="6" size="6" name="txtTRANSNO"></TD>
													<TD class="SEARCHLABEL" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"
														width="90"><FONT face="����">��������</FONT></TD>
													<TD class="SEARCHDATA" width="312"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="���μ���" style="WIDTH: 96px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="100" size="10" name="txtPRINTDAY"><IMG id="ImgPRINTDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
															name="ImgPRINTDAY">
													</TD>
												</TR>
												<tr>
													<TD class="SEARCHLABEL" style="WIDTH: 91px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME1)">������
													</TD>
													<TD class="SEARCHDATA" colspan="3"><INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 72px; HEIGHT: 22px"
															type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
															border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 209px; HEIGHT: 22px"
															type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME1"></TD>
												</tr>
											</TABLE>
										</FONT>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 794px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="right" border="0">
											<TR>
												<TD class="LABEL" width="90"><FONT face="����">������</FONT></TD>
												<TD class="DATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME">
												</TD>
												<TD class="LABEL" width="90"><FONT face="����">���μ�</FONT></TD>
												<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="DEPT_NAME" class="NOINPUT_L" id="txtDEPT_NAME" title="�귣���ڵ�" style="WIDTH: 172px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtDEPT_NAME"></FONT>
												</TD>
												<TD class="LABEL" width="90"><FONT face="����">û������</FONT></TD>
												<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 104px; HEIGHT: 22px"
															accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="12" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
															style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></FONT></TD>
											</TR>
											<TR>
												<TD class="LABEL"><FONT face="����">���ް���</FONT></TD>
												<TD class="DATA"><FONT face="����"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="�����ݾ�" style="WIDTH: 135px; HEIGHT: 22px"
															dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="16" name="txtAMT"></FONT>
												</TD>
												<TD class="LABEL"><FONT face="����">�ΰ���</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="�ΰ���" style="WIDTH: 135px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="17" name="txtVAT"></TD>
												<TD class="LABEL"><FONT face="����">��</FONT></TD>
												<TD class="DATA"></FONT></FONT><INPUT dataFld="SUMAMTVAT" class="NOINPUT_R" id="txtSUMAMTVAT" title="��" style="WIDTH: 154px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="20" name="txtSUMAMTVAT"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End--></TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 791px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 794px; HEIGHT: 336px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 786px; HEIGHT: 336px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20796">
									<PARAM NAME="_ExtentY" VALUE="8890">
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
								<OBJECT id="sprSht_SUM" style="WIDTH: 786px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20796">
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
							<DIV id="pnlTab2" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 786px; HEIGHT: 336px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="20796">
									<PARAM NAME="_ExtentY" VALUE="8890">
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
					<!--tr>
						<td class="BRANCHFRAME" vAlign="middle">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;�� 
							�� :&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT class="NOINPUT_R" id="txtSUM" title="�ݾ�" style="WIDTH: 128px; HEIGHT: 19px" accessKey="NUM"
								readOnly type="text" size="16" name="txtSUM">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</td>
					</tr-->
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 794px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
					</TR>
					<!--BodySplit End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 794px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>