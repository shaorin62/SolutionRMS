<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMOUTDOOR.aspx.vb" Inherits="MD.MDCMOUTDOOR" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� û�����</title>
		<meta content="False" name="vs_showGrid">
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
Dim mobjOUTDOOR '�����ڵ�, Ŭ����
Dim mobjMDCMGET	
Dim mstrCheck
CONST meTAB = 9
mstrCheck = True

'=============================
' �̺�Ʈ ���ν��� 
'=============================
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


Sub imgGETPOTALDATA_onclick
	Dim vntRet
	Dim vntInParams
	Dim strSPONSOR
	
	with frmThis
		vntInParams = array(.txtYEARMON1.value) 
		vntRet = gShowModalWindow("MDCMOUTDOORBATCH.aspx",vntInParams , 1060,670)
		
		if vntRet <> "" then
			.txtYEARMON1.value = vntRet
			selectRtn
		end if
	End with
End Sub


'****************************************************************************************
' �˾� �̺�Ʈ, ������, ��ü��, ��ü��
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value))
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))       ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))      ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE                  ' gSetChangeFlag objectID	 Flag ���� �˸�
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
			
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME1.value))
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,0))
					.txtCLIENTNAME1.value = trim(vntData(1,0))
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
' ��ü���ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE_POP
	dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMREALMEDPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			if .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME1.value = trim(vntRet(1,0))     ' �ڵ�� ǥ��
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
			vntData = mobjMDCMGET.GetREALMEDNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME1.value))
			if not gDoErrorRtn ("GetREALMEDNO") then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,0))
					.txtREAL_MED_NAME1.value = trim(vntData(1,0))
				Else
					Call REAL_MED_CODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'VAT��
Sub chkVATYES_onchange
	DIM strVATYES
	if frmThis.chkVATYES.checked = true then 
		strVATYES = "Y"
	ELSE 
		strVATYES = "N"
	end if
	
	if frmThis.sprSht.ActiveRow >0 Then	
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_TAX_FLAG",frmThis.sprSht.ActiveRow, strVATYES
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'VAT��
Sub chkVATYES_onClick
	DIM strVATYES
	if frmThis.chkVATYES.checked = true then 
		strVATYES = "Y"
	ELSE 
		strVATYES = "N"
	end if
	
	if frmThis.sprSht.ActiveRow >0 Then	
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_TAX_FLAG",frmThis.sprSht.ActiveRow, strVATYES
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'VAT��
Sub chkVATNO_onchange
	DIM strVATNO
	if frmThis.chkVATNO.checked = true then 
		strVATNO = "N"
	ELSE 
		strVATNO = "Y"
	end if
	
	if frmThis.sprSht.ActiveRow >0 Then	
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_TAX_FLAG",frmThis.sprSht.ActiveRow, strVATNO
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'VAT��
Sub chkVATNO_onClick
	DIM strVATNO
	if frmThis.chkVATNO.checked = true then 
		strVATNO = "N"
	ELSE 
		strVATNO = "Y"
	end if
	
	if frmThis.sprSht.ActiveRow >0 Then	
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMMI_TAX_FLAG",frmThis.sprSht.ActiveRow, strVATNO
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'�Ϲݸ��ⱸ��
Sub chkGBN_FLAG1_onchange
	if frmThis.sprSht.ActiveRow >0 Then	
		IF frmThis.chkGBN_FLAG1.checked = true then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "0"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end if
	gSetChange
End Sub
'�Ϲݸ��ⱸ��
Sub chkGBN_FLAG1_onClick
	if frmThis.sprSht.ActiveRow >0 Then	
		IF frmThis.chkGBN_FLAG1.checked = true then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "0"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end if
	gSetChange
End Sub
'���౸��
Sub chkGBN_FLAG2_onchange
	if frmThis.sprSht.ActiveRow >0 Then	
		IF frmThis.chkGBN_FLAG2.checked = true then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "0"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end if
	gSetChange
End Sub

'���౸��
Sub chkGBN_FLAG2_onClick
	if frmThis.sprSht.ActiveRow >0 Then	
		IF frmThis.chkGBN_FLAG1.checked = true then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "0"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN_FLAG",frmThis.sprSht.ActiveRow, "1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end if
	gSetChange
End Sub

Sub txtNOTE_onchange
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"NOTE",frmThis.sprSht.ActiveRow, frmThis.txtNOTE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		
	end if
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 and Col > 1 then		
			sprShtToFieldBinding Col,Row
		elseif Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
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
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'��Ʈ�� �������ѷο��� ������ ��� �ʴ��� ���ε�
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		
		.txtYEARMON.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtSEQ.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtREAL_MED_NAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_NAME",Row)
		.txtPROGNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",Row)
		.txtCLIENTSUBNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTSUBNAME",Row)
		.txtSUBSEQNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQNAME",Row)
		.txtDEPT_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtTOTALAMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"TOTALAMT",Row)
		.txtAMT.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtNOTE.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"NOTE",Row)
		.txtCOMMI_RATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_RATE",Row)
		.txtCOMMISSION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		
		.txtTBRDSTDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row)
		.txtTBRDEDDATE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row)
		.txtCOMMISSION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION",Row)
		
		.txtMED_GBN.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"MED_GBN",Row)
		.txtLOCATION.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"LOCATION",Row)
		.txtOUT_AMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"OUT_AMT",Row)
		
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TAX_FLAG",Row) = "Y" THEN
			.chkVATYES.checked = TRUE
			.chkVATNO.checked = FALSE
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TAX_FLAG",Row) = "N" THEN
			.chkVATYES.checked = FALSE
			.chkVATNO.checked = TRUE
		ELSE
			.chkVATYES.checked = FALSE
			.chkVATNO.checked = FALSE
		END IF
		
'		IF mobjSCGLSpr.GetTextBinding(.sprSht,"GBN_FLAG",Row) = "1" THEN
'			.chkGBN_FLAG1.checked = TRUE
'			.chkGBN_FLAG2.checked = FALSE
'		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"GBN_FLAG",Row) = "0" THEN
'			.chkGBN_FLAG1.checked = FALSE
'			.chkGBN_FLAG2.checked = TRUE
'		ELSE
'			.chkGBN_FLAG1.checked = FALSE
'			.chkGBN_FLAG2.checked = FALSE
'		END IF
		
   	end with
   	call gFormatNumber(frmThis.txtTOTALAMT,0,true)
	call gFormatNumber(frmThis.txtAMT,0,true)
	call gFormatNumber(frmThis.txtCOMMISSION,0,true)
	call gFormatNumber(frmThis.txtOUT_AMT,0,true)
End Function

Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht1, NewTop, NewLeft
End Sub

sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht1	
	End with
end sub
'=============================
' UI���� ���ν��� 
'=============================
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	set mobjOUTDOOR	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOOR")
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "278px"
	pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 29, 0, 1, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|YEARMON|SEQ|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|DEPT_NAME|TBRDSTDATE|TBRDEDDATE|TITLE|PROGNAME|TOTALAMT|AMT|OUT_AMT|COMMI_RATE|COMMISSION|MED_GBN|LOCATION|NOTE|GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|����|������|�����|��ü��|�귣��|�μ�|��������|���������|����|�����|�Ѱ��ݾ�|��û���ݾ�|�����ֺ�|������|������|��������|���|���|�Ϲݴ��౸��|CONTIDX|CYEAR|CMONTH|�Ϲݸ���ŷ�������ȣ|�Ϲݸ��⼼�ݰ�꼭��ȣ|�ΰ�����������|����ŷ�������ȣ|���༼�ݰ�꼭��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|   0|   0|    15|	15|	   15|    15|   0|         8|         8|    15|    18|        10|        10|     10|      6|    10|      10|  10|  18|           0|	     0|    0|     0|                     0|                     0|             0|                 0|                 0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "TBRDSTDATE|TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOTALAMT|AMT|OUT_AMT|COMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "NOTE", -1, -1, 20
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "YEARMON|SEQ|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|DEPT_NAME|TBRDSTDATE|TBRDEDDATE|TITLE|PROGNAME|TOTALAMT|AMT|OUT_AMT|COMMI_RATE|COMMISSION|MED_GBN|LOCATION|GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO"
		mobjSCGLSpr.ColHidden .sprSht, "GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO", true
		
		'�հ� ǥ�� �׸��� �⺻ȭ�� ����
		gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 29, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht1, "CHK|YEARMON|SEQ|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|DEPT_NAME|TBRDSTDATE|TBRDEDDATE|TITLE|PROGNAME|TOTALAMT|AMT|OUT_AMT|COMMI_RATE|COMMISSION|MED_GBN|LOCATION|NOTE|GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO"
		mobjSCGLSpr.SetText .sprSht1, 4, 1, "��      ��"
	    mobjSCGLSpr.SetScrollBar .sprSht1, 0
	    mobjSCGLSpr.SetBackColor .sprSht1,"1|2|4",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "TOTALAMT|AMT|OUT_AMT|COMMISSION", -1, -1, 0
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht1, "GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO", true
		
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht1
	    	    
	    .sprSht.style.visibility  = "visible"
	    .sprSht1.style.visibility = "visible"
	End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjOUTDOOR = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.txtYEARMON1.value = Mid(gNowDate,1,4) & MID(gNowDate,6,2)
		
'		txtYEARMON_onchange
'		txtTBRDSTDATE_onchange
'		txtTBRDEDDATE_onchange
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON
	Dim strSEQ 
	
	with frmThis
   		'������ Validation
		if DataValidation =false then exit sub
		On error resume next

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|SEQ|CLIENTNAME|CLIENTSUBNAME|REAL_MED_NAME|SUBSEQNAME|DEPT_NAME|TBRDSTDATE|TBRDEDDATE|TITLE|PROGNAME|TOTALAMT|AMT|OUT_AMT|COMMI_RATE|COMMISSION|MED_GBN|LOCATION|NOTE|GBN_FLAG|CONTIDX|CYEAR|CMONTH|COMMI_TRANS_NO|COMMI_TAX_NO|COMMI_TAX_FLAG|TRU_TRANS_NO|TRU_TAX_NO")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strYEARMON=""
		strSEQ=0
		
		intRtn = mobjOUTDOOR.ProcessRtn(gstrConfigXml,strMasterData,vntData, strYEARMON, strSEQ)
   		
   		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
   		
   	End with
	DataValidation = true
End Function

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCLIENTFLAG, strCLIENTCODE,strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strREAL_MED_NAME
   	Dim i, strCols
	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON1.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME1.value
		strREAL_MED_CODE	= .txtREAL_MED_CODE.value
		strREAL_MED_NAME	= .txtREAL_MED_NAME1.value
		
		vntData = mobjOUTDOOR.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strCLIENTCODE, strREAL_MED_CODE)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt >0 then
				Call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   		
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				sprShtToFieldBinding 2, 1
   				AMT_SUM
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   				PreSearchFiledValue strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME
   			end if
   		end if
   	end with
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME)
	frmThis.txtYEARMON1.value = strYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME1.value = strREAL_MED_NAME
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, strAMT, strAMTSUM
	Dim strCOMMISSION, strCOMMISSIONSUM
	
	With frmThis
		strAMTSUM = 0
		strCOMMISSIONSUM = 0
		
		For lngCnt = 1 To .sprSht.MaxRows
			strAMT = 0
			strCOMMISSION = 0
			strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			strCOMMISSION = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMISSION", lngCnt)
			strAMTSUM = strAMTSUM + strAMT
			strCOMMISSIONSUM = strCOMMISSIONSUM + strCOMMISSION
		Next
		
		mobjSCGLSpr.SetTextBinding .sprSht1,"AMT",1, strAMTSUM
		mobjSCGLSpr.SetTextBinding .sprSht1,"COMMISSION",1, strCOMMISSIONSUM
		'.txtSUM.value = gRound(strAMTSUM,0)
		'Call gFormatNumber(.txtSUM,0,true)
	End With

End Sub

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i, strint
	dim strYEARMON, strSEQ
	Dim lngchkCnt
	Dim intCnt
	Dim strCONTIDX
	Dim strCYEAR
	Dim strCMONTH
	
	lngchkCnt = 0
	with frmThis
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		for i = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMI_TRANS_NO",i) <> "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"TRU_TRANS_NO",i) <> "" Then
					gErrorMsgBox i & "���� �������� �ŷ������� ���� �մϴ�." & vbcrlf & "�켱 �ŷ����� �� ���� �Ͻʿ�","�����ȳ�!"
					exit Sub
				else 
					lngchkCnt = lngchkCnt +1
				end if
			end if
		next
		
		IF lngchkCnt = 0 THEN
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT SUB
		END IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				strCONTIDX = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTIDX",i)
				strCYEAR = mobjSCGLSpr.GetTextBinding(.sprSht,"CYEAR",i)
				strCMONTH = mobjSCGLSpr.GetTextBinding(.sprSht,"CMONTH",i)
			
				intRtn = mobjOUTDOOR.DeleteRtn(gstrConfigXml,strYEARMON, strSEQ, strCONTIDX, strCYEAR, strCMONTH)
					
				IF not gDoErrorRtn ("DeleteRtn") then
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   					
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
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
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;���� û�����</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" style="WIDTH: 1040px;" cellSpacing="0" cellPadding="0"
							width="792" border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON1, '')"
												width="70">�� ��</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtYEARMON1" title="��ȸ���" style="WIDTH: 87px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" size="9" name="txtYEARMON1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE) "
												width="70">������
											</TD>
											<TD class="SEARCHDATA" width="290"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 212px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="29" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 53px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_CODE, txtREAL_MED_NAME1)"
												width="70">��ü��
											</TD>
											<TD class="SEARCHDATA" width="290"><INPUT class="INPUT_L" id="txtREAL_MED_NAME1" title="�ڵ��" style="WIDTH: 190px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="26" name="txtREAL_MED_NAME1"><IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE"><INPUT class="INPUT" id="txtREAL_MED_CODE" title="�ڵ���ȸ" style="WIDTH: 55px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtREAL_MED_CODE">
											</TD>
											<TD class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 1040px; HEIGHT: 72px" vAlign="top" align="center">
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
														<td class="TITLE">&nbsp;&nbsp;����û�� ��ȸ �� ����</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="0" border="0">
													<TR>
														<td><IMG id="imgGETPOTALDATA" onmouseover="JavaScript:this.src='../../../images/imgGETPOTALDATAOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgGETPOTALDATA.gIF'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgGETPOTALDATA.gIF" align="right"
																border="0" name="imgGETPOTALDATA"></td>
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" align="right" border="0" name="imgSave"></td>
														<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" align="right" border="0"
																name="imgDelete"></td>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
									<table>
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 7px"><FONT face="����"></FONT></TD>
										</TR>
									</table>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="HEIGHT: 25px" width="70">��&nbsp; ��</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="200"><INPUT dataFld="YEARMON" class="NOINPUT" id="txtYEARMON" title="���" style="WIDTH: 88px; HEIGHT: 22px"
													accessKey="MON" dataSrc="#xmlBind" readOnly type="text" maxLength="6" size="9" name="txtYEARMON">&nbsp;<INPUT dataFld="SEQ" class="NOINPUT_R" id="txtSEQ" title="�Ϸù�ȣ" style="WIDTH: 48px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="2" name="txtSEQ">
											</TD>
											<TD class="LABEL" style="HEIGHT: 25px" width="70">���Ⱓ</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="200"><INPUT dataFld="TBRDSTDATE" class="NOINPUT" id="txtTBRDSTDATE" title="����Ⱓ" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="8" name="txtTBRDSTDATE">&nbsp;~
												<INPUT dataFld="TBRDEDDATE" class="NOINPUT" id="txtTBRDEDDATE" title="����Ⱓ" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="10" size="9" name="txtTBRDEDDATE"></TD>
											<TD class="LABEL" style="HEIGHT: 25px" width="70">�Ѱ��ݾ�</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="180"><INPUT dataFld="TOTALAMT" class="NOINPUT_R" id="txtTOTALAMT" title="�Ѱ��ݾ�" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="13" size="22" name="txtTOTALAMT"></TD>
											<TD class="LABEL" style="HEIGHT: 25px" width="70">��������</TD>
											<TD class="DATA" style="HEIGHT: 19pt" width="180"><INPUT dataFld="MED_GBN" class="NOINPUT_L" id="txtMED_GBN" title="��ü���ڵ�" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="24" name="txtMED_GBN"></TD>
										</TR>
										<TR>
											<TD class="LABEL">������</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="26" name="txtCLIENTNAME">
											</TD>
											<TD class="LABEL">��ü��</TD>
											<TD class="DATA" style="HEIGHT: 19pt"><INPUT dataFld="REAL_MED_NAME" class="NOINPUT_L" id="txtREAL_MED_NAME" title="��ü���" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="2" name="txtREAL_MED_NAME"></TD>
											<TD class="LABEL">��û���ݾ�</TD>
											<TD class="DATA"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="�ݾ�" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="13" size="1" name="txtAMT"></TD>
											<TD class="LABEL">���ֺ�</TD>
											<TD class="DATA"><INPUT dataFld="OUT_AMT" class="NOINPUT_L" id="txtOUT_AMT" title="���ֺ�" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100000" size="7" name="txtOUT_AMT"></TD>
										</TR>
										<TR>
											<TD class="LABEL">�귣��</TD>
											<TD class="DATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="��������" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="5" name="txtSUBSEQNAME"></TD>
											<TD class="LABEL">��&nbsp;&nbsp; ��</TD>
											<TD class="DATA"><INPUT dataFld="DEPT_NAME" class="NOINPUT_L" id="txtDEPT_NAME" title="���μ���" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="4" name="txtDEPT_NAME"></TD>
											<TD class="LABEL">��������</TD>
											<TD class="DATA"><INPUT dataFld="COMMI_RATE" class="NOINPUT_R" id="txtCOMMI_RATE" title="��������" style="WIDTH: 136px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="3" size="17" name="txtCOMMI_RATE">(%)</TD>
											<TD class="LABEL">������</TD>
											<TD class="DATA"><INPUT dataFld="COMMISSION" class="NOINPUT_R" id="txtCOMMISSION" title="������" style="WIDTH: 179px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="13" size="2" name="txtCOMMISSION"></TD>
										</TR>
										<TR>
											<TD class="LABEL">�����</TD>
											<TD class="DATA"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="����θ�" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="4" name="txtCLIENTSUBNAME"></TD>
											<TD class="LABEL">�����</TD>
											<TD class="DATA"><INPUT dataFld="PROGNAME" class="NOINPUT_L" id="txtPROGNAME" title="�����" style="WIDTH: 197px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100000" size="13" name="txtPROGNAME"></TD>
											<TD class="LABEL">���</TD>
											<TD class="DATA" colSpan="3"><INPUT dataFld="LOCATION" class="NOINPUT_L" id="txtLOCATION" title="���" style="WIDTH: 250px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="36" name="txtLOCATION"></TD>
										</TR>
									</TABLE>
									<table>
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 7px"><FONT face="����"></FONT></TD>
										</TR>
									</table>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<!--TD class="LABEL" style="WIDTH: 63px" width="63">��&nbsp;&nbsp; ��</TD>
													<TD class="DATA" style="WIDTH: 200px" width="200">&nbsp; <INPUT id="chkGBN_FLAG1" type="radio" value="1" name="chkGBN_FLAG">&nbsp;�Ϲݸ���&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="chkGBN_FLAG2" type="radio" value="0" name="chkGBN_FLAG">&nbsp;����</TD-->
											<TD class="LABEL" style="WIDTH: 62px" width="62">VAT����</TD>
											<TD class="DATA" style="WIDTH: 202px">&nbsp; <INPUT id="chkVATYES" type="radio" value="Y" name="chkVAT">&nbsp;��&nbsp;&nbsp; 
												&nbsp;<INPUT id="chkVATNO" type="radio" value="N" name="chkVAT">&nbsp;��</TD>
											<TD class="LABEL" style="WIDTH: 61px">���</TD>
											<TD class="DATA"><INPUT dataFld="NOTE" class="INPUT_L" id="txtNOTE" title="����" style="WIDTH: 428px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100000" size="65" name="txtNOTE"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
			<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
				</tr>
				<TR>
					<TD>
						<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 522px" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 498px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27464">
											<PARAM NAME="_ExtentY" VALUE="13176">
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
										<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
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
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
