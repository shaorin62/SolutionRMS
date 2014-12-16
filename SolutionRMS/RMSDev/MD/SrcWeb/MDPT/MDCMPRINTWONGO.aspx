<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMPRINTWONGO.aspx.vb" Inherits="MD.MDCMPRINTWONGO" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�μⱤ�� �������</title>
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
'HISTORY    :1) 2008/08/18 By Kim Tae Ho
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
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMPRINTWONGO, mobjMDCOGET
Dim mstrCheck
Dim mALLCHECK
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
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub


Sub imgClose_onclick ()
	Window_OnUnload
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
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
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
		'On error resume Next
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
'-----------------------------------------------------------------------------------------
' ��ü���˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,425)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,1))
					.txtREAL_MED_NAME.value = trim(vntData(1,1))
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'����ó
Sub txtCRE_NAME_onchange
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CRE_NAME",frmThis.sprSht.ActiveRow, frmThis.txtCRE_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'���ó
Sub txtDELIVER_NAME_onchange
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVER_NAME",frmThis.sprSht.ActiveRow, frmThis.txtDELIVER_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'���ð�
Sub cmbEND_TIME_onchange
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"END_TIME",frmThis.sprSht.ActiveRow, frmThis.cmbEND_TIME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'��������
Sub txtOLD_DATE_onchange
	DIM strdate 
	DIM strOLD_DATE
	strdate = ""
	strOLD_DATE =""
	
	strdate=frmThis.txtOLD_DATE.value
	'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
	if len(strdate) = 4 then
		strOLD_DATE = Mid(gNowDate2,1,4) & strdate
	elseif len(strdate) = 10 then
		strOLD_DATE = Mid(strdate,1 , 4) & Mid(strdate,6 , 2) & Mid(strdate,9 , 2)
	elseif len(strdate) = 3 then
		strOLD_DATE = Mid(gNowDate2,1,4) & "0" & strdate
	else
		strOLD_DATE = strdate
	end if
	
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OLD_DATE",frmThis.sprSht.ActiveRow, strOLD_DATE
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

'���
Sub txtNOTE_onchange
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"NOTE",frmThis.sprSht.ActiveRow, frmThis.txtNOTE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


'CONTACT_FLAG ����
Sub chkCONTACT_FLAG1_onClick
	Dim strCONTACT_FLAG
	Dim strCONTACT_FLAGNAME
	WITH frmThis
		IF .chkCONTACT_FLAG1.checked = TRUE THEN
			strCONTACT_FLAG = "Y"
			strCONTACT_FLAGNAME = "��"
		ELSEIF .chkCONTACT_FLAG1.checked = FALSE THEN
			strCONTACT_FLAG = "N"
			strCONTACT_FLAGNAME = "��"
		END IF
		
		if frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTACT_FLAG",frmThis.sprSht.ActiveRow, strCONTACT_FLAG
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTACT_FLAGNAME",frmThis.sprSht.ActiveRow, strCONTACT_FLAGNAME
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if		
	end with
End Sub

Sub chkCONTACT_FLAG2_onClick
	Dim strCONTACT_FLAG
	Dim strCONTACT_FLAGNAME
	WITH frmThis
		IF .chkCONTACT_FLAG2.checked = TRUE THEN
			strCONTACT_FLAG = "Y"
			strCONTACT_FLAGNAME = "��"
		ELSEIF .chkCONTACT_FLAG2.checked = FALSE THEN
			strCONTACT_FLAG = "N"
			strCONTACT_FLAGNAME = "��"
		END IF
		
		if frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTACT_FLAG",frmThis.sprSht.ActiveRow, strCONTACT_FLAG
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTACT_FLAGNAME",frmThis.sprSht.ActiveRow, strCONTACT_FLAGNAME
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if		
	end with
End Sub

'�ű� ����
Sub chkGUBUN1_onClick
	Dim strGUBUN
	Dim strGUBUN_NAME
	WITH frmThis
		IF .chkGUBUN1.checked = TRUE THEN
			strGUBUN = "N"
			strGUBUN_NAME = "��"
			document.getElementById("lblchange").innerHTML="������ð�"
			pnlEND_TIME.style.display = "inline"
			pnlOLD_DATE.style.display = "none"
			
			if .sprSht.MaxRows >0 then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, strGUBUN
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN_NAME",frmThis.sprSht.ActiveRow, strGUBUN_NAME
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OLD_DATE",frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
		ELSEIF .chkGUBUN1.checked = FALSE THEN
			strGUBUN = "O"
			strGUBUN_NAME = "��"
			document.getElementById("lblchange").innerHTML="��������"
			pnlEND_TIME.style.display = "none"
			pnlOLD_DATE.style.display = "inline"
			if .sprSht.MaxRows >0 then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, strGUBUN
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN_NAME",frmThis.sprSht.ActiveRow, strGUBUN_NAME
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"END_TIME",frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
		END IF
	end with
End Sub

Sub chkGUBUN2_onClick
	Dim strGUBUN
	Dim strGUBUN_NAME
	WITH frmThis
		IF .chkGUBUN2.checked = TRUE THEN
			strGUBUN = "O"
			strGUBUN_NAME = "��"
			document.getElementById("lblchange").innerHTML="��������"
			pnlEND_TIME.style.display = "none"
			pnlOLD_DATE.style.display = "inline"
			
			if .sprSht.MaxRows >0 then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, strGUBUN
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN_NAME",frmThis.sprSht.ActiveRow, strGUBUN_NAME
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"END_TIME",frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
		ELSEIF .chkGUBUN2.checked = FALSE THEN
			strGUBUN = "N"
			strGUBUN_NAME = "��"
			document.getElementById("lblchange").innerHTML="������ð�"
			pnlEND_TIME.style.display = "inline"
			pnlOLD_DATE.style.display = "none"
			if .sprSht.MaxRows >0 then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, strGUBUN
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN_NAME",frmThis.sprSht.ActiveRow, strGUBUN_NAME
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OLD_DATE",frmThis.sprSht.ActiveRow, ""
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			end if
		END IF
	end with
End Sub


Sub chkOUTFLAG_onClick
	Dim strGUBUN
	Dim strGUBUN_NAME
	WITH frmThis
		IF .chkOUTFLAG.checked = TRUE THEN
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUTFLAG",frmThis.sprSht.ActiveRow, "��"
		ELSE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUTFLAG",frmThis.sprSht.ActiveRow, ""
		END IF
		if frmThis.sprSht.ActiveRow >0 Then
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end with
End Sub

'****************************************************************************************
' �޷�
'****************************************************************************************
Sub imgCalEndar1_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtOLD_DATE,frmThis.imgCalEndar1,"txtOLD_DATE_onchange()"
	if frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OLD_DATE",frmThis.sprSht.ActiveRow, frmThis.txtOLD_DATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
End Sub
'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row > 0 and Col > 0 then		
			sprShtToFieldBinding Col,Row
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
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
	set mobjMDCMPRINTWONGO	= gCreateRemoteObject("cMDPT.ccMDPTPRINTWONGO")
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
		
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 25, 0, 1, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht,   "PUB_DATE | MEDNAME  | CLIENTNAME  | MATTERNAME   | STD_STEP  | STD_CM   | COL_DEG  | PUB_FACE  | YEARMON  | SEQ  | CLIENTCODE  | MEDCODE  | REAL_MED_CODE | MED_FLAG  | END_TIME  | OLD_DATE  | CRE_NAME  | DELIVER_NAME  | CONTACT_FLAGNAME  | CONTACT_FLAG  | GUBUN_NAME  | GUBUN  | OUTFLAG  | NOTE  | WONGOYEARMON"
		mobjSCGLSpr.SetHeader .sprSht,		   "������|��ü��|������|�����|��|CM|����|û���|���|����|�������ڵ�|��ü�ڵ�|��ü���ڵ�|��ü����|���ð�|������|��������ó|�������ó|��������|���������ڵ�|��/��|�ű��ڵ�|��|���|������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "     8|	   12|    12|    13| 4| 5|   4|     8|   0|   0|         0|       0|         0|       0|       7|     8|         9|         9|       5|           0|    5|       0| 5|  11|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "18"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PUB_DATE | OLD_DATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "STD_CM", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | STD_STEP", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "MEDNAME  | CLIENTNAME  | MATTERNAME   | STD_STEP  | STD_CM   | COL_DEG  | PUB_FACE  | YEARMON  | SEQ  | CLIENTCODE  | MEDCODE  | REAL_MED_CODE | MED_FLAG  | END_TIME  | OLD_DATE  | CRE_NAME  | DELIVER_NAME  | CONTACT_FLAGNAME  | CONTACT_FLAG  | GUBUN_NAME  | GUBUN  | OUTFLAG  | NOTE  | WONGOYEARMON", -1, -1, 100
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PUB_DATE | MEDNAME  | CLIENTNAME  | MATTERNAME   | STD_STEP  | STD_CM   | COL_DEG  | PUB_FACE  | YEARMON  | SEQ  | CLIENTCODE  | MEDCODE  | REAL_MED_CODE | MED_FLAG  | END_TIME  | OLD_DATE  | CRE_NAME  | DELIVER_NAME  | CONTACT_FLAGNAME  | CONTACT_FLAG  | GUBUN_NAME  | GUBUN  | OUTFLAG  | NOTE  | WONGOYEARMON"
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | CLIENTCODE | MEDCODE | REAL_MED_CODE", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "MEDNAME | CLIENTNAME | MATTERNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "OUTFLAG",-1,-1,2,2,false
			
		.sprSht.style.visibility = "visible"	
    End With    

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
End Sub

Sub EndPage()
	set mobjMDCMPRINTWONGO = Nothing
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
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
		document.getElementById("lblchange").innerHTML="������ð�"
		pnlEND_TIME.style.display = "inline"
		pnlOLD_DATE.style.display = "none"
		.sprSht.MaxRows = 0	
		
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strREAL_MED_NAME
	Dim strMED_FLAG
	Dim strGFLAG
	strCLIENTCODE = ""
	strCLIENTNAME = ""
	
	with frmThis
   		'������ Validation
   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PUB_DATE | MEDNAME  | CLIENTNAME  | MATTERNAME   | STD_STEP  | STD_CM   | COL_DEG  | PUB_FACE  | YEARMON  | SEQ  | CLIENTCODE  | MEDCODE  | REAL_MED_CODE | MED_FLAG  | END_TIME  | OLD_DATE  | CRE_NAME  | DELIVER_NAME  | CONTACT_FLAGNAME  | CONTACT_FLAG  | GUBUN_NAME  | GUBUN  | OUTFLAG  | NOTE  | WONGOYEARMON")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strYEARMON		= .txtYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		strREAL_MED_CODE= .txtREAL_MED_CODE.value
		strREAL_MED_NAME= .txtREAL_MED_NAME.value
		strMED_FLAG		= .cmbMED_FLAG.value
		strGFLAG		= .cmbGFLAG.value
		
		
		intRtn = mobjMDCMPRINTWONGO.ProcessRtn(gstrConfigXml,strMasterData,vntData)
   		
   		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			InitPageData
			gOkMsgBox "�������� ����Ǿ����ϴ�..","Ȯ��"
			
			If intRtn <> 0  Then
				.txtYEARMON.value = strYEARMON
				.txtCLIENTCODE.value = strCLIENTCODE
				.txtCLIENTNAME.value = strCLIENTNAME
				.txtREAL_MED_CODE.value = strREAL_MED_CODE
				.txtREAL_MED_NAME.value = strREAL_MED_NAME
				.cmbMED_FLAG.value = strMED_FLAG
				.cmbGFLAG.value = strGFLAG
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
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim vntData2
	Dim strYEARMON, strCLIENTCODE,strCLIENTNAME, strMED_FLAG
	Dim strREAL_MED_CODE, strREAL_MED_NAME
	Dim strGFLAG
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ txtYEARMON
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		strREAL_MED_CODE= .txtREAL_MED_CODE.value
		strREAL_MED_NAME= .txtREAL_MED_NAME.value
		strMED_FLAG		= .cmbMED_FLAG.value
		strGFLAG		= .cmbGFLAG.value
		
		
		vntData = mobjMDCMPRINTWONGO.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strCLIENTCODE, strREAL_MED_CODE,strMED_FLAG, strGFLAG)

		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt >0 then
				Call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				sprShtToFieldBinding 2, 1
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   				PreSearchFiledValue strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strMED_FLAG, strGFLAG
   			end if
   		end if
   	end with
End Sub



Sub PreSearchFiledValue (strYEARMON,strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE,strREAL_MED_NAME, strMED_FLAG, strGFLAG)
	frmThis.txtYEARMON.value = strYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME.value = strCLIENTNAME
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
	frmThis.cmbMED_FLAG.value = strMED_FLAG
	frmThis.cmbGFLAG.value = strGFLAG
End Sub


Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"OLD_DATE",Row) ="O" THEN
			pnlEND_TIME.style.display = "none"
			pnlOLD_DATE.style.display = "inline"
		ELSE
			pnlEND_TIME.style.display = "inline"
			pnlOLD_DATE.style.display = "none"
		END IF
		
		
		.txtCRE_NAME.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"CRE_NAME",Row)
		.txtDELIVER_NAME.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVER_NAME",Row)
		.cmbEND_TIME.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"END_TIME",Row)
		.txtOLD_DATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"OLD_DATE",Row)
		.txtNOTE.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"NOTE",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONTACT_FLAG",Row) = "Y" THEN
			.chkCONTACT_FLAG1.checked = TRUE
			.chkCONTACT_FLAG2.checked = FALSE
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"CONTACT_FLAG",Row) = "N" THEN
			.chkCONTACT_FLAG1.checked = FALSE
			.chkCONTACT_FLAG2.checked = TRUE
		ELSE
			.chkCONTACT_FLAG1.checked = FALSE
			.chkCONTACT_FLAG2.checked = FALSE
		END IF
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",Row) = "N" THEN
			.chkGUBUN1.checked = TRUE
			.chkGUBUN2.checked = FALSE
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",Row) = "O" THEN
			.chkGUBUN1.checked = FALSE
			.chkGUBUN2.checked = TRUE
		ELSE
			.chkGUBUN1.checked = FALSE
			.chkGUBUN2.checked = FALSE
		END IF
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"OUTFLAG",Row) = "��" THEN
			.chkOUTFLAG.checked = TRUE
		ELSE
			.chkOUTFLAG.checked = FALSE
		END IF
		
		
   	end with
End Function

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON
	Dim strSEQ

	with frmThis
	
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit Sub
		End IF
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",vntData(i))
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)))
				
				intRtn = mobjMDCMPRINTWONGO.DeleteRtn(gstrConfigXml,strYEARMON, strSEQ)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				'mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				gWriteText "", "[" & strSEQ & "] �ڷ��� �������� �����Ǿ����ϴ�."
   			End IF
		next
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
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
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="85" background="../../../images/back_p.gIF"
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
											<td class="TITLE">�μ� �������</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 302px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="90">���</TD>
											<TD class="SEARCHDATA" width="350"><INPUT class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 98px; HEIGHT: 22px" accessKey="MON"
													type="text" maxLength="6" size="6" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" width="90">��ü����
											</TD>
											<td class="SEARCHDATA" colSpan="2"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
													align="right" border="0" name="imgQuery"><SELECT id="cmbMED_FLAG" title="��������" style="WIDTH: 80px" name="cmbMED_FLAG">
													<OPTION value="MP01" selected>�Ź�</OPTION>
													<OPTION value="MP02">����</OPTION>
												</SELECT><SELECT id="cmbGFLAG" title="��������" style="WIDTH: 120px" name="cmbGFLAG">
													<OPTION value="" selected>���౸��-��ü</OPTION>
													<OPTION value="M">����</OPTION>
													<OPTION value="B">����</OPTION>
													<OPTION value="J">����</OPTION>
													<OPTION value="S">����</OPTION>
												</SELECT>
											</td>
										</TR>
										<tr>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 193px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="30" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF"  align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 55px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="3" name="txtCLIENTCODE"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_CODE, txtREAL_MED_NAME)">��ü��
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtREAL_MED_NAME" title="�ڵ��" style="WIDTH: 193px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="26" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE"> <INPUT class="INPUT" id="txtREAL_MED_CODE" title="�ڵ���ȸ" style="WIDTH: 55px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" name="txtREAL_MED_CODE"></TD>
										</tr>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle"></td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="ImgSave" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gif'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/ImgSave.gif'" height="20" alt="���泻���� �����մϴ�."
																src="../../../images/ImgSave.gif" align="absMiddle" border="0" name="ImgSave"></td>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
																name="imgDelete"></TD>
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
											<TD class="BODYSPLIT" style="WIDTH: 100%"></TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD class="LABEL" width="90">��������ó</TD>
											<TD class="DATA" width="257"><INPUT dataFld="CRE_NAME" class="INPUT_L" id="txtCRE_NAME" title="�����ָ�" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="37" name="txtCRE_NAME">
											</TD>
											<TD class="LABEL" width="90">��/�� ����</TD>
											<TD class="DATA" width="256">&nbsp;&nbsp;&nbsp;<INPUT id="chkGUBUN1" type="radio" value="N" name="chkGUBUN">
												&nbsp;�Ű�&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="chkGUBUN2" type="radio" value="O" name="chkGUBUN">&nbsp;����</TD>
											<TD class="LABEL" id="lblchange" width="90" onclick="vbscript:IF pnlOLD_DATE.style.visibility = 'visible' THEN  Call gCleanField(txtOLD_DATE, '') END IF"></TD>
											<TD class="DATA">
												<DIV id="pnlEND_TIME" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><SELECT id="cmbEND_TIME" title="��������" style="WIDTH: 108px" name="cmbEND_TIME">
														<OPTION value="" selected></OPTION>
														<OPTION value="00:00">00:00</OPTION>
														<OPTION value="01:00">01:00</OPTION>
														<OPTION value="02:00">02:00</OPTION>
														<OPTION value="03:00">03:00</OPTION>
														<OPTION value="04:00">04:00</OPTION>
														<OPTION value="05:00">05:00</OPTION>
														<OPTION value="06:00">06:00</OPTION>
														<OPTION value="07:00">07:00</OPTION>
														<OPTION value="08:00">08:00</OPTION>
														<OPTION value="09:00">09:00</OPTION>
														<OPTION value="10:00">10:00</OPTION>
														<OPTION value="11:00">11:00</OPTION>
														<OPTION value="12:00">12:00</OPTION>
														<OPTION value="13:00">13:00</OPTION>
														<OPTION value="14:00">14:00</OPTION>
														<OPTION value="15:00">15:00</OPTION>
														<OPTION value="16:00">16:00</OPTION>
														<OPTION value="17:00">17:00</OPTION>
														<OPTION value="18:00">18:00</OPTION>
														<OPTION value="19:00">19:00</OPTION>
														<OPTION value="20:00">20:00</OPTION>
														<OPTION value="21:00">21:00</OPTION>
														<OPTION value="22:00">22:00</OPTION>
														<OPTION value="23:00">23:00</OPTION>
														<OPTION value="24:00">24:00</OPTION>
													</SELECT></DIV>
												<DIV id="pnlOLD_DATE" style="DISPLAY: none; WIDTH: 200px; POSITION: relative; HEIGHT: 24px"
													ms_positioning="GridLayout"><INPUT dataFld="" class="INPUT" id="txtOLD_DATE" title="������" style="WIDTH: 100px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="" type="text" maxLength="10" name="txtOLD_DATE"><IMG id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
														border="0" name="imgCalEndar1"></DIV>
											</TD>
										</TR>
										<TR>
											<TD class="LABEL">�������ó</TD>
											<TD class="DATA"><INPUT dataFld="DELIVER_NAME" class="INPUT_L" id="txtDELIVER_NAME" title="�����ָ�" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="37" name="txtDELIVER_NAME">
											</TD>
											<TD class="LABEL">��������</TD>
											<TD class="DATA">&nbsp;&nbsp;&nbsp;<INPUT id="chkCONTACT_FLAG1" type="radio" value="Y" name="chkCONTACT_FLAG">
												&nbsp;�� &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT id="chkCONTACT_FLAG2" type="radio" value="N" name="chkCONTACT_FLAG">&nbsp;��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
												&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <INPUT dataFld="OUTFLAG" id="chkOUTFLAG" title="����" dataSrc="#xmlBind" type="checkbox"
													name="chkOUTFLAG">&nbsp;��</TD>
											<TD class="LABEL">���</TD>
											<TD class="DATA"><INPUT dataFld="NOTE" class="INPUT_L" id="txtNOTE" title="���" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" type="text" maxLength="100" size="37" name="txtNOTE"></TD>
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
										<PARAM NAME="_ExtentX" VALUE="31856">
										<PARAM NAME="_ExtentY" VALUE="13256">
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
					<TD class="BOTTOMSPLIT" id="lblStatus"></TD>
				</TR>
			</TABLE>
			</TD></TR></TABLE></FORM>
	</body>
</HTML>
