<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRICLIST.aspx.vb" Inherits="MD.MDCMELECTRICLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����Ź ����ó��</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ó�� ǥ�� ����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
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
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCMELECTRICLIST 
Dim mobjMDCMGET
Dim mstrCheck
Dim mstrGUBUN
mstrCheck = True
mstrGUBUN = "KOBACO"
Dim mstrGFLAG
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
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�",""
		exit Sub
	end if
	CALL SelectRtn_PRESUSU (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub imgSetting_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmCancel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ImgFind_onclick()
	with frmThis
		initpageData
		if mstrGUBUN = "KOBACO" then
			.sprSht.MaxRows = 400
			msgbox "������ ��Ŀ� �°� �����͸� �ٿ� ��������." & vbcrlf & "����� �ϴ��� �Ѱ� �÷��� �߰� �Ͻø� �ȵ˴ϴ�.."
		else
			.sprSht_SBS.MaxRows = 400
			msgbox "������ ��Ŀ� �°� �����͸� �ٿ� ��������." & vbcrlf & "����� �ϴ��� �Ѱ� �÷��� �߰� �Ͻø� �ȵ˴ϴ�.."
		end if
	End with
End Sub

Sub ImgSave_onclick
	gFlowWait meWAIT_ON
	if mstrGUBUN = "KOBACO" then
		call ProcessRtn_TEMP(frmThis.sprSht)
	else
		call ProcessRtn_TEMP(frmThis.sprSht_SBS)
	end if
	
	gFlowWait meWAIT_OFF	
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
			CALL SelectRtn_PRESUSU (mstrGUBUN)
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
			vntData = mobjMDCMGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					CALL SelectRtn_PRESUSU (mstrGUBUN)
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtYEARMON_onkeydown
	If window.event.keyCode = meEnter Then
		CALL SelectRtn_PRESUSU (mstrGUBUN)
		frmThis.txtCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'��ó�� (�ڹ���)
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	frmThis.btnTab2.style.backgroundImage = meURL_TAB
		
	pnlTab_KOBACO.style.visibility = "visible" 
	pnlTab_SBS.style.visibility = "hidden" 	
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "KOBACO"
	CALL SelectRtn_PRESUSU (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'��ó�� (SBS)
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgroundImage = meURL_TAB
	frmThis.btnTab2.style.backgroundImage = meURL_TABON
	
	pnlTab_KOBACO.style.visibility = "hidden" 
	pnlTab_SBS.style.visibility = "visible" 
	
		
	gFlowWait meWAIT_ON
	mstrGUBUN = "SBS"
	CALL SelectRtn_PRESUSU (mstrGUBUN)
	gFlowWait meWAIT_OFF
	
	mobjSCGLCtl.DoEventQueue
End Sub

'�������� ��Ʈ �̺�Ʈ
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

sub sprSht_SBS_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_SBS, ""
		end if
	end with
end sub

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
	set mobjMDCMELECTRICLIST = gCreateRemoteObject("cMDET.ccMDETELECTRICLIST")
	set mobjMDCMGET			 = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		.imgSetting.disabled = True
		.ImgConfirmCancel.disabled = True
	   
		InitPageData
		
		btnTab1_onclick
	end with
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'�ʱ� ������ ����
	with frmThis
		Grid_init
		gridLayOut
		.sprSht.MaxRows = 0	
		.sprSht_SBS.MaxRows = 0	
		
		.txtYEARMON.value =  Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		.txtYEARMON.focus()
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub EndPage()
	set mobjMDCMELECTRICLIST = Nothing
	set mobjMDCMGET = Nothing
	gEndPage
End Sub

Sub Grid_init ()
	Dim intCnt
	with frmThis
		'KOBACO �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht, "TEMPSEQ"
		mobjSCGLSpr.SetHeader .sprSht,		 ""
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " "
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "TEMPSEQ", -1, -1, 20
		mobjSCGLSpr.SetCellAlign2 .sprSht, "TEMPSEQ",-1,-1,2,2,false
		
		'SBS�׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_SBS
		mobjSCGLSpr.SpreadLayout .sprSht_SBS, 1, 0, 0, 0,5
		mobjSCGLSpr.SpreadDataField .sprSht_SBS, "TEMPSEQ"
		mobjSCGLSpr.SetHeader .sprSht_SBS,		 ""
		mobjSCGLSpr.SetColWidth .sprSht_SBS, "-1", " "
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "0", "20"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SBS, "TEMPSEQ", -1, -1, 20
		mobjSCGLSpr.SetCellAlign2 .sprSht_SBS, "TEMPSEQ",-1,-1,2,2,false
	End With
End Sub

Sub gridLayOut
	mstrGFLAG = "T"
	With frmThis
		'KOBACO �׸���
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0
		mobjSCGLSpr.SpreadDataField .sprSht,   "TEMPSEQ | DIVFLAG | KOBACOCODE | CLIENTNAME | MGBN | TOT | M140 | M144 | M142 | M141 | M143 | M145 | YEARMON | SEQ | CLIENTCODE | ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|����ǿ�|�������ڵ�|�����ָ�|��Ź����|�հ�|����|�λ�����|�뱸����|��������|��������|��������|���|SEQ|CLIENTCODE|����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|8       |9         |14      |8       |12  |11  |11      |11      |      11|      11|      11|   0|  0|       0  |30"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOT | M140 | M144 | M142 | M141 | M143 | M145", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, false, "TOT | M140 | M144 | M142 | M141 | M143 | M145"
		mobjSCGLSpr.SetCellTypeEdit2    .sprSht, " TEMPSEQ | DIVFLAG | KOBACOCODE | CLIENTNAME | MGBN | ERRMSG", , ,200
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | CLIENTCODE", true
		mobjSCGLSpr.ColHidden .sprSht, "TEMPSEQ | KOBACOCODE | ERRMSG", false
		
		'SBS�׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_SBS
		mobjSCGLSpr.SpreadLayout .sprSht_SBS, 10, 0
		mobjSCGLSpr.SpreadDataField .sprSht_SBS,   "TEMPSEQ | DIVFLAG | KOBACOCODE | CLIENTNAME | MGBN | TOT | YEARMON | SEQ | CLIENTCODE | ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht_SBS,		   "����|����ǿ�|�������ڵ�|�����ָ�|��Ź����|�հ�|���|SEQ|CLIENTCODE|����"
		mobjSCGLSpr.SetColWidth .sprSht_SBS, "-1", "   4|8       |20        |14      |8       |12  |   0|  0|       0  |30"
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SBS, "TOT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_SBS, false, "TOT"
		mobjSCGLSpr.SetCellTypeEdit2    .sprSht_SBS, " TEMPSEQ | DIVFLAG | KOBACOCODE | CLIENTNAME | MGBN | ERRMSG", , ,200
		mobjSCGLSpr.ColHidden .sprSht_SBS, "YEARMON | SEQ | CLIENTCODE", true
		mobjSCGLSpr.ColHidden .sprSht_SBS, "TEMPSEQ | KOBACOCODE | ERRMSG", false
    End With
End Sub

Sub gridSelectLayOut
	mstrGFLAG = "F"
	With frmThis
		'KOBACO �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 16, 0
		mobjSCGLSpr.SpreadDataField .sprSht,   "TEMPSEQ | KOBACOCODE | CLIENTNAME | DIVFLAG | MGBN | TOT | M140 | M144 | M142 | M141 | M143 | M145 | YEARMON | SEQ | CLIENTCODE | ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|�������ڵ�|�����ָ�|����ǿ�|��Ź����|�հ�|����|�λ�����|�뱸����|��������|��������|��������|���|SEQ|CLIENTCODE|����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   0|0         |28      |8       |8       |12  |11  |11      |11      |      11|      11|      11|   0|  0|       0  |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "TOT | M140 | M144 | M142 | M141 | M143 | M145"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "TOT | M140 | M144 | M142 | M141 | M143 | M145", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " DIVFLAG | CLIENTNAME | MGBN", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON | SEQ | CLIENTCODE | TEMPSEQ | KOBACOCODE | ERRMSG", true
		mobjSCGLSpr.CellGroupingEach .sprSht, "CLIENTNAME | DIVFLAG"
		
		'SBS�׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_SBS
		mobjSCGLSpr.SpreadLayout .sprSht_SBS, 10, 0
		mobjSCGLSpr.SpreadDataField .sprSht_SBS,   "TEMPSEQ | KOBACOCODE | CLIENTNAME | DIVFLAG | MGBN | TOT | YEARMON | SEQ | CLIENTCODE | ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht_SBS,		   "����|�������ڵ�|�����ָ�|����ǿ�|��Ź����|�հ�|���|SEQ|CLIENTCODE|����"
		mobjSCGLSpr.SetColWidth .sprSht_SBS, "-1", "   0|0         |28      |8       |10      |18  |   0|  0|       0  |0"
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_SBS, "0", "15"
		mobjSCGLSpr.SetCellsLock2 .sprSht_SBS, true, "TOT"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SBS, "TOT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht_SBS, " DIVFLAG | CLIENTNAME | MGBN", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_SBS, "YEARMON | SEQ | CLIENTCODE | TEMPSEQ | KOBACOCODE | ERRMSG", true
		mobjSCGLSpr.CellGroupingEach .sprSht_SBS, "CLIENTNAME | DIVFLAG"
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn_PRESUSU (strGUBUN)
	Dim strYEARMON
	Dim vntData
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim IngCOMMITColCnt, IngCOMMITRowCnt
	Dim intCnt, intCnt2
	
	with frmThis
		gridSelectLayOut
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		
		If strYEARMON = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		strCLIENTCODE = .txtCLIENTCODE.value
		strCLIENTNAME = .txtCLIENTNAME.value 
		
		Call PreConfirm ()
		
		vntData = mobjMDCMELECTRICLIST.SelectRtn_PRESUSU(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTCODE,strCLIENTNAME, strGUBUN)
		
		If not gDoErrorRtn ("SelectRtn_PRESUSU") then
			if mstrGUBUN = "KOBACO" then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			else
				mobjSCGLSpr.SetClipBinding .sprSht_SBS, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			end if
			
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		
   		
   		if mstrGUBUN = "KOBACO" then
			For intCnt = 1 To .sprSht.MaxRows	
				For intCnt2 = 6 To .sprSht.MaxCols-2
					if mobjSCGLSpr.GetTextBinding(.sprSht, intCnt2, intCnt) <> "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, intCnt2,intCnt2, intCnt, intCnt, rgb(255,173,173), False   'CHOOSE THE PINK..
					End if
				Next	
			Next	
		else
			For intCnt = 1 To .sprSht_SBS.MaxRows	
				For intCnt2 = 6 To .sprSht_SBS.MaxCols-2
					if mobjSCGLSpr.GetTextBinding(.sprSht_SBS, intCnt2, intCnt) <> "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht_SBS, intCnt2,intCnt2, intCnt, intCnt, rgb(255,173,173), False   'CHOOSE THE PINK..
					End if
				Next	
			Next	
		end if
		
   				
	end with   	
End SUb

Sub PreConfirm
	Dim vntDataConfirm
	Dim IngCOMMITColCnt, IngCOMMITRowCnt
	Dim strYEARMON
		
	with frmThis
		IngCOMMITColCnt=clng(0) : IngCOMMITRowCnt=clng(0)
		
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		strYEARMON		= .txtYEARMON.value
		vntDataConfirm = mobjMDCMELECTRICLIST.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON, mstrGUBUN)
		
		'Ȯ���� �Ǿ��ִٸ�
		If IngCOMMITRowCnt > 0 Then
			.ImgFind.disabled = true
			.imgSave.disabled = true
			.imgSetting.disabled = true
			.ImgConfirmCancel.disabled = false
		Else
			.ImgFind.disabled = false
			.imgSave.disabled = false	
			.imgSetting.disabled = false
			.ImgConfirmCancel.disabled = true
		End If
	End With
End Sub

Sub ProcessRtn_TEMP(sprSht)
	Dim intCnt
	Dim intCnt2
	Dim intCnt3
	Dim vntData
	Dim vntData1
	Dim strYEARMON
	Dim intRtn
	Dim intCnt4
	Dim strGUBUN
	with frmThis
		If mstrGFLAG = "F" Then
			msgbox "�ʱ�ȭ��ư �� �����ð� ������ �����͸� �ٿ������ʽÿ�." & vbcrlf & "������ ��ȸ���¿����� �������� �����ϴ�."
			Exit Sub
		End IF
		
		
		For intCnt = 1 to sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(sprSht,5,intCnt) = "" then
				mobjSCGLSpr.DeleteRow sprSht,intCnt
			Else
				mobjSCGLSpr.SetTextBinding sprSht,"YEARMON",intCnt, .txtYEARMON.value
				mobjSCGLSpr.SetTextBinding sprSht,"SEQ",intCnt, intCnt
			End If
		Next
		
		For intCnt2 = 1 to sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(sprSht,"DIVFLAG",intCnt2) = "" then
				mobjSCGLSpr.SetTextBinding sprSht,"DIVFLAG",intCnt2, mobjSCGLSpr.GetTextBinding(sprSht,"DIVFLAG",intCnt2-1)
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"DIVFLAG",intCnt2) = "" then
				mobjSCGLSpr.SetTextBinding sprSht,"DIVFLAG",intCnt2, mobjSCGLSpr.GetTextBinding(sprSht,"DIVFLAG",intCnt2-1)
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"TEMPSEQ",intCnt2) = "" then
				mobjSCGLSpr.SetTextBinding sprSht,"TEMPSEQ",intCnt2, mobjSCGLSpr.GetTextBinding(sprSht,"TEMPSEQ",intCnt2-1)
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"KOBACOCODE",intCnt2) = "" then
				mobjSCGLSpr.SetTextBinding sprSht,"KOBACOCODE",intCnt2, mobjSCGLSpr.GetTextBinding(sprSht,"KOBACOCODE",intCnt2-1)
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",intCnt2) = "" then
				mobjSCGLSpr.SetTextBinding sprSht,"CLIENTNAME",intCnt2, mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",intCnt2-1)
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",intCnt2) = "�Ѱ�" then
				mobjSCGLSpr.SetTextBinding sprSht,"DIVFLAG",intCnt2, "�Ѱ�"
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",intCnt2) = "�Ѱ�" then
				mobjSCGLSpr.SetTextBinding sprSht,"CLIENTNAME",intCnt2, "�Ѱ�"
			End If
			If mobjSCGLSpr.GetTextBinding(sprSht,"CLIENTNAME",intCnt2) = "�Ѱ�" then
				mobjSCGLSpr.SetTextBinding sprSht,"KOBACOCODE",intCnt2, "�Ѱ�"
			End If
		Next
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		For intCnt3 = 1 To sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(sprSht,"KOBACOCODE",intCnt3) <> "�Ѱ�" then
				vntData = mobjMDCMELECTRICLIST.GetClient(gstrConfigXml,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(sprSht,"KOBACOCODE",intCnt3))
				if mlngRowCnt = 0 Then
					mobjSCGLSpr.SetTextBinding sprSht,"ERRMSG",intCnt3, "KOBACO�������ڵ� �� ����Ͻʽÿ�" 
				Else
					mobjSCGLSpr.SetTextBinding sprSht,"CLIENTCODE",intCnt3, vntData(0,1)
				end If
			end if
		Next
		
		'MD_ELECTRIC_PRESUSU ����
		mobjSCGLSpr.SetFlag  sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸��� �����ϴ�.","�������"
   			Exit Sub
   		End If
   		
		For intCnt4 = 1 To sprSht.maxRows
			If mobjSCGLSpr.GetTextBinding(sprSht,"ERRMSG",intCnt4) <> "" Then
			gErrorMsgbox "�����׸��� Ȯ���Ͻʽÿ�.","����ȳ�!"
			Exit Sub
			End If
		Next
		
		'ó�� ������ü ȣ��
		strYEARMON = .txtYEARMON.value
		
		
		if mstrGUBUN = "KOBACO" then
			strGUBUN = ""
			vntData1 = mobjSCGLSpr.GetDataRows(sprSht,"YEARMON | SEQ | DIVFLAG | CLIENTCODE | MGBN | TOT | M140 | M144 | M142 | M141 | M143 | M145")
		else
			strGUBUN = "SBS"
			vntData1 = mobjSCGLSpr.GetDataRows(sprSht,"YEARMON | SEQ | DIVFLAG | CLIENTCODE | MGBN | TOT")
		end if
		
		intRtn = mobjMDCMELECTRICLIST.ProcessRtn_PRESUSU(gstrConfigXml,vntData1,strYEARMON, strGUBUN)

		if not gDoErrorRtn ("ProcessRtn_PRESUSU") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  sprSht,meCLS_FLAG
			'InitPageData
			gOkMsgBox "����� �����ڷᰡ ���� �Ǿ����ϴ�.","Ȯ��"
			CALL SelectRtn_PRESUSU (mstrGUBUN)
   		end if
	End with 
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn (strGUBUN)
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strYEARMON
	Dim intCnt
	Dim intYNRtn
	with frmThis
		'�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS
		
		IF strGUBUN = "KOBACO" THEN 
			If .sprSht.MaxRows = 0 Then
   				gErrorMsgBox "���׸��� �����ϴ�.","Ȯ������"
   				Exit Sub
   			End If
   		ELSE
   			If .sprSht_SBS.MaxRows = 0 Then
   				gErrorMsgBox "���׸��� �����ϴ�.","Ȯ������"
   				Exit Sub
   			End If
		END IF 
   		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strYEARMON = .txtYEARMON.value
		
		IF strGUBUN = "KOBACO" THEN 
			intYNRtn = gYesNoMsgbox("�ڹ��� �����͸� Ȯ�� �Ͻðڽ��ϱ�?","Ȯ��Ȯ��")
			IF intYNRtn <> vbYes then exit Sub
		ELSE
			intYNRtn = gYesNoMsgbox("SBS �����͸� Ȯ�� �Ͻðڽ��ϱ�?","Ȯ��Ȯ��")
			IF intYNRtn <> vbYes then exit Sub
		END IF 
		
		intRtn = mobjMDCMELECTRICLIST.ProcessRtn(gstrConfigXml, strMasterData,vntData,strYEARMON, mstrGUBUN)

		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			'InitPageData
			gOkMsgBox "����� �����ڷᰡ Ȯ���Ǿ����ϴ�.","Ȯ��"
			CALL SelectRtn_PRESUSU (mstrGUBUN)
   		end if
   	end with
End Sub

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON

	with frmThis
		intSelCnt = 0
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		IF mstrGUBUN = "KOBACO" THEN 
			If .sprSht.MaxRows = 0 Then
   				gErrorMsgBox "���׸��� �����ϴ�.","Ȯ������"
   				Exit Sub
   			End If
   		ELSE
   			If .sprSht_SBS.MaxRows = 0 Then
   				gErrorMsgBox "���׸��� �����ϴ�.","Ȯ������"
   				Exit Sub
   			End If
		END IF 
   		
   		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
   		strYEARMON = .txtYEARMON.value
   		intRtn = mobjMDCMELECTRICLIST.SelectRtn_CANCEL(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, mstrGUBUN)
   		if mlngRowCnt > 0 then
   			gErrorMsgBox "�ŷ������� ������ �����ʹ� Ȯ����Ұ� �ȵ˴ϴ�." & vbcrlf & "����Ͻ÷��� �ش����� ��� �ŷ������� ���� �Ͻʽÿ�.","Ȯ����ҿ���"
   			Exit Sub
   		end if
		
		intRtn = gYesNoMsgbox("Ȯ���� ����� �����ڷḦ ����Ͻðڽ��ϱ�?","Ȯ����� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		strYEARMON = .txtYEARMON.value
	
		intRtn = mobjMDCMELECTRICLIST.DeleteRtn(gstrConfigXml,strYEARMON, mstrGUBUN)
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gOkMsgBox  "Ȯ���� ����� �����ڷᰡ ��ҵǾ����ϴ�.","Ȯ��"
			CALL SelectRtn_PRESUSU (mstrGUBUN)
   		End IF
	End with
	err.clear	
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<P dir="ltr" style="MARGIN-RIGHT: 0px">
				<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 98%" cellSpacing="0" cellPadding="0" border="0">
					<TR>
						<TD>
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
								border="0">
								<TR>
									<TD align="left" width="400" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="163" background="../../../images/back_p.gIF"
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
												<td class="TITLE">���� �����ڷ� ���� �� Ȯ��</td>
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
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 183px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="50" border="0">
											<TR>
												<TD><!--<IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose">--></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<!--���̺��� �������°��� �����ش�-->
							<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
								<TR>
									<TD align="left" width="100%" height="1"></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD style="WIDTH: 100%" vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">�� 
													��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 65px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 64px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="5" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 77px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">�����ָ�</TD>
												<TD class="SEARCHDATA" width="313"><INPUT dataFld="CLIENTNAME" class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 224px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" size="32" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0"
														name="ImgCLIENTCODE"> <INPUT dataFld="CLIENTCODE" class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 64px; HEIGHT: 22px"
														accessKey=",M" dataSrc="#xmlBind" type="text" size="5" name="txtCLIENTCODE"></TD>
												<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 25px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 100%">
							<!--�׽�Ʈ ����-->
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
											type="button" value="KOBACO" name="btnTab1"> <INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
											type="button" size="20" value="SBS" name="btnTab2">
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
											<TR>
												<TD><IMG id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" height="20" alt="Loading"
														src="../../../images/imgCho.gif" width="64" border="0" name="imgFind"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
												<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
														height="20" alt="Ȯ���մϴ�." src="../../../images/imgSetting.gIF" width="54" border="0"
														name="imgSetting"></TD>
												<TD><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gif'"
														height="20" alt="Ȯ������մϴ�." src="../../../images/ImgConfirmCancel.gIF" border="0"
														name="ImgConfirmCancel"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<!--�׽�Ʈ ��--></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab_KOBACO" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
								<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 100%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 VIEWASTEXT>
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="_ExtentX" VALUE="31829">
	<PARAM NAME="_ExtentY" VALUE="21061">
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
							<DIV id="pnlTab_SBS" style="LEFT: 7px; VISIBILITY: hidden; WIDTH: 100%; POSITION: absolute; HEIGHT: 100%"
										ms_positioning="GridLayout">
								<OBJECT id="sprSht_SBS" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31803">
									<PARAM NAME="_ExtentY" VALUE="12462">
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
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
					</TR>
				</TABLE>
			</P>
		</form>
	</body>
</HTML>
