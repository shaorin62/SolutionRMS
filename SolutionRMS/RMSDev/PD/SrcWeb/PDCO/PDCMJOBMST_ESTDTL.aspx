<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_ESTDTL.aspx.vb" Inherits="PD.PDCMJOBMST_ESTDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBMST_ESTDTL.aspx
'��      �� : JOBMST�� �ι�° �� - ��/�� �������� ���� �� ���� �Ѵ�. 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
'****************************************************************************************
-->
		<meta content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script id="clientEventHandlersVBS" language="vbscript">

option explicit

Dim mobjPDCOPREESTDTL, mobjPDCOGET
Dim mlngTempRowCnt
Dim mlngTempColCnt
Dim mlngRowCnt, mlngColCnt		
Dim mstrFIRSTPRODUCTIONCHECK
Dim mstrSELECT		'������ȸ�� �ش� JOBNO �� �������� JOBNO �� ���̸�  Ȯ��  �ٸ��� "F" ������ "T"
Dim mstrCheck

mstrFIRSTPRODUCTIONCHECK = "N"
CONST meTAB = 9
mstrCheck = TRUE

'=============================
' �̺�Ʈ ���ν��� 
'=============================
Sub window_onload
	window.setTimeout "call Initpage()",1000 
End Sub

Sub Window_OnUnload() 
	EndPage
End Sub

'�⺻���� ����
Sub ImgBasicFormat_onclick
	Dim vntData
	Dim vntRet, vntInParams
	Dim intCnt
	Dim strPREESTNO
	
	with frmThis
		If .sprSht.MaxRows <> 0 Or .txtPREESTGBN.value = "������" Then
			gErrorMsgBox "������ �� �󼼳��� ����� ���������� �����Ҽ� �����ϴ�." & vbcrlf & "�󼼳��� ���� �� �������� ���������� ��ȯ�Ͽ� ó���Ͻʽÿ�.","ó���ȳ�"
			Exit Sub
		End If
		
		vntInParams = array("","")
		vntRet = gShowModalWindow("PDCMESTTYPEPOP.aspx",vntInParams , 590,490)
		if isArray(vntRet) then
			'�׸��� �ϴ� �ʱ�ȭ
			.sprSht.MaxRows = 0
			If .txtPREESTNO.value = "" Then
				strPREESTNO = "9999999999"
			Else
				strPREESTNO = .txtPREESTNO.value 
			End If
			.txtSUMAMT.value = trim(vntRet(3,0))
			.txtAMT.value = trim(vntRet(4,0))
			.txtSUSURATE.value  = trim(vntRet(5,0))
			.txtSUSUAMT.value = trim(vntRet(6,0))
			.txtCOMMITION.value = trim(vntRet(7,0))
			.txtNONCOMMITION.value = trim(vntRet(8,0))
			
			txtSUSUAMT_onblur
			txtSUMAMT_onblur
			txtAMT_onblur
			
			txtESTSUSUAMT_onblur
			txtESTSUMAMT_onblur
			txtESTAMT_onblur
			
			txtCOMMITION_onblur
			
			
			txtNONCOMMITION_onblur
			
			
		    '�ϴ� Sheet ����
		    mlngRowCnt=clng(0): mlngColCnt=clng(0)
		    vntData = mobjPDCOPREESTDTL.SelectRtn_ProcEST(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(vntRet(0,0)),strPREESTNO)
			
			IF not gDoErrorRtn ("SelectRtn_ProcEST") then
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
				'���⼭ ���� Detail ��ư ����
				
	
				For intCnt =1 To .sprSht.MaxRows 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",intCnt) = "Y"  Then
						'�ܰ�,����,�ݾ� lock / ��ư�� �Է»��� - QTY|PRICE|AMT
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",intCnt,intCnt,false
						'��ư���·� ����
						If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
							mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�������Է�","DETAIL_BTN",intCnt,intCnt,,false
						Else
							mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�󼼰���","DETAIL_BTN",intCnt,intCnt,,false
						End If
						
					Else
						'�ܰ�,����,�ݾ� �Է¹����� �ֵ��� ���� / ��ư�� lock
						'mobjSCGLSpr.SetCellsLock2 .sprSht,true,"DETAIL_BTN|QTY|PRICE|AMT",Row,Row,false	
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",intCnt,intCnt,false
						'�Ϲ����·� ����
						'mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "DETAIL_BTN",Row,Row,255,,,,,False
						mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",intCnt,intCnt,0,,,,,,,,False
					End If
					sprSht_Change 1,intCnt
				Next
			End If
		end if
	End with
End Sub	


Sub imgBonSave_onclick
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ ���������� �����ϱ����ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�.","����������ȳ�!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	ExeProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'������ư
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	with frmThis
		'msgbox .txtPREESTNO.value & "--" & .txtPREESTNOVIEW.value
		
		'  "" �ΰ��� �ٸ����� ����� //  <>"" �ΰ��� �ش�job�� ���� ������
		if .txtPREESTNO.value <> "" then 
			if frmThis.txtENDFLAG.value = "T" Then
				gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
				Exit Sub
			End If
		end if 
	end with
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDel_onclick
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ �����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�.","�����ȳ�!"
		exit sub
	else
		if frmThis.txtENDFLAG.value = "T" Then
			gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ������� �Ұ��� �մϴ�.","�����ȳ�!"
			Exit Sub
		End If
	end if 
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgCFInput_onclick
	Dim vntInParams
	Dim vntRet
	Dim strPREESTNO
	Dim i, SaveCHK  '���� üũ
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "������ ���� ���� �����Ͱ� �ֽ��ϴ�. ������ �Ϸ��Ͻ� �� �����ϼ���.! ","CF ���� �Է� �ȳ�.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ �����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�. ","CF ���� �Է� �ȳ�.!"
		exit sub
	end if 
	
	with frmThis
	
		If .txtPREESTNO.value = "" Then
			strPREESTNO = "9999999999"
		Else
			strPREESTNO = .txtPREESTNO.value 
		End If
			
		vntInParams = array(strPREESTNO)
		vntRet = gShowModalWindow("PDCMJOBMST_CFINPUT.aspx",vntInParams , 1149,400)
	End With
End Sub

Sub imgTableUP_onclick
	Dim strRow
	Dim intCnt
	Dim i
	
	
	with frmThis
		
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				intCnt = intCnt + 1
			End if 
		next
		
		if intCnt > 1 then
			gErrormsgbox "������ �̵��� �� �����͸� �����ϼž� �մϴ�.","�̵��ȳ�!"
			exit sub
		end if
			
			
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�̵��ȳ�!"
			Exit Sub
		end if 
		if strRow = 1 then exit sub
		
		'�ڱ��ڽ��� �ѱ��.
		sprSht_UpCopy strRow
	
	end with
End Sub

Sub imgTableDown_onclick
	Dim strRow
	Dim intCnt
	Dim i	
	
	with frmThis
		
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				intCnt = intCnt + 1
			End if 
		next
		
		if intCnt > 1 then
			gErrormsgbox "������ �̵��� �� �����͸� �����ϼž� �մϴ�.","�̵��ȳ�!"
			exit sub
		end if	
	
		for i=1 to .sprSht.Maxrows
			if mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",i) = "1" then
				strRow = i
				exit for
			End if 
		Next
		
		if strRow = 0 then 
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�̵��ȳ�!"
			Exit Sub
		end if 
		
		if strRow = (.sprSht.MaxRows) then exit sub	
		sprSht_DownCopy strRow
	end with
End Sub

Sub sprSht_UpCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	Dim i
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row����	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'msgbox strRow
		'���鼭 �ڽŰ� printseq��ŭ ������ ����
		for i=1 to .sprSht_copy.MaxRows
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow- strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow-strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow -strPRINT_SEQ)

			mobjSCGLSpr.CellChanged frmThis.sprSht_copy, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)			
			strPRINT_SEQ = strPRINT_SEQ -1

		next

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
			End if
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)	
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


Sub sprSht_DownCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	Dim i
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'row����	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		'���鼭 �ڽŰ� printseq��ŭ ������ ����
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow+ strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow+strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DIVNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CLASSNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"ITEMCODENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"FAKENAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"FAKENAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"STD",i, mobjSCGLSpr.GetTextBinding( .sprSht,"STD",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"COMMIFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"COMMIFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUSUAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUSUAMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"GBN",i, mobjSCGLSpr.GetTextBinding( .sprSht,"GBN",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBDETAIL",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBDETAIL",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"IMESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"DETAILYNFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"INDIRECFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRODUCTIONCOMMISSION",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRODUCTIONCOMMISSION",strRow +strPRINT_SEQ)
			
			mobjSCGLSpr.CellChanged frmThis.sprSht_copy, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DIVNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CLASSNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"ITEMCODENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"FAKENAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"STD",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"STD",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"COMMIFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUSUAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"GBN",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"GBN",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBDETAIL",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"IMESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"DETAILYNFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"DETAILYNFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"INDIRECFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRODUCTIONCOMMISSION",i)
				
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 		
		next
		.sprSht_copy.MaxRows = 0
	end with
End Sub


'������ ���
Sub imgPrintEstBasic_onclick
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,SaveCHK		'�μ��� ���� üũ
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim intRtn
	Dim strUSERID
	
	SaveCHK = 0
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "������ ���� ���� �����Ͱ� �ֽ��ϴ�. ������ �Ϸ��Ͻ� �� ����ϼ���.! ","��� ������ �ȳ�.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ ����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�. ","��� ������ �ȳ�.!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "PD"
		
		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow)
		
		if cdate(.txtAGREEYEARMON.value) < cdate("2013-01-31") then
			ReportName = "ESTIMATE_ONE.rpt"
		else
			ReportName = "ESTIMATE_ONE_P.rpt"
		end if 
		
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		strUSERID = ""
		vntData = mobjPDCOPREESTDTL.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, 1, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		
		window.setTimeout "call printSetTimeout('" & strPREESTNO & "')", 10000
		
	end with
	gFlowWait meWAIT_OFF
End Sub

'������ ���
Sub imgPrintEst_onclick
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,SaveCHK		'�μ��� ���� üũ
	Dim strPREESTNO
	Dim vntData
	Dim vntDataTemp
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	for i = 1 to frmThis.sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(frmthis.sprSht,"SAVEFLAG",i) = "Y" then
			gErrorMsgBox "������ ���� ���� �����Ͱ� �ֽ��ϴ�. ������ �Ϸ��Ͻ� �� ����ϼ���.! ","��� ������ �ȳ�.!"
			exit sub
		end if 
	next
	
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ ����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�. ","��� ������ �ȳ�.!"
		exit sub
	end if 
	
	gFlowWait meWAIT_ON
	with frmThis
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)

		'md_trans_temp���� ��
		
		ModuleDir = "PD"
		
		strPREESTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",.sprSht.ActiveRow)
		
		IF .cmbESTTYPE.value = 1 THEN
			IF .txtPREESTGBN.value = "������" THEN
				if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
					ReportName = "ESTIMATE.rpt"
				else
					ReportName = "ESTIMATE_P.rpt"
				end if
			ELSE
				if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
					ReportName = "ESTIMATE_ESTBACK.rpt"
				else
					ReportName = "ESTIMATE_ESTBACK_P.rpt"
				end if
				
			END IF 
		ELSEIF .cmbESTTYPE.value = 2 THEN
			IF .txtPREESTGBN.value = "������" THEN
				gErrorMsgBox "�������� ��� �����մϴ�.","�μ�ȳ�!"
				EXIT SUB
			END IF 
			
			if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
				ReportName = "ESTIMATE_ACTUAL.rpt"
			else
				ReportName = "ESTIMATE_ACTUAL_P.rpt"
			end if
			
		ELSE
			IF .txtPREESTGBN.value = "������" THEN
				gErrorMsgBox "�������� ��� �����մϴ�.","�μ�ȳ�!"
				EXIT SUB
			END IF
			
			if cdate(.txtAGREEYEARMON.value) <= cdate("2013-01-31") then
				ReportName = "ACTUAL.rpt"
			else
				ReportName = "ACTUAL_P.rpt"
			end if
			
		END IF

		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		strUSERID = ""
		vntData = mobjPDCOPREESTDTL.ProcessRtn_TEMP(gstrConfigXml,strPREESTNO, 1, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		
		window.setTimeout "call printSetTimeout('" & strPREESTNO & "')", 10000
		
	end with
	gFlowWait meWAIT_OFF
End Sub


'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout(strPREESTNO)
	Dim intRtn, intRtn2
	with frmThis
		intRtn = mobjPDCOPREESTDTL.DeleteRtn_temp(gstrConfigXml)
		'intRtn2 = mobjPDCOPREESTDTL.DeleteRtnUpdate_ATTR07(gstrConfigXml, strTRANSYEARMON, strTRANSNO)
	end with
end sub

Sub imgCalEndarAGREE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtAGREEYEARMON,frmThis.imgCalEndarAGREE,"txtAGREEYEARMON_onchange()"
		gSetChange
	end with
End Sub

Sub imgimgCalEndarCREDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgimgCalEndarCREDAY,"txtPRINTDAY_onchange()"
		gSetChange
	end with
End Sub

'���߰��� �������=================================================================================================
'�������ΰ��� txtEndflag �� û�������� �ִ� ��� �߰��Ұ�
'������ �� �ű��Է½ô� �ٷ� �߰� ���� << ������ �����Ĵ� �ش� JOB �� û�������� ��ȸ�Ͽ� endflag �� �ٽ� ��������,
'���������� ���� �Ѵٸ� frmThis.txtENDFLAG.value = "T" ������ ����� ���� �Ұ��̴�,
'==================================================================================================================

Sub imgRowAdd_onclick ()
	'���߰��� ������ ���� �������̹Ƿ� ���������� Ȯ���ÿ��� �Ǵ�����... �������� �������� ��������� ������ ��� ����!!
	'�������̶��
	if mstrSELECT = "F" then
		gErrorMsgBox "�ٸ� JOBNO �� ������ �����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�.","�߰��ȳ�!"
		exit sub
	end if 
	
	If frmThis.txtPREESTGBN.value = "������" And frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ���߰��� �Ұ��� �մϴ�.","ó���ȳ�!"
		Exit Sub
	End IF
	
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DETAIL_BTN")  Then
			If frmThis.txtPREESTGBN.value = "������" Then
				If frmThis.txtENDFLAG.value <> "T" Then
					intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
					DefaultValue
				end if
			Else
				intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
				DefaultValue
			End If
		End If
	Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select
	End If
End Sub



Sub DefaultValue
	Dim i
	Dim imeseqCHECK
	Dim highCnt, cnt1, cnt2
	Dim highvalue
	
	imeseqCHECK = true
	highCnt = 0	
	cnt2 = 0
	
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
		mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",.sprSht.ActiveRow, "N"
		If .sprSht.MaxRows = 1 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow,1
			mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow,1
		Else
			'IMESEQ �� ��� ���߿� ITEMCODESEQ �� ���� ���������Ѵ�. ���̺��� �ִ밪���� �����Ѵ�.
			for i = 1  to .sprSht.MaxRows
				cnt1 = 0
				if mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i) <> "" then 
					cnt1 = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i))
				end if
				if cnt2 < cnt1 then 
					highcnt = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i) + 1)
					cnt2 = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i))
				end if 
			next
				highvalue = cstr(highcnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",.sprSht.ActiveRow, highvalue
			
			mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"PRINT_SEQ",.sprSht.MaxRows -1)+1
		End If
	End With
End Sub


'-----------------------------------------
'��Ʈ ���� Ű�� �������� ���� �ݾ� �ջ�. 
'-----------------------------------------
Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	Dim strCOLUMN
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub
	'Ű �����϶� ���ε�

	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") or  _
																		   mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) or _ 
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT")) Then
				
				
				
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
	End With
END SUB

'-----------------------------------
'��Ʈ���� ���콺�� �����ö� �̺�Ʈ
'-----------------------------------
Sub sprSht_Mouseup(KeyCode, Shift, X,Y)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		strSUM = 0
		intSelCnt = 0
		intSelCnt1 = 0
	
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") or _
																			   .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT")  Then
																			
				
					vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
					vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)
				
				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" then
						FOR j = 0 TO intSelCnt1 -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					End If
				Next
				
				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if	
		
	End With
End Sub

'-----------------------------
' UI
'-----------------------------	
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
		
		If .txtSUSURATE.value = 0 Then
			.txtSUSUAMT.readOnly = false
		Else
			.txtSUSUAMT.readOnly = true
		End If
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtSUSURATE_onblur
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		
	End with
End Sub
Sub CleanFieldRate
	with frmThis
		.txtSUSURATE.value = 0
	End With
End Sub
Sub txtCOMMITION_onfocus

	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
		
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
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

Sub txtESTAMT_onfocus
	with frmThis
		.txtESTAMT.value = Replace(.txtESTAMT.value,",","")
	end with
End Sub
Sub txtESTAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTAMT,0,true)
	end with
End Sub

Sub txtESTSUMAMT_onfocus
	with frmThis
		.txtESTSUMAMT.value = Replace(.txtESTSUMAMT.value,",","")
	end with
End Sub
Sub txtESTSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTSUMAMT,0,true)
	end with
End Sub

Sub txtESTSUSUAMT_onfocus
	with frmThis
		.txtESTSUSUAMT.value = Replace(.txtESTSUSUAMT.value,",","")
	end with
End Sub
Sub txtESTSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtESTSUSUAMT,0,true)
	end with
End Sub

Sub txtESTSUSURATE_onblur
	with frmThis
		If .txtESTSUSURATE.value = "" Then
			.txtESTSUSURATE.value = 0
		End If
		
	End with
End Sub




Sub txtAGREEYEARMON_onchange
	gSetChange
End Sub


Sub txtPRINTDAY_onchange
	gSetChange
End Sub

Sub txtSUSURATE_onchange
	with frmThis
		If .txtSUSURATE.value = "" Then
			.txtSUSURATE.value = 0
		End If
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM

		gSetChangeFlag .txtSUSURATE  
	End with
End Sub
Sub txtESTSUSURATE_onchange
	with frmThis
		If .txtESTSUSURATE.value = "" Then
			.txtESTSUSURATE.value = 0
		End If
		ESTSUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtESTSUSURATE  
	End with
End Sub


Sub txtSUSUAMT_onchange
	with frmThis
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSUAMT  
	End with
End Sub

Sub txtESTSUSUAMT_onchange
	with frmThis
		ESTSUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSUAMT  
	End with
End Sub





'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
'���������� ��� �׸��� ���� ���� �ɶ� �߻� �ϴ� �̺�Ʈ �Դϴ�.
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, strCols
	Dim strCode, strCodeName
	Dim strQTY, strPRICE, strAMT
	Dim lngPrice
	Dim lngVALUE
	Dim lngVALUE1
	Dim lngVALUE2

	with frmThis
				'Long Type�� ByRef ������ �ʱ�ȭ
				mlngRowCnt=clng(0)
				mlngColCnt=clng(0)
				strCode = ""
				strCodeName = ""
				IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
					strCode = ""
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",.sprSht.ActiveRow)
					vntData = mobjPDCOGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",strCodeName)
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntData(3,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntData(4,1)	
						mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntData(7,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntData(8,1)
						
						'�����׸�󼼱����� "Y" �̰�, JOB ������ �����϶��� �󼼰��� ��ư�� ����
						If vntData(7,1) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
						
							'�ܰ�,����,�ݾ� lock / ��ư�� �Է»��� - QTY|PRICE|AMT
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
							mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
							'����,�ܰ�,�ݾ��� 0ó��
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
							'��ư���·� ����(������� �ι� �Է����� �б�ó��!)
							If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�������Է�","DETAIL_BTN",Row,Row,,false
							Else
								mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�󼼰���","DETAIL_BTN",Row,Row,,false
							End If
						Else
							'�ܰ�,����,�ݾ� �Է¹����� �ֵ��� ���� / ��ư�� lock
							mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
							'�Ϲ����·� ����
							mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
						End If	
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
						
						'SUSUAMT_CHANGEVALUE2
						IF .txtPREESTGBN.value ="������" THEN
							Call ESTSUSUAMT_CHANGEVALUE2
						ELSE	
							Call SUSUAMT_CHANGEVALUE2
						END IF
						
						BUDGET_AMT_SUM
						
						
						If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
							mstrFIRSTPRODUCTIONCHECK = "Y"
						
							call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
						end if 
					Else
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End If
					.txtSUSURATE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
					.sprSht.Focus	
					mobjSCGLSpr.ActiveCell .sprSht, Col, Row
				'��������	
				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strPRICE <> "" And strAMT = "" Then
   						lngVALUE = strQTY * strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE
   					ElseIf strPRICE = "" And strAMT <> "" Then
   						lngVALUE1 = gRound(strAMT/strQTY,0)
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngVALUE1
   					ElseIf strPRICE <> "" And strAMT <> "" Then
   						lngVALUE2 = strQTY * strPRICE
   						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE2
   					End IF
   					
   					IF .txtPREESTGBN.value ="������" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
					
   					BUDGET_AMT_SUM
   				'�ܰ� ����
   				ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
   					strQTY		= mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",.sprSht.ActiveRow)
					strPRICE   = mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",.sprSht.ActiveRow)
					strAMT = strQTY * strPRICE
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT	
					
					'Call SUSUAMT_CHANGEVALUE(Row)
					IF .txtPREESTGBN.value ="������" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
					BUDGET_AMT_SUM
				'�ݾ׷���	
   				ElseIf  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
   					strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   					strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   					strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   					If strAMT = 0 Then
   						mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, strAMT
   						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, strAMT
   					Else 
   						If strQTY <> 0  Then
   							lngPrice = gRound(strAMT/strQTY,0)
   							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngPrice
   						End IF
   					End IF
   					'Call SUSUAMT_CHANGEVALUE(Row)
   					IF .txtPREESTGBN.value ="������" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
   					BUDGET_AMT_SUM
   				Elseif Col = mobjSCGLSpr.CnvtDataField(.sprSht,"COMMIFLAG") Then
   					'Call SUSUAMT_CHANGEVALUE2
   					IF .txtPREESTGBN.value ="������" THEN
						Call ESTSUSUAMT_CHANGEVALUE(Row)
					ELSE	
						Call SUSUAMT_CHANGEVALUE(Row)
					END IF
   					BUDGET_AMT_SUM
				END IF
	end with
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub SUSUAMT_CHANGEVALUE(ByVal Row)
	Dim strAMT,strCOMMIFLAG
	with frmThis
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		if strCOMMIFLAG = "1" Then
			'�������� ������ .txtSUSURATE �� Null �� ��� ����
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), Row
	End with
End SUb

Sub ESTSUSUAMT_CHANGEVALUE(ByVal Row)
	Dim strAMT,strCOMMIFLAG
	with frmThis
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		if strCOMMIFLAG = "1" Then
			'�������� ������ .txtSUSURATE �� Null �� ��� ����
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * .txtESTSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), Row
	End with
End SUb


Sub SUSUAMT_CHANGEVALUE2
Dim intCnt
Dim strAMT,strCOMMIFLAG
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		if strCOMMIFLAG = "1" Then
		'�������� ������ .txtSUSURATE �� Null �� ��� ����
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * .txtSUSURATE.value /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), intCnt
	Next
	
	End with
End Sub

Sub ESTSUSUAMT_CHANGEVALUE2
	Dim intCnt
	Dim strAMT,strCOMMIFLAG
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
	strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
	strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		if strCOMMIFLAG = "1" Then
		'�������� ������ .txtSUSURATE �� Null �� ��� ����
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * .txtESTSUSURATE.value /100),0)
		
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
		mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SUSUAMT"), intCnt
	Next
	
	End with
End Sub


'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	
	With frmThis
	
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		
		If .sprSht.MaxRows > 0 Then
			.txtSUMAMT_TOTAL.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT_TOTAL,0,True)
			
		else
			.txtSUMAMT_TOTAL.value = 0
		End If
		
	End With
End Sub


Sub BUDGET_AMT_SUM
	'���հ� ����
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM, intAMTSUB
	Dim lngSUSU
	'������ ��� ����
	Dim intCnt,intSUSU,intSUSUSUM 
	'commition ��� ����
	Dim intCnt1,intCOM,intCOMSUM 
	'noncommition ��꺯��
	Dim intCnt2,intNON,intNONSUM 
	
	Dim intESTSUSU , intESTSUSUSUM
	Dim IntESTAMT,IntESTAMTSUM, intESTAMTSUB
	
	with frmThis
	
		IntAMTSUM = 0
		IntPRICESUM = 0
		intSUSU = 0
		intSUSUSUM = 0
		intCOM = 0
		intCOMSUM = 0
		intNON = 0
		intNONSUM = 0
		intAMTSUB = 0
		
		intESTSUSU = 0
		intESTSUSUSUM = 0
		IntESTAMT = 0
		IntESTAMTSUM = 0
		intESTAMTSUB = 0
	
		IF .txtPREESTGBN.value = "������" THEN
			
			For intCnt = 1 To .sprSht.MaxRows
				intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
				intSUSUSUM = intSUSUSUM + intSUSU
			Next
			
			'If .txtESTSUSURATE.value <> 0 Then
				.txtESTSUSUAMT.value = intSUSUSUM
			'End If
			
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			intAMTSUB = IntAMTSUM
		
			If .txtESTSUSURATE.value <> 0 Then
				IntAMTSUM = IntAMTSUM + intSUSUSUM
			Else
				IntAMTSUM = IntAMTSUM + replace(.txtESTSUSUAMT.value,",","")
			End If
			
			.txtESTAMT.value = intAMTSUB
			.txtESTSUMAMT.value = IntAMTSUM
		ELSE
			
			For intCnt = 1 To .sprSht.MaxRows
				intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
				intSUSUSUM = intSUSUSUM + intSUSU
			Next
			
			'If .txtSUSURATE.value <> 0 Then
				.txtSUSUAMT.value = intSUSUSUM
			'End If
			
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			intAMTSUB = IntAMTSUM
		
			If .txtSUSURATE.value <> 0 Then
				IntAMTSUM = IntAMTSUM + intSUSUSUM
			Else
				IntAMTSUM = IntAMTSUM + replace(.txtSUSUAMT.value,",","")
			End If
			
			.txtAMT.value = intAMTSUB
			.txtSUMAMT.value = IntAMTSUM
		END IF 
		
		
		For intCnt1 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG", intCnt1) = "1" Then
				
				intCOM = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intCOMSUM = intCOMSUM + intCOM
			Else
				intNON = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intNONSUM = intNONSUM + intNON
			end if
		Next
		.txtCOMMITION.value = intCOMSUM
		.txtNONCOMMITION.value = intNONSUM
		
		txtSUSUAMT_onblur
		txtSUMAMT_onblur
		txtAMT_onblur
		
		txtESTSUSUAMT_onblur
		txtESTSUMAMT_onblur
		txtESTAMT_onblur
		
		txtCOMMITION_onblur
		txtNONCOMMITION_onblur
		
	End With
End Sub

'=================================
'���������� ���� Ŭ�� �� �߻�
'=================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim strSUM
	Dim i, j
	Dim vntData_col, vntData_row
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			If mstrCheck = True Then 
				mstrCheck = False
			elseif mstrCheck = False Then 
				mstrCheck = True
			End If
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			Next
		End If
	end With
End Sub

'=================================
'���������� ���� ���� Ŭ�� �� �߻�
'=================================
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'=========================================================
'���������� �׸��� ���ҽ� ��� �Լ��� �¿���� �Ҷ� ���
'=========================================================
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"ITEMCODENAME") Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding( sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)	
				
				'�����׸�󼼱����� "Y" �̰�, JOB ������ �����϶��� �󼼰��� ��ư�� ����
				If vntRet(7,0) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
				
					'�ܰ�,����,�ݾ� lock / ��ư�� �Է»��� - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
					'����,�ܰ�,�ݾ��� 0ó��
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'��ư���·� ����(������� �ι� �Է����� �б�ó��!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�������Է�","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�󼼰���","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'�ܰ�,����,�ݾ� �Է¹����� �ֵ��� ���� / ��ư�� lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
					'�Ϲ����·� ����
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If	
				'mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
	
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				IF .txtPREESTGBN.value ="������" THEN
					ESTSUSUAMT_CHANGEVALUE2
				ELSE	
					SUSUAMT_CHANGEVALUE2
				END IF 
				
				BUDGET_AMT_SUM
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
					mstrFIRSTPRODUCTIONCHECK = "Y"
					call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
				end if 
			End IF
			.txtSUSURATE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row +1
		end if
	End With
End Sub

'=================================================
'�������� �� ��ư�� Ŭ�� �Ͽ����� �߻� �ϴ� �̺�Ʈ
'=================================================
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strPREESTNO
	Dim strGBN
	Dim strITEMCODE
	Dim dblAMT, dblESTAMT
	Dim vntData, vntData_Temp
	Dim strRTN , strRTN2
	Dim intCnt,intRtn
	Dim strChk
	Dim intTempChkCnt
	Dim dblSUSUAMT  , dblESTSUSUAMT
	Dim intColorCnt
	Dim strITEMCODESEQ
	Dim returnPOP		'�󼼳��� �˾� ���� �� �޴� ����
	Dim indirecPOP		'������ �˾� ���� �� ����
	Dim processData		'������ �ӽ� ���� ����
	Dim intProductionCnt'������ �ο�� [2�� �̻��� ���´�.]
	
	Dim strOldITEMCODE
	with frmThis
	    '�����׸��ڵ� ���� ��ư �κ�
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then
			
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
			strOldITEMCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row)
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBDETAIL",Row, vntRet(7,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRODUCTIONCOMMISSION",Row, vntRet(8,0)
				'�����׸�󼼱����� "Y" �̰�, JOB ������ �����϶��� �󼼰��� ��ư�� ����
				
				If vntRet(7,0) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
					'�ܰ�,����,�ݾ� lock / ��ư�� �Է»��� - QTY|PRICE|AMT
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",Row,Row,false
					mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY | PRICE | AMT",Row,Row,false
					'����,�ܰ�,�ݾ��� 0ó��
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, 0
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
					'��ư���·� ����(������� �ι� �Է����� �б�ó��!)
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�������Է�","DETAIL_BTN",Row,Row,,false
					Else
						mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�󼼰���","DETAIL_BTN",Row,Row,,false
					End If
				Else
					'�ܰ�,����,�ݾ� �Է¹����� �ֵ��� ���� / ��ư�� lock
					mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY | PRICE | AMT",Row,Row,false
					'�Ϲ����·� ����
					mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",Row,Row,0,,,,,,,,False
				End If
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				
				IF .txtPREESTGBN.value ="������" THEN
					ESTSUSUAMT_CHANGEVALUE2
				ELSE	
					SUSUAMT_CHANGEVALUE2
				END IF
				BUDGET_AMT_SUM
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then
					mstrFIRSTPRODUCTIONCHECK = "Y"
					call sprSht_ButtonClicked (18,.sprSht.ActiveRow,"DETAIL_BTN")
				end if 
			End IF
			.txtSUSURATE.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row+1
			
		'CF������ �ι��� ���ý� ��ưó��....	
		'�󼼰��� ���� �� ��ȸ �˾� ȣ��
'------------------------------------------------------------------------------------------------------------			
'***********************1.   ������ �ι� ó���� �󼼳��� ó���� ������.
'------------------------------------------------------------------------------------------------------------
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DETAIL_BTN") Then
			'���δ��� ���� ������ �κ�
			If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row) = "242001" Then 
			
				'�ٸ� JOBNO �� ������ �״�� ���� �ϰ���� ��� �ϴ� ������ �����Ѵ�.
				if mstrSELECT = "F" then
					gErrorMsgBox "�ٸ� JOBNO �� ������ �����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�.","������ ���� �ȳ�!"
					exit sub
				end if 
				
				intProductionCnt = 0
				'������ ��Ʈ�� ���� �����´�.
				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) <> "242001" Then
						If mobjSCGLSpr.GetTextBinding( .sprSht,"INDIRECFLAG",intCnt) = "T" Then
							mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",intCnt, "F"
						Else
							mobjSCGLSpr.SetTextBinding .sprSht,"INDIRECFLAG",intCnt, "T"
						End If
						mobjSCGLSpr.CellChanged frmThis.sprSht, mobjSCGLSpr.CnvtDataField(.sprSht,"INDIRECFLAG"), intCnt
					ELSE
						'������ �ο�üũ
						intProductionCnt = intProductionCnt + 1
					End If
				Next
				
				if intProductionCnt > 1 then
					gErrorMsgBox "�̹� �ٸ� ������ �����Ͱ� �ֽ��ϴ�. ������� �Ѱ����� �Է� �����մϴ�." & vbcrlf & "�����͸� Ȯ���ϼ���.!" ,"������ �ȳ�"
					mobjSCGLSpr.DeleteRow .sprSht,Row
					exit sub
				end if 
				'indirectflag =  ��� ���� ������������,,, For ���� Cellchange �� �¿��.
				'PD_PRODUCTIONCOMMI_INPUT ���̺� ó������============================================================
				'������ ������ �˾��� ��� �˾����� �Ͼ�� ��� ó���� INPUT���̺��� �ذ�ȴ�.
				'���� ���� ������ ��� INPUT ���� �̷������ ���� ȭ���� ������ ���� ��� ������ ���� ������ �ȴ�.
				'���� ������ �Է��ϴ� ��ȸ�� �ϰų� �ٸ� �۾��� �Ұ�� �����Ͱ� ������� ���� ���� ���� �����ʹ� ������ ��ġ�� �ʴ´�.
				'====================================================================================================
				strPREESTNO = .txtPREESTNO.value 
				
				strChk = 0
				For intTempChkCnt=1 To .sprSht.MaxRows
					'���δ��� ������ �� ������ �ٸ� ������ ���� ����Ǿ� ���� ������ [�̸��� �� ������ȣ�� ���ٸ� ������ �Է��Ҽ� ���ٴ� ��..]
					If  (mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",intTempChkCnt) = "" and _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intTempChkCnt) <> "242001") or _
													 (mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intTempChkCnt) <> "242001" and _
													 mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",intTempChkCnt) = "Y" ) Then 
						strChk = strChk +1
					End If
				Next
				
				If strChk <> 0 Then
					if mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) = "" THEN
						gErrorMsgBox "������ ����������� ����� ���� Ȯ���ϰų� ������ ���� �Ͻʽÿ�." & vbcrlf & "������� ����Ǿ�� ������ �Է��ϽǼ� �ֽ��ϴ�.","ó���ȳ�"
						mobjSCGLSpr.DeleteRow .sprSht,Row
					END IF
					'������ �Ǿ��ִ� �������ǰ�� �޽��� �ڽ��� ���µ� ���� ���� ���� ����
					gErrorMsgBox "������ ����������� ����� ���� Ȯ���ϰų� ������ ���� �Ͻʽÿ�." & vbcrlf & "������� ����Ǿ�� ������ �Է��ϽǼ� �ֽ��ϴ�.","ó���ȳ�"
					EXIT SUB
				Else
				
					'��Ʈ�� ����� �����͸� �����´�.
					vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | ITEMCODE | AMT | PRODUCTIONCOMMISSION | PRINT_SEQ")
				
					if IsArray(vntData) then
						processData = vntData
					else 
						gErrorMsgBox "ó���� �����Ͱ� �����ϴ�." & vbcrlf & "����������� ������ �� ���������Ͱ� �ʿ��մϴ�.","ó���ȳ�!"
						Exit Sub
					end if
					
					dblAMT = 0
					dblAMT = mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",Row)
					mstrFIRSTPRODUCTIONCHECK = "Y"
					
					vntInParams = array(strPREESTNO,Trim(.txtENDFLAG.value), mstrFIRSTPRODUCTIONCHECK,.txtPREESTGBN.value, processData)
					
					vntRet = gShowModalWindow("PDCMJOBMST_INDIRECTCOST.aspx",vntInParams , 1149,650)
					
					indirecPOP = vntRet(0)
					
					'�˾�â���� ������ ���ٸ� ���� �������� ������ ���� �ʴ´�.
					if indirecPOP <> "False" then
						'�ش� �˾��� ����� �����͸� �����Ͽ� ������������ �ٸ��ٸ� ������ ������ ���� �Ҽ� �ֵ��� �����Ѵ�.
						mlngTempRowCnt=clng(0): mlngTempColCnt=clng(0)
						vntData = mobjPDCOPREESTDTL.SelectRtn_CommiHDRSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",Row))			

						If mlngTempRowCnt > 0 Then
							strRTN = Cstr(vntData(0,1)) '������
							strRTN2 = Cstr(vntData(1,1)) '������
						Else 
							strRTN = "0"
							strRTN2 = "0"
						End If

						IF .txtPREESTGBN.value = "������" THEN
							If strRTN <> Cstr(indirecPOP) Then
								'���� �޾ƿ��� ���� �ƴ϶�, PD_SUBITEM_DTL �Ǵ� PD_SUBITEM_INPUT ���� �����´�.
								mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
								If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
									dblSUSUAMT = indirecPOP * .txtESTSUSURATE.value * 0.01
									mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
								End If
								
								BUDGET_AMT_SUM
								
								'���� ����Ǿ����� ǥ��
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							End IF
						ELSE
							If strRTN2	<> Cstr(indirecPOP) Then
								'���� �޾ƿ��� ���� �ƴ϶�, PD_SUBITEM_DTL �Ǵ� PD_SUBITEM_INPUT ���� �����´�.
								mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, indirecPOP
								mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
								If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
									dblSUSUAMT = indirecPOP * .txtSUSURATE.value * 0.01
									mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
								End If
								BUDGET_AMT_SUM
								
								'���� ����Ǿ����� ǥ��
								If vntRet(1) = "T" then
									mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
								else 
									mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "N"
								end if  
							End IF
						END IF 
					
						mobjSCGLSpr.CellChanged .sprSht, Col,Row
						
						For intColorCnt = 1 To .sprSht.MaxRows
							If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
							mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
							Else
								If intColorCnt Mod 2 = 0 Then
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
								Else
								mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
								End If
							End if
						Next
					end if
				End If	
'------------------------------------------------------------------------------------------------------------			
'***********************2.   TV-CF �� ���� ó�� 
'------------------------------------------------------------------------------------------------------------
			Else
				strRTN = 0
				strRTN2 = 0
				dblAMT = 0
				returnPOP = 0
				
				
				'�ٸ� JOBNO �� ������ �״�� ���� �ϰ���� ��� �ϴ� ������ �����Ѵ�.
				
				if mstrSELECT = "F" then
					gErrorMsgBox "�ٸ� JOBNO �� ������ �����ϱ� ���ؼ��� �ϴ� ������ ���� �ؾ� �մϴ�.","�� ���� �ȳ�!"
					exit sub
				end if 
				
				For intCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
						If mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",intCnt) = "Y" Then
							gErrorMsgBox "������ ���� �Ǿ��ֽ��ϴ�. �����Ͻ��� �� ������ �Է��ϼ���.","�� ���� �ȳ�!"
							exit sub
						end if
					End If
				Next
				
				If .txtPREESTNO.value = "" Then
					strPREESTNO = "9999999999"
					
				Else
					strPREESTNO = .txtPREESTNO.value 
				End If	
				
				If .txtPREESTGBN.value = "������" Then
					strGBN="F"
				ElseIf .txtPREESTGBN.value = "������" Then
					strGBN = "T"
				End If
				
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row) = "" Then
				'ITEMCODESEQ �� ���ٸ�[������� �ʰ� �󼼳������� �Է��� ���̶��] IMESEQ �� �����´�.[���Ŀ� ITEMCODESEQ �� IMESEQ �� ���� ��������.!]
					strITEMCODESEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row)
				Else
					strITEMCODESEQ = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODESEQ",Row)
				End If
				
				dblAMT = mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",Row)
				vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"DIVNAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"CLASSNAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row), _
									mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",Row), _
									strPREESTNO,strGBN,strITEMCODESEQ,Trim(.txtENDFLAG.value), _
									.txtJOBNO.value)
				
				vntRet = gShowModalWindow("PDCMJOBMST_SUBITEM.aspx",vntInParams , 1149,650)
			
				'�˾����� �����ǰų� ����Ȱ�
				returnPOP = vntRet(0)
				
				'�˾����� ������ �̷���� input�� ���� �ִٸ� ���� �ƴϸ� false �� ��ȯ	
				if returnPOP <> "False" then		
					
					mlngTempRowCnt=clng(0): mlngTempColCnt=clng(0)
					vntData = mobjPDCOPREESTDTL.SelectRtn_DtlSum(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt, _
																.txtJOBNO.value, _
																mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",Row), _
																mobjSCGLSpr.GetTextBinding( .sprSht,"IMESEQ",Row))
					
					If mlngTempRowCnt > 0 Then
						strRTN = Cstr(vntData(0,1)) 
						strRTN2 = Cstr(vntData(1,1)) 
					Else 
						strRTN = "0"
						strRTN2 = "0"
					End If
		
					IF .txtPREESTGBN.value = "������" THEN
						If strRTN <> Cstr(returnPOP) Then
							'���� �޾ƿ��� ���� �ƴ϶�, PD_SUBITEM_DTL �Ǵ� PD_SUBITEM_INPUT ���� �����´�.
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtESTSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							'���� ����Ǿ����� ǥ��
							mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
						'���� ���ٸ� ���� ���� �ѷ��൵ ��� ���� �ٸ� �˾����� �����ߴٰ� �ٽ� ������� ������� 
						else 
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtESTSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'�ݾ��� ������� �� ������ �ٸ��� �ֱ⶧���� �˾����� �̺�Ʈ�� �־��ٸ� ������ �����Ѵ�.
							If vntRet(1) = "T" then
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							end if  
						End IF
					ELSE
						If strRTN2	<> Cstr(returnPOP) Then
							'���� �޾ƿ��� ���� �ƴ϶�, PD_SUBITEM_DTL �Ǵ� PD_SUBITEM_INPUT ���� �����´�.
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'���� ����Ǿ����� ǥ��
							mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
						
						else 
							mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, returnPOP
							mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
							If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
								dblSUSUAMT = returnPOP * .txtSUSURATE.value * 0.01
								mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, dblSUSUAMT
							End If
							
							BUDGET_AMT_SUM
							
							'�ݾ��� ������� �� ������ �ٸ��� �ֱ⶧���� �˾����� �̺�Ʈ�� �־��ٸ� ������ �����Ѵ�.							
							If vntRet(1) = "T" then
								mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"
							end if 
						End IF
					END IF 
				else
					'�˾��� ���� false �̸�[input �� ���� ������] [CHANGEFLAG = "T" �����̳� ������ �ߴٸ�.](��ü �����ϰ����)
					IF vntRet(1) = "T" then
						mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, 0
						mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, 0
						mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "1"
						If mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row) = "1" Then
							mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",Row, 0
						End If
						
						BUDGET_AMT_SUM
						
						'���� ����Ǿ����� ǥ��
						mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",Row, "Y"					
					end if 
				End if 
					
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next
				
				mobjSCGLSpr.CellChanged .sprSht,mobjSCGLSpr.CnvtDataField(.sprSht,"SAVEFLAG"),Row
			End If
		end if
	End with
End Sub

'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	Dim strMSG
	Dim strPREESTNO
	
	'����������ü ����	
	set mobjPDCOPREESTDTL = gCreateRemoteObject("cPDCO.ccPDCOPREESTDTL")
	set mobjPDCOGET		  = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 24, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | PRINT_SEQ | PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | DETAIL_BTN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION"
		mobjSCGLSpr.SetHeader .sprSht,		  "����|�μ�|��������ȣ|����|��з�|�ߺз�|�����׸��ڵ�|�����׸��|������|����|Ŀ�̼�|����|�ܰ�|�ݾ�|������ݾ�|���屸��|�󼼰���|�󼼰�������|��¥����|�����忩��|�󼼺κп���|���̷�Ʈ�÷���|������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|   4|        10|   4|     8|    12|        8 |2|        15|12    |  18|     6|  6|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     "
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRINT_SEQ | QTY | PRICE | AMT | SUSUAMT | PRODUCTIONCOMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY", -1, -1, 1
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | ITEMCODENAME | FAKENAME | STD | GBN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PRINT_SEQ | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | PREESTNO | DETAIL_BTN | IMESEQ | SAVEFLAG | DETAILYNFLAG"
		mobjSCGLSpr.ColHidden .sprSht, "GBN | SUBDETAIL | INDIRECFLAG | PRODUCTIONCOMMISSION | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION | PREESTNO", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE | ITEMCODESEQ",-1,-1,2,2,false


		'���� ������ ������ ���� �׸���
		gSetSheetColor mobjSCGLSpr, .sprSht_copy
		mobjSCGLSpr.SpreadLayout .sprSht_copy, 24, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht_copy, 7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_copy, "CHK | PRINT_SEQ | PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | DETAIL_BTN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION"
		mobjSCGLSpr.SetHeader .sprSht_copy,		  "����|�μ�|��������ȣ|����|��з�|�ߺз�|�����׸��ڵ�|�����׸��|������|����|Ŀ�̼�|����|�ܰ�|�ݾ�|������ݾ�|���屸��|�󼼰���|�󼼰�������|��¥����|�����忩��|�󼼺κп���|���̷�Ʈ�÷���|������"
		mobjSCGLSpr.SetColWidth .sprSht_copy, "-1","   4|   4|        10|   4|     8|    12|        8 |2|        15|12    |  18|     6|  6|  13|13  |10        |0       |10      |0           |10      |13          |10          |0             |0     "
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_copy, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_copy,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_copy, "CHK | COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "PRINT_SEQ | QTY | PRICE | AMT | SUSUAMT | PRODUCTIONCOMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "QTY", -1, -1, 1
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_copy, "PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | ITEMCODENAME | FAKENAME | STD | GBN | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht_copy, true, "ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | PREESTNO | DETAIL_BTN | IMESEQ | SAVEFLAG | DETAILYNFLAG"
		mobjSCGLSpr.ColHidden .sprSht_copy, "GBN | SUBDETAIL | INDIRECFLAG | PRODUCTIONCOMMISSION | SUBDETAIL | IMESEQ | SAVEFLAG | DETAILYNFLAG | INDIRECFLAG | PRODUCTIONCOMMISSION | PREESTNO", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "ITEMCODE | ITEMCODESEQ",-1,-1,2,2,false	    

		.sprSht.style.visibility  = "visible"
		.sprSht_copy.style.visibility  = "visible"

		InitPageData
		
		'�θ�â�� ������ ��������
		.txtJOBNO.value =  parent. document.forms("frmThis").txtJOBNO.value 
		
		
		strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value
		.txtPREESTNO.value = strPREESTNO

		'������ ���� ��� strPREESTNO �κп� .txtPREESTNO.value�� �׳� �������� ������ �ɰ� ������
		'�ʵ忡 �ԷµǴ°ź��� �Ʒ� ������ ������ Ÿ�� �� ����.
		'�׷��� �θ�â�� PREESTNO�� ������ �޾Ƽ� ������ ������ ����...KTY 20110504
		'������ ���ٸ�
		if strPREESTNO = "" Then
			document.getElementById("strMsgBox").innerHTML = "- ���������� �����ϴ�."
			window.setTimeout "call NewDataSet()",1000 
		Else
			'������ �ִٸ�
			SelectRtn

			'������ ������, �������� ���� ���� �������� �ִ� ��� û���� ���곻���� ��ȸ��.
			if .txtSETCONFIRMFLAG.value = "F" Then
				strMSG = "- �������� �����ϴ�."
			Else
				if .txtENDFLAG.value = "T" AND .txtENDFLAGEXE.value = "T" Then
					strMSG = "- û��,���� ������ �ֽ��ϴ�"
				Elseif .txtENDFLAG.value = "T" AND .txtENDFLAGEXE.value = "F" Then
					strMSG = "- û�� ������ �ֽ��ϴ�"
				Elseif .txtENDFLAG.value = "F" AND .txtENDFLAGEXE.value = "T" Then
					strMSG = "- ���� ������ �ֽ��ϴ�"
				Elseif .txtENDFLAG.value = "F" AND .txtENDFLAGEXE.value = "F" Then
					strMSG = "- û��,���� ������ �����ϴ�."
				End If
			End IF
			document.getElementById("strMsgBox").innerHTML = strMSG
		End If
	End With
End Sub

Sub EndPage()
	set mobjPDCOPREESTDTL = Nothing
	set mobjPDCOGET = Nothing
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	with frmThis

		.sprSht.MaxRows = 0
		.sprSht_copy.MaxRows = 0
		.txtPRINTDAY.value = gNowDate
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'���ǻ��� ���� �������� û������ ��� ���������� ���� �ɼ� ����. - �������������� ��ư�� ���� ��ȸ�� �Ͽ� �ش� JOBNO �� û���Ǿ� �ִ����� Ȯ�� �ؾ� ��....
Sub NewDataSet
	with frmThis
		.txtAGREEYEARMON.value = gNowDate
		.txtPREESTGBN.value ="������"
		.txtPREESTNO.value = ""
		.txtMEMO.value = ""
		.txtPREESTNAME.value = ""
		.txtAMT.value = 0
		.txtCOMMITION.value = 0
		.txtNONCOMMITION.value = 0
		.txtSUSUAMT.value = 0
		.txtSUMAMT.value = 0
		.txtSUSURATE.value = 10
		.txtESTAMT.value = 0
		.txtESTSUMAMT.value = 0
		.txtESTSUSUAMT.value = 0
		.txtESTSUSURATE.value = 10
		'������ ������,��,�귣�带 �����Ѵ�.
		.txtCLIENTCODE.value = parent.document.forms("frmThis").txtCLIENTCODE.value
		.txtTIMCODE.value = parent.document.forms("frmThis").txtTIMCODE.value
		.txtSUBSEQ.value  = parent.document.forms("frmThis").txtSUBSEQ.value

		.txtJOBNAME.value  = parent.document.forms("frmThis").txtPRIJOBNAME.value
		.txtJOBNO.value  = parent.document.forms("frmThis").txtJOBNO.value 
		.txtPREESTNAME.value = parent.document.forms("frmThis").txtPRIJOBNAME.value
		txtSUSUAMT_onfocus
	End With
End Sub

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	window.setTimeout "call SelectRtn_Real()", 1000
End Sub

Sub SelectRtn_Real()
	Dim strCODE
	Dim strCHK
	With frmThis
		
		strCODE = parent.document.forms("frmThis").txtPREESTNO.value
		
		IF not SelectRtn_Head (strCODE) Then Exit Sub
	
		'��Ʈ ��ȸ
		CALL SelectRtn_Detail (strCODE)
		
		txtSUSUAMT_onblur
		txtSUMAMT_onblur
		txtAMT_onblur
		
		txtESTSUSUAMT_onblur
		txtESTSUMAMT_onblur
		txtESTAMT_onblur
		
		txtCOMMITION_onblur
		txtNONCOMMITION_onblur
	
		'�ŷ����� û���ǿ� ���� �������� �������� �Ҽ� �ִ�.
		If .txtENDFLAG.value = "T" AND .txtPREESTGBN.value = "������" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
			.txtAGREEYEARMON.className = "NOINPUT"
			.txtAGREEYEARMON.readOnly = true
			.imgCalEndarAGREE.disabled = true
	
			
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = true
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = true
			
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			
		ElseIF  .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "������" Then
			
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtSUSUAMT.className = "INPUT_R"
			.txtSUSUAMT.readOnly = FALSE
			.txtSUSURATE.className = "INPUT_R"
			.txtSUSURATE.readOnly = FALSE
			
			
		ELSEIF .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "������" Then
			
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = TRUE
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = TRUE
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtESTSUSUAMT.className = "INPUT_R"
			.txtESTSUSUAMT.readOnly =  FALSE
			.txtESTSUSURATE.className = "INPUT_R"
			.txtESTSUSURATE.readOnly = FALSE
			
		End If
			
		.txtSELECTAMT.value = 0
		.txtSUMAMT_TOTAL.value = 0
		AMT_SUM
	End with
	
	'��������Ʈ���� �������� �ƴ� �ٸ����� �ҷ��ð��[�����ҷ���] F �� ����ִ�.
	mstrSELECT = parent. document.forms("frmThis").txtSELECT.value 
	
	
	
	If mstrSELECT = "F" Then
		Est_Copy
	End If
	
	CFINPUT_VISIBLE
		
End SUb

'���� ����Ʈ���� �ٸ������� �����ð��.....
Sub Est_Copy
	Dim intCnt
	Dim intRtn
	Dim intSaveRtn
	Dim strOldCode
	Dim strJOBNO
	
	with frmThis		
		'���� �����ִ� ������ ������ ����
		intRtn = mobjPDCOPREESTDTL.ProcessRtn_TempDelete(gstrConfigXml)
				
		if gDoErrorRtn ("ProcessRtn_TempDelete") then
			gErrorMsgBox "���ʰ����ۼ��� �󼼿����׸� Temp���� ����µ� �����Ͽ����ϴ�." & vbcrlf & "�����ڿ��� ���� �Ͻʽÿ�.","ó���ȳ�!"
		End If
	
		'�ؽ�Ʈ �ʵ��� Ȱ��ȭ
		.txtPREESTNAME.className = "INPUT_L"
		.txtPREESTNAME.readOnly = false
		.txtSUSURATE.className = "INPUT_R"
		.txtSUSURATE.readOnly = false
		.txtSUSUAMT.className = "INPUT_R"
		.txtSUSUAMT.readOnly = false
			
		'�������縦 �ϱ����Ͽ� �غ��Ѵ�.	
		.txtAGREEYEARMON.value = gNowDate
		.txtPREESTGBN.value ="������"
		.txtPREESTNO.value = ""
		.txtMEMO.value = ""
		.txtPREESTNAME.value = ""
		
		'������ ������,��,�귣�带 �����Ѵ�.
		.txtCLIENTCODE.value = parent.document.forms("frmThis").txtCLIENTCODE.value
		.txtTIMCODE.value = parent.document.forms("frmThis").txtTIMCODE.value
		.txtSUBSEQ.value  = parent.document.forms("frmThis").txtSUBSEQ.value
		
		.txtJOBNAME.value  = parent.document.forms("frmThis").txtPRIJOBNAME.value
		.txtJOBNO.value  = parent.document.forms("frmThis").txtJOBNO.value 
		
		'���׸� ���� pd_subitem_input �� ����
		strOldCode = parent.document.forms("frmThis").txtPREESTNO.value
		
		'�����Ϸ��� ���� ��ȸ�� ����� ������ ���� ���� ������ jobno �� �ڴ´�.
		strJOBNO = parent.document.forms("frmThis").txtJOBNO.value
		
		intSaveRtn = mobjPDCOPREESTDTL.ProcessRtn_TempInsert(gstrConfigXml,strOldCode,strJOBNO)
		
		'�󼼰����� 
		For intCnt = 1 To .sprSht.MaxRows 
			mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",intCnt, ""
			
			'mobjSCGLSpr.SetTextBinding .sprSht,"IMESEQ",intCnt, ""
			
			'mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",intCnt, ""
			'If mobjSCGLSpr.GetTextBinding( .sprSht,"DETAILYNFLAG",intCnt) = "Y" Then
			'	mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",intCnt, "Y"
			'End IF
			'sprSht_Change 1,intCnt
		Next
		'�ݾ� ����
		IF .txtPREESTGBN.value ="������" THEN
			Call ESTSUSUAMT_CHANGEVALUE2
		ELSE	
			Call SUSUAMT_CHANGEVALUE2
			
		END IF
		BUDGET_AMT_SUM	
		
	End with
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'==============================================
'---������ �����Ŀ� �ش� �����͸� ����ȸ ��----
'==============================================
Sub SelectRtn_ProcessRtn (strCODE)
	Dim strCHK
	
	IF not SelectRtn_Head (strCODE) Then Exit Sub
	CALL SelectRtn_Detail (strCODE)
	
	txtSUSUAMT_onblur
	txtSUMAMT_onblur
	txtAMT_onblur
	
	txtESTSUSUAMT_onblur
	txtESTSUMAMT_onblur
	txtESTAMT_onblur
	
	txtCOMMITION_onblur
	txtNONCOMMITION_onblur
	
	with frmThis
		If .txtAGREEYEARMON.value = "" Then
		'	.txtAGREEYEARMON.value = gNowDate
		End if
		
		If .txtENDFLAG.value = "T" AND .txtPREESTGBN.value = "������" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
			.txtAGREEYEARMON.className = "NOINPUT"
			.txtAGREEYEARMON.readOnly = true
			.imgCalEndarAGREE.disabled = true
	
			
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = true
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = true
			
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			
		ElseIF  .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "������" Then
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtESTSUSUAMT.className = "NOINPUT_R"
			.txtESTSUSUAMT.readOnly =  true
			.txtESTSUSURATE.className = "NOINPUT_R"
			.txtESTSUSURATE.readOnly = true
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			
			.txtSUSUAMT.className = "INPUT_R"
			.txtSUSUAMT.readOnly = FALSE
			.txtSUSURATE.className = "INPUT_R"
			.txtSUSURATE.readOnly = FALSE
			
			
		ELSEIF .txtENDFLAG.value = "F" AND .txtPREESTGBN.value = "������" Then
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
			.txtSUSURATE.className = "NOINPUT_R"
			.txtSUSURATE.readOnly = TRUE
			.txtSUSUAMT.className = "NOINPUT_R"
			.txtSUSUAMT.readOnly = TRUE
			.txtAGREEYEARMON.className = "INPUT"
			.txtAGREEYEARMON.readOnly = false
			.imgCalEndarAGREE.disabled = false
			
			.txtESTSUSUAMT.className = "INPUT_R"
			.txtESTSUSUAMT.readOnly =  FALSE
			.txtESTSUSURATE.className = "INPUT_R"
			.txtESTSUSURATE.readOnly = FALSE
		End If
		
		
	End with
	CFINPUT_VISIBLE
End Sub

Sub CFINPUT_VISIBLE
	with frmThis
		
		If parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" AND .txtPREESTNO.value <> "" Then
			.imgCFInput.style.visibility = "visible"
		Else
			.imgCFInput.style.visibility = "hidden"
		End If
	End with
End Sub

Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCOPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			'gErrorMsgBox "������ �������� ���Ͽ�" & meNO_DATA, "" '��û�����̳�, ���� �ǵ��� ��ȸ�� MSG â�� ����ڿ��� ȥ���� ��,, �����ϰ� ����...By TH
			frmThis.sprSht.MaxRows = 0 
			NewDataSet			
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function

'���� ���̺� ��ȸ
Function SelectRtn_Detail (ByVal strCODE)
	Dim vntData
	Dim intCnt
	Dim intCnt_SubDetail
	Dim strRows
	Dim intColorCnt
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCOPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intColorCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"SAVEFLAG",intColorCnt) = "Y" Then
					mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HCCFFFF, &H000000,False
					Else
						If intColorCnt Mod 2 = 0 Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HF4EDE3, &H000000,False
						Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intColorCnt, intColorCnt,&HFFFFFF, &H000000,False
						End If
					End if
				Next	
				
				Detail_Yn
				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
			window.setTimeout "AMT_SUM",100	
			
		End with
	End IF
End Function

Sub Detail_Yn()
	Dim intCnt
	with frmThis
		For intCnt =1 To .sprSht.MaxRows 
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DETAILYNFLAG",intCnt) = "Y" AND parent.document.forms("frmThis").txtJOBGUBN.value = "PA02" Then
				'�ܰ�,����,�ݾ� lock / ��ư�� �Է»��� - QTY|PRICE|AMT
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"DETAIL_BTN",intCnt,intCnt,false
				mobjSCGLSpr.SetCellsLock2 .sprSht,true,"QTY|PRICE|AMT",intCnt,intCnt,false
				'��ư���·� ����
				If mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODE",intCnt) = "242001" Then
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�������Է�","DETAIL_BTN",intCnt,intCnt,,false
				Else
					mobjSCGLSpr.SetCellTypeButton2 .sprSht,"�󼼰���","DETAIL_BTN",intCnt,intCnt,,false
				End If
			Else
				'�ܰ�,����,�ݾ� �Է¹����� �ֵ��� ���� / ��ư�� lock
				mobjSCGLSpr.SetCellsLock2 .sprSht,false,"QTY|PRICE|AMT",intCnt,intCnt,false
				'�Ϲ����·� ����
				mobjSCGLSpr.SetCellTypeFloat2 .sprSht,"DETAIL_BTN",intCnt,intCnt,0,,,,,,,,False
		End If
		Next
	End With
End Sub

'------------------------------------------
' ������ ����
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim strPREESTNO
	Dim intCnt
	Dim vntData
	Dim strAGREEYEARMON
	Dim strJOBNO
	Dim strCHKCONFIRM
	Dim strNEWPREESTNO
	Dim strPRODUCTIONCHK
	Dim i
	
	if DataValidation =false then exit sub
	strMasterData = gXMLGetBindingData (xmlBind)
	
	with frmThis
	
			
		'��� �����͸� �����Ѵ�. print_seq �� ����� ���� ���� �ʴ°�찡 �߻��Ͽ� ������ �����̴� ��츦 ���´� _20120323_ SH
		for i = 1 to .sprSht.maxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		next
		
		'������ ������� ó��
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | IMESEQ | SAVEFLAG | PRINT_SEQ")		
		
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
			Else
				gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
				exit sub
			End If
		End If
		
		'��� Insert Update �б�ó��	
		If .txtPREESTNO.value <> "" Then   
		
			'�������ϰ�� Validation ����
			If .txtPREESTGBN.value  = "������" Then
				strCHKCONFIRM = "F"	
				
			Elseif .txtPREESTGBN.value  = "������" Then
			'�������ϰ�� ���� Validation �ʿ� - �ŷ������� �ۼ��Ǿ� û���� �Ȼ��¶� ���� �ɼ� �ִ� ������ ����.......
				strCHKCONFIRM = "T"
				if .txtENDFLAG.value = "T" Then
					gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
					Exit Sub
				End If
			End If
			strPREESTNO = .txtPREESTNO.value 
			
			'������ �ִ��� ������ äũ
			If PRODUCTIONCHK = True Then
				strPRODUCTIONCHK = "T"
			Else
				strPRODUCTIONCHK = "F"
			End IF
			
			'�̰��� ������ üũ�� �ϸ� �ȴ�.
			
			intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,strCHKCONFIRM,"UU",strPRODUCTIONCHK,"PROCESS")
				
			if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
				strJOBNO = .txtJOBNO.value
				SelectRtn_ProcessRtn(strPREESTNO)
				'1���� ����ȸ
				parent.jobMst_Tab1Search
				parent.jobMst_Tab5Search
			End If
		Else
		'������ȣ�� ���� ��. ���� ����Ʈ���� �ٸ� JOB �� ������ ���� �Ѱ��̴�.
			'�ϴ� ���� ������ȣ�� �ִ� �̹� ����Ǿ��ִ� �ٸ� JOB �� �ҷ����� ���...(���� ���� ��ȣ�� �����͸� ��ȸ�� �ü� �ִ�.)
			if parent.document.forms("frmThis").txtPREESTNO.value <> "" Then 
			 
				'������ JOB �� �ٸ��Ƿ� Copy Rule �� ������.	
				strPREESTNO = parent.document.forms("frmThis").txtPREESTNO.value
				strNEWPREESTNO = ""
				strJOBNO = .txtJOBNO.value
				
				intRtn = mobjPDCOPREESTDTL.ProcessRtn_HDRLESS_COPY(gstrConfigXml, strMasterData, strPREESTNO,strNEWPREESTNO, strJOBNO, strAGREEYEARMON)
				
				if not gDoErrorRtn ("ProcessRtn_HDRLESS_COPY") then
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
					gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
					strJOBNO = .txtJOBNO.value
					SelectRtn_ProcessRtn(strNEWPREESTNO)
					'1���� ����ȸ
					
					parent.jobMst_Tab1Search_EstCopy
					parent.jobMst_Tab5Search
				End If	
			
			else
				'������ JOB �� �ٸ��Ƿ� Copy Rule �� ������.	
				strPREESTNO = ""
				
				intRtn = mobjPDCOPREESTDTL.ProcessRtn_HDRLESS(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON)
				
				if not gDoErrorRtn ("ProcessRtn_HDRLESS") then
					mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
					gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
					strJOBNO = .txtJOBNO.value
					SelectRtn_ProcessRtn(strPREESTNO)
					'1���� ����ȸ
					
					parent.jobMst_Tab1Search_EstCopy
					parent.jobMst_Tab5Search
				End If	
			end if
		End If
		mstrSELECT = "T"
		mstrFIRSTPRODUCTIONCHECK = "N"
	End with
End Sub

'------------------------------------------
' û�������� ������(�������) ���� ����
'------------------------------------------
Sub ExeProcessRtn
	Dim intCnt
	Dim intRtn
	Dim strPREESTNO
	Dim intRtnSave
	Dim vntData, vntInParams, vntRet
	Dim intRtnChk
	Dim strPRODUCTIONCHK
	Dim strJOBNO, strAGREEYEARMON
	Dim strMasterData
	
	With frmThis
		if DataValidation =false then exit sub
		strMasterData = gXMLGetBindingData (xmlBind)
		
		
   		If PRODUCTIONCHK = True Then
			strPRODUCTIONCHK = "T"
		Else
			strPRODUCTIONCHK = "F"
		End IF

		strPREESTNO = .txtPREESTNO.value 
		strJOBNO = .txtJOBNO.value 
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'���� �����ʹ� �ٸ��� jobno�θ� �ϴ� ������ 
		'�ش����� ���� ����(preestno)�� ������ ���������� �������ΰ͸� û�������� �����ϱ⶧���� 
		'�ش����� ���� �����߿���  confrimflag�� 3�̻��� ���� �ִٸ� �������� ���������� �ٲܼ� ����.
		vntData = mobjPDCOPREESTDTL.DeleteRtn_Check(gstrConfigXml,mlngRowCnt,mlngColCnt, strJOBNO)
		
		If mlngRowCnt > 0  Then
			gOkMsgBox "[������]�� ���λ��� �Ǵ� û���� ����Ǿ����ϴ�. �����Ҽ� �����ϴ�","�����ȳ�!"
			Exit Sub
		end if 
		
		
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO | ITEMCODESEQ | DIVNAME | CLASSNAME | ITEMCODE | BTN | ITEMCODENAME | FAKENAME | STD | COMMIFLAG | QTY | PRICE | AMT | SUSUAMT | GBN | IMESEQ | SAVEFLAG | PRINT_SEQ")		
	
		if  not IsArray(vntData)  then 
			If  gXMLIsDataChanged (xmlBind) Then
			Else
				'gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
				intRtnChk = gYesNoMsgbox("������ڷᰡ �����ϴ�. ���������� �ݿ��Ͻðڽ��ϱ�?","����ȳ�")
				If intRtnChk <> vbYes then 
					exit sub
				End If
			End If
		End If
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		intCnt = mobjPDCOPREESTDTL.SelectRtn_ExeCount(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
	
		IF not gDoErrorRtn ("SelectRtn_ExeCount") then
			If intCnt = "" then
				
				'�������� ���°�� (Type1. �������� ���������� ��� ���� ��, Type2. ������ ������ �ۼ��ϴ� ���� ���������� ����)
				If .txtPREESTNO.value <> "" Then
				'Type1.
					intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
					if not gDoErrorRtn ("ProcessRtn_Confirm") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
						
						SelectRtn_ProcessRtn(strPREESTNO)
						'1���� ����ȸ
						parent.jobMst_Tab1Search
						parent.jobMst_Tab5Search
					End If
				Else
				'�������� �ִ� ���
				'Type2.
				'�̰�� ������ �ִ��� üũ �ؾ� ��
					strPREESTNO = ""
					intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","I",strPRODUCTIONCHK,"EXPROCESS")
					if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
						strJOBNO = .txtJOBNO.value
						SelectRtn_ProcessRtn(strPREESTNO)
						'1���� ����ȸ
						parent.jobMst_Tab1Search
						parent.jobMst_Tab5Search
					End If
				End If
			else
				
				'�������� �ִ°�� -
				if .txtENDFLAG.value = "T" Then
					gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ������ ������ �Ұ��� �մϴ�.","����ȳ�!"
					Exit Sub
				End If
				
				intRtnSave = gYesNoCancelMsgBox("�̹� �������� �ֽ��ϴ�. ���� ���������� Ȯ���Ͻðڽ��ϱ�?" & vbcrlf & vbcrlf & "[��:Ȯ�������� / �ƴϿ�:���������� �����]","������Ȯ������ �� ���")
				If intRtnSave = 6 then
					'���������� Ȯ�� �˾� ������,,,�������� �ִµ��� ������ ���� ��� �������� ������ ����.
					'�Ʒ� intRtnSave = 7 �ΰ��� ������ ���μ����� Ÿ����,,,, �˾��� ȣ��Ǿ� ���Ҽ� �־�� ��. 
					'�÷��׿��� ���� �������������� ���� �ϸ�Ʒ����μ����� �����ϵ��� ó��.,,,��Ҹ� �����ٸ� ���������� ���� ���� �ƴ��� ���°� �״�� �ݿ�
					'intCnt �� ���� ��������ȣ ��.
				
					if .txtPREESTNO.value = "" then 
						gErrorMsgBox " ������ ����Ǿ� ���� �ʽ��ϴ�. " & vbcrlf & " ������ ������ �� ������ ������ ���ؾ� �մϴ�.","����ȳ�" 
						exit sub
					else 
						vntInParams = array(Trim(.txtPREESTNO.value))
					end if
					
					vntRet = gShowModalWindow("PDCMJOBMST_PREESTCONFIRM.aspx",vntInParams, 1149,650)
					
					If vntRet = "TRUE" Then
							'�ٷ� ���������� ����
						If .txtPREESTNO.value <> "" Then
							intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
							if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
								mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
								gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
								strJOBNO = .txtJOBNO.value
								SelectRtn_ProcessRtn(strPREESTNO)
								'1���� ����ȸ
								parent.jobMst_Tab1Search
								parent.jobMst_Tab5Search
							End If
						Else
							
						End If
					
					End If
					
				Elseif intRtnSave = 7 then
					
					'�ٷ� ���������� ����
					If .txtPREESTNO.value <> "" Then
					'Type1. ������ ���� �����ϸ� �������ϴ� ������ ���������� ���� �ϴ°�,,,,
						intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","U",strPRODUCTIONCHK,"EXPROCESS")
						if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
							mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
							gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
							strJOBNO = .txtJOBNO.value
							SelectRtn_ProcessRtn(strPREESTNO)
							'1���� ����ȸ
							parent.jobMst_Tab1Search
							parent.jobMst_Tab5Search
						End If
					Else
					'Type2. ������ ���� ���� ���� �ƴ��ϸ�, ���� �Է�ī�� �Ȱ��� ���������� ��� ���� �ϴ°�
						strPREESTNO = ""
						intRtn = mobjPDCOPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON,"T","I",strPRODUCTIONCHK,"EXPROCESS")
						
						if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
							mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
							gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
							strJOBNO = .txtJOBNO.value
							
							msgbox "����Ȱ�����ȣ��: " & strPREESTNO
							SelectRtn_ProcessRtn(strPREESTNO)
							'1���� ����ȸ
							parent.jobMst_Tab1Search
							parent.jobMst_Tab5Search
						End If
					End If
				End If
			End IF
		End IF
		mstrSELECT = "T"
	End With	
End Sub

'===============================
'������ ���Կ��θ� �˾ƺ��� �Լ�
'===============================
Function PRODUCTIONCHK
	Dim intCnt_prod
	Dim intCnt_ProdChk
	intCnt_prodChk = 0
	with frmThis
		For intCnt_prod = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt_prod) = "242001" Then
				intCnt_prodChk = intCnt_prodChk + 1
			End If
		Next
		
		If intCnt_prodChk = 0  Then
			PRODUCTIONCHK = False	
		Else
			PRODUCTIONCHK = True
		End If
	End with
End Function

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt, intcnt2
   	Dim intRtnChk
   	
	'On error resume next
	with frmThis
  		'Field �ʼ� �Է� �׸� �˻�
  		If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "�������� �Է��Ͻʽÿ�.","����ȳ�"
			Exit Function
		End If
		If .txtAGREEYEARMON.value = "" Then
			gErrorMsgBox "�������� �Է��Ͻʽÿ�.","����ȳ�"
			Exit Function
		End If
		
		'Sheet �ʼ� �Է� �׸� �˻� 
		If .sprSht.MaxRows = 0 Then
				gErrorMsgBox "������ �� ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
				Exit Function
		End IF
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		intcnt2 = 0
   		for intCnt = 1 to .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNAME",intCnt) = "" _
				Or mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNAME",intCnt) = "" _
				Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt) = "" Or _
				mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODENAME",intCnt) = "" Then 
				
				gErrorMsgBox intCnt & " ��° ���� �����׸� ���� �� Ȯ���Ͻʽÿ�","�Է¿���"
				Exit Function
			End if
		next
   	End with
	DataValidation = true
End Function

'================================================
'--------------------�ڷ����--------------------
'================================================
Sub DeleteRtn ()
	Dim intRtn ,intRtn2
	Dim i, lngCnt
	Dim strPREESTNO
	Dim strJOBNO
	Dim strITEMCODESEQ
	Dim strITEMCODE
	Dim strSUBDETAIL
	Dim intCnt_Prod
	Dim strPRODUCTSUSUCHK
	Dim strCHKCONFIRM
	Dim intCount
	Dim intCount2
	
	with frmThis
		
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","�ڷ� ���� �ȳ�"
			Exit Sub
		end if
		
		'üũ�� ������ Ȯ�� 
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				intCount = intCount + 1
			end if
		next
		
		if intCount = 0 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, "�����ȳ�"
			Exit Sub
		end if

		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
	
		If .txtPREESTGBN.value  = "������" Then
			strCHKCONFIRM = "F"	
			
		Elseif .txtPREESTGBN.value  = "������" Then'txtJOBNO
		'�������ϰ�� ���� Validation �ʿ� - �ŷ������� �ۼ��Ǿ� û���� �Ȼ��¶� ���� �ɼ� �ִ� ������ ����.......
			strCHKCONFIRM = "T"
		End If
		
		strJOBNO = .txtJOBNO.value
		
		lngCnt =0
		intRtn2 = 0
		intCnt_Prod = 0
				
		for i = .sprSht.MaxRows to 1 step -1
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN	
				If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",i) <> ""  Then
					strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)
					strITEMCODESEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",i))
					strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
					strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
					For intCount2 = 1 to .sprSht.MaxRows
						If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCount2) = "242001" Then
							intCnt_Prod = intCnt_Prod +1
						End If
					Next
					
					If intCnt_Prod <> 0 Then
						strPRODUCTSUSUCHK = "T"
					Else 
						strPRODUCTSUSUCHK = "F"
					End If
					
					intRtn2 = mobjPDCOPREESTDTL.DeleteRtn(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL,strPRODUCTSUSUCHK,strJOBNO,strCHKCONFIRM)
					
					IF not gDoErrorRtn ("DeleteRtn") then
						lngCnt = lngCnt +1
						mobjSCGLSpr.DeleteRow .sprSht, i
					end if
				else
					strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",i)
					
					If strPREESTNO = "" Then 
						strPREESTNO = "9999999999"
						
						strITEMCODESEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i)
						
						strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
						strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
						intRtn2 = mobjPDCOPREESTDTL.DeleteRtn_TempDel(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL)
						IF not gDoErrorRtn ("DeleteRtn_TempDel") then
							lngCnt = lngCnt +1
							mobjSCGLSpr.DeleteRow .sprSht, i					
						end if
					else 
					
						strITEMCODESEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"IMESEQ",i)
						strITEMCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",i)
						strSUBDETAIL = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBDETAIL",i)
					
						
						intRtn2 = mobjPDCOPREESTDTL.DeleteRtn_TempDel(gstrConfigXml,strPREESTNO, strITEMCODESEQ, strITEMCODE,strSUBDETAIL)
						IF not gDoErrorRtn ("DeleteRtn_TempDel") then
							lngCnt = lngCnt +1
							mobjSCGLSpr.DeleteRow .sprSht, i					
						end if
						
					end if 
				end if	
			end if
		next
		'�������
		IF .txtPREESTGBN.value ="������" THEN
			Call ESTSUSUAMT_CHANGEVALUE2
		ELSE	
			Call SUSUAMT_CHANGEVALUE2
		END IF
		BUDGET_AMT_SUM
		'����Ǿ��ִ� ���� ������ DB �� ������� ���� ���� 
		If intRtn2 = 0 Then
   		Else
			DelProc
		End If
		'1���̶� �������� �ִٸ� �޼��� ���
		If lngCnt <> 0 Then
			gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
		End If
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
	End with
	err.clear
End Sub

'�����Ŀ� ��� �ٽ� ���
Sub DelProc
	Dim intHDR
	Dim strMasterData
	Dim strPREESTNO
	Dim strCHKCONFIRM
	Dim strAGREEYEARMON
	Dim strJOBNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		If .txtPREESTGBN.value  = "������" Then
			strCHKCONFIRM = "F"	
			
		Elseif .txtPREESTGBN.value  = "������" Then
		'�������ϰ�� ���� Validation �ʿ� - �ŷ������� �ۼ��Ǿ� û���� �Ȼ��¶� ���� �ɼ� �ִ� ������ ����.......
			strCHKCONFIRM = "T"
		End If
		strPREESTNO = .txtPREESTNO.value
		strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
		'intHDR = mobjPDCOPREESTDTL.ProcessRtn_DelProc(gstrConfigXml,strMasterData,strPREESTNO,strAGREEYEARMON,strCHKCONFIRM,"U")
		
		if not gDoErrorRtn ("ProcessRtn_DelProc") then
			strJOBNO = .txtJOBNO.value
			SelectRtn_ProcessRtn(strPREESTNO)
			'1���� ����ȸ
			parent.jobMst_Tab1Search
		End If
	End with
End Sub

Function CleanField (ByVal objField, ByVal objField1)
	if isobject(objField) then objField.value = ""
	if isobject(objField1) then objField1.value = ""
end Function


		</script>
	</HEAD>
	<body style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px; MARGIN-RIGHT: 0px" class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE border="0" cellSpacing="1" cellPadding="0" width="100%" align="left" height="98%">
				<TR>
					<TD>
						<TABLE id="tblTitle1" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD0" height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<TR>
											<td class="TITLE">û������
											</td>
										</TR>
									</table>
								</TD>
								<TD style="WIDTH: 100%" height="20" vAlign="middle" align="right">
									<!--Common Button Start--></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px; HEIGHT: 4px" colSpan="2"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD vAlign="top">
						<TABLE id="tblDATA" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
							align="left">
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80">��/��</TD>
								<TD class="SEARCHDATA" width="150"><INPUT style="WIDTH: 148px; HEIGHT: 22px" id="txtPREESTGBN" dataSrc="#xmlBind" class="NOINPUTB_L"
										title="��/�� ��������" dataFld="PREESTGBN" readOnly maxLength="10" size="32" name="txtPREESTGBN"></TD>
								<TD style="WIDTH: 92px; CURSOR: hand" class="SEARCHLABEL" width="92">��ǥJOB��</TD>
								<TD class="SEARCHDATA" width="260"><INPUT accessKey=",M" style="WIDTH: 256px; HEIGHT: 22px" id="txtJOBNAME" dataSrc="#xmlBind"
										class="NOINPUTB_L" title="JOB��" dataFld="JOBNAME" readOnly maxLength="255" size="37" name="txtJOBNAME"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL"><span id="strMsg_Amt">�������ݾ�</span></TD>
								<TD class="SEARCHDATA" width="105"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUMAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="������ + �ݾ�" dataFld="ESTSUMAMT" readOnly maxLength="20" size="32" name="txtESTSUMAMT"></SPAN></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL">��������ݾ�</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtSUMAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="������ + �ݾ�" dataFld="SUMAMT" readOnly maxLength="20" size="32" name="txtSUMAMT">
								</TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80">�����ڵ�</TD>
								<TD class="SEARCHDATA" width="150"><INPUT style="WIDTH: 148px; HEIGHT: 22px" id="txtPREESTNO" dataSrc="#xmlBind" class="NOINPUTB_L"
										title="�����ڵ�" dataFld="PREESTNO" readOnly maxLength="10" size="32" name="txtPREESTNO">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call CleanField('','')"
									width="94">������</TD>
								<TD class="SEARCHDATA" width="260"><INPUT accessKey=",M" style="WIDTH: 256px; HEIGHT: 22px" id="txtPREESTNAME" dataSrc="#xmlBind"
										class="INPUT_L" title="������" dataFld="PREESTNAME" maxLength="255" size="37" name="txtPREESTNAME"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80" align="right">(��)�ݾ�</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="�ݾ��հ�" dataFld="ESTAMT" readOnly maxLength="20" size="32" name="txtESTAMT"></TD>
								<TD style="CURSOR: hand" class="SEARCHLABEL" width="80" align="right">(��)�ݾ�</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtAMT" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="�ݾ��հ�" dataFld="AMT" readOnly maxLength="20" size="32" name="txtAMT"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')">������</TD>
								<TD class="SEARCHDATA"><INPUT accessKey="DATE,M" style="WIDTH: 72px; HEIGHT: 22px" id="txtAGREEYEARMON" dataSrc="#xmlBind"
										class="INPUT" title="����������" dataFld="AGREEYEARMON" maxLength="10" size="6" name="txtAGREEYEARMON">
									<IMG style="CURSOR: hand" id="imgCalEndarAGREE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarAGREE"
										align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')">���</TD>
								<TD class="SEARCHDATA"><TEXTAREA style="WIDTH: 256px; HEIGHT: 22px" id="txtMEMO" dataSrc="#xmlBind" dataFld="MEMO"
										wrap="hard" cols="10" name="txtMEMO"></TEXTAREA></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')"
									width="80">(��)������</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUSUAMT" dataSrc="#xmlBind"
										class="INPUT_R" title="������ݾ��հ�" dataFld="ESTSUSUAMT" maxLength="20" size="32" name="txtESTSUSUAMT"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField('', '')"
									width="80">(��)������</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 100px; HEIGHT: 22px" id="txtSUSUAMT" dataSrc="#xmlBind"
										class="INPUT_R" title="������ݾ��հ�" dataFld="SUSUAMT" maxLength="20" size="32" name="txtSUSUAMT"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtENDFLAG" dataSrc="#xmlBind" dataFld="ENDFLAG"
										size="1" type="hidden" name="txtENDFLAG"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtENDFLAGEXE" dataSrc="#xmlBind" dataFld="ENDFLAGEXE"
										size="1" type="hidden" name="txtENDFLAGEXE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtSETCONFIRMFLAG" dataSrc="#xmlBind" dataFld="SETCONFIRMFLAG"
										size="1" type="hidden" name="txtSETCONFIRMFLAG"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">Commission</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 148px; HEIGHT: 22px" id="txtCOMMITION" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="��������ݾ�" dataFld="COMMITION" readOnly maxLength="20" size="32" name="COMMITION">
								</TD>
								<TD style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">NonCommission</TD>
								<TD class="SEARCHDATA"><INPUT accessKey=",NUM" style="WIDTH: 256px; HEIGHT: 22px" id="txtNONCOMMITION" dataSrc="#xmlBind"
										class="NOINPUTB_R" title="���������ܱݾ�" dataFld="NONCOMMITION" readOnly maxLength="20" size="37" name="txtNONCOMMITION"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField(txtMEMO, '')"
									width="80">(��)��������</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 100px; HEIGHT: 22px" id="txtESTSUSURATE" dataSrc="#xmlBind" class="INPUT_R"
										title="��/�� ��������" dataFld="ESTSUSURATE" maxLength="20" size="37" name="txtESTSUSURATE"></TD>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL" onclick="vbscript:Call CleanField(txtMEMO, '')"
									width="80">(��)��������</TD>
								<TD class="SEARCHDATA"><INPUT style="WIDTH: 100px; HEIGHT: 22px" id="txtSUSURATE" dataSrc="#xmlBind" class="INPUT_R"
										title="��/�� ��������" dataFld="SUSURATE" maxLength="20" size="37" name="txtSUSURATE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtJOBNO" dataSrc="#xmlBind" dataFld="JOBNO"
										size="1" type="hidden" name="txtJOBNO"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtCREDAY" dataSrc="#xmlBind" dataFld="CREDAY"
										size="1" type="hidden" name="txtCREDAY"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtTIMCODE" dataSrc="#xmlBind" dataFld="TIMCODE"
										size="1" type="hidden" name="txtTIMCODE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtCLIENTCODE" dataSrc="#xmlBind" dataFld="CLIENTCODE"
										size="1" type="hidden" name="txtCLIENTCODE"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtSUBSEQ" dataSrc="#xmlBind" dataFld="SUBSEQ"
										size="1" type="hidden" name="txtSUBSEQ"></TD>
							</TR>
							<TR>
								<TD style="CURSOR: hand; HEIGHT: 25px" class="SEARCHLABEL">�ΰ��׸�</TD>
								<TD class="SEARCHDATA" colSpan="7"><INPUT accessKey="DATE" style="WIDTH: 72px; HEIGHT: 22px" id="txtPRINTDAY" class="INPUT"
										title="�����������" maxLength="10" size="6" name="txtPRINTDAY"> <IMG style="CURSOR: hand" id="imgimgCalEndarCREDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgimgCalEndarCREDAY" align="absMiddle" src="../../../images/btnCalEndar.gIF"
										height="15">&nbsp;<SELECT style="WIDTH: 132px" id="cmbESTTYPE" title="������������" name="cmbESTTYPE">
										<OPTION selected value="1">ESTIMATE</OPTION>
										<OPTION value="2">ESTIMATE/ACTUAL</OPTION>
										<OPTION value="3">ACTUAL</OPTION>
									</SELECT><IMG style="CURSOR: hand" id="imgPrintEst" onmouseover="JavaScript:this.src='../../../images/imgPrintEstOn.gIF'"
										title="������ �� ���������� �����Ͻþ� ������ �������� ����մϴ�" onmouseout="JavaScript:this.src='../../../images/imgPrintEst.gif'"
										border="0" name="imgPrintEst" alt="���������(��)." align="absMiddle" src="../../../images/imgPrintEst.gIF"
										width="100" height="20">&nbsp;<IMG style="CURSOR: hand" id="imgPrintEstBasic" onmouseover="JavaScript:this.src='../../../images/imgPrintEstBasicOn.gIF'"
										title="������ �� ���������� �����Ͻþ� ������ �⺻�������� ����մϴ�" onmouseout="JavaScript:this.src='../../../images/imgPrintEstBasic.gif'"
										border="0" name="imgPrintEstBasic" alt="���������(�⺻)." align="absMiddle" src="../../../images/imgPrintEstBasic.gIF" width="120"
										height="20">&nbsp;<IMG style="CURSOR: hand" id="imgCFInput" onmouseover="JavaScript:this.src='../../../images/imgCFInputOn.gIF'"
										title="CF ���ֳ��� �׸��� �����Ͽ� �������� �����մϴ�." onmouseout="JavaScript:this.src='../../../images/imgCFInput.gif'"
										border="0" name="imgCFInput" alt="" align="absMiddle" src="../../../images/imgCFInput.gIF" height="20"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 25px" id="spacebar" class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD>
						<TABLE id="tblTitle3" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD id="TD1" height="20" width="120" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="200">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="53" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td height="3"></td>
										</tr>
										<tr>
											<td class="TITLE">���γ���&nbsp;<span id="strMsgBox"></span>
											</td>
										</tr>
									</table>
								</TD>
								<td class="TITLE">�ݾ��հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT_TOTAL" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
										accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT_TOTAL">&nbsp; 
									�����հ� : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
								<TD height="20" vAlign="middle" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
										<TR>
											<td width="62" align="left"><input accessKey="NUM," style="VISIBILITY: hidden; WIDTH: 5px" id="txtPRINT_SEQ" value="1"
													maxLength="2" name="txtPRINT_SEQ"><IMG style="CURSOR: hand" id="imgTableUp" border="0" name="imgTableUp" alt="�ڷḦ �ø��ϴ�."
													align="absMiddle" src="../../../images/imgTableUp.gif"> <IMG style="CURSOR: hand" id="imgTableDown" border="0" name="imgTableDown" alt="�ڷḦ �����ϴ�."
													align="absMiddle" src="../../../images/imgTableDown.gif"></td>
											<TD><IMG style="CURSOR: hand" id="ImgBasicFormat" onmouseover="JavaScript:this.src='../../../images/ImgBasicFormatOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/ImgBasicFormat.gIF'" border="0"
													name="ImgBasicFormat" alt="����Ÿ�Ժ� �⺻���� �����մϴ�" src="../../../images/ImgBasicFormat.gIF"
													height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" border="0" name="imgRowAdd"
													alt="�ڷ��Է��� ���� �����߰��մϴ�." src="../../../images/imgRowAdd.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" border="0" name="imgRowDel"
													alt="������ ���������մϴ�." src="../../../images/imgRowDel.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgSave"
													alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgBonSave" onmouseover="JavaScript:this.src='../../../images/imgBonSaveOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgBonSave.gIF'" border="0" name="imgBonSave"
													alt="��������������" src="../../../images/imgBonSave.gIF" height="20"></TD>
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" height="20"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 99%" vAlign="top" align="center">
						<DIV style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" id="pnlTab1"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="9737">
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
				</tr>
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 1%" vAlign="top" align="center">
						<DIV style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%" id="pnlTab2"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_copy" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" >
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="31962">
								<PARAM NAME="_ExtentY" VALUE="503">
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
				</tr>
				<TR>
					<TD style="WIDTH: 1040px" id="lblStatus" class="BOTTOMSPLIT"></TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
