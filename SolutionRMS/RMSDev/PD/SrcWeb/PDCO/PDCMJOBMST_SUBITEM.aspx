<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_SUBITEM.aspx.vb" Inherits="PD.PDCMJOBMST_SUBITEM" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�󼼰�������</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBMST_SUBITEM.aspx
'��      �� : JOBMST�� �ι�° �� PDCMJOBMST_ESTDTL �� �󼼰��� ��ư�� Ŭ���Ͽ����� ó�� 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/19 By KimTH
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">
		
Dim mlngRowCnt,mlngColCnt
Dim mobjPDCOPREESTSUB		'cc ��ü ���������� ����
Dim mstrITEMCODE			'���� �׸� �ڵ�
Dim mstrITEMCODESEQ			'���� �׸� �ڵ� ����
Dim mstrSAVEFLAG			'�󼼳����� ����Ǿ��ִ��� �������� ����
Dim mstrPREESTNO			'���� ��ȣ
Dim mstrCheck	
Dim mstrGBN					'������ / ������
Dim mlngTempRowCnt,mlngTempColCnt
Dim mstrSAVEGBN
Dim mstrJOBNO				'JOBNO
Dim mstrCHANGEFALG			'����Ȯ�� �÷���(������ �Ұ�� ��ü ������ ����� ���ܸ� ó���Ѵ�.)  [ T/F  (T �Ϲ� �̺�Ʈ�� / F ���� �̺�Ʈ�� �߻��Ұ��)]


mstrCheck = True	
mstrCHANGEFALG = "F"

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	Dim lngPRICEAMT
	Dim vntData
	Dim INPUTAMT
	'�������� ��� ���� �ݾ��� ���� �������� ��� ���� ������ �����Ѵ�.
	with frmThis
		
		'������ �Ǿ��ִ��� �ƴ��� ���� [������ �Ǿ��ٴ°� INPUT �� �ݾ��� �ִٴ� ��!]
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		'�ݱ� ��ư�� ������� window_onunload �� �ݱ� ��ư�� ���ؼ� 1��Ÿ�� ���� �����찡 �����鼭 �ѹ�(�ι�) Ÿ�°Ű���. �׷��� ��ü ���� ����
		'set mobjPDCOPREESTSUB = gCreateRemoteObject("cPDCO.ccPDCOPREESTSUB")
		
		vntData = mobjPDCOPREESTSUB.SelectRtn_INPUTAMT(gstrConfigXml,mlngRowCnt,mlngColCnt, mstrJOBNO, mstrITEMCODE, mstrITEMCODESEQ)
		
		'Set mobjPDCOPREESTSUB = Nothing

		'�������ϰ��
		If mstrGBN = "F" Then
			if vntData(0,1) = "" then
				INPUTAMT = split("False;" & mstrCHANGEFALG ,";")
			ELSE 
				INPUTAMT = split(vntData(0,1) & ";" & mstrCHANGEFALG ,";")
			END IF		
		'�������ϰ��
		Elseif mstrGBN = "T" Then
		
			if vntData(1,1) = "" then
				INPUTAMT = split("False;" & mstrCHANGEFALG ,";")
			ELSE 
				INPUTAMT = split(vntData(1,1) & ";" & mstrCHANGEFALG ,";")
			END IF
		End If
		window.returnvalue = INPUTAMT
		
	End with
	Set mobjPDCOPREESTSUB = Nothing
End Sub

Sub imgClose_onclick ()
	EndPage
	'Window_OnUnload
End Sub

'-----------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------

Sub imgRowAdd_onclick ()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ���߰��� �Ұ��� �մϴ�.","���߰��ȳ�!"
		Exit Sub
	End If
	call sprSht_Keydown(meINS_ROW, 0)
End Sub


Sub imgSave_onclick ()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	ProcessRtn
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

Sub imgRowDel_onclick()
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ���� �Ұ��� �մϴ�.","������ȳ�!"
		Exit Sub
	End If
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgAllType_onclick()
	Dim intCnt
	Dim strSEQ
	Dim strSEQString
	Dim strLen
	Dim vntData
	Dim dblRow
	
	if mstrSAVEGBN = "T" Then
		gErrorMsgBox "û����û �� �ŷ����� �������̹Ƿ� ��üŸ���׸� ���Ⱑ �Ұ��� �մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	With frmThis
		
		For intCnt = 1 To .sprSht.MaxRows
			strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBITEMCODESEQ",intCnt)
			strSEQString = strSEQString & strSEQ & ","
		Next 
		
		dblRow = .sprSht.MaxRows + 1
		if strSEQString <> "" then
			strSEQString = MID(strSEQString,1, Len(strSEQString) -1)
		else 
			strSEQString = "999999"
		end if 
		
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		vntData = mobjPDCOPREESTSUB.SelectRtn_Fill(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,strSEQString)
		if not gDoErrorRtn ("SelectRtn_Fill") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, dblRow, mlngColCnt, mlngRowCnt, True
				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end If
   		end if
   					
	End With
End Sub

Sub ImgMoveData_onclick
	Dim intCnt
	with frmThis
		If mstrGBN = "F" Then 
			gErrorMsgBox "������ ���´� ��������� �Է��Ҽ� �����ϴ�.","�Է¾ȳ�!"
			EXIT SUB
		END IF 
		
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"QTY", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"TERM", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt,"1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
	End with
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
	Dim strCheckRow
	
	with frmThis
	
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		'PRINT_SEQ �� 1���� ������ΰ� �ƴҼ�������. ��ü�� ���鼭 ���� ���� ���� ã�Ƴ���.
		strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",1)
		for i=1 to .sprSht.MaxRows-1
			for j=i+1 to .sprSht.MaxRows
				IF strCheckRow > mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j) then
					strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j)
				end if
			Next
		Next	
		strCheckRow= strCheckRow -1		
		
		
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
		
		for i=1 to .sprSht_copy.MaxRows
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow- strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow-strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMCODESEQ",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMNAME",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"TERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"TERM",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"MEMO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEPRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEPRICE",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEQTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEQTY",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXETERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXETERM",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEAMT",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEMEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEMEMO",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow -strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"NEWFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"NEWFLAG",strRow -strPRINT_SEQ)
			
			'�⺻���� Y�� �ڴ´�
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


Sub sprSht_DownCopy(strRow)
	Dim strPRINT_SEQ 
	Dim strCopyRow
	Dim strCopySeq
	
	
	with frmThis
		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		
		'PRINT_SEQ �� 1���� ������ΰ� �ƴҼ�������. ��ü�� ���鼭 ���� ���� ���� ã�Ƴ���.
		strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",1)
		for i=1 to .sprSht.MaxRows-1
			for j=i+1 to .sprSht.MaxRows
				IF strCheckRow > mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j) then
					strCheckRow = mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",j)
				end if
			Next
		Next	
		strCheckRow= strCheckRow -1	
		
		
		'row����	
		.sprSht_copy.MaxRows = strPRINT_SEQ+1
	
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"CHK",i, mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",strRow+ strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PREESTNO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",strRow+strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRINT_SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRINT_SEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMCODESEQ",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMCODESEQ",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SUBITEMNAME",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SUBITEMNAME",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"PRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"QTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"TERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"TERM",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"AMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"AMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"MEMO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEPRICE",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEPRICE",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEQTY",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEQTY",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXETERM",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXETERM",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEAMT",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEAMT",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"EXEMEMO",i, mobjSCGLSpr.GetTextBinding( .sprSht,"EXEMEMO",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"SAVEFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"SAVEFLAG",strRow +strPRINT_SEQ)
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"NEWFLAG",i, mobjSCGLSpr.GetTextBinding( .sprSht,"NEWFLAG",strRow +strPRINT_SEQ)
			
			mobjSCGLSpr.SetTextBinding .sprSht_copy,"MOVEFLAG",i, "Y"
			
			strPRINT_SEQ = strPRINT_SEQ -1
		next
		

		strPRINT_SEQ = .txtPRINT_SEQ.value
		
		for i=1 to .sprSht_copy.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		for i=1 to .sprSht_copy.MaxRows
			if mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i) = "1" then
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow+strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			else
				mobjSCGLSpr.SetTextBinding .sprSht,"CHK",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"CHK",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PREESTNO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMCODESEQ",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMCODESEQ",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SUBITEMNAME",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SUBITEMNAME",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"QTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"TERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"TERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"AMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"MEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEPRICE",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEQTY",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXETERM",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEAMT",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"EXEMEMO",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"EXEMEMO",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"SAVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"SAVEFLAG",i)
				mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"NEWFLAG",i)
				
				mobjSCGLSpr.SetTextBinding .sprSht,"MOVEFLAG",mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)-strCheckRow-strPRINT_SEQ , mobjSCGLSpr.GetTextBinding( .sprSht_copy,"MOVEFLAG",i)
				
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, mobjSCGLSpr.GetTextBinding( .sprSht_copy,"PRINT_SEQ",i)
			End if 		
			
			
		next
		
		.sprSht_copy.MaxRows = 0
		
	end with
End Sub


'=============================================================
'Sheet Event
'=============================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if
	end with	
End Sub


Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim lngPRICE
	Dim lngQTY
	Dim lngTERM
	Dim lngAMT
	Dim lngCalCul
	
	Dim lngEXEPRICE
	Dim lngEXEQTY
	Dim lngEXETERM
	Dim lngEXEAMT
	Dim lngEXECalCul
	
	Dim lngOldAMT
	Dim lngOldPRICE
	Dim lngOldExeAmt
	Dim lngOldExePRICE
	
	With frmThis	
			
			lngOldAMT		= mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			lngOldPRICE		= mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
			lngOldExeAmt	= mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
			lngOldExePRICE	= mobjSCGLSpr.GetTextBinding(.sprSht,"EXEPRICE",Row)
				
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Or  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"QTY") Or _
																	Col = mobjSCGLSpr.CnvtDataField(.sprSht,"TERM") Then
   				lngPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   				lngQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   				lngTERM = mobjSCGLSpr.GetTextBinding(.sprSht,"TERM",Row)
		   		
   				lngCalCul = lngPRICE * lngQTY * lngTERM
   				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngCalCul
		   		
   				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   				.txtDIVNAME.focus
				.sprSht.focus
		   		
   					'�󼼳����ݾ��� 0���� �����Ұ�� ����Ǿ���..._20110830_SH
   		   		IF lngOldAMT <> mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) THEN
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				ELSE
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   				END IF 
	   	
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
   				lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
	   			
   				If lngAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"TERM",Row, "1"
   				End If
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Or  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEQTY") Or _
   																		Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXETERM") Then
	   			
   				lngEXEPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEPRICE",Row)
   				lngEXEQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEQTY",Row)
   				lngEXETERM = mobjSCGLSpr.GetTextBinding(.sprSht,"EXETERM",Row)
		   		
   				lngEXECalCul = lngEXEPRICE * lngEXEQTY * lngEXETERM
   				mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",Row, lngEXECalCul
		   		
   				lngEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
   				.txtDIVNAME.focus
				.sprSht.focus


   				IF lngOldEXEAMT <> mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) THEN
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				ELSE
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   				END IF 
	   				   			
   			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
   				lngEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row)
	   			
   				If lngEXEAMT <> 0 Then
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "-1"
   				Else
   					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",Row, "0"
   					mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",Row, "1"
   				End If
   			End If
   		AMT_SUM
   	End with 
   	'���� Sprsht ���濡 ���� �÷��� ó��
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
   	
End Sub



Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"NEWFLAG")  Then
			
				intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
				DefaultValue
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
	with frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, mstrPREESTNO
		
		mobjSCGLSpr.SetTextBinding .sprSht,"NEWFLAG",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"NEWFLAG",.sprSht.MaxRows -1)
		
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",.sprSht.ActiveRow, 0
		mobjSCGLSpr.SetTextBinding .sprSht,"QTY",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"TERM",.sprSht.ActiveRow, 1
		mobjSCGLSpr.SetTextBinding .sprSht,"EXEPRICE",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"EXEQTY",.sprSht.ActiveRow, 0	
		mobjSCGLSpr.SetTextBinding .sprSht,"EXETERM",.sprSht.ActiveRow, 1
		
		mobjSCGLSpr.SetTextBinding .sprSht,"PRINT_SEQ",.sprSht.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht,"PRINT_SEQ",.sprSht.MaxRows -1)+1	
	End With
End Sub




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

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		'sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow	
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") Then
				strCOLUMN = "PRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
				strCOLUMN = "EXEPRICE"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
				strCOLUMN = "EXEAMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE"))  Then
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
		If .sprSht.MaxRows >0 Then
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"PRICE") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEPRICE") Then
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

Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCOPREESTSUB = gCreateRemoteObject("cPDCO.ccPDCOPREESTSUB")

	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue

	gSetSheetDefaultColor
	with frmThis
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����

		for i = 0 to intNo
			select case i
				case 0 : .txtDIVNAME.value = vntInParam(i)			'��з���
				case 1 : .txtCLASSNAME.value = vntInParam(i)		'�ߺз���
				case 2 : .txtITEMCODENAME.value = vntInParam(i)		'�����׸��
				case 3 : mstrITEMCODE = vntInParam(i)				'�����׸��ڵ�
				case 4 : mstrSAVEFLAG = vntInParam(i)				'��ȸ�� �����Է� ��ȸ����, ����ȳ��� ��ȸ���� ����
				case 5 : mstrPREESTNO = vntInParam(i)				'������ȣ�� �����´�.
				case 6 : mstrGBN = vntInParam(i)					'T:�������� ó��, F:��������ó��
				case 7 : mstrITEMCODESEQ = vntInParam(i)			'�����׸��ڵ� ���� ���� ���ϰ� �Ѿ�ý� ��¥ ���������[IMESEQ]
				case 8 : mstrSAVEGBN = vntInParam(i)				'û�� Ȯ�ο� �÷���
				case 9 : mstrJOBNO = vntInParam(i)					'�θ�â�� JOBNO
			end select
		next
			
	'**************************************************
	'***Sum Sheet ������
	'**************************************************	
	gSetSheetColor mobjSCGLSpr, .sprSht
	mobjSCGLSpr.SpreadLayout .sprSht, 19, 0
	mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG "
	mobjSCGLSpr.SetHeader .sprSht,		    "����|������ȣ|�̵�|����|�ڵ�|�󼼰����׸�|�ܰ�|����|�Ⱓ|�ݾ�|���|����ܰ�|����|�Ⱓ|����ݾ�|������|���屸��|�ű����Կ���|�̵�����"
	mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   4|      10|   4|   4|   4|          21|  12|   6|   4|  12|  12|      12|   6|   4|      12|      12|      10|          10|       6"
	mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRINT_SEQ|SUBITEMCODESEQ|PRICE|AMT|EXEPRICE|EXEAMT", -1, -1, 0
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|TERM|EXEQTY|EXETERM", -1, -1, 1
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht,true,"SAVEFLAG|SEQ"
	mobjSCGLSpr.SetCellAlign2 .sprSht, "SUBITEMNAME|MEMO|EXEMEMO",-1,-1,0,2,false ' ����
	mobjSCGLSpr.SetCellAlign2 .sprSht, "PRINT_SEQ | SEQ | SUBITEMCODESEQ",-1,-1,2,2,false ' ����
	
	
	gSetSheetColor mobjSCGLSpr, .sprSht_copy
	mobjSCGLSpr.SpreadLayout .sprSht_copy, 19, 0
	mobjSCGLSpr.SpreadDataField .sprSht_copy,    "CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG "
	mobjSCGLSpr.SetHeader .sprSht_copy,		    "����|������ȣ|�̵�|����|�ڵ�|�󼼰����׸�|�ܰ�|����|�Ⱓ|�ݾ�|���|����ܰ�|����|�Ⱓ|����ݾ�|������|���屸��|�ű����Կ���|�̵�����"
	mobjSCGLSpr.SetColWidth .sprSht_copy, "-1",  "   4|      10|   4|   4|   4|          21|  12|   6|   4|  12|  12|      12|   6|   4|      12|      12|      10|          10|     6"
	mobjSCGLSpr.SetRowHeight .sprSht_copy, "0", "15"
	mobjSCGLSpr.SetRowHeight .sprSht_copy, "-1", "13"
	mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_copy, "CHK "
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "PRINT_SEQ|SUBITEMCODESEQ|PRICE|AMT|EXEPRICE|EXEAMT", -1, -1, 0
	mobjSCGLSpr.SetCellTypeFloat2 .sprSht_copy, "QTY|TERM|EXEQTY|EXETERM", -1, -1, 1
	mobjSCGLSpr.SetCellTypeEdit2 .sprSht_copy, "PREESTNO|SUBITEMNAME|MEMO|EXEMEMO", -1, -1, 255
	mobjSCGLSpr.SetCellsLock2 .sprSht_copy,true,"PRINT_SEQ|SAVEFLAG|SEQ|NEWFLAG|SUBITEMCODESEQ"
	mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "SUBITEMNAME|MEMO|EXEMEMO",-1,-1,0,2,false ' ����
	mobjSCGLSpr.SetCellAlign2 .sprSht_copy, "PRINT_SEQ | SEQ | SUBITEMCODESEQ",-1,-1,2,2,false ' ����
	mobjSCGLSpr.ColHidden .sprSht, "SAVEFLAG | NEWFLAG | MOVEFLAG", true
	
	'�������̶�� ���������� �����Ұ�
	If mstrGBN = "F" Then
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO"

	'�������̶�� ���������� �����Ұ�
	Else
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PRICE | QTY | TERM | AMT | MEMO"
	
	End If

	pnlTab1.style.visibility = "visible" 
	pnlTab2.style.visibility = "visible" 
	End with
	
	'ȭ�� �ʱⰪ ����
	SelectRtn
	
End Sub

Sub EndPage
	gEndPage
End Sub

'=============================================================
'��ȸ
'=============================================================
Sub SelectRtn
	Dim vntData
   	Dim vntData_Temp
   	Dim vntData_Dtl
    Dim intCnt
    Dim intRtn
	
    
	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		mlngTempRowCnt=clng(0)
		mlngTempColCnt=clng(0)
			
			'�������� �� ������ ���� ��ȸ�Ѵ�.(���� INPUT�� ����Ǿ��ִ� ��)
			intCnt = mobjPDCOPREESTSUB.SelectRtn_TempCount(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt, mstrJOBNO, mstrITEMCODE, mstrITEMCODESEQ,mstrPREESTNO)
			IF not gDoErrorRtn ("SelectRtn_ExeCount") then
				If mlngTempRowCnt > 0 Then
					'�󼼳����� �ִٸ� �󼼳����� �����ش�.
					vntData_Temp = mobjPDCOPREESTSUB.SelectRtn_Temp(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ,mstrPREESTNO)
					if not gDoErrorRtn ("SelectRtn_Temp") then
						if mlngRowCnt > 0 Then
							mobjSCGLSpr.SetClipbinding .sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True
							gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
						Else	
							.sprSht.MaxRows = 0
							gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   						end If
   					end if
				ELSE
					'�� ������ ������ ����Ǿ��ִ� ������ ��������
					intRtn = mobjPDCOPREESTSUB.SelectRtn_DtlCount(gstrConfigXml,mlngTempRowCnt,mlngTempColCnt,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ,mstrPREESTNO)
					IF not gDoErrorRtn ("SelectRtn_DtlCount") then
						If mlngTempRowCnt > 0 Then
							vntData_Dtl = mobjPDCOPREESTSUB.SelectRtn_Dtl(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,mstrITEMCODESEQ,mstrJOBNO)
							if not gDoErrorRtn ("SelectRtn_Dtl") then
								if mlngRowCnt > 0 Then
									mobjSCGLSpr.SetClipbinding .sprSht, vntData_Dtl, 1, 1, mlngColCnt, mlngRowCnt, True
									gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   								Else	
   									.sprSht.MaxRows = 0
   									gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   								end If
   							end if
						Else
							'�󼼳����� ���ٸ� ��ü ������ �����ش�.
							vntData = mobjPDCOPREESTSUB.SelectRtn_Empty(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO,mstrITEMCODE,mstrITEMCODESEQ)
							if not gDoErrorRtn ("SelectRtn_Empty") then
								if mlngRowCnt > 0 Then
									mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
									gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   								Else	
   									.sprSht.MaxRows = 0
   									gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   								end If
   							end if
						End If
					END IF
				END IF 
			END IF 
		
	window.setTimeout "AMT_SUM",1	
	.txtSELECTAMT.value = 0
	mstrCHANGEFALG = "F"
   	end with
End Sub


Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	
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
		
		IntEXEAMT = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtEXESUMAMT.value = 0
		else
			.txtEXESUMAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtEXESUMAMT,0,True)
		End If
	End With
End Sub
'------------------------------------------------------------------
'---------------------------------����----------------------------- 
'------------------------------------------------------------------
Sub processRtn
	Dim vntData
	Dim intRtn
	Dim intCnt
	Dim ProcessCHK
	Dim i
	Dim strNEWFALG

	with frmThis
		if .sprSht.MaxRows = 0 Then 
			gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
			EXIT Sub
		end if
		
		'ù��° NEWFLAG �� �����´� (������ �������)
		strNEWFALG = mobjSCGLSpr.GetTextBinding(.sprSht,"NEWFLAG",1)
		
		if strNEWFALG = "Y" then
			For i = 1 To .sprSht.MaxRows 
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				intCnt = intCnt + 1
				end if 
			Next
			
			If intCnt = 0 Then
				gErrorMsgBox "�����Ͻ�(�����) �ڷᰡ �����ϴ� ���� �Ͻ� �ڷḦ �����ϼ���.","�������"
				EXIT Sub
			End If
		end if 
		
		'�ڷḦ �����ϱ����� �ѹ��� Ȯ���Ѵ�.[������ �����ϱ� ������ ������ �ڷ�� ����� �ڷ��� ������ ��Ȯ�� �ؾ� �Ѵ�.]
		ProcessCHK = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ����� Ȯ��")
		IF ProcessCHK <> vbYes then exit Sub
		
				'insert �÷��� ���� 
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
   		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | PREESTNO | PRINT_SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | SEQ | NEWFLAG | MOVEFLAG")
		
		
		intRtn = mobjPDCOPREESTSUB.ProcessRtn(gstrConfigXml,vntData, mstrJOBNO,mstrPREESTNO, mstrITEMCODE, mstrITEMCODESEQ, mstrGBN,strNEWFALG)
	
		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
			.sprSht.focus()
			mstrCHANGEFALG = "T"
		End If
	End with
End Sub

'----------------------------------------
'--------------------����----------------
'----------------------------------------
Sub DeleteRtn
	Dim vntData
	Dim intCnt, intRtn, i
	'����Key ����
	Dim dblSEQ
	
	with frmThis
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		'���õ� �ڷḦ ������ ���� ����
		
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				mobjSCGLSpr.DeleteRow .sprSht,i
				intCnt = intCnt + 1
			end if 
		next
		
				'insert �÷��� ���� 
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
   		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | PREESTNO | PRINT_SEQ | SEQ | SUBITEMCODESEQ | SUBITEMNAME | PRICE | QTY | TERM | AMT | MEMO | EXEPRICE | EXEQTY | EXETERM | EXEAMT | EXEMEMO | SAVEFLAG | NEWFLAG | MOVEFLAG ")
		
		'�ش� INPUT�� ��� ������ ������ ���� 
		intRtn = mobjPDCOPREESTSUB.DeleteRtn_Temp(gstrConfigXml,vntData,mstrJOBNO,mstrITEMCODE,mstrITEMCODESEQ, mstrGBN)
		
		If not gDoErrorRtn ("DeleteRtn_Temp") Then
			gOkMsgBox  intCnt & "���� ����" & mePROC_DONE, "���� �ȳ�!"
   		End IF
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		
		if intRtn = 0 then 
			'���� Ȯ���÷���(��ü ���������� �÷���)
			mstrCHANGEFALG = "T"
			'��ü ���� ������� ���°��� �ƴ϶� ��� Ÿ���� �����Ҽ� �ֵ��� �Ѵ�.
			call ImgAllType_onclick
		end if 
		.sprSht.focus()	
	End with
	err.clear	
End Sub

		</script>
	</HEAD>
	<BODY class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table style="WIDTH: 100%; HEIGHT: 24px" cellSpacing="0" cellPadding="0" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="76" background="../../../images/back_p.gIF"
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
								<td class="TITLE">�󼼰�������&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="SEARCHDATA" width="100%">
				<tr>
					<td class="searchDATA" colSpan="7">&nbsp;��з� <INPUT class="NOINPUTB_L" id="txtDIVNAME" title="��з���" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="32" name="txtDIVNAME">&nbsp;&nbsp;&nbsp;�ߺз�
						<INPUT class="NOINPUTB_L" id="txtCLASSNAME" title="�ߺз���" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="29" name="txtCLASSNAME"> &nbsp;&nbsp;&nbsp;&nbsp;�����׸�
						<INPUT class="NOINPUTB_L" id="txtITEMCODENAME" title="�����׸��" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="30" name="txtITEMCODENAME">&nbsp;</td>
					<td align="right" width="54"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="ȭ���� �ݽ��ϴ�."
							src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
				</tr>
			</table>
			</TABLE> 
			<!--Top Define Table Start-->
			<TABLE cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF"
				border="0">
				<TR>
					<TD align="left" width="100%" height="10"></TD>
				</TR>
			</TABLE>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">�� �� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSUMAMT"> <INPUT class="NOINPUTB_R" id="txtEXESUMAMT" title="�հ�ݾ�" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtEXESUMAMT">&nbsp; <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<TD align="right" width="700">
						<input id="txtPRINT_SEQ" style="VISIBILITY: hidden; WIDTH: 5px" type="text" maxLength="2"
							value="1" name="txtPRINT_SEQ" accessKey="NUM,"><IMG id="imgTableUp" style="CURSOR: hand" alt="�ڷḦ �ø��ϴ�." src="../../../images/imgTableUp.gif"
							align="absMiddle" border="0" name="imgTableUp"> <IMG id="imgTableDown" style="CURSOR: hand" alt="�ڷḦ �����ϴ�." src="../../../images/imgTableDown.gif"
							align="absMiddle" border="0" name="imgTableDown"> <IMG id="ImgAllType" onmouseover="JavaScript:this.src='../../../images/ImgAllTypeOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAllType.gIF'" height="20" alt="�����׸��������ݴϴ�." src="../../../images/ImgAllType.gIF"
							align="absMiddle" border="0" name="ImgAllType">&nbsp;<IMG id="ImgMoveData" onmouseover="JavaScript:this.src='../../../images/ImgMoveDataOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMoveData.gIF'" height="20" alt="���������� �� ����������� �����մϴ�." src="../../../images/ImgMoveData.gIF"
							align="absMiddle" border="0" name="ImgMoveData">&nbsp;<IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" height="20" alt="�ڷ��Է��� ���� �����߰��մϴ�." src="../../../images/imgRowAdd.gIF"
							align="absMiddle" border="0" name="imgRowAdd">&nbsp;<IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="������ ���������մϴ�." src="../../../images/imgRowDel.gIF"
							align="absMiddle" border="0" name="imgRowDel">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
							onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" align="absMiddle"
							border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�."
							src="../../../images/imgExcel.gIF" width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</TD>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--����-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
								VIEWASTEXT>
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="26009">
								<PARAM NAME="_ExtentY" VALUE="11933">
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
				<tr>
					<TD style="WIDTH: 100%; HEIGHT: 0%" vAlign="top" align="center">
						<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id="sprSht_copy" height="100%" width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
								<PARAM NAME="_Version" VALUE="393216">
								<PARAM NAME="_ExtentX" VALUE="26009">
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
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
				</TR>
			</table>
		</form>
	</BODY>
</HTML>
