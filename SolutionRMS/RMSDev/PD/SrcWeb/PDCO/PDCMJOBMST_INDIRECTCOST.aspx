<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST_INDIRECTCOST.aspx.vb" Inherits="PD.PDCMJOBMST_INDIRECTCOST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���������</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMJOBMST_SUBITEM.aspx
'��      �� : JOBMST�� �ι�° �� PDCMJOBMST_ESTDTL �� ������ó�� ��ư�� Ŭ���Ͽ����� ó�� 
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/10/28 By KimTH
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
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
Dim mlngTempRowCnt,mlngTempColCnt
Dim mobjPDCOPREESTINDIRECTCOST
Dim mstrPREESTNO			'������ȣ
Dim mstrCheck	
Dim mstrGBN					'�������� ������ ����
Dim mstrSAVEGBN				'û����û ������ ���� �÷���
Dim mstrFIRSTPRODUCTIONCHECK'�ڵ� ������ ���� �÷���
Dim mstrProcessData			'���� ��ȸ�� �󼼳����� �����͸� �ӽ� ���̺� ���� �ϱ�����.
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
Dim vntData
Dim returnAMT

	with frmThis
	
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		'set mobjPDCOPREESTINDIRECTCOST = gCreateRemoteObject("cPDCO.ccPDCOPREESTINDIRECTCOST")
		'HDRINPUT �� ���� �ִٴ°��� ���� �ߴٴ� ��.
		vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_returnAMT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
	
		'Set mobjPDCOPREESTINDIRECTCOST = Nothing
		
		if mstrGBN = "������" then
			if vntData(0,1) = "" then
				'returnAMT = False
				returnAMT = split("False;" & mstrCHANGEFALG ,";")
			else
				'returnAMT = vntData(0,1) 
				returnAMT = split(vntData(0,1) & ";" & mstrCHANGEFALG ,";")
			end if
			
		elseif mstrGBN = "������" then 
		
			if vntData(1,1) = "" then
				'returnAMT = False
				returnAMT = split("False;" & mstrCHANGEFALG ,";")
			else
				'returnAMT = vntData(1,1)
				returnAMT = split(vntData(1,1) & ";" & mstrCHANGEFALG ,";")
			end if
		end if 
		window.returnvalue = returnAMT
	
	end with
	Set mobjPDCOPREESTINDIRECTCOST = Nothing
End Sub

Sub imgClose_onclick ()
	EndPage
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

Sub ImgMoveData_onclick
	Dim intCnt
	with frmThis
		If mstrGBN = "������" Then 
			gErrorMsgBox "������ ���´� ��������� �Է��Ҽ� �����ϴ�.","�Է¾ȳ�!"
			EXIT SUB
		END IF 
		.txtEXECOMMIRATE.value = .txtCOMMIRATE.value
		.txtEXEAMT.value = .txtAMT.value
		
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"EXECHK",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"CHK", intCnt)
			mobjSCGLSpr.SetTextBinding .sprSht,"EXEAMT",intCnt, mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt)
			'mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt,"1"
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
	End with
End Sub

Sub InitPage()
	'����������ü ����	
	Dim vntInParam
	Dim intNo,i
									  
	set mobjPDCOPREESTINDIRECTCOST = gCreateRemoteObject("cPDCO.ccPDCOPREESTINDIRECTCOST")
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
				case 0 : mstrPREESTNO = vntInParam(i)			'������ȣ
				case 1 : mstrSAVEGBN = vntInParam(i)			'F/T
				case 2 : mstrFIRSTPRODUCTIONCHECK = vntInParam(i) '��ư Ŭ�� �̺�Ʈ�� �߻��ϸ� Y �� �Ѿ�´� ������ �Ǹ� N ���� ���� 
				case 3 : mstrGBN = vntInParam(i)				'������ ������
				case 4 : mstrProcessData = vntInParam(i)		'�˾� ���� ���� ������  
			end select
		next
	
		'**************************************************
		'***Sum Sheet ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 12, 0
		mobjSCGLSpr.SpreadDataField .sprSht,    "CHK | EXECHK | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | AMT | EXEAMT | PRINT_SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		    "������|������|������ȣ|����|�׸�|�ڵ�|������з�|�����ߺз�|�����׸�|�ݾ�|����ݾ�|���屸��|���ļ���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "     4|     4|      10|       4|   4|  10|  10|        10|        15|      12|  12|       0|       0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | EXECHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | EXEAMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | PRINT_SEQ"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CLASSNAME | ITEMCODENAME",-1,-1,0,2,false ' ����
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | PRINT_SEQ",-1,-1,2,2,false '���
		mobjSCGLSpr.ColHidden .sprSht, "PRINT_SEQ", true
		
		IF mstrGBN = "������" THEN
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"EXECHK | EXEAMT"
		ELSE
			mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CHK | AMT"
		END IF 

		pnlTab1.style.visibility = "visible" 
		
		if .txtAMT.value <> "0" then
			mstrCheck= true
		else 
			mstrCheck= false
		end if
		
		if .txtEXEAMT.value <> "0" then
			mstrCheck= true
		else 
			mstrCheck= false
		end if
	End with

	'ȭ�� �ʱⰪ ����
	InitPageData
	'���� ������ �˾��� ���� ������ ��� �����͸� �����ͼ� ������ �������� ������� �Ѵ�.
	initpageProcess
	
	'��ȸ 
	SelectRtn
End Sub	

Sub EndPage
	gEndPage
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitpageData
	with frmThis

		.txtCOMMIRATE.value = 10
		.txtAMT.value = 0
		.txtEXECOMMIRATE.value = 10
		.txtEXEAMT.value = 0
		
		'.txtEXEAMT.value = 0
		'.txtEXECOMMIRATE.value = 10
		.txtPREESTNO.style.visibility = "hidden"
		
		If mstrSAVEGBN = "T" AND mstrGBN = "������" Then
			.txtCOMMIRATE.className = "NOINPUT_R"
			.txtCOMMIRATE.readOnly = true
			.txtAMT.className = "NOINPUT_R"
			.txtAMT.readOnly = true
			.txtEXECOMMIRATE.className = "NOINPUT_R"
			.txtEXECOMMIRATE.readOnly = true
			.txtEXEAMT.className = "NOINPUT_R"
			.txtEXEAMT.readOnly = true
			
		ElseIF  mstrSAVEGBN = "F" AND mstrGBN = "������" Then
			.txtCOMMIRATE.className = "NOINPUT_R"
			.txtCOMMIRATE.readOnly = true
			.txtAMT.className = "NOINPUT_R"
			.txtAMT.readOnly = true
			
			.txtEXECOMMIRATE.className = "INPUT_R"
			.txtEXECOMMIRATE.readOnly = false
			.txtEXEAMT.className = "INPUT_R"
			.txtEXEAMT.readOnly = false

		ELSEIF mstrSAVEGBN = "F" AND mstrGBN = "������" Then
			.txtCOMMIRATE.className = "INPUT_R"
			.txtCOMMIRATE.readOnly = false
			.txtAMT.className = "INPUT_R"
			.txtAMT.readOnly = false
			
			.txtEXECOMMIRATE.className = "NOINPUT_R"
			.txtEXECOMMIRATE.readOnly = true
			.txtEXEAMT.className = "NOINPUT_R"
			.txtEXEAMT.readOnly = true
		End If
		
	End with
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'-----------------------------------------
'���� �˾� �� ���� ���� ������ ������ ���� 
'-----------------------------------------
sub initpageProcess
	with frmthis
	'���� �ʱ�ȭ
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
		'�������̺��� �����Ͱ� ����Ǿ� �ִ��� Ȯ���Ͽ� ����Ȱ��� ���� ��쿡�� ���� ���ΰ��������� ���������� �����Ѵ�.
		vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)	
		if not gDoErrorRtn ("SelectRtn_TempCnt") then
			if mlngRowCnt > 0 Then
			'������ ���� ���� 
   			Else	
   				vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Cnt(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)	
   				if mlngRowCnt > 0 then 
   					'���� ���̺��� ���� ����Ǿ� ���� �ʴٸ� [���� �����ִ� �����!]
   				else 
   					intRtn = mobjPDCOPREESTINDIRECTCOST.ProcessRtn_Indirect(gstrConfigXml,mstrProcessData)
   				end if  
   			end If
   		end if	
	end with
end sub

'================================================================
'UI
'================================================================
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

Sub txtCOMMIRATE_onchange
	Dim intCnt
	Dim dblAMT

	with frmThis
		dblAMT = 0
		For intCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
					dblAMT = dblAMT +  (mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * .txtCOMMIRATE.value * 0.01)
			End If
		Next
		.txtAMT.value = dblAMT
	End with
End Sub

Sub txtEXEAMT_onfocus
	with frmThis
		.txtEXEAMT.value = Replace(.txtEXEAMT.value,",","")
	end with
End Sub
Sub txtEXEAMT_onblur
	with frmThis
		call gFormatNumber(.txtEXEAMT,0,true)
	end with
End Sub

Sub txtEXECOMMIRATE_onchange
	Dim intCnt
	Dim dblEXEAMT

	with frmThis
		dblEXEAMT = 0
		For intCnt = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
					dblEXEAMT = dblEXEAMT +  (mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * .txtEXECOMMIRATE.value * 0.01)
			End If
		Next
		.txtEXEAMT.value = dblEXEAMT
	End with
End Sub

'------------------------
'-----���� �ݾ� �ջ� ----
'------------------------
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
		
	End With
End Sub

'=============================================================
'Sheet Event
'=============================================================
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
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) or _
											(vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT")) Then
				
				
				
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") Then
																			
				
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


Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim dblAMT
	
	with frmThis
		If mstrGBN = "������" Then
			
			if Row = 0 and Col = 1 then
			
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
				dblAMT = 0
				.txtAMT.value=0
				for intcnt =1 to .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intcnt) = 1 THEN
						dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * 0.01)
					Else
						dblAMT = dblAMT	- (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intcnt) * 0.01)
					End If
					mobjSCGLSpr.CellChanged .sprSht, Col, intcnt
				next
				'���࿡ dblamt �� ���̳ʽ��� 0���� ....
				if dblAMT < 0 then 
					.txtAMT.value = 0
				else
					.txtAMT.value = dblAMT
				end if
				
				'�÷��� ����
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
			end if
			
		ELSE
		
			if Row = 0 and Col = 2 then
				
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 2, 2, , , "", , , , , mstrCheck
				dblAMT = 0
				.txtEXEAMT.value=0
				
				for intcnt =1 to .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"EXECHK",intcnt) = 1 THEN
						dblAMT = dblAMT	+ (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * 0.01)
					Else
						dblAMT = dblAMT	- (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intcnt) * 0.01)
					End If
					mobjSCGLSpr.CellChanged .sprSht, Col, intcnt
				next
				'���࿡ dblamt �� ���̳ʽ��� 0���� ....
				if dblAMT < 0 then 
					.txtEXEAMT.value = 0
				else
					.txtEXEAMT.value = dblAMT
				end if
				
				'�÷��� ����
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
			end if
		END IF 
		mobjSCGLSpr.CellChanged .sprSht, Col, Row
	end with	
End Sub




Sub sprSht_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		End If
	End With
End Sub

'-----------------------------------
'----------SprSht change------------
'-----------------------------------
Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim dblAMT
	Dim intCnt 
	
	with frmThis
		If mstrGBN = "������" Then
			
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then
				dblAMT = .txtAMT.value 
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 THEN
					dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) * 0.01)
				Else
					dblAMT = dblAMT	- (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row) * 0.01)
				End If
				.txtAMT.value = dblAMT
				
			ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") THEN
			
				dblAMT =0
				
				FOR intCnt=1 to .sprSht.Maxrows
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 then
						dblAMT = dblAMT + mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
					end if 
				Next
				
				dblAMT = (.txtCOMMIRATE.value * dblAMT * 0.01)
				.txtAMT.value = dblAMT
			End If
		
		ELSE
			dblAMT = .txtEXEAMT.value 
			If  Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXECHK") Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"EXECHK",Row) = 1 THEN
					dblAMT = dblAMT	+ (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) * 0.01)
				Else
					dblAMT = dblAMT	- (.txtEXECOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",Row) * 0.01)
				End If
				.txtEXEAMT.value = dblAMT
				
			ELSEIF Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EXEAMT") THEN
				dblAMT =0
				
				FOR intCnt=1 to .sprSht.Maxrows
					if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 then
						dblAMT = dblAMT + mobjSCGLSpr.GetTextBinding(.sprSht,"EXEAMT",intCnt)
					end if 
				Next
				
				dblAMT = (.txtCOMMIRATE.value * dblAMT * 0.01)
				.txtEXEAMT.value = dblAMT
			End If
		End If
		mobjSCGLSpr.CellChanged .sprSht,.sprSht.ActiveCol+1,.sprSht.ActiveRow
		txtAMT_onblur
		txtEXEAMT_onblur
	End with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
   	
End Sub



Sub RATESUM
	Dim dblAMT
	Dim intCnt
	with frmThis
		dblAMT = 0
			For intCnt = 1 To .sprSht.MaxRows
				
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 THEN
					dblAMT = dblAMT	+ (.txtCOMMIRATE.value * mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) * 0.01)
				End If
				.txtAMT.value = dblAMT
			
			Next
	End with
End Sub
'=============================================================
'��ȸ
'=============================================================
Sub SelectRtn
	IF not SelectRtn_Head () Then Exit Sub
	CALL SelectRtn_Detail ()
	RATESUM
	txtAMT_onblur
	txtEXEAMT_onblur
	mstrCHANGEFALG = "F"
End Sub

'=============================================================
'------------------����� �ؽ�Ʈ�ڽ� ��ȸ---------------------
'=============================================================
Function SelectRtn_Head()
	Dim vntData
	Dim vntData_temp
	
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	'�ӽ� ���̺��� ��ȸ�Ѵ�.
		vntData_temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempHDRCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		if mlngRowCnt > 0 then 
		
			vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempHDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		else
			'�ӽ� ���̺��� ���ٸ� ���� ����� ���̺��� ������ ��ȸ�Ѵ�.
			vntData_temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_HDRCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			if mlngRowCnt > 0 then
				'���� ����Ǿ��ִ� ���̺��� �ִٸ� ����Ǿ��ִ� ���̺��� �����´�
				vntData = mobjPDCOPREESTINDIRECTCOST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			end if 
			'������ ���� ����Ǿ��ִ� ������ ���ٸ� ����� ȭ�鿡�� ����ȭ���� ���� �������ִ´�.
		end if 
	
	IF not gDoErrorRtn ("SelectRtn_TempHDR") then
		'��ȸ�� �����͸� ���ε�
		If mlngRowCnt > 0 Then
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			
		End If
		SelectRtn_Head = True
	End IF
End Function

'=============================================================
'------------------�ϴ��� �׸��� ��ȸ---------------------
'=============================================================
Function SelectRtn_Detail()
	Dim vntData
   	Dim vntData_Temp
   	Dim vntData_TempCNT
    
	'On error resume next
	with frmThis
	
	'Long Type�� ByRef ������ �ʱ�ȭ
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

		'������ ������ ������ ���� Ȯ��
		vntData_TempCNT = mobjPDCOPREESTINDIRECTCOST.SelectRtn_TempCNT(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
		if mlngRowCnt > 0 then
			vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Temp(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			
			if not gDoErrorRtn ("SelectRtn_Temp") then
				if mlngRowCnt > 0 Then
					call mobjSCGLSpr.SetClipbinding (.sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True)
					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				Else	
   					.sprSht.MaxRows = 0
   					gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   				end If
   			end if
		else
			vntData_Temp = mobjPDCOPREESTINDIRECTCOST.SelectRtn_Detail(gstrConfigXml,mlngRowCnt,mlngColCnt,mstrPREESTNO)
			
			if not gDoErrorRtn ("SelectRtn_Detail") then
				if mlngRowCnt > 0 Then
					call mobjSCGLSpr.SetClipbinding (.sprSht, vntData_Temp, 1, 1, mlngColCnt, mlngRowCnt, True)
					gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				Else	
   					.sprSht.MaxRows = 0
   					gWriteText lblStatus, 0 & "���� �ڷᰡ �˻�" & mePROC_DONE
   				end If
   			end if		
		end if 
		window.setTimeout "AMT_SUM",1	
		txtAMT_onblur
		txtEXEAMT_onblur
   	end with
End Function

'======================================
'---------------����-------------------
'======================================
Sub processRtn
	Dim vntData
	Dim intRtn
	with frmThis
		
		'XML������ ����� �ڽ��� �����͸� �����´�.
		strMasterData = gXMLGetBindingData (xmlBind)

		'insert �÷��� ���� [��� ������ ��������]
   		mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | EXECHK | PREESTNO | SEQ | ITEMCODESEQ | ITEMCODE | DIVNAME | CLASSNAME | ITEMCODENAME | AMT | EXEAMT | PRINT_SEQ")
		
		if  not IsArray(vntData)  then
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			Exit Sub
		End If

		' �������� ���嵵 ��� input ������ȴ�.�����̺� ������ �Ǿ��ٰ� ���� ������ ���Ұ�� �ݾ��� �ٸ��κ��� ���´�.
		'DELETE INSERT 
		intRtn = mobjPDCOPREESTINDIRECTCOST.ProcessRtn(gstrConfigXml,strMasterData,vntData,mstrPREESTNO)

		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
			.sprSht.focus()
			mstrCHANGEFALG = "T"
		End If

	End with
	mstrFIRSTPRODUCTIONCHECK = "N"
End Sub

		</script>
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 10px">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<table cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
				border="0">
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td align="left">
									<TABLE cellSpacing="0" cellPadding="0" width="82" background="../../../images/back_p.gIF"
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
								<td class="TITLE">CF���������</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
			<table class="SEARCHDATA" width="100%">
				<tr>
					<td class="SEARCHLABEL" width="50">��������
					</td>
					<td class="SEARCHDATA" width="70"><INPUT dataFld="COMMIRATE" class="INPUT_R" id="txtCOMMIRATE" title="��������" style="WIDTH: 70px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtCOMMIRATE">&nbsp;%</td>
					<td class="SEARCHLABEL" width="40">������</td>
					<td class="SEARCHdata" width="112"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="������" style="WIDTH: 112px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="13" name="txtAMT"></td>
					<td class="SEARCHLABEL" width="80">���ణ������
					</td>
					<td class="SEARCHDATA" width="70"><INPUT dataFld="EXECOMMIRATE" class="INPUT_R" id="txtEXECOMMIRATE" title="��������" style="WIDTH: 70px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtEXECOMMIRATE">&nbsp;%</td>
					<td class="SEARCHLABEL" width="70">���ణ����</td>
					<td class="SEARCHdata" width="112"><INPUT dataFld="EXEAMT" class="INPUT_R" id="txtEXEAMT" title="������" style="WIDTH: 112px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="13" name="txtEXEAMT"></td>
					<td class="SEARCHLABEL" width="80">�󼼰������</td>
					<td class="SEARCHDATA"><INPUT dataFld="MEMO" class="INPUT_L" id="txtMEMO" title="�󼼰������" style="WIDTH: 200px; HEIGHT: 20px"
							dataSrc="#xmlBind" type="text" maxLength="255" size="65" name="txtMEMO"></td>
					<td class="SEARCHDATA" width="54"><INPUT dataFld="PREESTNO" class="INPUT" id="txtPREESTNO" title="������" style="WIDTH: 48px; HEIGHT: 20px"
							accessKey=",NUM" dataSrc="#xmlBind" type="text" maxLength="15" size="2" name="txtPREESTNO"></td>
					<td class="SEARCHDATA" width="54"><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="ȭ���� �ݽ��ϴ�."
							src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose">&nbsp;</td>
				</tr>
			</table>
			</TABLE><BR>
			<table cellSpacing="0" cellPadding="0" width="100%" border="0">
				<tr>
					<td class="TITLE">�� �� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSUMAMT">
					</td>
					<td class="TITLE">�����հ� : <INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="�հ�ݾ�" style="HEIGHT: 22px" accessKey="NUM"
							readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
					</td>
					<TD align="right" width="600"><IMG id="ImgMoveData" onmouseover="JavaScript:this.src='../../../images/ImgMoveDataOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgMoveData.gIF'" height="20" alt="���������� �� ����������� �����մϴ�."
							src="../../../images/ImgMoveData.gIF" align="absMiddle" border="0" name="ImgMoveData">&nbsp;<IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
							onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" align="absMiddle" border="0" name="imgSave">&nbsp;<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" align="absMiddle" border="0" name="imgExcel">&nbsp;
					</TD>
				</tr>
			</table>
			<table height="500" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR vAlign="top" align="left">
					<!--����-->
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
						<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
							ms_positioning="GridLayout">
							<OBJECT id=sprSht classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 width="100%" height="100%">
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="_ExtentX" VALUE="31750">
	<PARAM NAME="_ExtentY" VALUE="21060">
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
					<TD class="BOTTOMSPLIT" id="lbltext" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
				</TR>
			</table>
		</form>
	</body>
</HTML>
