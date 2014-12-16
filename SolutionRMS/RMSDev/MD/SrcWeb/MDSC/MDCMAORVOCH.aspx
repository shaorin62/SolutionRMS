<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMAORVOCH.aspx.vb" Inherits="MD.MDCMAORVOCH" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� ��ǥ����</title>
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
'HISTORY    :1) 2012.06.07 OH SE HOON
'			(������ AOR ����� ��û���� ���ϰ� ��ǥ ó�� ��û���� ���Ŀ� ������ �ʿ� ��...)
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
'�������� ����
Dim mlngRowCnt,mlngColCnt
Dim mobjMDSCAORVOCH
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrGUBUN
Dim mstrPROCESS
Dim vntData_ProcesssRtn

mstrGUBUN = "M"
mstrPROCESS = ""
mstrCheck=True

CONST meTAB = 9
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

'�������� ��ư �����
Sub Set_delete(byVal strmode)
	With frmThis
		IF .rdT.checked = TRUE then 
			document.getElementById("imgVochDelco").style.DISPLAY = "BLOCK"
		else
			document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
		end if
	End With
End Sub

'��ȸ��ư
Sub imgQuery_onclick
	If frmThis.txtYEARMON.value = "" Then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	End If

	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'����
Sub btnTab1_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TABON
	'frmThis.btnTab2.style.backgrounDimage = meURL_TAB
	
	pnlTab_susu.style.visibility = "visible" 
	'pnlTab_gen.style.visibility = "hidden" 

	gFlowWait meWAIT_ON
	mstrGUBUN = "M"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
	mobjSCGLCtl.DoEventQueue
End Sub

'����
Sub btnTab2_onclick
	frmThis.btnTab1.style.backgrounDimage = meURL_TAB
	frmThis.btnTab2.style.backgrounDimage = meURL_TABON
	
	pnlTab_susu.style.visibility = "hidden" 
	pnlTab_gen.style.visibility = "visible"
	
	gFlowWait meWAIT_ON
	mstrGUBUN = "B"
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF

	mobjSCGLCtl.DoEventQueue
End Sub

Sub imgExcel_onclick()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		
		if mstrGUBUN = "M"  then 
			mobjSCGLSpr.ExportExcelFile .sprSht_SUSU
		elseif mstrGUBUN = "B"  then  
			mobjSCGLSpr.ExportExcelFile .sprSht_OUT
		end if
		
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'��ǥ���� Ŭ��
Sub ImgvochCre_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Create"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� Ŭ��
Sub imgVochDel_onclick ()
	gFlowWait meWAIT_ON
	mstrPROCESS = "Delete"
	ProcessRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'��ǥ���� ���� Ŭ��
Sub imgVochDelco_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn(mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'������ǥ����Ŭ��
Sub ImgErrVochDel_onclick()
	gFlowWait meWAIT_ON
	ErrVochDeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'���� ����
Sub ImgSUMMApp_onclick()
	Dim intCnt,intCnt2, intCnt3
	Dim intSumCnt
	Dim intRtn
	with frmThis
		intSumCnt = 0
		
		if mstrGUBUN = "M"  then 
			For intCnt = 1 To .sprSht_SUSU.MaxRows
				If mobjSCGLSpr.GetTextBinding( .sprSht_SUSU,"CHK",intCnt) = 1 Then 
					intSumCnt = intSumCnt +1
				end If
			Next
			
			If intSumCnt = 0  Then 	
				Exit Sub		
			Elseif Trim(.txtSUMM.value) <> "" Then 
				intRtn = gYesNoMsgbox("���並 �����Ͻðڽ��ϱ�?","����ȳ�!")
				
				IF intRtn <> vbYes then exit Sub
				
				For intCnt3 = 1 To .sprSht_SUSUDTL.MaxRows
					mobjSCGLSpr.SetTextBinding .sprSht_SUSUDTL,"SUMM",intCnt3, .txtSUMM.value 
				Next
			End If
		elseif mstrGUBUN = "B"  then  
			For intCnt = 1 To .sprSht_OUT.MaxRows
				If mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"CHK",intCnt) = 1 Then 
					intSumCnt = intSumCnt +1
				end If
			Next
			
			If intSumCnt = 0  Then 	
				Exit Sub		
			Elseif Trim(.txtSUMM.value) <> "" Then 
				intRtn = gYesNoMsgbox("���並 �����Ͻðڽ��ϱ�?","����ȳ�!")
				
				IF intRtn <> vbYes then exit Sub
				
				For intCnt2 = 1 To .sprSht_OUT.MaxRows
					If mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"CHK",intCnt2) = 1 Then 
						mobjSCGLSpr.SetTextBinding .sprSht_OUT,"SUMM",intCnt2, .txtSUMM.value 
					End If
				Next
			End If
		end if
	End With
End Sub

'-----------------------------------------------------------------------------------------
' �������˾�(��ȸ)
'-----------------------------------------------------------------------------------------
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,425)	
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value),"A")

			if not gDoErrorRtn ("txtCLIENTNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'�̿Ϸ�üũ
Sub rdF_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub
'����üũ
Sub rdE_onclick
	gFlowWait meWAIT_ON
	CALL SelectRtn (mstrGUBUN)
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------------------------------------------------------------
' Field Event
'-----------------------------------------------------------------------------------------
Sub txtSUMM_onchange
	Dim blnByteCHk
	Dim intRtn
	blnByteCHk =  checkBytes(frmThis.txtSUMM.value)
	
	If blnByteCHk  > 23 Then
		intRtn = gYesNoMsgbox("������ ũ��� 23Byte �� ������ �����ϴ�. �ʱ�ȭ �Ͻðڽ��ϱ�?","ó���ȳ�!")
		IF intRtn <> vbYes then exit Sub
		
		frmThis.txtSUMM.value = ""
	End If
End Sub

function checkBytes(expression)
	Dim VLength
	Dim temp
	Dim EscTemp
	Dim i
	
	VLength=0
	temp = expression
	
	if temp <> "" then
		for i=1 to len(temp) 
			if mid(temp,i,1) <> escape(mid(temp,i,1))  then
				EscTemp=escape(mid(temp,i,1))
				if (len(EscTemp)>=6) then
					VLength = VLength +2
				else
				VLength = VLength +1
				end if
			else
				VLength = VLength +1
			end if
		Next
	end if
	checkBytes = VLength
end function

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_SUSU_Change(ByVal Col, ByVal Row)
	mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, Col, Row
End Sub

Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
	with frmThis
		if	Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"paycode") then
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "G" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404150"
			ELSEif mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE", Row) = "T" THEN
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404100"
			else
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "DEBTOR", Row, Row, 255
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEBTOR",Row, "404103"
			end if 
		end if 

		If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",Row) = "Y" Then
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",Row,Row,false
		Else
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",Row,Row,false
			mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",Row,Row,false
		End If
	End With
	
	mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, Col, Row
End Sub

'-----------------------------------
' SpreadSheet Ŭ��
'-----------------------------------
Sub sprSht_SUSU_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			.sprSht_SUSUDTL.MaxRows = 0
			for intCnt = 1 To .sprSht_SUSU.MaxRows
				mobjSCGLSpr.SetCellTypeCheckBox .sprSht_SUSU, 1, 1, intCnt, intCnt, "", , , , , mstrCheck
			Next    

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_SUSU.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_SUSU, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intCnt, i
	Dim lngSUMAMT,lngAMT,lngTOT

	With frmThis
		if Col = 1 and Row = 0 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck

			if mstrCheck = True then  
				for intCnt = 1 To .sprSht_OUT.MaxRows
					mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
				Next    
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if 
	End With
End Sub 

Sub sprSht_SUSU_ButtonClicked (Col,Row,ButtonDown)
	if Col = 1 and Row > 0 then 
		if mobjSCGLSpr.GetTextBinding( frmThis.sprSht_SUSU,"CHK",Row) = 1 THEN
			SelectRtn_SUSUDTL Col,Row
		ELSE
			call DeleteRtn_SUSUDTL(Row)
		END IF
	end if
End Sub

Sub DeleteRtn_SUSUDTL (Row)
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strSEQ	

	With frmThis
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_SUSUDTL.MaxRows to 1 step -1
			strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
			strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)

			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",i) = strTAXYEARMON and _
			   mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",i) = strTAXNO then	
			   mobjSCGLSpr.DeleteRow .sprSht_SUSUDTL,i
			end if				
		next
	End With
	err.clear	
End Sub

'-----------------------------------
' SpreadSheet ���� Ŭ��
'-----------------------------------
sub sprSht_SUSU_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_SUSU, ""
		end if
	end with
end sub

sub sprSht_OUT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_OUT, ""
		end if
	end with
end sub

'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()
	'����������ü ����	
	Set mobjMDSCAORVOCH	 = gCreateRemoteObject("cMDSC.ccMDSCAORVOCH")
	Set mobjSCCOGET		 = gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	'�� ��ġ ���� �� �ʱ�ȭ
	mobjSCGLCtl.DoEventQueue
    Dim strComboPREPAYMENT
	Dim strSemuComboListB, strSemuComboListA

	gSetSheetDefaultColor
    with frmThis
		strComboPREPAYMENT =  "Y" & vbTab & " "
		strSemuComboListB =  "B5" & vbTab & "BR"
		strSemuComboListA =  "A0" & vbTab & "AI" & vbTab & "A8" & vbTab & "AZ"

		'**************************************************
		'���� ��Ʈ ������ hdr
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSU
		mobjSCGLSpr.SpreadLayout .sprSht_SUSU, 12, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSU,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | TAXYEARMON | TAXNO | VOCHNO | RMSNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.SetHeader .sprSht_SUSU,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|�ݾ�|�ΰ���|RMS���|RMS��ȣ|��ǥ��ȣ|RMSNO|�����ڵ�|�����޼���"
		mobjSCGLSpr.SetColWidth .sprSht_SUSU, "-1",  "  4|       8|        10|    20|  11|    11|      7|      7|      10|    0|       0|        15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSU, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_SUSU, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSU, "POSTINGDATE "
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CUSTOMERCODE | TAXYEARMON | TAXNO | VOCHNO | RMSNO",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSU, "CUSTNAME | ERRMSG",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSU, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSU,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | VAT | TAXYEARMON | TAXNO | VOCHNO | RMSNO | ERRCODE | ERRMSG"
		mobjSCGLSpr.ColHidden .sprSht_SUSU, "ERRCODE", true

		'**************************************************
		'���� ��Ʈ ������ dtl
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_SUSUDTL
		mobjSCGLSpr.SpreadLayout .sprSht_SUSUDTL, 33, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_SUSUDTL,    "MEDFLAGNAME | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | DEPT_NAME | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | VENDOR | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | TAXNOSEQ | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG | RMSNO"
		mobjSCGLSpr.SetHeader .sprSht_SUSUDTL,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����|�������|�ڽ�Ʈ����|���μ�|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������|VENDOR|����|��������|����|������|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|RMS�ι�ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|AMTGBN|MEDFLAG|RMSNO"
		mobjSCGLSpr.SetColWidth .sprSht_SUSUDTL, "-1",  "   6|       8|        10|    15|  20|       5|         8|      10|  10|    10|       7| 5|       8|    10|     0|   0|       7|   7|     8|        10|            13|            13|      20|      7|      7|        7|       9|       0|        10|    0|     0|      0|    0"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_SUSUDTL, "-1", "13"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_SUSUDTL, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "MEDFLAGNAME | CUSTOMERCODE | CUSTNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | CUSTOMERCODE|DEBTOR | ACCOUNT ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "CUSTNAME | SUMM | ERRMSG",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_SUSUDTL, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_SUSUDTL, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_SUSUDTL,mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_SUSUDTL,"SEMU"),-1,-1,strSemuComboListB,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_SUSUDTL, "CUSTOMERCODE | BA | SEMU | BP | DEBTOR | ACCOUNT | TAXYEARMON | TAXNO | GBN | VOCHNO",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellsLock2 .sprSht_SUSUDTL,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME | AMT | BP | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | MEDFLAG"
		mobjSCGLSpr.ColHidden .sprSht_SUSUDTL, "GBN | GFLAG | DEMANDDAY | AMTGBN | TAXNOSEQ | MEDFLAG | RMSNO", true 
		
		'**************************************************
		'���� ��Ʈ ������
		'**************************************************	
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 31, 0, 4
		mobjSCGLSpr.SpreadDataField .sprSht_OUT,    "CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PAYCODE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		    "����|��ǥ����|�ŷ�ó�ڵ�|�ŷ�ó|����ó|����|�������|�ڽ�Ʈ����|�ݾ�|�ΰ���|�����ڵ�|BP|���ޱ���|������|����|��������|����|������|���޹��|�����ݱ���|������(������)|������(������)|����TEXT|RMS���|RMS��ȣ|��ǥ��ȣ|�����ڵ�|�����޼���|GFLAG|���ⱸ��| AMTGBN"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1", "   4|       8|        10|    15|    15|  20|       5|         8|  10|    10|       7| 5|       8|    10|   0|       7|   7|     8|      20|        10|            13|            13|      20|      7|      7|       9|       0|        10|    0|      10|      0"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_OUT, "POSTINGDATE | DEMANDDAY | DOCUMENTDATE | FROMDATE | TODATE | DUEDATE"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "CUSTOMERCODE | CUSTNAME | VENDORNAME | BA | COSTCENTER | BP | GBN | ACCOUNT | DEBTOR | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN", -1, -1, 200
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMMTEXT", -1, -1, 50
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "SUMM", -1, -1, 25
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, "PAYCODE", -1, -1, 255
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "AMT | VAT", -1, -1, 0 '������
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"PREPAYMENT"),-1,-1,strComboPREPAYMENT,,80
		mobjSCGLSpr.SetCellTypeComboBox .sprSht_OUT,mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"SEMU"),-1,-1,strSemuComboListA,,50
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTOMERCODE | BA | SEMU | BP | TAXYEARMON | TAXNO | GBN | VOCHNO | DEBTOR | ACCOUNT ",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CUSTNAME | SUMM | ERRMSG | VENDORNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"POSTINGDATE | CUSTOMERCODE | CUSTNAME  | AMT | BP | GBN | DOCUMENTDATE | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | JOBBASE"
		mobjSCGLSpr.ColHidden .sprSht_OUT, "GBN | GFLAG | JOBBASE | DUEDATE | AMTGBN", true 

	End with
	pnlTab_susu.style.visibility = "visible" 
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate2,1,4) & Mid(gNowDate2,6,2)
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.txtYEARMON.focus	
		
		Get_COMBO_VALUE	
		'ó���� ���� ���� ����
		document.getElementById("imgVochDelco").style.DISPLAY = "NONE"
	End with
End Sub

Sub EndPage()
	set mobjMDSCAORVOCH = Nothing
	Set mobjSCCOGET = Nothing
	gEndPage	
End Sub

Sub Get_COMBO_VALUE ()		
	Dim vntData
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht_SUSU.MaxRows = 0
		.sprSht_OUT.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		vntData = mobjMDSCAORVOCH.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_PAYCODE")
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_OUT, "PAYCODE",,,vntData,,160
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

Sub SelectRtn (strVOCH_TYPE)	
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		.sprSht_SUSUDTL.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		
		IF strVOCH_TYPE = "M" THEN
			CALL SelectRtn_SUSU()
		ELSEIF strVOCH_TYPE = "B" THEN
			CALL SelectRtn_OUT()
		END IF
   	end with
End Sub

Sub SelectRtn_SUSU ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
	
	with frmThis
		.sprSht_SUSU.MaxRows = 0
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF 

		vntData = mobjMDSCAORVOCH.SelectRtn_SUSU(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
													strCLIENTCODE, strCLIENTNAME,  _
													strGBN)

		if not gDoErrorRtn ("SelectRtn_SUSU") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "M"
				mobjSCGLSpr.SetClipbinding .sprSht_SUSU, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			ELSE
				.txtSELECTAMT.value = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub SelectRtn_SUSUDTL (Col, Row)
	Dim vntData
   	Dim i, strCols
    Dim intCnt
    Dim strTAXYEARMON
    Dim strTAXNO
    Dim strRow
    
	with frmThis
		'Sheet�ʱ�ȭ
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",Row)
		strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",Row)
		
		if .rdF.checked then
			vntData = mobjMDSCAORVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		else
			vntData = mobjMDSCAORVOCH.SelectRtn_SUSUDTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strTAXYEARMON, strTAXNO)
		end if

		If not gDoErrorRtn ("SelectRtn_SUSUDTL") Then
			If mlngRowCnt >0 Then
				strRow = 0
				strRow = .sprSht_SUSUDTL.MaxRows + 1
				Call mobjSCGLSpr.SetClipBinding (.sprSht_SUSUDTL,vntData, 1, strRow, mlngColCnt, mlngRowCnt,True)
   			End If
   		End If
   	end with
End Sub

Sub SelectRtn_OUT ()
   	Dim vntData
    Dim intCnt
    Dim strYEARMON, strCLIENTCODE, strCLIENTNAME, strGBN
    Dim strOUTSCODE, strOUTSNAME
    Dim strPROGBN
	
	with frmThis
		.sprSht_OUT.MaxRows = 0

		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE.value
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		IF .rdT.checked THEN
			strGBN = .rdT.value
		ELSEIF .rdF.checked THEN
			strGBN = .rdF.value
		ELSEIF .rdE.checked THEN
			strGBN = .rdE.value
		END IF

		vntData = mobjMDSCAORVOCH.SelectRtn_OUT(gstrConfigXml, mlngRowCnt, mlngColCnt, strYEARMON, _
												  strCLIENTCODE, strCLIENTNAME, strOUTSCODE, strOUTSNAME,  _
												  strGBN)
		if not gDoErrorRtn ("SelectRtn_OUT") then
			if mlngRowCnt > 0 Then
				mstrGUBUN = "B"
				
				mobjSCGLSpr.SetClipbinding .sprSht_OUT, vntData, 1, 1, mlngColCnt, mlngRowCnt, True

				For intCnt = 1 To .sprSht_OUT.MaxRows
					If  .rdT.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,true,"DUEDATE",intCnt,intCnt,false
					elseif .rdF.checked or .rdE.checked then
						mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1,1,intCnt,intCnt,,0,1,2,2,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DEMANDDAY",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"DUEDATE",intCnt,intCnt,false
					End If

					'������ ó����
					If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) = "Y" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,false,"TODATE",intCnt,intCnt,false
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"FROMDATE",intCnt,intCnt,false
						mobjSCGLSpr.SetCellsLock2 .sprSht_OUT,True,"TODATE",intCnt,intCnt,false
					End If	
				Next
				Call AMT_SUM (.sprSht_OUT)
			ELSE
				.txtSELECTAMT.value = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   	end with
End Sub

Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM
	With frmThis
		IntAMTSUM = 0

		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

Function DataValidation_SUSU ()
	DataValidation_SUSU = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	
	With frmThis
		For intCnt =1  To .sprSht_SUSUDTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_SUSU = True
End Function

Function DataValidation_GEN ()
	DataValidation_GEN = false	
	Dim intCnt, intCnt2
	Dim chkcnt
	
	intCnt = 0
	With frmThis
		For intCnt =1  To .sprSht_OUT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",intCnt) = "1" AND mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"duedate",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ������û���� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
		Next
	End With
	DataValidation_GEN = True
End Function

'�������
Sub ProcessRtn(strVOCH_TYPE)
	Dim intRtn

	with frmThis
		IF mstrPROCESS = "Create" THEN
			IF NOT .rdF.checked THEN
				gErrorMsgBox "�̿Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if

		IF mstrPROCESS = "Delete" THEN
			IF NOT .rdT.checked THEN
				gErrorMsgBox "�Ϸ���ȸ�� �����մϴ�.","�����׻���"
				exit sub
			end IF 
		end if 
		
		IF strVOCH_TYPE = "M" THEN
			if DataValidation_SUSU =false then exit sub
			CALL ProcessRtn_SUSU()
		ELSEIF strVOCH_TYPE = "B" THEN
			if DataValidation_GEN =false then exit sub
			CALL ProcessRtn_OUT()
		END IF
   	end with
End Sub

Sub ProcessRtn_SUSU()
	Dim intRtn
	Dim strTAXYEARMON
	Dim strTAXNO
	
	with frmThis
		mobjSCGLSpr.SetFlag frmThis.sprSht_SUSUDTL, meINS_FLAG

		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_SUSUDTL,"MEDFLAGNAME | POSTINGDATE | CUSTOMERCODE | CUSTNAME | SUMM | BA | COSTCENTER | DEPT_NAME | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | VENDOR | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | TAXNOSEQ | VOCHNO | ERRCODE | ERRMSG | GFLAG | AMTGBN | MEDFLAG | RMSNO")

		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		strTAXYEARMON = "" : strTAXNO = ""
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN
		IF_GUBUN = "RMS_0002"'����

		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1
		
				strRMS_DOC_TYPE = "M"

				if strIF_CNT = "1" then
					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt)
				Else

					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if
				
					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt)
				end if

				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_SUSUDTL.MaxRows
				strIF_CNT = strIF_CNT + 1

				strRMS_DOC_TYPE = "Z"
				
				if strIF_CNT = "1" then
					strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt)
				Else
					if strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) and _
						strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) THEN
						
						strHSEQ = strHSEQ
						strISEQ = strISEQ+1
					else 
						strHSEQ = strHSEQ + 1
						strISEQ = 1
					end if

					strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"AMTGBN",intCnt) + "|" + _
									"" + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"VOCHNO",intCnt)
				end if
				
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXYEARMON",intCnt)
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSUDTL,"TAXNO",intCnt)
			Next
		end if 

		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'�������
Sub ProcessRtn_OUT()
	Dim intRtn
	
	with frmThis
		vntData_ProcesssRtn = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | POSTINGDATE | CUSTOMERCODE | CUSTNAME | VENDORNAME | SUMM | BA | COSTCENTER | AMT | VAT | SEMU | BP | DEMANDDAY | DUEDATE  | GBN | ACCOUNT | DEBTOR | DOCUMENTDATE | PAYCODE | PREPAYMENT | FROMDATE | TODATE | SUMMTEXT | TAXYEARMON | TAXNO | VOCHNO | ERRCODE | ERRMSG | GFLAG | JOBBASE | AMTGBN")
		'ó�� ������ü ȣ��
		if  not IsArray(vntData_ProcesssRtn) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1
		Dim strISEQ : strISEQ = 1
		Dim strRMS_DOC_TYPE : strRMS_DOC_TYPE = "Z" '�ӽ���ǥ ���� �÷���
		
		intCol = ubound(vntData_ProcesssRtn, 1)
		intRow = ubound(vntData_ProcesssRtn, 2)
		
		Dim IF_GUBUN

'		if .rdPRO.checked then
'			IF_GUBUN = "RMS_0009" '���θ�� ����
'		else
'			IF_GUBUN = "RMS_0004"'�뿪����
'		'ELSE
'		'	IF_GUBUN = "RMS_0005" '�������
'		end if
		
		if mstrPROCESS = "Create" then
			For intCnt = 1 To .sprSht_OUT.MaxRows
				if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "O"

					if strIF_CNT = "1" then
						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					else
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					end if

					strHSEQ = strHSEQ+1
				end if 
			Next
		elseif mstrPROCESS = "Delete" then
			For intCnt = 1 To .sprSht_OUT.MaxRows
				if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"chk",intCnt) = "1" then		
					strIF_CNT = strIF_CNT + 1
			
					strRMS_DOC_TYPE = "Z"
		
					if strIF_CNT = "1" then
						strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					else
						strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
									cstr(strISEQ) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"POSTINGDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMM",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BA",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"COSTCENTER",intCnt) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMT",intCnt)) + "|" + _
									cstr(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VAT",intCnt)) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEMU",intCnt) + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"BP",intCnt) + "|" + _ 
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEMANDDAY",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CUSTOMERCODE",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"GFLAG",intCnt) + "|" + _
									strRMS_DOC_TYPE + "|" + _ 
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"ACCOUNT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEBTOR",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DOCUMENTDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PREPAYMENT",intCnt) + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"FROMDATE",intCnt),"-","") + "|" + _
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TODATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SUMMTEXT",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"AMTGBN",intCnt) + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"PAYCODE",intCnt) + "|" + _  
									replace(mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DUEDATE",intCnt),"-","") + "|" + _
									mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",intCnt)
					end if
					strHSEQ = strHSEQ+1
				end if 
			Next
		end if 
		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
   	end with
End Sub

'---------------------------------------------------
' ��ǥ���� �� ��ǥ��ȣ �޾ƿ��� �� ���� RMS������Ʈ
'---------------------------------------------------
Sub Set_VochValue (strRETURNLIST)
	gFlowWait meWAIT_ON
	Dim strDOC_STATUS
	Dim strDOC_MESSAGE
	Dim strVOCHNO

	With frmThis
		if mstrPROCESS ="Create" then
			if mstrGUBUN = "M" then
				intRtn = mobjMDSCAORVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN)
			elseif mstrGUBUN = "B" then
				intRtn = mobjMDSCAORVOCH.ProcessRtn(gstrConfigXml,vntData_ProcesssRtn, strRETURNLIST, mstrGUBUN)
			end if 

			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "M" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU, meCLS_FLAG
				ELSEIF mstrGUBUN = "B" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT, meCLS_FLAG
				END IF

				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				else
					gErrorMsgBox "������ �߻��߽��ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			end if

   		elseif mstrPROCESS ="Delete" then
   			intRtn = mobjMDSCAORVOCH.VOCHDELL(gstrConfigXml, strRETURNLIST, mstrGUBUN)
   			
   			if not gDoErrorRtn ("VOCHDELL") then
				'��� �÷��� Ŭ����
				IF mstrGUBUN = "M" THEN
					mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
				ELSEIF mstrGUBUN = "B" THEN
					mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
				END IF

				if intRtn > 0 Then
					gErrorMsgBox "��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
				End If
				SelectRtn(mstrGUBUN)
   			end if
   		end if 
   		IF mstrGUBUN = "M" THEN
			.sprSht_SUSU.focus()
		ELSEIF mstrGUBUN = "B" THEN
			.sprSht_OUT.focus()
		END IF
	End With
	gFlowWait meWAIT_OFF
End Sub

sub ErrVochDeleteRtn
	Dim intRtn
   	Dim vntData
	with frmThis
   		
   		IF NOT .rdE.checked THEN
			gErrorMsgBox "������ȸ�� �����մϴ�.","�����׻���"
			exit sub
		end if 
		
		IF mstrGUBUN = "M" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_SUSU,"CHK | TAXYEARMON | TAXNO | RMSNO | ERRCODE")
		ELSEIF mstrGUBUN = "B" THEN
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | TAXYEARMON | TAXNO | RMSNO | ERRCODE")
		END IF

		'ó�� ������ü ȣ��
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit sub
		End If
		
		intRtn = mobjMDSCAORVOCH.DeleteRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("DeleteRtn") then
			'��� �÷��� Ŭ����
			IF mstrGUBUN = "M" THEN
				mobjSCGLSpr.SetFlag  .sprSht_SUSU,meCLS_FLAG
			ELSEIF mstrGUBUN = "B" THEN
				mobjSCGLSpr.SetFlag  .sprSht_OUT,meCLS_FLAG
			END IF
			
			if intRtn > 0 Then
			gErrorMsgBox "���� ��ǥ�� �����Ǿ����ϴ�.","����ȳ�"
			End If
			
			SelectRtn(mstrGUBUN)
   		end if
   	end with
End Sub

'-----------------------------------------
'��ǥ ���� ����
'-----------------------------------------
Sub DeleteRtn (strGUBUN)
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strTAXYEARMON, strTAXNO
	Dim strVOCHNO
	Dim lngchkCnt
		
	lngchkCnt = 0
	With frmThis
		If mstrGUBUN = "M"  then  
			If .sprSht_SUSU.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_SUSU.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		ELSEIf mstrGUBUN = "B"  then
			If .sprSht_OUT.MaxRows = 0 then
				gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
				Exit Sub
			End If
			
			For i = 1 To .sprSht_OUT.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
					lngchkCnt = lngchkCnt + 1
				END IF
			next
			if lngchkCnt = 0 then
				gErrorMsgBox "�����Ͻ� �ڷᰡ �����ϴ�.","�����ȳ�!"
				exit sub
			end if
		END IF
	
		intRtn = gYesNoMsgbox("���������� SAP���� ���ε� ��ǥ�� SAP���� ����Ͽ� RMS�ʿ��� ������ �� ������ RMS�� ��ǥ�� ������ �����Ҷ� ����մϴ�. " & vbCrlf & "  " & vbCrlf & " ��ǥ�� ������ �����Ͻðڽ��ϱ�?","�������� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		
		intCnt = 0
		'���õ� �ڷḦ ������ ���� ����
		If mstrGUBUN = "M"  then  
			for i = .sprSht_SUSU.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_SUSU,"VOCHNO",i)
					
					intRtn = mobjMDSCAORVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN)

					If not gDoErrorRtn ("DeleteRtn_GANG") Then
						mobjSCGLSpr.DeleteRow .sprSht_SUSU,i
   					End If
		   				
   					intCnt = intCnt + 1
   				End If
			Next
		ELSEIf mstrGUBUN = "B"  then
			for i = .sprSht_OUT.MaxRows to 1 step -1
				If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then
					strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXYEARMON",i)
					strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"TAXNO",i)
					strVOCHNO = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"VOCHNO",i)
					
					intRtn = mobjMDSCAORVOCH.DeleteRtn_GANG(gstrConfigXml,strTAXYEARMON, strTAXNO, strVOCHNO, mstrGUBUN)

					If not gDoErrorRtn ("DeleteRtn_GANG") Then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End If

   					intCnt = intCnt + 1
   				End If
			Next
		END IF
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
			SelectRtn (strGUBUN)
	End With
	err.clear	
End Sub

		</script>
		<script language="javascript">
		//##########################################################################################################################################
		//******************************************��1) frmSapCon ���� ������ �� �̿��Ͽ� Submit �ϴ� �Լ�
		//##########################################################################################################################################

		function Set_WebServer(strIF_CNT, strIF_GUBUN, strIF_USER, strITEMLIST) {		
			//���
			frmSapCon.document.getElementById("txtcnt").value = strIF_CNT;
			frmSapCon.document.getElementById("txtIF_GUBUN").value = strIF_GUBUN;
			frmSapCon.document.getElementById("txtIF_USER").value = strIF_USER;
			//dtl
			frmSapCon.document.getElementById("txtITEMLIST").value = strITEMLIST;
			window.frames[0].document.forms[0].submit();
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TR>
					<TD style="HEIGHT: 54px">
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<TD align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left">
												<TABLE cellSpacing="0" cellPadding="0" width="145" background="../../../images/back_p.gIF"
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
											<td class="TITLE">AOR ������� ��ǥó��</td>
										</tr>
									</table>
								</TD>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE id="tblWaitP" style="Z-INDEX: 101; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
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
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" style="WIDTH: 100%" height="93%" cellSpacing="0" cellPadding="0" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%" colSpan="2"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 15px" vAlign="top" align="center" colSpan="2">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
												width="70">&nbsp;û����
											</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtYEARMON" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													maxLength="8" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="75">&nbsp;û����
											</TD>
											<TD class="SEARCHDATA" style="WIDTH: 219px" width="219"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 142px; HEIGHT: 22px"
													maxLength="100" size="29" name="txtCLIENTNAME"> <IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgCLIENTCODE"> <INPUT class="INPUT_L" id="txtCLIENTCODE" title="�������ڵ�" style="WIDTH: 53px; HEIGHT: 22px"
													accessKey=",M" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<td class="SEARCHDATA">
												<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
													<TR>
														<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
													</TR>
												</TABLE>
											</td>
										</TR>
										<TR>
											<TD class="SEARCHLABEL">����
											</TD>
											<TD class="SEARCHDATA" colspan="4">
												<INPUT id="rdT" title="�Ϸ᳻����ȸ" type="radio" value="rdT" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;�Ϸ�&nbsp;
												<INPUT id="rdF" title="�̿Ϸ� ������ȸ" type="radio" value="rdF" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')"
													CHECKED>&nbsp;�̿Ϸ�&nbsp; <INPUT id="rdE" title="������ǥ ������ȸ" type="radio" value="rdE" name="rdGBN" onclick="vbscript:Call Set_delete('imgVochDelco')">&nbsp;����&nbsp;
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 15px"></TD>
							</TR>
							<TR>
								<TD vAlign="middle" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTABON" id="btnTab1" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
													type="button" value="����" name="btnTab1"> <!--INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTab.gIF)"
													type="button" size="20" value="����" name="btnTab2"-->
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" width="50" border="0">
													<TR>
														<td><IMG id="ImgvochCre" onmouseover="JavaScript:this.src='../../../images/ImgvochCreOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgvochCre.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/ImgvochCre.gIF" border="0" name="ImgvochCre"></td>
														<td><IMG id="imgVochDel" onmouseover="JavaScript:this.src='../../../images/imgVochDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDel.gIF'"
																height="20" alt="��ǥ�� �����մϴ�." src="../../../images/imgVochDel.gIF" border="0" name="imgVochDel"></td>
														<td><IMG id="ImgErrVochDel" onmouseover="JavaScript:this.src='../../../images/ImgErrVochDelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgErrVochDel.gIF'"
																height="20" alt="������ǥ �� �����մϴ�." src="../../../images/ImgErrVochDel.gIF" border="0"
																name="ImgErrVochDel"></td>
														<td><IMG id="imgVochDelco" onmouseover="JavaScript:this.src='../../../images/imgVochDelcoOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgVochDelco.gIF'"
																height="20" alt="��ǥ�� ������ �����մϴ�." src="../../../images/imgVochDelco.gIF" border="0"
																name="imgVochDelco" title="SAP���� ���������Ͽ� RMS���� ������ �� ������ RMS��ǥ�� ������ �����Ѵ�."></td>
														<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE class="SEARCHDATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 75px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSUMM,'')">��������
											</TD>
											<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtSUMM" title="��������" style="WIDTH: 402px; HEIGHT: 21px" size="61"
													name="txtSUMM"><IMG id="ImgSUMMApp" onmouseover="JavaScript:this.src='../../../images/ImgAppOn.gIF'"
													title="���並 �ϰ� �����մϴ�" style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgApp.gIF'"
													height="20" alt="���並 �ϰ� �����մϴ�" src="../../../images/ImgApp.gif" width="54" align="absMiddle" border="0"
													name="ImgSUMMApp">
											</TD>
											<TD align="right"><INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
													accessKey="NUM" readOnly maxLength="100" size="13" name="txtSUMAMT"><INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"></TD>
							</TR>
							<!--���� �� �׸���-->
							<TR vAlign="top" align="left">
								<!--����-->
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab_susu" style="POSITION: absolute; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden; LEFT: 7px"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 70%" id="sprSht_SUSU" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31855">
											<PARAM NAME="_ExtentY" VALUE="9022">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
										<OBJECT style="WIDTH: 100%; HEIGHT: 30%" id="sprSht_SUSUDTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31855">
											<PARAM NAME="_ExtentY" VALUE="3862">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
									<DIV id="pnlTab_gen" style="POSITION: absolute; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden; LEFT: 7px"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_OUT" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31855">
											<PARAM NAME="_ExtentY" VALUE="12885">
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
											<PARAM NAME="MaxCols" VALUE="11">
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
						</TABLE>
					</TD>
				</TR>
				<!--List End-->
				<!--Bottom Split Start-->
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus" style="WIDTH: 100%"></TD>
				</TR>
			</TABLE>
		</FORM>
		</TR></TABLE><iframe id="frmSapCon" style="WIDTH: 100%; DISPLAY: none; HEIGHT: 300px" name="frmSapCon"
			src="../../../MD/WebService/TRUVOCHWEBSERVICE.aspx"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
