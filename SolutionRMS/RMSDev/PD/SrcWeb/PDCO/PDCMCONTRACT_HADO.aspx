<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT_HADO.aspx.vb" Inherits="PD.PDCMCONTRACT_HADO" %>
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
'HISTORY    :1) 2009/11/21 By Ȳ����
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
Dim mobjPDCMCONTRACT_HADO, mobjPDCMGET
Dim mstrCheck
Dim mstrmode
Dim mstrCHKcheck

CONST meTAB = 9

mstrCheck = True
mstrmode = True
mstrCHKcheck = True
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If  mstrmode Then
			document.getElementById("tblBody3").style.display = "inline"
			document.getElementById("tblBody4").style.display = "inline"
			mstrmode = false
		ELSE 
			document.getElementById("tblBody3").style.display = "none"
			document.getElementById("tblBody4").style.display = "none"
			mstrmode = true
		END IF
		
	End With
End Sub

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

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

'�űԹ�ư
Sub imgREG_onclick ()
	if frmThis.cmbENDGBN.value <> "F" then
		gErrorMsgBox "�ű��߰��� �̿Ϸ� ���¿����� �߰��Ͻ� �� �ֽ��ϴ�.","�ű��߰��ȳ�"
		Exit Sub
	end if
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

Sub imgSave_onclick ()
'	if frmThis.cmbENDGBN.value <> "F" then
'		gErrorMsgBox "������ �̿Ϸ� ���¿����� �����Ͻ� �� �ֽ��ϴ�.","����ȳ�"
'		Exit Sub
'	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'Ȯ����ư �̺�Ʈ
Sub imgConfirm_onclick
	with frmThis
		if frmThis.cmbENDGBN.value <> "F" then
			gErrorMsgBox "Ȯ���� �̿Ϸ� ���¿����� Ȯ��ó���Ͻ� �� �ֽ��ϴ�.","Ȯ���ȳ�"
			Exit Sub
		end if
	End with
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub

'Ȯ����ҹ�ư �̺�Ʈ
Sub imgConfirmCancel_onclick
	with frmThis
		if frmThis.cmbENDGBN.value <> "T" then
			gErrorMsgBox "Ȯ����Ҵ� �Ϸ� ���¿����� Ȯ�����ó���Ͻ� �� �ֽ��ϴ�.","Ȯ����Ҿȳ�"
			Exit Sub
		end if
	End with
	gFlowWait meWAIT_ON
	DeleteRtn_HDR
	gFlowWait meWAIT_OFF
End Sub

'������ư �̺�Ʈ
Sub imgDelete_onclick
	with frmThis
		if frmThis.cmbENDGBN.value <> "F" then
			gErrorMsgBox "������ �̿Ϸ� ���¿����� ����ó���Ͻ� �� �ֽ��ϴ�.","�����ȳ�"
			Exit Sub
		end if
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim intRtn
	Dim i, j, intCount
	Dim strCONTRACTNO
	Dim strUSERID
	Dim vntDataTemp

	If frmThis.cmbENDGBN.value = "F" then 
		gErrorMsgBox "�Ϸ�Ȱ�༭�� �μⰡ �����մϴ�.","ó���ȳ�!"
		Exit Sub
	else
		'üũ�� �� �����Ͱ� �ִ��� ������ üũ�Ѵ�.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
				intCount = 1
			end if
		next
		
		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if intCount = 0 then
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�. �μ��� �����͸� üũ�Ͻÿ�",""
			Exit Sub
		end if
		
		gFlowWait meWAIT_ON
		with frmThis
			'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			intRtn = mobjPDCMCONTRACT_HADO.DeleteRtn_TEMP(gstrConfigXml)
			
			ModuleDir = "PD"
			
			IF .chkCONFLAG.checked THEN 
				
					
				if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2011-12-31") then
					ReportName = "PDCMCONTRACTHADO_CON.rpt"			
				ELSEif cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2013-01-31") then
					ReportName = "PDCMCONTRACTHADO_CON_NEW.rpt"
				ELSE
					ReportName = "PDCMCONTRACTHADO_CON_P.rpt"
				end if 
				
			END IF 
			
			IF .chkDIVFLAG.checked THEN
				if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2011-12-31") then
					ReportName = "PDCMCONTRACTHADO_DIV.rpt"
				ELSEif cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2013-01-31") then
					ReportName = "PDCMCONTRACTHADO_DIV_NEW.rpt"
				ELSE
					ReportName = "PDCMCONTRACTHADO_DIV_P.rpt"
				end if
			END IF 
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strCONTRACTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					strUSERID = ""
					vntDataTemp = mobjPDCMCONTRACT_HADO.ProcessRtn_TEMP(gstrConfigXml,strCONTRACTNO, i, strUSERID)
				END IF
			next
			
			Params = strUSERID 
			Opt = "A"
			
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	end if
End Sub	


'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMCONTRACT_HADO.DeleteRtn_TEMP(gstrConfigXml)
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
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			selectrtn
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
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
					selectrtn
				Else
					Call SEARCHOUT_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ������Ʈ�� �� �޷� /
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgTO_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgTO,"txtTo_onchange()"
		gSetChange
	end with
End Sub

Sub imgFROM2_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtSTDATE,frmThis.imgFROM,"txtSTDATE_onchange()"
		gSetChange
	end with
End Sub

Sub imgTO2_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtEDDATE,frmThis.imgTO,"txtEDDATE_onchange()"
		gSetChange
	end with
End Sub

Sub imgFROM3_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtTESTDAY,frmThis.imgFROM,"txtTESTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub imgTO3_onclick
	WITH frmThis
		gShowPopupCalEndar frmThis.txtTESTENDDAY,frmThis.imgTO,"txtTESTENDDAY_onchange()"
		gSetChange
	end with
End Sub


Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub imgDELIVERYDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtDELIVERYDAY,.imgDELIVERYDAY,"txtDELIVERYDAY_onchange()"
		gSetChange
	end with
End Sub


Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTo_onchange
	gSetChange
End Sub

Sub txtSTDATE_onchange
	IF frmthis.chkCONFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtSTDATE.value 	
	end if 
	gSetChange
End Sub

Sub txtEDDATE_onchange
	IF frmthis.chkDIVFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtEDDATE.value
	end if
	
	frmThis.txtDELIVERYDAY.value = frmThis.txtEDDATE.value
	frmthis.txtTESTDAY.value = frmThis.txtEDDATE.value
	frmthis.txtTESTENDDAY.value = frmthis.txtEDDATE.value
	gSetChange
End Sub

Sub txtTESTDAY_onchange
	gSetChange
End Sub

Sub txtTESTENDDAY_onchange
	gSetChange
End Sub

Sub txtCONTRACTDAY_onchange
	
	frmthis.txtSTDATE.value = frmthis.txtCONTRACTDAY.value 
	gSetChange
End Sub

Sub txtDELIVERYDAY_onchange
	gSetChange
End Sub

Sub cmbENDGBN_onchange
	Set_Layout (frmThis.cmbENDGBN.value)
	SelectRtn
End Sub

Sub chkCONFLAG_onClick
	with FrmThis
		IF .chkDIVFLAG.checked then
			.chkDIVFLAG.checked = false
		else
			.chkCONFLAG.checked = true
		end if 

		
		.txtCONTRACTDAY.value = .txtSTDATE.value
		
		if frmThis.sprSht.ActiveRow > 0 AND frmThis.cmbENDGBN.value = "T"  Then
			if frmThis.chkCONFLAG.checked = TRUE Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFLAG",frmThis.sprSht.ActiveRow, "1"
				
				.txtCONTRACTDAY.value = .txtSTDATE.value		
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value

			else
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFLAG",frmThis.sprSht.ActiveRow, "0"
			End if
			
				
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVFLAG",frmThis.sprSht.ActiveRow, "0"
			
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end with
	
	gSetChange
End Sub

Sub chkDIVFLAG_onClick
	with FrmThis
		IF .chkCONFLAG.checked then
			.chkCONFLAG.checked = false
		else
			.chkDIVFLAG.checked = true
		end if 
		
		.txtCONTRACTDAY.value = .txtEDDATE.value
		
		if frmThis.sprSht.ActiveRow > 0   AND frmThis.cmbENDGBN.value = "T"  Then
			if frmThis.chkDIVFLAG.checked = TRUE Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVFLAG",frmThis.sprSht.ActiveRow, "1"
				
				.txtCONTRACTDAY.value = .txtEDDATE.value
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value			
			else
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVFLAG",frmThis.sprSht.ActiveRow, "0"
			End if
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFLAG",frmThis.sprSht.ActiveRow, "0"
			
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if
	end with
	gSetChange
End Sub

Sub chkCONFIRMFLAG_onClick
	if frmThis.sprSht.ActiveRow > 0   AND frmThis.cmbENDGBN.value = "T"  Then
		if frmThis.chkCONFIRMFLAG.checked = TRUE Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "1"
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "0"
		End if
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub
'-----------------------------------
'�ʵ��߰� 
'------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)			
	End If
	
	if frmThis.cmbENDGBN.value = "F" then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"SEQ",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
	end if
	
End Sub




Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	with frmThis
		If Row = 0 and Col = 1  then 
			mstrCHKcheck = false
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mstrCHKcheck = true
			
			if mstrCheck then 
				.txtAMT.value = .txtSUMAMT.value
				.txtTESTAMT.value = .txtSUMAMT.value
			else
				.txtAMT.value = 0
				.txtTESTAMT.value = 0
			End if
			
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			
			for intcnt = 1 to .sprSht.MaxRows
			
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intcnt
			next
			
		ELSE
			If .cmbENDGBN.value = "T" Then
   				sprShtToFieldBinding Col, Row
   			END IF
		End if		
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	DIM vntData
	Dim dblAMT
	Dim intcnt, intcount
	DIM strYEARMON
	DIM strCode		
	DIM strCodeName
	
	with frmThis
	
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",trim(strCodeName))

				If not gDoErrorRtn ("GetEXECUSTNO") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						
						.txtOUTSNAME.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME"), Row
						.txtOUTSNAME.focus()
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
   	End with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OUTSNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row)))
			
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		if .cmbENDGBN.value = "F" then
			'�� ��Ͽ����� �������� �˾�
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"JOBNAME") Then	
					
				vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",Row)), _
									MID(REPLACE(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTDAY",Row)),"-",""),1,6),"")
				
				vntRet = gShowModalWindow("PDCMJOBNOPOP_SALE.aspx",vntInParams , 413,435)
				If isArray(vntRet) Then
					mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",Row, vntRet(1,0)		
					mobjSCGLSpr.SetTextBinding .sprSht,"JOBNAME",Row, vntRet(2,0)
					
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				End If
			End If
		END IF
		.sprSht.Focus
	End With
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet
	Dim vntInParams
	Dim dblAMT
	
	with frmThis
		if .cmbENDGBN.value = "F" then
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then			
		
				vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row)))
				
				vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
				If isArray(vntRet) Then
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)		
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
					
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				End If
			END IF
		
			'�� ��Ͽ����� �������� �˾�
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTNJOB") Then	
					
				vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",Row)), _
									MID(REPLACE(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTDAY",Row)),"-",""),1,6),"")
				
				vntRet = gShowModalWindow("PDCMJOBNOPOP_SALE.aspx",vntInParams , 413,435)
				If isArray(vntRet) Then
					mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",Row, vntRet(1,0)		
					mobjSCGLSpr.SetTextBinding .sprSht,"JOBNAME",Row, vntRet(2,0)
					
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				End If
			End If
			
			if .sprSht.MaxRows > 0 then
				if mstrCHKcheck then
					
					If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then 
						if .txtAMT.value <> "" then
							dblAMT = .txtAMT.value
						else 
							dblAMT = 0
						end if
						
						if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = "1"  THEN
							dblAMT = dblAMT	+ mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
						Else
							dblAMT = dblAMT	- mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
						End if
						
						.txtAMT.value = dblAMT
						.txtTESTAMT.value = dblAMT
						
						call gFormatNumber(.txtAMT,0,true)
						call gFormatNumber(.txtTESTAMT,0,true)
						
						'������� ���ε�
						.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
						.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
						.txtTESTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
					End if
				End if
			End if
		END IF
		.sprSht.Focus
	End with
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
		
	with frmThis		
	
		If KeyCode = 229 Then Exit Sub
		
		If KeyCode <> meCR and KeyCode <> meTab _
			and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
			and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
			and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

		If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
			If .cmbENDGBN.value  = "T" Then
				sprShtToFieldBinding .sprSht.ActiveCol, .sprSht.ActiveRow
			End IF
		End If
		
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")  Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
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

'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub Init_Layout()
	mobjSCGLCtl.DoEventQueue
    with frmThis
		gSetSheetDefaultColor()   
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 1, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht,    "GUBUN"
		mobjSCGLSpr.SetHeader .sprSht,		    ""
		mobjSCGLSpr.SetColWidth .sprSht, "-1",  "   0"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "GUBUN", -1, -1, 255
	End with
End Sub

Sub Set_Layout (strGUBUN)
	
	Init_Layout '�׸��� �ʱ�ȭ �Ŀ� �Ϸᱸ�п� ���� �׸��� ��
	
	gSetSheetDefaultColor() 
	
	With frmThis
		IF strGUBUN = "F" THEN '��༭ �̿Ϸ�
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 11, 0, 1
			mobjSCGLSpr.AddCellSpan  .sprSht, 3, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.AddCellSpan  .sprSht, 7, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK | SEQ | OUTSCODE | BTN | OUTSNAME | CONTRACTDAY | JOBNO | BTNJOB | JOBNAME | AMT | MEMO"
			mobjSCGLSpr.SetHeader .sprSht,		 "����|����|�ڵ�|����ó��|�����|JOBNO|JOB��|�ݾ�|���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|   0| 8|2|	   18|     8|  6|2|   15|  10|  20"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK "
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN | BTNJOB"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | AMT", -1, -1, 0
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONTRACTDAY", -1, -1, 10
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "OUTSCODE | OUTSNAME | JOBNO | JOBNAME | MEMO", -1, -1, 255
			mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNO",-1,-1,2,2,false
			'mobjSCGLSpr.ColHidden .sprSht, "SEQ", true
			mobjSCGLSpr.CellGroupingEach .sprSht," OUTSNAME"
			
		ELSE '��༭ �Ϸ�, ��ü
			gSetSheetColor mobjSCGLSpr, .sprSht
			mobjSCGLSpr.SpreadLayout .sprSht, 31, 0, 3
			mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONTRACTNO | CONTRACTNAME | OUTSNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY | CONFLAG | DIVFLAG"
			mobjSCGLSpr.SetHeader .sprSht,		"����|��༭��ȣ|��༭��|����ó��|�����|�������|��������|���������|���ݾ�|��ǰ��|�˼���|������޹��|�˼��ǰ�|Ư�����|����|���ޱ���|���ޱ�|�����ޱ���|�����ޱ�|��ȸ������|��ȸ����|�ܱ���|�ܱ�|��������|���ں�����|�����|�˻�����Ⱓ|�˻�ݾ�|��ü�ϼ�"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|		  13|	   25|      15|	 	8|	    13|         8|	       8|      11|     8|     8|           9|       9|      20|   4|       0|     0|         0|       0|         0|       0|     0|   0|         0|         0|     0|           0|		 0|       0"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | CONFIRMFLAG | CONFLAG | DIVFLAG"
			mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONTRACTDAY | STDATE | EDDATE | DELIVERYDAY | TESTDAY | TESTENDDAY"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | TESTAMT | LOSTDAY", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRERATE | ENDRATE | THISRATE | BALANCERATE", -1, -1, 2
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONTRACTNO | CONTRACTNAME | OUTSNAME | LOCALAREA | PAYMENTGBN | TESTMENT | COMENT | MANAGER", -1, -1, 300
			
			mobjSCGLSpr.SetCellsLock2 .sprSht, True, "CONTRACTNO | CONTRACTNAME | OUTSNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY | CONFLAG | DIVFLAG"
			
			mobjSCGLSpr.ColHidden .sprSht, " PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT", true
			mobjSCGLSpr.SetCellAlign2 .sprSht, "CONTRACTNO | MANAGER",-1,-1,2,2,false

		END IF
		
    End With    
End Sub


Sub InitPage()
	'����������ü ����	
	set mobjPDCMCONTRACT_HADO	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT_HADO")
	set mobjPDCMGET				= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"	
	
	mobjSCGLCtl.DoEventQueue
	
	With frmThis
	
		
	End With
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	pnlTab1.style.visibility = "visible"
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT_HADO = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub


'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	With frmThis
		.sprSht.MaxRows = 0
		
		frmThis.txtCONTRACTDAY.value= gNowDate
		frmThis.txtFROM.value		= Mid(gNowDate,1,4) & "-"  & Mid(gNowDate,6,2) & "-" & "01"
		frmThis.txtSTDATE.value		= gNowDate
		frmThis.txtTESTDAY.value	= gNowDate
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean3 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		
		.txtCONTRACTDAY.value = gNowDate
		.txtLOCALAREA.value = "SK Planet"
		.txtCOMENT.value  = ""
		.txtPAYMENTGBN.value = ""
		.txtAMT.value  = 0
		.txtBALANCERATE.value = 0
		.txtBALANCEAMT.value = 0
		.txtDELIVERYGUARANTY.value = 0
		.txtFAULTGUARANTY.value = 0
		.txtTHISRATE.value = 0
		.txtTHISAMT.value = 0
		.txtENDRATE.value = 0
		.txtENDAMT.value =0
		.txtPRERATE.value = 0 
		.txtPREAMT.value = 0
		.txtTESTAMT.value = 0
		.txtTESTMENT.value = "�հ�"
		
		.cmbENDGBN.value = "F"
		.cmbCONFIRM.value = ""
		
		Call Set_Layout(.cmbENDGBN.value)
		Field_Lock
	End With
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	
	frmThis.chkCONFLAG.checked = TRUE
End Sub


'û���� ��ȸ���� ����
Sub DateClean (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtTo.value = date2
	end if
End Sub

Sub DateClean2 (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtEDDATE.value = date2
		frmThis.txtDELIVERYDAY.value = date2
	end if
End Sub


Sub DateClean3 (strYEARMON)
	Dim date1
	Dim date2
	Dim strDATE
	
	if strYEARMON <> "" then
		strDATE = MID(strYEARMON,1,4) & "-" & MID(strYEARMON,5,2)
	
		date1 = Mid(strDATE,1,7)  & "-01"
		date2 = DateAdd("d", -1, DateAdd("m", 1, date1))
	
		frmThis.txtTESTENDDAY.value = date2
	end if
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		.txtCONTRACTNAME.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
		.txtCONTRACTDAY.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
		.txtLOCALAREA.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"LOCALAREA",Row)
		.txtSTDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
		.txtEDDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
		.txtAMT.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		.txtDELIVERYDAY.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYDAY",Row)
		.txtTESTDAY.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"TESTDAY",Row)
		.txtPAYMENTGBN.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"PAYMENTGBN",Row)
		
		if mobjSCGLSpr.GetTextBinding(.sprSht,"TESTMENT",Row) = "" then
			.txtTESTMENT.value = "�հ�"
			mobjSCGLSpr.SetTextBinding .sprSht,"TESTMENT",.sprSht.ActiveRow, .txtTESTMENT.value
		else
			.txtTESTMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTMENT",Row)
		end if
		
		.txtCOMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMENT",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",Row) = "1" THEN
			.chkCONFIRMFLAG.checked = TRUE
		ELSE
			.chkCONFIRMFLAG.checked = FALSE
		END IF
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONFLAG",Row) = "1" THEN
			.chkCONFLAG.checked = TRUE
		ELSE
			.chkCONFLAG.checked = FALSE
		END IF
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"DIVFLAG",Row) = "1" THEN
			.chkDIVFLAG.checked = TRUE
		ELSE
			.chkDIVFLAG.checked = FALSE
		END IF
		
		.txtPRERATE.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"PRERATE",Row)
		.txtPREAMT.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"PREAMT",Row)
		.txtENDRATE.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"ENDRATE",Row)
		.txtENDAMT.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"ENDAMT",Row)
		.txtTHISRATE.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"THISRATE",Row)
		.txtTHISAMT.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"THISAMT",Row)
		.txtBALANCERATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"BALANCERATE",Row)
		.txtBALANCEAMT.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"BALANCEAMT",Row)
		.txtDELIVERYGUARANTY.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYGUARANTY",Row)
		.txtFAULTGUARANTY.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"FAULTGUARANTY",Row)
		.txtMANAGER.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"MANAGER",Row)
		.txtTESTENDDAY.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"TESTENDDAY",Row)
		.txtTESTAMT.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"TESTAMT",Row)
		.txtLOSTDAY.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"LOSTDAY",Row)
		
		Field_Lock
			
	End with
End Function

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	With frmThis
		IntAMTSUM = 0
		
		If .sprSht.MaxRows > 0 Then
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		ELSE
			.txtSUMAMT.value = 0
		END IF
	End With
End Sub


'-----------------------------------------------------------------------------------------
' Field_Lock  �ŷ�������ȣ�� ���ݰ�꼭 ��ȣ�� ������ �����Ҽ� ������ �ʵ带 ReadOnlyó��
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",.sprSht.ActiveRow) <> ""  Then
			
				.txtCONTRACTNAME.className		= "NOINPUT_L" : .txtCONTRACTNAME.readOnly	= True 
				.txtMANAGER.className			= "NOINPUT_L" : .txtMANAGER.readOnly		= True
				.txtAMT.className				= "NOINPUT_R" : .txtAMT.readOnly			= True
				.txtCONTRACTDAY.className		= "NOINPUT"	  : .txtCONTRACTDAY.readOnly	= True
				.txtPRERATE.className			= "NOINPUT_R" : .txtPRERATE.readOnly		= True
				.txtPREAMT.className			= "NOINPUT_R" : .txtPREAMT.readOnly			= True
				.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
				.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
				.txtLOCALAREA.className			= "NOINPUT_L" : .txtLOCALAREA.readOnly		= True 
				.txtDELIVERYGUARANTY.className	= "NOINPUT_R" : .txtDELIVERYGUARANTY.readOnly= True
				.txtDELIVERYDAY.className		= "NOINPUT"   : .txtDELIVERYDAY.readOnly	= True
				.txtPAYMENTGBN.className		= "NOINPUT_L" : .txtPAYMENTGBN.readOnly		= True
				.txtFAULTGUARANTY.className		= "NOINPUT_R" : .txtFAULTGUARANTY.readOnly	= True
				.txtCOMENT.className			= "NOINPUT"   : .txtCOMENT.readOnly			= True 
				
				.txtTESTDAY.className			= "NOINPUT"   : .txtTESTDAY.readOnly		= True
				.txtTESTENDDAY.className		= "NOINPUT"   : .txtTESTENDDAY.readOnly		= True
				.txtENDRATE.className			= "NOINPUT_R" : .txtENDRATE.readOnly		= True
				.txtENDAMT.className			= "NOINPUT_R" : .txtENDAMT.readOnly			= True
				.txtTESTAMT.className			= "NOINPUT_R" : .txtTESTAMT.readOnly		= True
				.txtTHISRATE.className			= "NOINPUT_R" : .txtTHISRATE.readOnly		= True
				.txtTHISAMT.className			= "NOINPUT_R" : .txtTHISAMT.readOnly		= True
				.txtLOSTDAY.className			= "NOINPUT_R" : .txtLOSTDAY.readOnly		= True
				.txtBALANCERATE.className		= "NOINPUT_R" : .txtBALANCERATE.readOnly	= True
				.txtBALANCEAMT.className		= "NOINPUT_R" : .txtBALANCEAMT.readOnly		= True
				.txtTESTMENT.className			= "NOINPUT_L" : .txtTESTMENT.readOnly		= True 
				
				.ImgCONTRACTDAY.disabled = true
				.imgFROM2.disabled = true
				.imgTO2.disabled = true
				.imgDELIVERYDAY.disabled = true
				.ImgFROM3.disabled = true
				.ImgTO3.disabled = true
				
			else 
				.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
				.txtMANAGER.className			= "INPUT_L" : .txtMANAGER.readOnly			= False
				.txtAMT.className				= "INPUT_R" : .txtAMT.readOnly				= False
				.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
				.txtPRERATE.className			= "INPUT_R" : .txtPRERATE.readOnly			= False
				.txtPREAMT.className			= "INPUT_R" : .txtPREAMT.readOnly			= False
				.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
				.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
				.txtLOCALAREA.className			= "INPUT_L" : .txtLOCALAREA.readOnly		= False 
				.txtDELIVERYGUARANTY.className	= "INPUT_R" : .txtDELIVERYGUARANTY.readOnly = False
				.txtDELIVERYDAY.className		= "INPUT"   : .txtDELIVERYDAY.readOnly		= False
				.txtPAYMENTGBN.className		= "INPUT_L" : .txtPAYMENTGBN.readOnly		= False
				.txtFAULTGUARANTY.className		= "INPUT_R" : .txtFAULTGUARANTY.readOnly	= False
				.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
				
				.txtTESTDAY.className			= "INPUT"   : .txtTESTDAY.readOnly			= False
				.txtTESTENDDAY.className		= "INPUT"   : .txtTESTENDDAY.readOnly		= False
				.txtENDRATE.className			= "INPUT_R" : .txtENDRATE.readOnly			= False
				.txtENDAMT.className			= "INPUT_R" : .txtENDAMT.readOnly			= False
				.txtTESTAMT.className			= "INPUT_R" : .txtTESTAMT.readOnly			= False
				.txtTHISRATE.className			= "INPUT_R" : .txtTHISRATE.readOnly			= False
				.txtTHISAMT.className			= "INPUT_R" : .txtTHISAMT.readOnly			= False
				.txtLOSTDAY.className			= "INPUT_R" : .txtLOSTDAY.readOnly			= False
				.txtBALANCERATE.className		= "INPUT_R" : .txtBALANCERATE.readOnly		= False
				.txtBALANCEAMT.className		= "INPUT_R" : .txtBALANCEAMT.readOnly		= False
				.txtTESTMENT.className			= "INPUT_L" : .txtTESTMENT.readOnly			= False 
				
				.ImgCONTRACTDAY.disabled = False
				.imgFROM2.disabled = False
				.imgTO2.disabled = False
				.imgDELIVERYDAY.disabled = False
				.ImgFROM3.disabled = False
				.ImgTO3.disabled = False
			End If
		else
			.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
			.txtMANAGER.className			= "INPUT_L" : .txtMANAGER.readOnly			= False
			.txtAMT.className				= "INPUT_R" : .txtAMT.readOnly				= False
			.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
			.txtPRERATE.className			= "INPUT_R" : .txtPRERATE.readOnly			= False
			.txtPREAMT.className			= "INPUT_R" : .txtPREAMT.readOnly			= False
			.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
			.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
			.txtLOCALAREA.className			= "INPUT_L" : .txtLOCALAREA.readOnly		= False 
			.txtDELIVERYGUARANTY.className	= "INPUT_R" : .txtDELIVERYGUARANTY.readOnly = False
			.txtDELIVERYDAY.className		= "INPUT"   : .txtDELIVERYDAY.readOnly		= False
			.txtPAYMENTGBN.className		= "INPUT_L" : .txtPAYMENTGBN.readOnly		= False
			.txtFAULTGUARANTY.className		= "INPUT_R" : .txtFAULTGUARANTY.readOnly	= False
			.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
			
			.txtTESTDAY.className			= "INPUT"   : .txtTESTDAY.readOnly			= False
			.txtTESTENDDAY.className		= "INPUT"   : .txtTESTENDDAY.readOnly		= False
			.txtENDRATE.className			= "INPUT_R" : .txtENDRATE.readOnly			= False
			.txtENDAMT.className			= "INPUT_R" : .txtENDAMT.readOnly			= False
			.txtTESTAMT.className			= "INPUT_R" : .txtTESTAMT.readOnly			= False
			.txtTHISRATE.className			= "INPUT_R" : .txtTHISRATE.readOnly			= False
			.txtTHISAMT.className			= "INPUT_R" : .txtTHISAMT.readOnly			= False
			.txtLOSTDAY.className			= "INPUT_R" : .txtLOSTDAY.readOnly			= False
			.txtBALANCERATE.className		= "INPUT_R" : .txtBALANCERATE.readOnly		= False
			.txtBALANCEAMT.className		= "INPUT_R" : .txtBALANCEAMT.readOnly		= False
			.txtTESTMENT.className			= "INPUT_L" : .txtTESTMENT.readOnly			= False 
			
			.ImgCONTRACTDAY.disabled = False
			.imgFROM2.disabled = False
			.imgTO2.disabled = False
			.imgDELIVERYDAY.disabled = False
			.ImgFROM3.disabled = False
			.ImgTO3.disabled = False
		End If
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strFROM, strTO
	Dim strOUTSCODE, strOUTSNAME, strJOBNAME, strENDGBN, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1
	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM			= .txtFrom.value
		strTO			= .txtTo.value
		strOUTSCODE		= .txtOUTSCODE.value
		strOUTSNAME		= .txtOUTSNAME.value
		strJOBNAME		= .txtJOBNAME.value
		strENDGBN		= .cmbENDGBN.value
		strCONFIRM		= .cmbCONFIRM.value
		strCONTRACTNO	= .txtCONTRACTNO.value
		strCONTRACTNAME1 = .txtCONTRACTNAME1.value
                              
		If .cmbENDGBN.value = "F" Then '�̿Ϸ���ȸ
			vntData = mobjPDCMCONTRACT_HADO.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
													  strFROM,strTO,strOUTSCODE,strOUTSNAME, strJOBNAME)
		ELSE 
			vntData = mobjPDCMCONTRACT_HADO.SelectRtn_EXIST(gstrConfigXml,mlngRowCnt,mlngColCnt, _
															strFROM,strTO, _
															strOUTSCODE, strOUTSNAME,strJOBNAME, _
															strCONFIRM, strCONTRACTNO, _
															strCONTRACTNAME1)
		End If


		If not gDoErrorRtn ("SelectRtn") Then
			InitPageData
   			PreSearchFiledValue strFROM,strTO, strOUTSCODE, strOUTSNAME, strJOBNAME, strENDGBN, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1
   			Call Set_Layout(.cmbENDGBN.value)
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
				
   				'�˻��ÿ� ù���� MASTER�� ���ε� ��Ű�� ����
   				If .cmbENDGBN.value = "T" Then
   					sprShtToFieldBinding 2, 1
   				END IF
   				
   				AMT_SUM
   			End If
   		End If
   	end With
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strFROM,strTO, strOUTSCODE, strOUTSNAME, strJOBNAME, strENDGBN, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1)
	With frmThis
		.txtFrom.value			= strFROM
		.txtTo.value			= strTO
		.txtOUTSCODE.value		= strOUTSCODE
		.txtOUTSNAME.value		= strOUTSNAME
		.txtJOBNAME.value		= strJOBNAME
		.cmbENDGBN.value		= strENDGBN
		.cmbCONFIRM.value		= strCONFIRM
		.txtCONTRACTNO.value	= strCONTRACTNO
		.txtCONTRACTNAME1.value = strCONTRACTNAME1
	End With
End Sub

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim vntData
	Dim strENDFLAG
	Dim strDataCHK
	Dim lngCol
	Dim lngRow
	
	
	with frmThis
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
		
		
		if strENDFLAG = "F" then
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "OUTSCODE | CONTRACTDAY",lngCol, lngRow, False) 
			
			If strDataCHK = False Then
				gErrorMsgBox lngRow & " ���� ����ó/����� �� �ʼ� �Է»����Դϴ�.","����ȳ�"
				Exit Sub		 
			End If
		end if
		
		strENDFLAG = .cmbENDGBN.value
		if strENDFLAG = "F" then
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | SEQ | OUTSCODE | BTN | OUTSNAME | CONTRACTDAY | JOBNO | BTNJOB | JOBNAME | AMT | MEMO")
		else
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONTRACTNO | CONTRACTNAME | OUTSNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY | CONFLAG | DIVFLAG")
		end if
		
		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		intRtn = mobjPDCMCONTRACT_HADO.ProcessRtn(gstrConfigXml, vntData, strENDFLAG)
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			
			'.cmbENDGBN.value = "T"
			SelectRtn
		End If
	End with
End Sub

'------------------------------------------
' Ȯ�� ������ ó��
'------------------------------------------
Sub ProcessRtn_HDR ()
   	Dim intRtn
   	Dim strMasterData
   	Dim vntData
	Dim lngchkCnt
	Dim i
	Dim strOUTSCODE
	Dim strCONTRACTNO
	
	With frmThis
		strMasterData = gXMLGetBindingData (xmlBind)
		
		lngchkCnt = 0
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",i)
				lngchkCnt = lngchkCnt +1
			End If
			
			if mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i) = ""  then 
				gErrorMsgBox "���� �����Ͱ� ���� ���� �ʾҽ��ϴ� ���� �����͸� ���� �Ͻ��� Ȯ���ϼ���.!.","Ȯ���ȳ�!"
				EXIT Sub
			END IF 
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "Ȯ���� �����͸� üũ�� �ּ���.","Ȯ���ȳ�!"
			EXIT Sub
		End If
		
		'�̿Ϸ� Ȯ����
		if DataValidation =false then exit sub
		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | SEQ | OUTSCODE | BTN | OUTSNAME | CONTRACTDAY | JOBNO | BTNJOB | JOBNAME | AMT | MEMO")
			
		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		
		intRtn = mobjPDCMCONTRACT_HADO.ProcessRtn_HDR(gstrConfigXml, strMasterData, vntData, strCONTRACTNO, strOUTSCODE)

		If not gDoErrorRtn ("ProcessRtn_HDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox strCONTRACTNO & " ��ȣ�� Ȯ�� �Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   		
   	end With
End Sub

'------------------------------------------
'���� ����ó���� �������� �Ǵ��ϱ�����
'-----------------------------------------
Function DataValidation ()
	DataValidation = false
	
   	Dim intCnt
   	Dim strOUTSCODE
   	Dim lngCnt
	'On error resume next
	with frmThis
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� 
   		IF not gDataValidation(frmThis) then exit Function
   		
   		strOUTSCODE = ""
   		for intCnt = 1 To .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
   				strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
   				Exit For
   			End if
   		Next
  
   		for intCnt = 1 to .sprSht.MaxRows
   			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
				if strOUTSCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) Then
					gErrorMsgBox intCnt & " ��° ���� ����ó�� Ȯ���Ͻʽÿ�." & vbcrlf & "���Ͽ���ó �ϰ�쿡�� ������ �����մϴ�.","�Է¿���"
					Exit Function
				End If
			End If
		next
  
   	End with
	DataValidation = true
End Function

'------------------------------------------
' ��༭ ���� ������ ó��
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strSEQ
	Dim strRow
	Dim lngchkCnt
	
	with frmThis
		lngchkCnt = 0
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","Ȯ����Ҿȳ�!"
			EXIT Sub
		End If
		
		'���õ� �ڷḦ ������ ���� ����
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		for i = .sprSht.MaxRows to 1 step -1
			strSEQ = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i) <> "" Then
					strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
					intRtn = mobjPDCMCONTRACT_HADO.DeleteRtn(gstrConfigXml, strSEQ)
					IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				else
					mobjSCGLSpr.DeleteRow .sprSht,i
				End IF
   			End If
   			
		next
		gWriteText lblstatus, "�ڷᰡ " & intRtn & " �� �����Ǿ����ϴ�."
	End with
	err.clear
End Sub

'------------------------------------------
' Ȯ�� ��� ������ ó��
'------------------------------------------
Sub DeleteRtn_HDR ()
   	Dim intRtn
   	Dim vntData
	Dim lngchkCnt
	Dim i
	Dim strCONTRACTNO
	
	With frmThis
		lngchkCnt = 0
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "Ȯ���� ����� �����͸� üũ�� �ּ���.","Ȯ����Ҿȳ�!"
			EXIT Sub
		End If
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | CONTRACTNO | CONTRACTNAME | OUTSNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY")
		
		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"Ȯ����Ҿȳ�"
			exit sub
		End If
		
		intRtn = mobjPDCMCONTRACT_HADO.DeleteRtn_HDR(gstrConfigXml, vntData)

		If not gDoErrorRtn ("ProcessRtn_HDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "Ȯ����� �Ǿ����ϴ�.","Ȯ����Ҿȳ�!"
			SelectRtn
   		End If  		
   	end With
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0">
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
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="150" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�ϵ��ް�༭(�ǰ���)����&nbsp;</td>
											</tr>
										</table>
									</TD>
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
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
							</TABLE>
							<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtfrom,txtTo)"
										width="56">���Ⱓ</TD>
									<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 24px"><INPUT class="INPUT" id="txtFrom" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtFrom"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtTo"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgTo">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 53px; CURSOR: hand; HEIGHT: 24px">�Ϸᱸ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 68px; CURSOR: hand; HEIGHT: 24px"><SELECT id="cmbENDGBN" style="WIDTH: 64px" name="cmbENDGBN">
											<OPTION value="F" selected>�̿Ϸ�</OPTION>
											<OPTION value="T">�Ϸ�</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 64px; CURSOR: hand">��༭Ȯ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 84px; CURSOR: hand"><SELECT id="cmbCONFIRM" style="WIDTH: 80px" name="cmbCONFIRM">
											<OPTION value="" selected>��ü</OPTION>
											<OPTION value="0">��༭ ��Ȯ��</OPTION>
											<OPTION value="1">��༭ Ȯ��</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 43px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)">����ó</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 24px"><INPUT class="INPUT_L" id="txtOUTSNAME" title="����ó�� ��ȸ" style="WIDTH: 160px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtOUTSNAME"> <IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
										<INPUT class="INPUT" id="txtOUTSCODE" title="����ó�ڵ���ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtOUTSCODE"></TD>
									<td class="SEARCHDATA" style="HEIGHT: 24px" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF"
											align="right" border="0" name="imgQuery"></td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')">��༭��ȣ</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCONTRACTNO" title="��༭��ȣ ��ȸ" style="WIDTH: 240px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME1, '')">����</TD>
									<TD class="SEARCHDATA" colSpan="3"><INPUT class="INPUT_L" id="txtCONTRACTNAME1" title="����� ��ȸ" style="WIDTH: 216px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="30" name="txtCONTRACTNAME1"></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)">JOB��</TD>
									<TD class="SEARCHDATA" colSpan="2"><INPUT class="INPUT_L" id="txtJOBNAME" title="JOB�� ��ȸ" style="WIDTH: 160px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtJOBNAME"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" width="300" height="20">
										<table id="TABLE1" cellSpacing="0" cellPadding="0" width="100%" border="0" runat="server">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="150" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�ϵ��ް�༭(�ǰ���)�Է�&nbsp;<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()">(�˻������߰��Է»���)</span></td>
											</tr>
										</table>
									</TD>
									<td>
										<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td class="TITLE" vAlign="middle" align="left" height="20">&nbsp;�հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
													<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 22px"
														readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</td>
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="�ű��ڷḦ �����մϴ�."
														src="../../../images/imgNew.gIF" border="0" name="imgREG"></TD>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
												<TD width="15"></TD>
												<TD><IMG id="imgConfirm" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
														height="20" alt="�ڷḦ Ȯ���մϴ�." src="../../../images/imgSetting.gIF" border="0" name="imgConfirm"></TD>
												<TD><IMG id="imgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/imgConfirmCancelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgConfirmCancel.gIF'"
														height="20" alt="�ڷḦ Ȯ������մϴ�." src="../../../images/imgConfirmCancel.gIF" border="0"
														name="imgConfirmCancel"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD>
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
							<TABLE id="tblBody" style="WIDTH: 100%" cellSpacing="0" cellPadding="0" border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 11px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" id="tblDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1"
											cellPadding="0" align="left" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')">����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 264px"><INPUT dataFld="CONTRACTNAME" class="INPUT_L" id="txtCONTRACTNAME" title="����" style="WIDTH: 240px; HEIGHT: 21px"
														accessKey=",M" dataSrc="#xmlBind" type="text" size="30" name="txtCONTRACTNAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtMANAGER,'')">�������</TD>
												<TD class="SEARCHDATA" style="WIDTH: 180px"><INPUT dataFld="MANAGER" class="INPUT_L" id="txtMANAGER" title="�����" style="WIDTH: 170px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="36" name="txtMANAGER"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 100px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtAMT,'')">���ݾ�</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="AMT" class="INPUT_R" id="txtAMT" title="���ݾ�" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="NUM,M" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTDAY, '')">�����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 264px"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="�����" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY">
													<IMG id="ImgCONTRACTDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" alt="ImgCONTRACTDAY" src="../../../images/btnCalEndar.gIF" align="absMiddle"
														border="0" name="ImgCONTRACTDAY">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;��༭����<INPUT dataFld="CONFIRMFLAG" id="chkCONFIRMFLAG" title="VAT����" dataSrc="#xmlBind" type="checkbox"
														value="" name="chkCONFIRMFLAG"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRERATE, '')">����������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PRERATE" class="INPUT_R" id="txtPRERATE" title="����������" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="33" name="txtPRERATE">(%)</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtPREAMT, '')">���ޱ�</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PREAMT" class="INPUT_R" id="txtPREAMT" title="���ޱ�" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="NUM,M" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtPREAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtSTDATE,txtEDDATE)">���Ⱓ</TD>
												<TD class="SEARCHDATA" style="WIDTH: 264px"><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="���Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtSTDATE">
													<IMG id="imgFROM2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgFROM2">&nbsp;~
													<INPUT dataFld="EDDATE" class="INPUT" id="txtEDDATE" title="���Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtEDDATE">
													<IMG id="imgTO2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgTO2">
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtLOCALAREA, '')">�������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="LOCALAREA" class="INPUT_L" id="txtLOCALAREA" title="�������" style="WIDTH: 170px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="36" name="txtLOCALAREA"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDELIVERYGUARANTY, '')">������� 
													������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="DELIVERYGUARANTY" class="INPUT_R" id="txtDELIVERYGUARANTY" title="������� ������"
														style="WIDTH: 150px; HEIGHT: 22px" accessKey="NUM,M" dataSrc="#xmlBind" type="text" maxLength="100" size="36"
														name="txtDELIVERYGUARANTY"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtDELIVERYDAY, '')">��ǰ��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 264px"><INPUT dataFld="DELIVERYDAY" class="INPUT" id="txtDELIVERYDAY" title="��ǰ��,�Ϸ����" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDELIVERYDAY">
													<IMG id="imgDELIVERYDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="imgDELIVERYDAY">
													��༭<INPUT dataFld="CONFLAG" id="chkCONFLAG" title="��༭" dataSrc="#xmlBind" type="checkbox"
														value="" name="chkCONFLAG">&nbsp;&nbsp;���꼭<INPUT dataFld="DIVFLAG" id="chkDIVFLAG" title="���꼭" dataSrc="#xmlBind" type="checkbox"
														value="" name="chkDIVFLAG">
												</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPAYMENTGBN, '')">������޹��</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PAYMENTGBN" class="INPUT_L" id="txtPAYMENTGBN" title="������޹��" style="WIDTH: 170px; HEIGHT: 22px"
														dataSrc="#xmlBind" type="text" maxLength="255" align="left" size="36" name="txtPAYMENTGBN"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFAULTGUARANTY,'')">���ں��� 
													������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="FAULTGUARANTY" class="INPUT_R" id="txtFAULTGUARANTY" title="���ں��� ������" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="NUM,M" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtFAULTGUARANTY"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCOMENT, '')">Ư�����</TD>
												<TD class="SEARCHDATA" colSpan="7"><TEXTAREA dataFld="COMENT" id="txtCOMENT" style="WIDTH: 778px" dataSrc="#xmlBind" name="txtCOMENT"
														wrap="hard" cols="10" ></TEXTAREA></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody3" style="DISPLAY: none" height="28" cellSpacing="0" cellPadding="0"
								width="100%" background="../../../images/TitleBG.gIF" border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="143" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�˻�����(��ǰ/�뿪/����)</td>
											</tr>
										</table>
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody4" style="DISPLAY: none" cellSpacing="0" cellPadding="0" width="100%"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD vAlign="middle" align="center">
										<TABLE class="SEARCHDATA" style="WIDTH: 100%; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="left" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(TESTDAY, txtTESTENDDAY)"
													width="66">�˻�Ⱓ</TD>
												<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 24px"><INPUT dataFld="TESTDAY" class="INPUT" id="txtTESTDAY" title="�˻�Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtTESTDAY">
													<IMG id="ImgFROM3" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="ImgFROM3">&nbsp;~
													<INPUT dataFld="TESTENDDAY" class="INPUT" id="txtTESTENDDAY" title="�˻�Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtTESTENDDAY">
													<IMG id="ImgTO3" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" src="../../../images/btnCalEndar.gIF" align="absMiddle" border="0" name="ImgTO3"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtENDRATE,'')"
													width="84">�����ޱ���</TD>
												<TD class="SEARCHDATA" style="WIDTH: 180px; HEIGHT: 24px"><INPUT dataFld="ENDRATE" class="INPUT_R" id="txtENDRATE" title="�����ޱ���" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtENDRATE">(%)</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 100px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtENDAMT,'')">�����ޱ�</TD>
												<TD class="SEARCHDATA" style="HEIGHT: 24px"><INPUT dataFld="ENDAMT" class="INPUT_R" id="txtENDAMT" title="�����ޱ�" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtENDAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtTESTAMT, '')">�˻�ݾ�</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="TESTAMT" class="INPUT_R" id="txtTESTAMT" title="�˻�ݾ�" style="WIDTH: 240px; HEIGHT: 21px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtTESTAMT"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand; HEIGHT: 24px" onclick="vbscript:Call gCleanField(txtTHISRATE, '')">��ȸ������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="THISRATE" class="INPUT_R" id="txtTHISRATE" title="��ȸ������" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtTHISRATE">(%)</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTHISAMT, '')">��ȸ����</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="THISAMT" class="INPUT_R" id="txtTHISAMT" title="��ȸ����" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtTHISAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtLOSTDAY, '')">��ü�ϼ�</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="LOSTDAY" class="INPUT_R" id="txtLOSTDAY" title="��ü�ϼ�" style="WIDTH: 226px; HEIGHT: 21px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="37" name="txtLOSTDAY">��
												</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 84px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtBALANCERATE, '')">�ܱ���</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="BALANCERATE" class="INPUT_R" id="txtBALANCERATE" title="�ܱ���" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtBALANCERATE">(%)</TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtBALANCEAMT, '')">�ܱ�</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="BALANCEAMT" class="INPUT_R" id="txtBALANCEAMT" title="�ܱ�" style="WIDTH: 150px; HEIGHT: 22px"
														accessKey="M,NUM" dataSrc="#xmlBind" type="text" maxLength="100" size="36" name="txtBALANCEAMT"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTESTMENT, '')">�˻��ǰ�</TD>
												<TD class="SEARCHDATA" colSpan="5"><INPUT dataFld="TESTMENT" class="INPUT_L" id="txtTESTMENT" title="�˻��ǰ�" style="WIDTH: 778px; HEIGHT: 21px"
														dataSrc="#xmlBind" type="text" maxLength="255" size="124" name="txtTESTMENT"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
					</TR>
					<tr>
						<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
								ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31829">
									<PARAM NAME="_ExtentY" VALUE="2805">
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
					</tr>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
	</body>
</HTML>
