<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT.aspx.vb" Inherits="PD.PDCMCONTRACT" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��༭ ��� �� Ȯ��</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/TR/���Ա� ��� ȭ��(TRLNREGMGMT0)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ���Աݿ� ���� MAIN ������ ��ȸ/�Է�/����/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/03/21 By kty
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<script id="clientEventHandlersVBS" language="vbscript">
		
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mcomecalender, mcomecalender2, mcomecalender3
Dim mobjPDCMCONTRACT, mobjPDCMGET
Dim mstrCheck
Dim mstrMEDGUBN
Dim mstrChk
Dim mstrmode
Dim mstrCHKcheck

CONST meTAB = 9
mcomecalender = FALSE
mcomecalender2 = FALSE
mcomecalender3 = FALSE

mstrMEDGUBN = ""
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

'�űԹ�ư
Sub imgREG_onclick ()
	if frmThis.cmbENDGBN.value <> "F" then
		gErrorMsgBox "�ű��߰��� �̿Ϸ� ���¿����� �߰��Ͻ� �� �ֽ��ϴ�.","�ű��߰��ȳ�"
		Exit Sub
	end if
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

'���� ��ư �̺�Ʈ
Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'��༭ ����
Sub imgContractCre_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_CONTRACT
	gFlowWait meWAIT_OFF
End Sub

'������ư �̺�Ʈ
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'�μ��ư �̺�Ʈ 
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
	Dim strOWNER

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
			intRtn = mobjPDCMCONTRACT.DeleteRtn_TEMP(gstrConfigXml)
		
			ModuleDir = "PD"
			
			IF .chkCONFLAG.checked THEN 
				if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2011-12-31") then
					strOWNER = "�̹���"
					ReportName = "PDCMCONTRACT_CON.rpt"
				elseif cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2013-01-31") then
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_CON.rpt"
				else 
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_CON_P.rpt"
				end if 
				
				
				
			END IF 
			
			IF .chkDIVFLAG.checked THEN
				if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2011-12-31") then
					strOWNER = "�̹���"
					ReportName = "PDCMCONTRACT_DIV.rpt"
				elseif cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2013-01-31") then
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_DIV.rpt"
				else 
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_DIV_P.rpt"
				end if 
				
				
			END IF 
			
			IF .chkEXEFLAG.checked THEN
				if cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2011-12-31") then
					strOWNER = "�̹���"
					ReportName = "PDCMCONTRACT_EXE_NEW.rpt"
				elseif cdate(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",.sprSht.ActiveRow)) <= cdate("2013-01-31") then
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_EXE_NEW.rpt"
				else 
					strOWNER = "������"
					ReportName = "PDCMCONTRACT_EXE_NEW_P.rpt"
				end if 
				
			END IF 
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strCONTRACTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					strUSERID = ""
					vntDataTemp = mobjPDCMCONTRACT.ProcessRtn_TEMP(gstrConfigXml,strCONTRACTNO, i, strUSERID)
				END IF
			next
			
			Params = strUSERID & ":" & strOWNER
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
		intRtn = mobjPDCMCONTRACT.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub



'������ư �̺�Ʈ
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub

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
			selectrtn
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
					selectrtn
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
		SELECTRTN
	end if
End Sub


'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCONTRACTNO_onkeydown
	if window.event.keyCode = meEnter then
		SELECTRTN
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ������Ʈ�� �� �޷� /
'-----------------------------------------------------------------------------------------
Sub imgFROM_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgTO,"txtTo_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgFROM2_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtSTDATE,frmThis.imgFROM,"txtSTDATE_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO2_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtEDDATE,frmThis.imgTO,"txtEDDATE_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgFROM3_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtTESTDAY,frmThis.imgFROM,"txtTESTDAY_onchange()"
		mcomecalender = false
		gSetChange
	end with
End Sub

Sub imgTO3_onclick
	WITH frmThis
		mcomecalender = true
		gShowPopupCalEndar frmThis.txtTESTENDDAY,frmThis.imgTO,"txtTESTENDDAY_onchange()"
		mcomecalender = false
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


'****************************************************************************************
' ��ȸ�ʵ� ü���� �̺�Ʈ
'****************************************************************************************
'�˻����� ������
Sub txtFROM_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	strdate = ""
	strFROM =""
	strFROM2 = ""
	With frmThis
		strdate=.txtFROM.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strFROM2 = strdate
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		'DateClean strFROM
		
	End With
	gSetChange
End Sub



Sub cmbENDGBN_onchange
	with frmThis
		If .cmbENDGBN.value = "T" Then
			.txtCONTRACTNO.style.visibility = "visible"
			.cmbCONFIRM.style.visibility = "visible"
			pnlFLAG1.style.visibility = "visible" 
			pnlFLAG2.style.visibility = "hidden" 
			.txtJOBNO.style.visibility = "hidden"
			.txtJOBNAME.style.visibility = "hidden"
			.ImgJOBNO.style.visibility = "hidden"
			
		Elseif  .cmbENDGBN.value = "F" Then
			.txtCONTRACTNO.style.visibility = "hidden"
			.cmbCONFIRM.style.visibility = "hidden"
			pnlFLAG1.style.visibility = "hidden" 
			pnlFLAG2.style.visibility = "visible" 
			.txtJOBNO.style.visibility = "visible"
			.txtJOBNAME.style.visibility = "visible"
			.ImgJOBNO.style.visibility = "visible"
			.cmbYEARMONGBN.value = "REGDATE"
			.cmbAMTFLAG.value = "1"
			
		End If
		
	End with
	SelectRtn
End Sub

Sub cmbYEARMONGBN_onchange
	with frmThis
		If .cmbYEARMONGBN.value = "CONTRACTDAY" Then
			IF .cmbENDGBN.value = "F" THEN
				gErrorMsgBox " �̿Ϸ� ������ ��쿡�� ���Ⱓ���� ��ȸ �Ͻ� �� �����ϴ�.","��ȸ�ȳ�"
				.cmbYEARMONGBN.value = "REGDATE"
				exit sub
			END IF
		End If
	End with
	SelectRtn
End Sub


Sub cmbCONFIRM_onchange
	SelectRtn
	gSetChange
End Sub



Sub cmbAMTFLAG_onchange
	SelectRtn
	gSetChange
End Sub

'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************
Sub txtTo_onchange
	gSetChange
End Sub

Sub txtCONTRACTNAME_onchange
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtCONTRACTDAY_onchange
	'frmthis.txtSTDATE.value = frmthis.txtCONTRACTDAY.value 
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtLOCALAREA_Onchange
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"LOCALAREA",frmThis.sprSht.ActiveRow, frmThis.txtLOCALAREA.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtSTDATE_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	strdate = ""
	strFROM =""
	strFROM2 = ""
	
	With frmThis
		strdate=.txtSTDATE.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender2 Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
	'	DateClean2 strFROM
	End With
	
	IF frmthis.chkCONFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtSTDATE.value 	
	end if 
	
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtEDDATE_onchange
	
	IF frmthis.chkDIVFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtEDDATE.value
	end if 
	frmThis.txtDELIVERYDAY.value = MID(frmThis.txtEDDATE.value,1,4) & "-" & MID(frmThis.txtEDDATE.value,6,2) & "-" & MID(frmThis.txtEDDATE.value,9,2)
	txtDELIVERYDAY_onchange
	
	frmthis.txtTESTDAY.value = frmThis.txtEDDATE.value
	txtTESTDAY_onchange
	frmthis.txtTESTENDDAY.value = frmthis.txtEDDATE.value
	txtTESTENDDAY_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVERYDAY",frmThis.sprSht.ActiveRow, frmThis.txtDELIVERYDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	
	gSetChange
End Sub

Sub txtAMT_Onchange
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTAMT",frmThis.sprSht.ActiveRow, frmThis.txtAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	IF frmThis.cmbENDGBN.value = "F" THEN
		frmThis.txtTESTAMT.value = frmThis.txtAMT.value	
		txtAMT_onblur
	END IF
	gSetChange
End Sub

Sub txtDELIVERYDAY_onchange
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVERYDAY",frmThis.sprSht.ActiveRow, frmThis.txtDELIVERYDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtTESTDAY_onchange
	Dim strdate 
	Dim strFROM, strFROM2
	strdate = ""
	strFROM =""
	strFROM2 = ""
	With frmThis
		
		strdate=.txtTESTDAY.value
		'�޷��˾��� ���� �����ʹ� 2000-01-01�̷������� ������ �����Է��� 20000101�̷������� �����Ƿ�
		If mcomecalender3 Then
			strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
			strFROM2 = strdate
		else
			If len(strdate) = 4 Then
				strFROM = Mid(gNowDate2,1,4) & Mid(strdate,1 , 2)
				strFROM2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,6 , 2)
				strFROM2 = strdate
			elseif len(strdate) = 3 Then
				strFROM = Mid(gNowDate2,1,4) & "0" & Mid(strdate,1 , 1)
				strFROM2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strFROM = Mid(strdate,1 , 4) & Mid(strdate,5 , 2)
				strFROM2 = strdate
			End If
		End If
		DateClean3 strFROM
	
	End With
	if frmThis.sprSht.ActiveRow >0 AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTDAY",frmThis.sprSht.ActiveRow, frmThis.txtTESTDAY.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTENDDAY",frmThis.sprSht.ActiveRow, frmThis.txtTESTENDDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtPAYMENTGBN_onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PAYMENTGBN",frmThis.sprSht.ActiveRow, frmThis.txtPAYMENTGBN.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


Sub txtTESTMENT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTMENT",frmThis.sprSht.ActiveRow, frmThis.txtTESTMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtCOMENT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMENT",frmThis.sprSht.ActiveRow, frmThis.txtCOMENT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
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

Sub chkCONFLAG_onClick
	with FrmThis
		IF .chkDIVFLAG.checked or .chkEXEFLAG.checked  then
			.chkDIVFLAG.checked = false
			.chkEXEFLAG.checked = false
		else
			.chkCONFLAG.checked = true
		end if 

		'��༭�� ���꼭 �� �ٸ��ٰ� �ϼż� �߰� ����
		.txtCONTRACTDAY.value = .txtSTDATE.value
		

		if frmThis.sprSht.ActiveRow > 0 AND frmThis.cmbENDGBN.value = "T"  Then
			if frmThis.chkCONFLAG.checked = TRUE Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFLAG",frmThis.sprSht.ActiveRow, "1"
				
				'��༭�� ���꼭 �� �ٸ��ٰ� �ϼż� �߰� ����
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
		IF .chkCONFLAG.checked or .chkEXEFLAG.checked then
			.chkCONFLAG.checked = false
			.chkEXEFLAG.checked = false
		else
			.chkDIVFLAG.checked = true
		end if 

		'��༭�� ���꼭 �� �ٸ��ٰ� �ϼż� �߰� ����
		.txtCONTRACTDAY.value = .txtEDDATE.value
		
		
		if frmThis.sprSht.ActiveRow > 0   AND frmThis.cmbENDGBN.value = "T"  Then
			if frmThis.chkDIVFLAG.checked = TRUE Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DIVFLAG",frmThis.sprSht.ActiveRow, "1"
				
				'��༭�� ���꼭 �� �ٸ��ٰ� �ϼż� �߰� ����
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

Sub chkEXEFLAG_onClick
	with FrmThis
		IF .chkCONFLAG.checked or .chkDIVFLAG.checked then
			.chkCONFLAG.checked = false
			.chkDIVFLAG.checked = false
		else
			.chkEXEFLAG.checked = true
		end if 
	end with
	gSetChange
End Sub

Sub txtOUTSCODE1_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OUTSCODE",frmThis.sprSht.ActiveRow, frmThis.txtOUTSCODE1.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


Sub txtPRERATE_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PRERATE",frmThis.sprSht.ActiveRow, frmThis.txtPRERATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtPREAMT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PREAMT",frmThis.sprSht.ActiveRow, frmThis.txtPREAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtENDRATE_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDRATE",frmThis.sprSht.ActiveRow, frmThis.txtENDRATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


Sub txtENDAMT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"ENDAMT",frmThis.sprSht.ActiveRow, frmThis.txtENDAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtTHISRATE_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"THISRATE",frmThis.sprSht.ActiveRow, frmThis.txtTHISRATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtTHISAMT_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"THISAMT",frmThis.sprSht.ActiveRow, frmThis.txtTHISAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtBALANCERATE_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BALANCERATE",frmThis.sprSht.ActiveRow, frmThis.txtBALANCERATE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtBALANCEAMT_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BALANCEAMT",frmThis.sprSht.ActiveRow, frmThis.txtBALANCEAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


Sub txtDELIVERYGUARANTY_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DELIVERYGUARANTY",frmThis.sprSht.ActiveRow, frmThis.txtDELIVERYGUARANTY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtFAULTGUARANTY_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"FAULTGUARANTY",frmThis.sprSht.ActiveRow, frmThis.txtFAULTGUARANTY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtMANAGER_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"MANAGER",frmThis.sprSht.ActiveRow, frmThis.txtMANAGER.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtTESTENDDAY_Onchange
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTENDDAY",frmThis.sprSht.ActiveRow, frmThis.txtTESTENDDAY.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub txtTESTAMT_Onchange
	
	if frmThis.sprSht.ActiveRow >0   AND frmThis.cmbENDGBN.value = "T"  Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"TESTAMT",frmThis.sprSht.ActiveRow, frmThis.txtTESTAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub


Sub txtLOSTDAY_Onchange
	if frmThis.sprSht.ActiveRow >0  AND frmThis.cmbENDGBN.value = "T"   Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"LOSTDAY",frmThis.sprSht.ActiveRow, frmThis.txtLOSTDAY.value
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
		.txtCONTRACTDAY.focus()
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

Sub txtPREAMT_onfocus
	with frmThis
		.txtPREAMT.value = Replace(.txtPREAMT.value,",","")
	end with
End Sub

Sub txtPREAMT_onblur
	with frmThis
		.txtSTDATE.focus()
		call gFormatNumber(.txtPREAMT,0,true)
	end with
End Sub


Sub txtDELIVERYGUARANTY_onfocus
	with frmThis
		.txtDELIVERYGUARANTY.value = Replace(.txtDELIVERYGUARANTY.value,",","")
	end with
End Sub

Sub txtDELIVERYGUARANTY_onblur
	with frmThis
		call gFormatNumber(.txtDELIVERYGUARANTY,0,true)
	end with
End Sub


Sub txtFAULTGUARANTY_onfocus
	with frmThis
		.txtFAULTGUARANTY.value = Replace(.txtFAULTGUARANTY.value,",","")
	end with
End Sub
Sub txtFAULTGUARANTY_onblur
	with frmThis
		.txtCOMENT.focus()
		call gFormatNumber(.txtFAULTGUARANTY,0,true)
	end with
End Sub

Sub txtTESTAMT_onfocus
	with frmThis
		.txtTESTAMT.value = Replace(.txtTESTAMT.value,",","")
	end with
End Sub


Sub txtTESTAMT_onblur
	with frmThis
		.txtTHISRATE.focus()
		call gFormatNumber(.txtTESTAMT,0,true)
	end with
End Sub

Sub txtENDAMT_onfocus
	with frmThis
		.txtENDAMT.value = Replace(.txtENDAMT.value,",","")
	end with
End Sub

Sub txtENDAMT_onblur
	with frmThis
		.txtTESTAMT.focus()
		call gFormatNumber(.txtENDAMT,0,true)
	end with
End Sub

Sub txtTHISAMT_onfocus
	with frmThis
		.txtTHISAMT.value = Replace(.txtTHISAMT.value,",","")
	end with
End Sub

Sub txtTHISAMT_onblur
	with frmThis
		.txtLOSTDAY.focus()
		call gFormatNumber(.txtTHISAMT,0,true)
	end with
End Sub

Sub txtBALANCEAMT_onfocus
	with frmThis
		.txtBALANCEAMT.value = Replace(.txtBALANCEAMT.value,",","")
	end with
End Sub

Sub txtBALANCEAMT_onblur
	with frmThis
		.txtTESTMENT.focus()
		call gFormatNumber(.txtBALANCEAMT,0,true)
	end with
End Sub



'****************************************************************************************
' �̺�Ʈ ó��
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	Dim dblAMT
	
	with frmThis
		If Row = 0 and Col = 1  then 
			mstrCHKcheck = false
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mstrCHKcheck = true
			
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
		End if
		
		if Row > 0 and Col > 0 then
			If .cmbENDGBN.value  = "T" Then
				sprShtToFieldBinding Col,Row
			End IF
			If .cmbENDGBN.value = "F" then
				.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",Row)
			End If
		end if
	end with
End Sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim dblAMT
	Dim intcnt
	Dim vntData
	Dim strCode
	Dim strCodeName
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		'����ó �������� �˾�
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
						.txtOUTSCODE1.value = vntData(0,0)
						
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
		
		if .cmbENDGBN.value = "F" then
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CONTRACTGUBUN") Then 
				strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTGUBUN",Row))
				'���ڸ� �ڵ常 �ҷ��´�.
					strCodeName = Mid(strCodeName,1,1)
				Call Get_JOBGUBUN_VALUE(strCodeName,Row)

		
			End If
		End If
	End with
	
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet
	Dim vntInParams
	Dim dblAMT
	
	with frmThis
		if .cmbENDGBN.value = "F" then
			if .sprSht.MaxRows > 0 then
				if mstrCHKcheck then
					If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CHK") Then 
						if .txtAMT.value <> "" then
							dblAMT = .txtAMT.value
						else 
							dblAMT = 0
						end if
						
						if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = "1"  THEN
							dblAMT = dblAMT	+ mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT",Row)
						Else
							dblAMT = dblAMT	- mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT",Row)
						End if
						
						.txtAMT.value = dblAMT
						.txtTESTAMT.value = dblAMT
						
						call gFormatNumber(.txtAMT,0,true)
						call gFormatNumber(.txtTESTAMT,0,true)
						
						'������� ���ε�
						.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REGDATE",Row)
						.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REGDATE",Row)
						.txtTESTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"REGDATE",Row)
					End if
				End if
			End if
		END IF
		.sprSht.Focus
	End with
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
				.txtOUTSCODE1.value = vntRet(0,0)	
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		.sprSht.Focus
	End With
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode = meINS_ROW Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,False,frmThis.sprSht.ActiveRow,1,13,True
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMTFLAG",frmThis.sprSht.ActiveRow, "1"
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht,true,"JOBNO"
	End If
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntInParams
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Elseif Row > 0 and Col > 1 then
			If .cmbENDGBN.value = "T" Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO", Row))
				gShowModalWindow "../PDCO/PDCMCONTRACT_DTLPOP.aspx",vntInParams , 810,580
			End IF
		end if
	end with
end sub

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
			If .cmbENDGBN.value = "F" then
				.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNAME",.sprSht.ActiveRow)
			End If
		End If
		
		IF .cmbENDGBN.value = "F" THEN
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")  Then
				strSUM = 0
				intSelCnt = 0
				intSelCnt1 = 0
				strCOLUMN = ""

				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
					strCOLUMN = "ADJAMT"
				End If
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

				FOR i = 0 TO intSelCnt -1
					If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT"))  Then
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
		ELSE
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
		END IF 
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
			IF .cmbENDGBN.value ="F" THEN
				If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
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
			ELSE
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
			END IF 
		else
			.txtSELECTAMT.value = 0
		End If
		Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		
			.txtCONTRACTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
			.txtCONTRACTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
			.txtLOCALAREA.value = mobjSCGLSpr.GetTextBinding(.sprSht,"LOCALAREA",Row)
			.txtSTDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
			.txtEDDATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
			.txtAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
			.txtDELIVERYDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYDAY",Row)
			.txtTESTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTDAY",Row)
			.txtPAYMENTGBN.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PAYMENTGBN",Row)
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
			.txtOUTSCODE1.value = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",Row)
			.txtPRERATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PRERATE",Row)
			.txtPREAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PREAMT",Row)
			.txtENDRATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDRATE",Row)
			.txtENDAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"ENDAMT",Row)
			.txtTHISRATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"THISRATE",Row)
			.txtTHISAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"THISAMT",Row)
			.txtBALANCERATE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BALANCERATE",Row)
			.txtBALANCEAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"BALANCEAMT",Row)
			.txtDELIVERYGUARANTY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DELIVERYGUARANTY",Row)
			.txtFAULTGUARANTY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"FAULTGUARANTY",Row)
			.txtMANAGER.value = mobjSCGLSpr.GetTextBinding(.sprSht,"MANAGER",Row)
			.txtTESTENDDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTENDDAY",Row)
			.txtTESTAMT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"TESTAMT",Row)
			.txtLOSTDAY.value = mobjSCGLSpr.GetTextBinding(.sprSht,"LOSTDAY",Row)
			
		If .txtAMT.value <> "" Then
			call gFormatNumber(.txtAMT,0,true)
			call gFormatNumber(.txtTESTAMT,0,true)
		End If
	End with
End Function

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


Sub Input_Layout
	gSetSheetDefaultColor() 
	With frmThis
		'******************************************************************
		'��༭ �̿Ϸ�
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONTRACTNO | OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT | JOBGUBN | CREPART | CONTRACTGUBUN | RANKTRANS | SEQ | AMTFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|��༭��ȣ|�ڵ�|����ó|�����|JOBNO|JOB��/����|�ݾ�|���ۺι�|���ۺз�|���з�|��ũ|����|�ϵ���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|        10|   6|    30|    10|   12|          20|  11|      10|      10|	   15|   0|   0|     4"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"	
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | AMTFLAG"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "REGDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONTRACTNO | OUTSCODE | OUTSNAME | JOBNO | JOBNAME | JOBGUBN | CREPART | CONTRACTGUBUN  ", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CONTRACTNO | OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT | JOBGUBN | CREPART | RANKTRANS | SEQ"
		mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO | JOBGUBN | CREPART | RANKTRANS | SEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CHK | CONTRACTNO | JOBNO | OUTSCODE | CONTRACTGUBUN ",-1,-1,2,2,false
		mobjSCGLSpr.CellGroupingEach .sprSht,"OUTSNAME | CONTRACTNO"

		Get_COMBO_VALUE

    End With    
End Sub

Sub Select_Layout
	Dim strComboList
	gSetSheetDefaultColor() 
	With frmThis
		strComboList =  "��༭ ��Ȯ��" & vbTab & "��༭ Ȯ��"
		'******************************************************************
		'��༭ �Ϸ�, ��ü
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 35, 0, 3
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | CONTRACTGUBUN | CONTRACTNO | OUTSNAME | CONTRACTNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | CONFLAG | DIVFLAG | OUTSCODE | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY | RANKTRANS | AMTFLAG"
		mobjSCGLSpr.SetHeader .sprSht,		"����|��༭����|��༭��ȣ|����ó��|����|�����|��ǰ���|�뿪������|�뿪������|���ݾ�|��ǰ��|�˼���|������޹��|�˼����|Ư�����|����|���|����|����ó�ڵ�|���ޱ���|���ޱ�|�����ޱ���|�����ޱ�|��ȸ������|��ȸ����|�ܱ���|�ܱ�|��������|���ں�����|�����|�˻�����Ⱓ|�˻�ݾ�|��ü�ϼ�|��ũ|�ϵ���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "6|        10|		 10|      17|    18|     8|       6|        10|        10|      10|     9|     9|           5|       6|       8|   4|   4|   4|         0|       0|     0|         0|       0|         0|       0|     0|   0|         0|         0|     0|           0|       0|      0 |   0|     4"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | CONFIRMFLAG | CONFLAG | DIVFLAG | AMTFLAG"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "STDATE | EDDATE | DELIVERYDAY | TESTDAY | TESTENDDAY | CONTRACTDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "PRERATE | ENDRATE | THISRATE | BALANCERATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | PREAMT | ENDAMT | THISAMT | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | TESTAMT | LOSTDAY", -1, -1, 0
		'mobjSCGLSpr.SetCellTypeComboBox .sprSht,18,18,-1,-1,strComboList
		mobjSCGLSpr.SetCellAlign2 .sprSht, "CONTRACTGUBUN | CONTRACTNO | CONTRACTNAME",-1,-1,2,2,false
	    mobjSCGLSpr.SetCellAlign2 .sprSht, "OUTSNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "LOCALAREA | PAYMENTGBN | TESTMENT | COMENT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"CONTRACTGUBUN | CONTRACTNO | OUTSNAME | CONTRACTNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | CONFLAG | DIVFLAG | OUTSCODE | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY"
		mobjSCGLSpr.ColHidden .sprSht, "OUTSCODE | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY", true
		mobjSCGLSpr.CellGroupingEach .sprSht,"CONTRACTNAME | LOCALAREA",,,,0
		
    End With    
End Sub

Sub Get_COMBO_VALUE ()
	Dim vntCONTRACTGUBUN
   	Dim i, strCols
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntCONTRACTGUBUN = mobjPDCMCONTRACT.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"PD_CONTRACT")
		
		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "CONTRACTGUBUN",,,vntCONTRACTGUBUN,,120
			mobjSCGLSpr.TypeComboBox = True 
   		End If
   	End With
End Sub

Sub Get_JOBGUBUN_VALUE(strCODE,strPos)							
	Dim vntData					
	With frmThis   					
		On error resume Next				
		'Long Type�� ByRef ������ �ʱ�ȭ				
		mlngRowCnt=clng(0)				
		mlngColCnt=clng(0)		

       	vntData = mobjPDCMCONTRACT.GetDataType_JOBGUBUN(gstrConfigXml, mlngRowCnt, mlngColCnt, strCODE)					
		If not gDoErrorRtn ("Get_JOBGUBUN_VALUE") Then 				
			
			if mlngRowCnt > 0 THEN
				.txtJOBGUBN.value = trim(vntData(0,0))
				.txtCREPART.value = trim(vntData(0,1))
			else 
				.txtJOBGUBN.value =	""
				.txtCREPART.value = ""
			END if
   		End If  				
   		gSetChange				
   	end With   					
End Sub	

Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	
	'����������ü ����	
	set mobjPDCMCONTRACT	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
	With frmThis
		'Sheet �⺻Color ����
		
		document.getElementById("tblBody3").style.display = "none"
		document.getElementById("tblBody4").style.display = "none"
		.txtCONTRACTNO.style.visibility = "hidden"
		.txtOUTSCODE1.style.visibility = "hidden"
		
		Input_Layout
		pnlTab1.style.visibility = "visible"
	End With
	
	'ȭ�� �ʱⰪ ����
	InitPageData
	
	'------------------------------------------------------------------------------------
	'�ʱ� ���� ���� (���ѱ� ����� ��û���� ���� ��ȸ�ø� ��¥ ���� �ϱ����...._20111209)
	'------------------------------------------------------------------------------------
	frmThis.txtFROM.value = Mid(gNowDate2,1,4) & "-"  & Mid(gNowDate2,6,2) & "-" & "01"
	frmThis.txtSTDATE.value = gNowDate2
'	frmThis.txtTESTDAY.value = gNowDate2
	DateClean Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	DateClean2 Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	DateClean3 Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	'------------------------------------------------------------------------------------
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	'�ʱ� ������ ����
	with frmThis
		
		'�ʱ� ������ ���� 
		.txtCONTRACTDAY.value = gNowDate2
		.txtLOCALAREA.value = "SK Planet"
		.cmbCONFIRM.style.visibility = "hidden"
		pnlFLAG1.style.visibility = "hidden"
		pnlFLAG2.style.visibility = "visible"
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

		'���ο� XML ���ε��� ����
		gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
		.chkCONFLAG.checked = true
		
	End with
End Sub


'*********************************
'û���� ��ȸ���� ����
'*********************************
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
	
	'	frmThis.txtTESTENDDAY.value = date2
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ��������ȸ
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim strYEARMONGBN
	Dim strFROM
	Dim strTO
	Dim strCONTRACTNO
	Dim strENDGBN
	Dim strCONFIRM
	Dim strCONTRACTCODE
	Dim strJOBNOCODE
	Dim strOUTSCODE
	Dim strOUTSNAME
	Dim strJOBNO
	Dim strJOBNAME
	Dim vntData
	Dim intCnt
	Dim strAMTFLAG
	'On error resume next
	
	with frmThis
		.sprSht.MaxRows = 0
		strFROM = .txtFROM.value
		strTO = .txtTo.value
		strCONTRACTNO = .txtCONTRACTNO.value 
		strENDGBN = .cmbENDGBN.value 
		strCONFIRM = .cmbCONFIRM.value
		strCONTRACTCODE = .cmbCONTRACTCODE.value
		strJOBNOCODE = .cmbJOBNOCODE.value  'JOB�� �ձ��ڸ� ���� �˻��Ҽ� �ִ�.
		strOUTSCODE = TRIM(.txtOUTSCODE.value)
		strOUTSNAME =  TRIM(.txtOUTSNAME.value)
		strJOBNO = TRIM(.txtJOBNO.value)
		strJOBNAME =  TRIM(.txtJOBNAME.value)
		strAMTFLAG = .cmbAMTFLAG.value
		
		strYEARMONGBN = .cmbYEARMONGBN.value
		
		If Len(strCONTRACTNO) = 10 Then
			strCONTRACTNO = MID(strCONTRACTNO,1,7) & "-" & MID(strCONTRACTNO,8,3)
		End if
		
		'========================================================================================================================================
		'�̿Ϸ���ȸ
		'========================================================================================================================================
		IF strENDGBN = "F" THEN
			if mstrMEDGUBN <> .cmbENDGBN.value then
				Call Init_Layout()
				Call Input_Layout()
			End if
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strAMTFLAG,strFROM,strTO,strOUTSCODE,strOUTSNAME,strJOBNO,strJOBNAME,strJOBNOCODE)

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
					
					PreSearchFiledValue strFROM, strTO, strCONTRACTNO, strENDGBN, strCONFIRM, strOUTSCODE, strOUTSNAME, strJOBNO,  strJOBNAME
   				Else
   					.sprSht.MaxRows = 0
   					initpageData
   				End If
   			end if
   		'========================================================================================================================================
   		'�Ϸ���ȸ
   		'========================================================================================================================================
		ELSEIF strENDGBN = "T" THEN
			if mstrMEDGUBN <> .cmbENDGBN.value then
				Call Init_Layout()
				Call Select_Layout()
			End if

			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjPDCMCONTRACT.SelectRtn_EXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strAMTFLAG,strFROM,strTO,strOUTSCODE,strOUTSNAME,strCONFIRM,strCONTRACTNO,strCONTRACTCODE,strYEARMONGBN)

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
   				mobjSCGLSpr.ColHidden .sprSht, "CONTRACTNO | CONTRACTNAME", false	
   				PreSearchFiledValue strFROM, strTO, strCONTRACTNO, strENDGBN, strCONFIRM, strOUTSCODE, strOUTSNAME, strJOBNO,  strJOBNAME
   			end if
		END IF
		AMT_SUM
		mstrMEDGUBN = .cmbENDGBN.value
   	end with
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		IF .cmbENDGBN.value = "F" THEN
			For lngCnt = 1 To .sprSht.MaxRows
				IntAMT = 0
				IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngCnt)
				IntAMTSUM = IntAMTSUM + IntAMT
			Next
			If .sprSht.MaxRows = 0 Then
				.txtSUMAMT.value = 0
			else
				.txtSUMAMT.value = IntAMTSUM
				Call gFormatNumber(frmThis.txtSUMAMT,0,True)
			End If
		ELSE
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
		END IF 
	End With
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strFROM, strTO, strCONTRACTNO, strENDGBN, strCONFIRM, strOUTSCODE, strOUTSNAME, strJOBNO, strJOBNAME)
	With frmThis
		.txtFrom.value = strFROM
		.txtTo.value =  strTO
		.txtCONTRACTNO.value = strCONTRACTNO
		.cmbENDGBN.value = strENDGBN
		.cmbCONFIRM.value = strCONFIRM
		.txtOUTSCODE.value = strOUTSCODE
		.txtOUTSNAME.value =  strOUTSNAME
		.txtJOBNO.value = strJOBNO
		.txtJOBNAME.value =  strJOBNAME
	End With
End Sub

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim strFROM, strTO
	Dim strJOBNO, strJOBNAME
	Dim strOUTSCODE, strOUTSNAME
	Dim strCONTRACTNO, strCONTRACTNAME
	Dim strENDGBN, strCONFIRM
	Dim strDataCHK
	Dim lngCol, lngRow
	Dim strSAVEFLAG , strAMTFLAG ,strCOMENT
	
	Dim strJOBGUBN,strCREPART
	
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		
		'�ʱⰪ ����
		If .cmbENDGBN.value  = "F" Then
			strSAVEFLAG = "F"
		Elseif .cmbENDGBN.value = "T" Then
			strSAVEFLAG = "T"
		End If
		
		If strSAVEFLAG = "F" Then '�̿Ϸ� ����
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, " OUTSCODE | OUTSNAME | REGDATE | JOBNAME | ADJAMT | CONTRACTGUBUN ",lngCol, lngRow, False) 

			If strDataCHK = False Then
				gErrorMsgBox lngRow & " ���� ����ó/�����/����/�ݾ�/��౸�� �� �ʼ� �Է»����Դϴ�.","����ȳ�"
				Exit Sub		 
			End If
		ELSE
			if DataValidation =false then exit sub
		END IF 
		
		'�������� �����ÿ� ��������ó�� �������� �ƴ��� �Ǵ� VALIDATION
		If strSAVEFLAG = "F" Then '�̿Ϸ� ����
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONTRACTNO | OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT | JOBGUBN | CREPART | CONTRACTGUBUN | RANKTRANS | SEQ | AMTFLAG")
		Elseif strSAVEFLAG = "T" then '�Ϸ� ����
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONTRACTNO | CONTRACTNAME | CONTRACTDAY | LOCALAREA | STDATE | EDDATE | AMT | DELIVERYDAY | TESTDAY | PAYMENTGBN | TESTMENT | COMENT | CONFIRMFLAG | CONFLAG | DIVFLAG | OUTSCODE | PRERATE | PREAMT | ENDRATE | ENDAMT | THISRATE | THISAMT | BALANCERATE | BALANCEAMT | DELIVERYGUARANTY | FAULTGUARANTY | MANAGER | TESTENDDAY | TESTAMT | LOSTDAY | AMTFLAG")
		End If
		
		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		strCOMENT = .txtCOMENT.value
		
		IF .txtJOBGUBN.value = "" AND .txtCREPART.value = "" THEN 
			strJOBGUBN = "J"
			strCREPART = "J"	
		ELSE
			strJOBGUBN = .txtJOBGUBN.value
			strCREPART = .txtCREPART.value 
		END IF
		
		intRtn = mobjPDCMCONTRACT.ProcessRtn(gstrConfigXml, strMasterData, vntData, strSAVEFLAG, strCOMENT,strJOBGUBN,strCREPART)
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			
			strFROM = .txtFROM.value
			strTO = .txtTo.value
			strCONTRACTNO = .txtCONTRACTNO.value 
			strCONFIRM = .cmbCONFIRM.value
			strOUTSCODE = TRIM(.txtOUTSCODE.value)
			strOUTSNAME =  TRIM(.txtOUTSNAME.value)
			strJOBNO = TRIM(.txtJOBNO.value)
			strJOBNAME =  TRIM(.txtJOBNAME.value)
			'���� �˻�� ��´�.
			PreSearchFiledValue strFROM, strTO, strCONTRACTNO, strENDGBN, strCONFIRM, strOUTSCODE, strOUTSNAME, strJOBNO,  strJOBNAME
			.cmbENDGBN.value  = strSAVEFLAG
			
			SelectRtn
		End If
	End with
End Sub

'------------------------------------------
' ��༭ ����
'------------------------------------------
Sub ProcessRtn_CONTRACT
	Dim intRtn
	Dim strMasterData
	Dim vntData
	Dim intCnt
	Dim lngCnt
	Dim strFROM, strTO
	Dim strJOBNAME, strJOBNO
	Dim strOUTSCODE, strOUTSNAME
	Dim strCONTRACTNO, strCONTRACTNAME,strCONTRACTCODE
	Dim strENDGBN, strCONFIRM
	Dim strSAVEFLAG , strAMTFLAG ,strCOMENT
	
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		'�ʱⰪ ����
		If .cmbENDGBN.value  <> "F" Then
			gErrorMsgBox "��༭������ �̿Ϸ� ��ȸ���¿��� �����մϴ�.","��༭�����ȳ�"
			Exit Sub
		End If
		
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","��༭�����ȳ�"
			Exit Sub
		end if
		
		lngCnt = TRUE
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", intCnt) = "1"  Then
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) = "" THEN
					gErrorMsgBox intCnt & " ��° �࿡ ����ó �ڵ带 Ȯ���ϼ���","��༭�����ȳ�"
					Exit Sub
				END IF
				strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
				lngCnt = FALSE
			End if
		Next
		
		If lngCnt Then
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�����ȳ�"
			Exit Sub
		End If
		
		if DataValidation =false then exit sub
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | CONTRACTNO | OUTSCODE | OUTSNAME | REGDATE | JOBNO | JOBNAME | ADJAMT | JOBGUBN | CREPART | CONTRACTGUBUN | RANKTRANS | SEQ | AMTFLAG")
		
		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If

		strCONTRACTNAME = .txtCONTRACTNAME.value 
		
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" Then
				strAMTFLAG =  mobjSCGLSpr.GetTextBinding(.sprSht,"AMTFLAG",intCnt)
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTGUBUN",intCnt) = "" THEN
						gErrorMsgBox intCnt & " ��° �࿡ ��༭ ������ Ȯ���ϼ���","��༭�����ȳ�"
						EXIT SUB
					ELSE
						strCONTRACTCODE =  MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTGUBUN",intCnt),1,1)
					END IF
				EXIT FOR
			End If
		Next

		strCOMENT = .txtCOMENT.value
		
		intRtn = mobjPDCMCONTRACT.ProcessRtn_CONTRACT(gstrConfigXml, strMasterData, vntData, strOUTSCODE, strCONTRACTNAME,strCONTRACTCODE, strCOMENT, strAMTFLAG )
		
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			
			strFROM = .txtFROM.value
			strTO = .txtTo.value
			strCONTRACTNO = .txtCONTRACTNO.value 
			strCONFIRM = .cmbCONFIRM.value
			strOUTSCODE = TRIM(.txtOUTSCODE.value)
			strOUTSNAME =  TRIM(.txtOUTSNAME.value)
			strJOBNO = TRIM(.txtJOBNO.value)
			strJOBNAME =  TRIM(.txtJOBNAME.value)
			strSAVEFLAG = .cmbENDGBN.value
			'���� �˻�� ��´�.
			PreSearchFiledValue strFROM, strTO, strCONTRACTNO, strENDGBN, strCONFIRM, strOUTSCODE, strOUTSNAME, strJOBNO,  strJOBNAME
			.cmbENDGBN.value  = strSAVEFLAG
			
			SelectRtn
		End If
	End with
End Sub

'------------------------------------------
'���� ����ó���� �������� �Ǵ��ϱ�����
'-----------------------------------------
Function DataValidation ()
	DataValidation = false
   	Dim intCnt
   	Dim strOUTSCODE
   	Dim lngCnt
   	Dim strCNT
   	
	'On error resume next
	with frmThis
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� 
   		IF not gDataValidation(frmThis) then exit Function
   		
		IF .cmbENDGBN.value = "F" THEN
			strOUTSCODE = ""
			strCNT = TRUE
   			for intCnt = 1 to .sprSht.MaxRows
   				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)  = "1" Then
   					IF strCNT THEN
   						strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt)
   						strCNT = FALSE
   					END IF   				
	   				
					if strOUTSCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) Then
						gErrorMsgBox intCnt & " ��° ���� ����ó�� Ȯ���Ͻʽÿ�." & vbcrlf & "���Ͽ���ó �ϰ�쿡�� ������ �����մϴ�.","�Է¿���"
						Exit Function
					End If
				End If
			next
		END IF
   	End with
	DataValidation = true
End Function

'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCONTRACTNO
	Dim lngCnt
	Dim lngSumCnt
	Dim strSEQ
	Dim strENDFLAG
	
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","��༭�����ȳ�"
			Exit Sub
		end if
		
		IF .cmbENDGBN.value = "F" THEN
			For i = 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", i) = "1"  Then
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",i) <> "" THEN
						gErrorMsgBox i & " ��° ���� �������꿡�� �Էµ� �������Դϴ�. �̿Ϸ� ������ �ű��Է��� �����͸� ���� �����մϴ�.","�����ȳ�"
						Exit Sub
					END IF
				End if
			Next
			
			intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
			IF intRtn <> vbYes then exit Sub
			
			for i = .sprSht.MaxRows to 1 step -1
				strENDFLAG = ""
				strENDFLAG = .cmbENDGBN.value
				IF strENDFLAG = "" THEN
					EXIT SUB
				END IF 
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
					If mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",i) = "" Then
						strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
						if strSEQ = "" then
							mobjSCGLSpr.DeleteRow .sprSht,i
						else
							
							intRtn = mobjPDCMCONTRACT.DeleteRtn(gstrConfigXml, strENDFLAG, strSEQ)
								   			
   							IF not gDoErrorRtn ("DeleteRtn") then
								mobjSCGLSpr.DeleteRow .sprSht,i
   							End IF
						end if 
					End IF
   				End If
			next
		ELSE
			For i = 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK", i) = "1"  Then
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "Y" Then
						gErrorMsgBox "���ε� ����� ������ �� �����ϴ�.","�����ȳ�"
						Exit Sub
					End if
				End if
			Next
			
			intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
			IF intRtn <> vbYes then exit Sub
			
			for i = .sprSht.MaxRows to 1 step -1
				strENDFLAG = ""
				strENDFLAG = .cmbENDGBN.value
				IF strENDFLAG = "" THEN
					EXIT SUB
				END IF
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
						strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
						intRtn = mobjPDCMCONTRACT.DeleteRtn(gstrConfigXml, strENDFLAG , strSEQ)
					End IF
					IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
   				End If
			next
		END IF
		
		gWriteText lblstatus, "�ڷᰡ " & intRtn & " �� �����Ǿ����ϴ�."
		'InitPageData
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
			<TABLE style="WIDTH: 100%" id="tblForm" border="0" cellSpacing="0" cellPadding="0" height="100%">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								height="28">
								<TR>
									<TD style="WIDTH: 400px" height="28" width="400" align="left">
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
											<tr>
												<td class="TITLE">������&nbsp;</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" height="28" vAlign="middle" align="right">
										<!--Wait Button Start-->
										<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 302px"
											id="tblWaitP" border="0" cellSpacing="1" cellPadding="1" width="75%">
											<TR>
												<TD style="Z-INDEX: 200" id="tblWait"><IMG style="CURSOR: wait" id="imgWaiting" border="0" name="imgWaiting" alt="ó�����Դϴ�."
														src="../../../images/Waiting.GIF" height="23">
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
								<TR>
									<TD height="1" width="100%" align="left"></TD>
								</TR>
							</TABLE>
							<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								height="13">
								<TR>
									<TD style="WIDTH: 1040px" class="TOPSPLIT"></TD>
								</TR>
							</TABLE>
							<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 60px; CURSOR: hand" class="SEARCHDATA"><SELECT style="WIDTH: 90px" id="cmbYEARMONGBN" name="cmbYEARMONGBN">
											<OPTION selected value="REGDATE">�����</OPTION>
											<OPTION value="CONTRACTDAY">�����</OPTION>
										</SELECT></TD>
									<TD style="WIDTH: 246px; HEIGHT: 24px" class="SEARCHDATA"><INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtFrom" class="INPUT" title="���˻� ��������"
											maxLength="10" size="9" name="txtFrom"> <IMG style="CURSOR: hand" id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgFrom" align="absMiddle" src="../../../images/btnCalEndar.gIF"
											height="15">&nbsp; ~&nbsp; <INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtTo" class="INPUT" title="���˻� ��������"
											maxLength="10" size="9" name="txtTo"> <IMG style="CURSOR: hand" id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgTo" align="absMiddle" src="../../../images/btnCalEndar.gIF"
											height="15">
									</TD>
									<TD style="WIDTH: 60px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL">�Ϸᱸ��</TD>
									<TD style="WIDTH: 68px; HEIGHT: 24px; CURSOR: hand" class="SEARCHDATA"><SELECT style="WIDTH: 110px" id="cmbENDGBN" name="cmbENDGBN">
											<OPTION selected value="F">�̿Ϸ�</OPTION>
											<OPTION value="T">�Ϸ�</OPTION>
										</SELECT></TD>
									<TD style="WIDTH: 61px; CURSOR: hand" class="SEARCHLABEL">�ϵ��ޱ���</TD>
									<TD style="WIDTH: 84px; CURSOR: hand" class="SEARCHDATA"><SELECT style="WIDTH: 100px" id="cmbAMTFLAG" name="cmbAMTFLAG">
											<OPTION selected value="1">�ϵ���</OPTION>
											<OPTION value="0">���ϵ���</OPTION>
										</SELECT></TD>
									<TD style="WIDTH: 46px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)">����ó</TD>
									<TD style="HEIGHT: 24px" class="SEARCHDATA"><INPUT style="WIDTH: 160px; HEIGHT: 22px" id="txtOUTSNAME" class="INPUT_L" title="����ó�� ��ȸ"
											maxLength="255" align="left" size="32" name="txtOUTSNAME"> <IMG style="CURSOR: hand" id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
										<INPUT style="WIDTH: 65px; HEIGHT: 22px" id="txtOUTSCODE" class="INPUT" title="����ó�ڵ���ȸ"
											maxLength="6" align="left" size="3" name="txtOUTSCODE"></TD>
									<td style="HEIGHT: 24px" class="SEARCHDATA" width="50"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ �˻��մϴ�." align="right" src="../../../images/imgQuery.gIF"
											height="20"></td>
								</TR>
								<TR>
									<TD style="WIDTH: 60px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')"
										width="60">��༭��ȣ</TD>
									<TD style="WIDTH: 246px" class="SEARCHDATA"><INPUT style="WIDTH: 240px; HEIGHT: 22px" id="txtCONTRACTNO" class="INPUT_L" title="��� �ڵ���ȸ"
											maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD style="WIDTH: 60px; CURSOR: hand" class="SEARCHLABEL">����ڵ�
									</TD>
									<TD style="WIDTH: 68px; CURSOR: hand" vAlign="top">
										<DIV id="pnlFLAG1" style="POSITION: absolute; WIDTH: 110px; VISIBILITY: hidden" ms_positioning="GridLayout">
											<SELECT style="WIDTH: 110px" id="cmbCONTRACTCODE" name="cmbCONTRACTCODE">
												<OPTION selected value="">��ü</OPTION>
												<OPTION value="B">B - �׷���-�纸</OPTION>
												<OPTION value="C">C - TV-��������</OPTION>
												<OPTION value="D">D - �귣��-����</OPTION>
												<OPTION value="G">G - �μ�</OPTION>
												<OPTION value="I">I - ���ͳ�</OPTION>
												<OPTION value="J">J - ���۱�</OPTION>
												<OPTION value="M">M - �𵨷�</OPTION>
												<OPTION value="O">O - ��Ÿ</OPTION>
												<OPTION value="P">P - PR</OPTION>
												<OPTION value="R">R - Radio</OPTION>
												<OPTION value="S">S - ���θ��</OPTION>
											</SELECT>
										</DIV>
										<DIV id="pnlFLAG2" style="POSITION: absolute; WIDTH: 110px; VISIBILITY: hidden" ms_positioning="GridLayout">
											<SELECT style="WIDTH: 110px" id="cmbJOBNOCODE" name="cmbJOBNOCODE">
												<OPTION selected value="">��ü</OPTION>
												<OPTION value="I">I
												</OPTION>
												<OPTION value="G">G
												</OPTION>
												<OPTION value="C">C
												</OPTION>
												<OPTION value="R">R
												</OPTION>
												<OPTION value="S">S
												</OPTION>
												<OPTION value="M">M
												</OPTION>
												<OPTION value="D">D
												</OPTION>
												<OPTION value="B">B
												</OPTION>
												<OPTION value="P">P
												</OPTION>
												<OPTION value="O">O
												</OPTION>
											</SELECT>
										</DIV>
									</TD>
									<TD style="WIDTH: 60px; CURSOR: hand" class="SEARCHLABEL">��༭Ȯ��</TD>
									<TD style="WIDTH: 90px; CURSOR: hand" class="SEARCHDATA"><SELECT style="WIDTH: 100px" id="cmbCONFIRM" name="cmbCONFIRM">
											<OPTION selected value="">��ü</OPTION>
											<OPTION value="0">��༭ ��Ȯ��</OPTION>
											<OPTION value="1">��༭ Ȯ��</OPTION>
										</SELECT></TD>
									<TD style="WIDTH: 46px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtJOBNAME, txtJOBNO)">JOB��</TD>
									<TD class="SEARCHDATA" colSpan="2"><INPUT style="WIDTH: 160px; HEIGHT: 22px" id="txtJOBNAME" class="INPUT_L" title="JOB�� ��ȸ"
											maxLength="255" align="left" size="32" name="txtJOBNAME"> <IMG style="CURSOR: hand" id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
										<INPUT style="WIDTH: 65px; HEIGHT: 22px" id="txtJOBNO" class="INPUT" title="JOBNO ��ȸ" maxLength="7"
											align="left" size="3" name="txtJOBNO"></TD>
								</TR>
							</TABLE>
							<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								height="13">
								<TR>
									<TD style="WIDTH: 1040px; HEIGHT: 25px" class="TOPSPLIT"></TD>
								</TR>
							</TABLE>
							<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								height="28"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD height="20" width="350" align="left">
										<table id="TABLE1" border="0" cellSpacing="0" cellPadding="0" width="100%" runat="server">
											<tr>
												<td align="left">
													<TABLE border="0" cellSpacing="0" cellPadding="0" width="145" background="../../../images/back_p.gIF">
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
												<td class="TITLE">����뿪 �� ����&nbsp;��༭&nbsp;&nbsp;&nbsp;<span style="CURSOR: hand" onclick="vbscript:Call Set_TBL_HIDDEN ()">(�˻������߰��Է»���)</span></td>
											</tr>
										</table>
									</TD>
									<td>
										<table border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
											<tr>
												<td class="TITLE" height="20" vAlign="middle" align="left">&nbsp;&nbsp;&nbsp;&nbsp;�հ� 
													: <INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 22px" id="txtSUMAMT" class="NOINPUTB_R"
														title="�հ�ݾ�" readOnly maxLength="100" size="13" name="txtSUMAMT"> <INPUT style="WIDTH: 120px; HEIGHT: 22px" id="txtSELECTAMT" class="NOINPUTB_R" title="���ñݾ�"
														readOnly maxLength="100" size="16" name="txtSELECTAMT">
												</td>
											</tr>
										</table>
									</td>
									<TD height="20" vAlign="middle" align="right">
										<!--Common Button Start-->
										<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
											<TR>
												<TD><IMG style="CURSOR: hand" id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" border="0" name="imgREG"
														alt="�ű��ڷḦ �����մϴ�." src="../../../images/imgNew.gIF"></TD>
												<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgSave"
														alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" height="20"></TD>
												<TD width="20"></TD>
												<TD><IMG style="CURSOR: hand" id="imgContractCre" onmouseover="JavaScript:this.src='../../../images/imgContractCreOn.gif'"
														onmouseout="JavaScript:this.src='../../../images/imgContractCre.gif'" border="0"
														name="imgContractCre" alt="��༭�� �����մϴ�." src="../../../images/imgContractCre.gIF"
														height="20"></TD>
												<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
														alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" height="20"></TD>
												<TD><IMG style="CURSOR: hand" id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
														onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" border="0" name="imgPrint"
														alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" height="20"></TD>
												<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
														onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
														alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE style="WIDTH: 100%" id="tblBody" border="0" cellSpacing="0" cellPadding="0">
								<TR>
									<TD style="WIDTH: 100%; HEIGHT: 11px" class="TOPSPLIT"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD vAlign="middle" align="center">
										<TABLE style="WIDTH: 100%; HEIGHT: 6px" id="tblDATA" class="SEARCHDATA" border="0" cellSpacing="1"
											cellPadding="0" align="left">
											<TR>
												<TD style="WIDTH: 53px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONTRACTNAME, '')">����</TD>
												<TD style="WIDTH: 291px" class="SEARCHDATA"><INPUT accessKey=",M" style="WIDTH: 240px; HEIGHT: 21px" id="txtCONTRACTNAME" dataSrc="#xmlBind"
														class="INPUT_L" title="����" dataFld="CONTRACTNAME" size="30" name="txtCONTRACTNAME"></TD>
												<TD style="WIDTH: 85px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtMANAGER,'')">�������</TD>
												<TD style="WIDTH: 175px" class="SEARCHDATA"><INPUT style="WIDTH: 170px; HEIGHT: 22px" id="txtMANAGER" dataSrc="#xmlBind" class="INPUT_L"
														title="�����" dataFld="MANAGER" maxLength="255" align="left" size="36" name="txtMANAGER"></TD>
												<TD style="WIDTH: 100px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtAMT,'')">���ݾ�</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="NUM,M" style="WIDTH: 150px; HEIGHT: 22px" id="txtAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="���ݾ�" dataFld="AMT" maxLength="100" size="19" name="txtAMT"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtJOBGUBN" dataSrc="#xmlBind" dataFld="JOBGUBN"
														type="hidden" name="txtJOBGUBN"><INPUT style="WIDTH: 8px; HEIGHT: 21px" id="txtCREPART" dataSrc="#xmlBind" dataFld="CREPART"
														size="1" type="hidden" name="txtCREPART"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 53px; HEIGHT: 23px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCONTRACTDAY, '')">�����</TD>
												<TD style="WIDTH: 291px; HEIGHT: 23px" class="SEARCHDATA"><INPUT accessKey="DATE,M" style="WIDTH: 88px; HEIGHT: 22px" id="txtCONTRACTDAY" dataSrc="#xmlBind"
														class="INPUT" title="�����" dataFld="CONTRACTDAY" maxLength="10" size="9" name="txtCONTRACTDAY">
													<IMG style="CURSOR: hand" id="Img1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="ImgCONTRACTDAY"
														alt="ImgCONTRACTDAY" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<INPUT style="WIDTH: 8px; HEIGHT: 22px" id="txtOUTSCODE1" dataSrc="#xmlBind" class="INPUT"
														title="����ó�ڵ���ȸ" dataFld="OUTSCODE1" maxLength="6" align="left" size="1" name="txtOUTSCODE1">
													&nbsp;&nbsp;&nbsp;&nbsp; &nbsp; &nbsp;��༭����<INPUT id="chkCONFIRMFLAG" dataSrc="#xmlBind" title="VAT����" dataFld="CONFIRMFLAG" value=""
														type="checkbox" name="chkCONFIRMFLAG">
												</TD>
												<TD style="HEIGHT: 23px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtPRERATE, '')">����������</TD>
												<TD style="WIDTH: 175px; HEIGHT: 23px" class="SEARCHDATA"><INPUT accessKey="NUM,M" style="WIDTH: 150px; HEIGHT: 22px" id="txtPRERATE" dataSrc="#xmlBind"
														class="INPUT_R" title="����������" dataFld="PRERATE" maxLength="100" size="33" name="txtPRERATE">(%)</TD>
												<TD style="HEIGHT: 23px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtPREAMT, '')">���ޱ�</TD>
												<TD style="HEIGHT: 23px" class="SEARCHDATA"><INPUT accessKey="NUM,M" style="WIDTH: 150px; HEIGHT: 22px" id="txtPREAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="���ޱ�" dataFld="PREAMT" maxLength="100" size="36" name="txtPREAMT"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 53px; CURSOR: hand" class="SEARCHLABEL">���Ⱓ</TD>
												<TD style="WIDTH: 291px" class="SEARCHDATA"><INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtSTDATE" dataSrc="#xmlBind"
														class="INPUT" title="���Ⱓ ������" dataFld="STDATE" maxLength="10" size="9" name="txtSTDATE">
													<IMG style="CURSOR: hand" id="imgFROM2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgFROM2"
														align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">&nbsp;~ <INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtEDDATE" dataSrc="#xmlBind"
														class="INPUT" title="���Ⱓ ������" dataFld="EDDATE" maxLength="10" size="9" name="txtEDDATE">
													<IMG style="CURSOR: hand" id="imgTO2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgTO2"
														align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">
												</TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtLOCALAREA, '')">�������</TD>
												<TD style="WIDTH: 175px" class="SEARCHDATA"><INPUT style="WIDTH: 170px; HEIGHT: 22px" id="txtLOCALAREA" dataSrc="#xmlBind" class="INPUT_L"
														title="�������" dataFld="LOCALAREA" maxLength="255" align="left" size="36" name="txtLOCALAREA"></TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtDELIVERYGUARANTY, '')">������� 
													������</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="NUM,M" style="WIDTH: 150px; HEIGHT: 22px" id="txtDELIVERYGUARANTY" dataSrc="#xmlBind"
														class="INPUT_R" title="������� ������" dataFld="DELIVERYGUARANTY" maxLength="100" size="36" name="txtDELIVERYGUARANTY"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 53px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtDELIVERYDAY, '')">��ǰ��</TD>
												<TD style="WIDTH: 291px" class="SEARCHDATA"><INPUT accessKey="DATE,M" style="WIDTH: 78px; HEIGHT: 22px" id="txtDELIVERYDAY" dataSrc="#xmlBind"
														class="INPUT" title="��ǰ��,�Ϸ����" dataFld="DELIVERYDAY" maxLength="10" size="9" name="txtDELIVERYDAY">
													<IMG style="CURSOR: hand" id="imgDELIVERYDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgDELIVERYDAY"
														align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15"> ��༭<INPUT id="chkCONFLAG" dataSrc="#xmlBind" title="��༭" dataFld="CONFLAG" value="" type="checkbox"
														name="chkCONFLAG"> ���꼭<INPUT id="chkDIVFLAG" dataSrc="#xmlBind" title="���꼭" dataFld="DIVFLAG" value="" CHECKED
														type="checkbox" name="chkDIVFLAG"> ���༭<INPUT id="chkEXEFLAG" dataSrc="#xmlBind" title="���༭" dataFld="EXEFLAG" value="" type="checkbox"
														name="chkEXEFLAG">
												</TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtPAYMENTGBN, '')">������޹��</TD>
												<TD style="WIDTH: 175px" class="SEARCHDATA"><INPUT style="WIDTH: 170px; HEIGHT: 22px" id="txtPAYMENTGBN" dataSrc="#xmlBind" class="INPUT_L"
														title="������޹��" dataFld="PAYMENTGBN" maxLength="255" size="32" name="txtPAYMENTGBN"></TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtFAULTGUARANTY,'')">���ں��� 
													������</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="NUM,M" style="WIDTH: 150px; HEIGHT: 22px" id="txtFAULTGUARANTY" dataSrc="#xmlBind"
														class="INPUT_R" title="���ں��� ������" dataFld="FAULTGUARANTY" maxLength="100" size="36" name="txtFAULTGUARANTY"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 53px" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCOMENT, '')">Ư�����</TD>
												<TD class="SEARCHDATA" colSpan="7"><TEXTAREA style="WIDTH: 778px" id="txtCOMENT" dataSrc="#xmlBind" dataFld="COMENT" wrap="hard"
														cols="10" name="txtCOMENT"></TEXTAREA></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody3" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								height="28"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD height="20" align="left">
										<table border="0" cellSpacing="0" cellPadding="0" width="100%">
											<tr>
												<td align="left">
													<TABLE border="0" cellSpacing="0" cellPadding="0" width="143" background="../../../images/back_p.gIF">
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
												<td class="TITLE">�˻�����(��ǰ/�뿪/����)</td>
											</tr>
										</table>
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody4" border="0" cellSpacing="0" cellPadding="0" width="100%">
								<TR>
									<TD style="WIDTH: 100%" class="TOPSPLIT"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD vAlign="middle" align="center">
										<TABLE style="WIDTH: 100%; HEIGHT: 6px" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0"
											align="left">
											<TR>
												<TD style="WIDTH: 52px; HEIGHT: 25px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(TESTDAY, txtTESTENDDAY)"
													width="52">�˻�Ⱓ</TD>
												<TD style="WIDTH: 250px" class="SEARCHDATA"><INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtTESTDAY" dataSrc="#xmlBind"
														class="INPUT" title="�˻�Ⱓ ������" dataFld="TESTDAY" maxLength="10" size="9" name="txtTESTDAY">
													<IMG style="CURSOR: hand" id="ImgFROM3" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="ImgFROM3"
														align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15">&nbsp;~ <INPUT accessKey="DATE" style="WIDTH: 88px; HEIGHT: 22px" id="txtTESTENDDAY" dataSrc="#xmlBind"
														class="INPUT" title="�˻�Ⱓ ������" dataFld="TESTENDDAY" maxLength="10" size="9" name="txtTESTENDDAY">
													<IMG style="CURSOR: hand" id="ImgTO3" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="ImgTO3"
														align="absMiddle" src="../../../images/btnCalEndar.gIF" height="15"></TD>
												<TD style="WIDTH: 87px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtENDRATE,'')"
													width="87">�����ޱ���</TD>
												<TD style="WIDTH: 177px" class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtENDRATE" dataSrc="#xmlBind"
														class="INPUT_R" title="�����ޱ���" dataFld="ENDRATE" maxLength="100" size="36" name="txtENDRATE">(%)</TD>
												<TD style="WIDTH: 100px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtENDAMT,'')">�����ޱ�</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtENDAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="�����ޱ�" dataFld="ENDAMT" maxLength="100" size="36" name="txtENDAMT"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 52px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTESTAMT, '')">�˻�ݾ�</TD>
												<TD style="WIDTH: 250px" class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 240px; HEIGHT: 21px" id="txtTESTAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="�˻�ݾ�" dataFld="TESTAMT" maxLength="100" size="36" name="txtTESTAMT"></TD>
												<TD style="WIDTH: 87px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTHISRATE, '')">��ȸ������</TD>
												<TD style="WIDTH: 177px" class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtTHISRATE" dataSrc="#xmlBind"
														class="INPUT_R" title="��ȸ������" dataFld="THISRATE" maxLength="100" size="36" name="txtTHISRATE">(%)</TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTHISAMT, '')">��ȸ����</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtTHISAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="��ȸ����" dataFld="THISAMT" maxLength="100" size="36" name="txtTHISAMT"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 52px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtLOSTDAY, '')">��ü�ϼ�</TD>
												<TD style="WIDTH: 250px" class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 226px; HEIGHT: 21px" id="txtLOSTDAY" dataSrc="#xmlBind"
														class="INPUT_R" title="��ü�ϼ�" dataFld="LOSTDAY" maxLength="100" size="37" name="txtLOSTDAY">��
												</TD>
												<TD style="WIDTH: 87px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtBALANCERATE, '')">�ܱ���</TD>
												<TD style="WIDTH: 177px" class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtBALANCERATE" dataSrc="#xmlBind"
														class="INPUT_R" title="�ܱ���" dataFld="BALANCERATE" maxLength="100" size="36" name="txtBALANCERATE">(%)</TD>
												<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtBALANCEAMT, '')">�ܱ�</TD>
												<TD class="SEARCHDATA"><INPUT accessKey="M,NUM" style="WIDTH: 150px; HEIGHT: 22px" id="txtBALANCEAMT" dataSrc="#xmlBind"
														class="INPUT_R" title="�ܱ�" dataFld="BALANCEAMT" maxLength="100" size="36" name="txtBALANCEAMT"></TD>
											</TR>
											<TR>
												<TD style="WIDTH: 52px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTESTMENT, '')">�˻��ǰ�</TD>
												<TD class="SEARCHDATA" colSpan="5"><INPUT style="WIDTH: 778px; HEIGHT: 21px" id="txtTESTMENT" dataSrc="#xmlBind" class="INPUT_L"
														title="�˻��ǰ�" dataFld="TESTMENT" maxLength="255" size="124" name="txtTESTMENT"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD style="WIDTH: 1040px" class="BODYSPLIT"></TD>
					</TR>
					<tr>
						<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
							<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
								ms_positioning="GridLayout">
								<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="31802">
									<PARAM NAME="_ExtentY" VALUE="2883">
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
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD style="WIDTH: 1040px" id="lblStatus" class="BOTTOMSPLIT"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
