<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCONTRACT_BASECONF.aspx.vb" Inherits="PD.PDCMCONTRACT_BASECONF" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�⺻ & �ܰ� ��༭ ü�� ��� �� ��ȸ</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : �⺻��༭ ü�� ��� �� ��ȸ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : �⺻ & �ܰ� ��༭ ü�� ��� �� ��ȸ/����/����
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
Dim mobjPDCMCONTRACT_BASE, mobjPDCMGET
Dim mstrCheck
Dim mstrmode
Dim mstrCHKcheck
Dim mstrCONFIRM

CONST meTAB = 9

mstrCheck = True
mstrmode = True
mstrCHKcheck = True
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
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

sub imgAgree_onclick ()
	mstrCONFIRM = "1"
	gFlowWait meWAIT_ON
	UpdateRtn_CONFIRM(mstrCONFIRM)
	gFlowWait meWAIT_OFF
end sub

sub imgAgreeCanCel_onclick ()
	mstrCONFIRM = "0"
	gFlowWait meWAIT_ON
	UpdateRtn_CONFIRM(mstrCONFIRM)
	gFlowWait meWAIT_OFF
end sub

'��༭ ���� �̺�Ʈ
Sub imgSave_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub

'������ư �̺�Ʈ
Sub imgDelete_onclick
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

		'üũ�� �� �����Ͱ� �ִ��� ������ üũ�Ѵ�.
		intCount = 0
		for i=1 to frmThis.sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CHK",i) = "1" THEN
				if mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"CONTRACTNO",i) = "" then
					gErrorMsgBox "��༭�� �������� �ʾҽ��ϴ�. �������� ����ϼ���"," ��༭ ��� �ȳ�!"
					Exit Sub	
				end if
				intCount = 1
			end if
		next

		'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
		if intCount = 0 then
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�. �μ��� �����͸� üũ�Ͻÿ�",""
			Exit Sub
		End if

		gFlowWait meWAIT_ON
		with frmThis
			'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn_TEMP(gstrConfigXml)

			ModuleDir = "PD"
			
			IF .rdCONFLAG.checked THEN 
				'�⺻ ��༭
				if .cmbGUBUN.value = "�⺻" then
					ReportName = "PDCMCONTRACTBASE_CON_N.rpt"			
				'�ܰ� ��༭
				elseif .cmbGUBUN.value = "�ܰ�" then
					ReportName = "PDCMCONTRACTPRICE_CON_N.rpt"			
				end if
			End if 

			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strCONTRACTNO	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					strUSERID = ""
					vntDataTemp = mobjPDCMCONTRACT_BASE.ProcessRtn_TEMP(gstrConfigXml,strCONTRACTNO, i, strUSERID)
				END IF
			next

			Params = strUSERID 
			Opt = "A"
			
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		End with
		gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn_TEMP(gstrConfigXml)
	end with
End sub

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
     	End if
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

Sub imgCONTRACTDAY_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar .txtCONTRACTDAY,.imgCONTRACTDAY,"txtCONTRACTDAY_onchange()"
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
	IF frmthis.rdCONFLAG.checked then
		frmThis.txtCONTRACTDAY.value = frmThis.txtSTDATE.value 	
	end if 
	gSetChange
End Sub

Sub txtCONTRACTNAME_onchange
	mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTNAME",frmThis.sprSht.ActiveRow, frmthis.txtCONTRACTNAME.value
	gSetChange
End Sub

Sub txtCONTRACTDAY_onchange
	frmthis.txtSTDATE.value = frmthis.txtCONTRACTDAY.value 
	gSetChange
End Sub

Sub txtDELIVERYDAY_onchange
	gSetChange
End Sub

Sub chkCONFIRMFLAG_onClick
	if frmThis.sprSht.ActiveRow > 0  Then
		if frmThis.chkCONFIRMFLAG.checked = TRUE Then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "1"
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONFIRMFLAG",frmThis.sprSht.ActiveRow, "0"
		End if
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	end if
	gSetChange
End Sub

Sub cmbGBN_onchange
	Dim strHTML
	with frmThis
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN",frmThis.sprSht.ActiveRow, frmthis.cmbGBN.value
		
		if .cmbGBN.value = "BTL" then
			document.getElementById("test").innerHTML = ""
			.cmbGUBUN.value = "�⺻"
			.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
			cmbGUBUN_onchange
		else
			document.getElementById("test").innerHTML = "�ܰ�"
			.cmbGUBUN.className				= "INPUT"   : .cmbGUBUN.disabled			= false 
			cmbGUBUN_onchange
		end if
		 
		
	end with
	gSetChange
End Sub

Sub cmbGUBUN_onchange
	with frmThis
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, frmthis.cmbGUBUN.value
		'�⺻��༭�� ��� ���Ⱓ�� ������ �ʴ´�.
		if frmThis.cmbGUBUN.value = "�⺻" then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, ""
			.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
			.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
		
			frmThis.txtSTDATE.value = ""
			frmThis.txtEDDATE.value = ""
		else
			frmThis.txtSTDATE.value		= gNowDate
			DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
			.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
			.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
		end if
		
	end with
	gSetChange
End Sub

Sub txtCOMENT_onchange
	mobjSCGLSpr.SetTextBinding frmThis.sprSht,"COMENT",frmThis.sprSht.ActiveRow, frmthis.txtCOMENT.value
	gSetChange
End Sub


Sub cmbGUBUN1_onchange
	with frmThis
		SelectRtn
	end with
end sub

'-----------------------------------
'�ʵ��߰� 
'------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		frmThis.sprSht.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)		
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CHK",frmThis.sprSht.ActiveRow, 1
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTNO",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CONTRACTDAY",frmThis.sprSht.ActiveRow, frmThis.txtCONTRACTDAY.value
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GUBUN",frmThis.sprSht.ActiveRow, "�⺻"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GBN",frmThis.sprSht.ActiveRow, "ATL"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"AMT",frmThis.sprSht.ActiveRow, "0"
		
		mobjSCGLSpr.SetCellsLock2 frmThis.sprSht, false, "OUTSCODE | BTN  | OUTSNAME"
		
		sprShtToFieldBinding 2, 1
		
		if frmThis.cmbGUBUN.value = "�⺻" then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, ""
		
			frmThis.txtSTDATE.value = ""
			frmThis.txtEDDATE.value = ""
		else
			frmThis.txtSTDATE.value		= gNowDate
			DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"STDATE",frmThis.sprSht.ActiveRow, frmThis.txtSTDATE.value
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EDDATE",frmThis.sprSht.ActiveRow, frmThis.txtEDDATE.value
			
			frmThis.sprSht.focus()
		end if	
	End If
End Sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
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
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intcnt
			next
			
		ELSE
			if Row > 0 then
				sprShtToFieldBinding Col, Row
			end if
   			
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
		
		.sprSht.Focus
	End With
End Sub

Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet
	Dim vntInParams
	Dim dblAMT
	
	with frmThis
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") Then			
		
				vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row)))
				
				vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
				If isArray(vntRet) Then
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)		
					mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
					
					mobjSCGLSpr.CellChanged .sprSht, Col,Row
					mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				End If
			End if
			
			if Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CONFIRMFLAG") then
				if mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "1" then
					.chkCONFIRMFLAG.checked = true
				else
					.chkCONFIRMFLAG.checked = false
				end if
			end if 
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
				sprShtToFieldBinding .sprSht.ActiveCol, .sprSht.ActiveRow
		End If
		
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")  Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""

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

			'.txtSELECTAMT.value = strSUM
			'Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			'.txtSELECTAMT.value = 0
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
						'.txtSELECTAMT.value = 0
						exit Sub
					End If
					
					FOR j = 0 TO intSelCnt1 -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(.sprSht,strCol,vntData_row(j))
						End If
					Next
					'.txtSELECTAMT.value = strSUM
				End If
			else
				'.txtSELECTAMT.value = 0
			End If
		else
			'.txtSELECTAMT.value = 0
		End If
		'Call gFormatNumber(.txtSELECTAMT,0,True)
	End With
End Sub

'ȭ�� �ʱ� ��Ʈ �̹��� ���� �� ���� �ʱ�ȭ 
Sub InitPage()
	'����������ü ����	
	set mobjPDCMCONTRACT_BASE	= gCreateRemoteObject("cPDCO.ccPDCOCONTRACT_BASE")
	set mobjPDCMGET				= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 1
		mobjSCGLSpr.AddCellSpan  .sprSht, 6, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | GBN | GUBUN | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG "
		mobjSCGLSpr.SetHeader .sprSht,		 "����|�������|��༭����|��༭��ȣ|����ó�ڵ�|����ó��|����|�ݾ�|�����|��������|���������|Ư�����|����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","  4|       8|         9|        10|         9|2|    15|    16|   8|     8|        10|        10|      12|   4"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK | CONFIRMFLAG"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"��", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CONTRACTDAY | STDATE | EDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CONTRACTNO | OUTSCODE | OUTSNAME | COMENT ", -1, -1, 255
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "CONTRACTNO | GBN | GUBUN"
		'mobjSCGLSpr.ColHidden .sprSht, "SEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "GBN | GUBUN | OUTSCODE",-1,-1,2,2,False
		mobjSCGLSpr.CellGroupingEach .sprSht," OUTSNAME"

		.sprSht.style.visibility = "visible"
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	

	pnlTab1.style.visibility = "visible"
End Sub

Sub EndPage()
	set mobjPDCMCONTRACT_BASE = Nothing
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
		DateClean Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		DateClean2 Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		
		.txtCONTRACTDAY.value = gNowDate
		.txtCOMENT.value  = ""
		.cmbCONFIRM.value = ""
		
		Field_Lock
	End With
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
	
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
	end if
End Sub

Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	with frmThis
		if .sprSht.MaxRows = 0 then exit function '�׸��� �����Ͱ� ������ ������.
		.txtCONTRACTNAME.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNAME",Row)
		.txtCONTRACTDAY.value	= mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTDAY",Row)
		.txtSTDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"STDATE",Row)
		.txtEDDATE.value		= mobjSCGLSpr.GetTextBinding(.sprSht,"EDDATE",Row)
		
		.cmbGUBUN.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"GUBUN",Row)
		.cmbGBN.value			= mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",Row)
		
		.txtCOMENT.value = mobjSCGLSpr.GetTextBinding(.sprSht,"COMENT",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",Row) = "1" THEN
			.chkCONFIRMFLAG.checked = TRUE
		ELSE
			.chkCONFIRMFLAG.checked = FALSE
		END IF
		
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
' Field_Lock  ��Ȳ�� ���� �����Ҽ� ������ �ʵ带 ReadOnlyó��
'-----------------------------------------------------------------------------------------
Sub Field_Lock ()
	With frmThis
		If .sprSht.MaxRows > 0 Then
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",.sprSht.ActiveRow) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",.sprSht.ActiveRow) = "1" Then
			
				.txtCONTRACTNAME.className		= "NOINPUT_L" : .txtCONTRACTNAME.readOnly	= True 
				.txtCONTRACTDAY.className		= "NOINPUT"	  : .txtCONTRACTDAY.readOnly	= True
				.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
				.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
				.txtCOMENT.className			= "NOINPUT"   : .txtCOMENT.readOnly			= True 
				.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
				.cmbGBN.className				= "NOINPUT"   : .cmbGBN.disabled			= True 
				
				.ImgCONTRACTDAY.disabled = true
				.imgFROM2.disabled = true
				.imgTO2.disabled = true
			
			'��༭ ��ȣ�� �ְ� ������ ���� ������������ ��� ��౸�а� ������ ��ٴ�.
			Elseif mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",.sprSht.ActiveRow) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",.sprSht.ActiveRow) = "0" then
			'�⺻��༭�� ��� ���Ⱓ�� ������ �ʴ´�.
				if .cmbGUBUN.value = "�⺻" then
					.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
					.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
					
				else
					.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
					.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
					
				end if
				.cmbGUBUN.className				= "NOINPUT"   : .cmbGUBUN.disabled			= True 
				.cmbGBN.className				= "NOINPUT"   : .cmbGBN.disabled			= True 
				.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
				.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
				.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
			
				.ImgCONTRACTDAY.disabled = False
				.imgFROM2.disabled = False
				.imgTO2.disabled = False
			Else 
				'�⺻��༭�� ��� ���Ⱓ�� ������ �ʴ´�.
				if .cmbGUBUN.value = "�⺻" then
					.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
					.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
					
				else
					.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
					.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
					
				end if
				.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
				.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
				.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
				.cmbGUBUN.className				= "INPUT"   : .cmbGUBUN.disabled			= False
				.cmbGBN.className				= "INPUT"   : .cmbGBN.disabled			= False
				
				.ImgCONTRACTDAY.disabled = False
				.imgFROM2.disabled = False
				.imgTO2.disabled = False
			End If
		Else
			'�⺻��༭�� ��� ���Ⱓ�� ������ �ʴ´�.
			if .cmbGUBUN.value = "�⺻" then
				.txtSTDATE.className			= "NOINPUT"   : .txtSTDATE.readOnly			= True
				.txtEDDATE.className			= "NOINPUT"   : .txtEDDATE.readOnly			= True
			else
				.txtSTDATE.className			= "INPUT"   : .txtSTDATE.readOnly			= False
				.txtEDDATE.className			= "INPUT"   : .txtEDDATE.readOnly			= False
			end if
			.txtCONTRACTNAME.className		= "INPUT_L" : .txtCONTRACTNAME.readOnly		= False 
			.txtCONTRACTDAY.className		= "INPUT"	: .txtCONTRACTDAY.readOnly		= False
			.txtCOMENT.className			= "INPUT"   : .txtCOMENT.readOnly			= False 
			.cmbGUBUN.className				= "INPUT"	: .cmbGUBUN.disabled			= False
			.cmbGBN.className				= "INPUT"	: .cmbGBN.disabled			= False
			

			.ImgCONTRACTDAY.disabled = False
			.imgFROM2.disabled = False
			.imgTO2.disabled = False
		End If
	End With
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strFROM, strTO
	Dim strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN,strcmbGBN
	Dim i,j,strRows

	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		j = 1

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strFROM			= .txtFrom.value
		strTO			= .txtTo.value
		strOUTSCODE		= .txtOUTSCODE.value
		strOUTSNAME		= .txtOUTSNAME.value
		strCONFIRM		= .cmbCONFIRM.value
		strCONTRACTNO	= .txtCONTRACTNO.value
		strCONTRACTNAME1 = .txtCONTRACTNAME1.value
		strcmbGBN		= .cmbGBN1.value
		strcmbGUBUN		= .cmbGUBUN1.value

		vntData = mobjPDCMCONTRACT_BASE.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, strFROM,strTO, _ 
												  strOUTSCODE,strOUTSNAME,strCONFIRM,strCONTRACTNO, _ 
												  strCONTRACTNAME1,strcmbGBN, strcmbGUBUN)

		If not gDoErrorRtn ("SelectRtn") Then
			InitPageData
   			PreSearchFiledValue strFROM,strTO, strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE

   				for i = 1 to .sprSht.MaxRows
   					If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" and mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "1" then
   					
						If j = 1 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
						j = j + 1
						
					End If
   				next
   				mobjSCGLSpr.SetCellsLock2 .sprSht,True,strRows,2,11,True

   				'AMT_SUM
   				sprShtToFieldBinding 2, 1
   			else
   				.sprSht.MaxRows = 0
   			End If
   		End If
   	end With
End Sub

'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strFROM,strTO, strOUTSCODE, strOUTSNAME, strCONFIRM, strCONTRACTNO, strCONTRACTNAME1,strcmbGUBUN)
	With frmThis
		.txtFrom.value			= strFROM
		.txtTo.value			= strTO
		.txtOUTSCODE.value		= strOUTSCODE
		.txtOUTSNAME.value		= strOUTSNAME
		.cmbCONFIRM.value		= strCONFIRM
		.txtCONTRACTNO.value	= strCONTRACTNO
		.txtCONTRACTNAME1.value = strCONTRACTNAME1
		.cmbGUBUN1.value		= strcmbGUBUN
	End With
End Sub

'------------------------------------------
' ��༭ ����
'------------------------------------------
Sub ProcessRtn_HDR ()
   	Dim intRtn
   	Dim strMasterData
   	Dim vntData
	Dim lngchkCnt
	Dim i
	Dim strOUTSCODE
	Dim strCONTRACTNO
	Dim strGUBUN
	
	With frmThis
		strMasterData = gXMLGetBindingData (xmlBind)
		lngchkCnt = 0 :  strCONTRACTNO = "" : strGUBUN = ""

		For i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
				strOUTSCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",i)
				lngchkCnt = lngchkCnt +1
				'������ ��Ʈ�� ���� �÷��׸� �����Ѵ�.
				mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
			End If 
			
		Next
		
		if strOUTSCODE = "" then
			gErrorMsgBox "����ó �Է��� �ʼ� �����Դϴ�.","��༭ ���� �ȳ�!"
			Exit Sub
		end if 

		If lngchkCnt = 0 Then
			gErrorMsgBox "������ ��༭�� üũ�� �ּ���.","Ȯ���ȳ�!"
			EXIT Sub
		End If

		'������ ������ üũ
		if DataValidation (strOUTSCODE) = false then exit sub
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | GBN | GUBUN | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG ")

		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If

		intRtn = mobjPDCMCONTRACT_BASE.ProcessRtn(gstrConfigXml, strMasterData, vntData, strCONTRACTNO, strOUTSCODE)

		If not gDoErrorRtn ("ProcessRtn_HDR") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox strCONTRACTNO & " ��ȣ�� Ȯ�� �Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   	End With
End Sub

'-------------------------------------------------
'������ �⺻��༭�� �ܰ���༭�� ���� �ϴ��� Ȯ��
'-------------------------------------------------
Function DataValidation (strOUTSCODE)
	DataValidation = false
	
	Dim vntData
	Dim intYNRtn
	
	'On error resume next
	with frmThis
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� 
   		'IF not gDataValidation(frmThis) then exit Function
   		
   		mlngRowCnt=clng(0): mlngColCnt=clng(0)
   		
   		vntData = mobjPDCMCONTRACT_BASE.SelctRtn_validation(gstrConfigXml, mlngRowCnt,mlngColCnt, strOUTSCODE)
   		
   		If not gDoErrorRtn ("SelctRtn_validation") Then
			
			'�ش� �����Ͱ� �ϴ��ִٸ� ��༭ ������ Ȯ���ϰ� �޽����� ����.
			If mlngRowCnt >0 Then
				
				'�⺻ ��༭�ϰ��
				if vntData(0,1) = "�⺻" then
					if .cmbGUBUN.value = "�⺻" then
						gErrorMsgBox "�ش� ����ó�� �⺻��༭�� ���� �մϴ�." ,"����ȳ�"
						exit function
					elseif .cmbGUBUN.value = "�ܰ�" then
						intYNRtn = gYesNoMsgbox("�⺻��༭�� ���� �մϴ� �ܰ���༭�� �߰� ���� �Ͻðڽ��ϱ�?","Ȯ��Ȯ��")
						IF intYNRtn <> vbYes then exit function	
					end if
	
				elseif vntData(0,1) = "�ܰ�" then
					if .cmbGUBUN.value = "�⺻" then
						gErrorMsgBox "�ش� ����ó�� �ܰ���༭�� ���� �մϴ�." ,"����ȳ�"
						exit function
					elseif .cmbGUBUN.value = "�ܰ�" then
						gErrorMsgBox "�ش� ����ó�� �ܰ���༭�� ���� �մϴ�." ,"����ȳ�"
						exit function
					end if
				end if
   			End If
   		End If
   		
   	End with
	DataValidation = true
End Function

'------------------------------------------
' ��༭ ���� ������ ó��
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	Dim strCONTRACTNO
	Dim lngchkCnt
	
	with frmThis
		lngchkCnt = 0
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				lngchkCnt = lngchkCnt +1
				
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = "1" then
					gErrorMsgBox "���ε� �����ʹ� ���� �Ͻ� �� �����ϴ�..","�ڷ� ���� �ȳ�!"
					EXIT Sub		
				end if
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","Ȯ����Ҿȳ�!"
			EXIT Sub
		End If

		'���õ� �ڷḦ ������ ���� ����
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub

		For i = .sprSht.MaxRows to 1 step -1
			strCONTRACTNO = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i) <> "" Then
					strCONTRACTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"CONTRACTNO",i)
					intRtn = mobjPDCMCONTRACT_BASE.DeleteRtn(gstrConfigXml, strCONTRACTNO)
					IF not gDoErrorRtn ("DeleteRtn") then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End IF
				else
					mobjSCGLSpr.DeleteRow .sprSht,i
				End IF
   			End If
		Next
		gOkMsgBox "�ڷᰡ ���� �Ǿ����ϴ�.","���� �ȳ�!"
		gWriteText lblstatus, "�ڷᰡ " & intRtn & " �� �����Ǿ����ϴ�."
	End with
	err.clear
End Sub

Sub UpdateRtn_CONFIRM (confirm)
   	Dim i, lngchkCnt
   	Dim intRtn
   	Dim vntData
   	Dim strMSG

	With frmThis
		
		If confirm = 1 then
			strMSG = "����"
		Else 
			strMSG = "���� ���"
		End if 

		For i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				if mobjSCGLSpr.GetTextBinding(.sprSht,"CONFIRMFLAG",i) = confirm then
					gErrorMsgBox "�̹� " & strMSG & " �� ������ �Դϴ�..","������ ���� �ȳ�!"
					EXIT Sub
				end if 
				lngchkCnt = lngchkCnt +1
			End If 
		Next

		If lngchkCnt = 0 Then
			gErrorMsgBox strMSG & "�� ��༭�� üũ�� �ּ���.","Ȯ���ȳ�!"
			EXIT Sub
		End If

		vntData = mobjSCGLSpr.GetDataRows(.sprSht," CHK | CONTRACTNO | OUTSCODE | BTN  | OUTSNAME | CONTRACTNAME | AMT | CONTRACTDAY | STDATE | EDDATE | COMENT | CONFIRMFLAG | GUBUN")

		if  not IsArray(vntData)  then 
			gErrorMsgBox "����� �Է��ʵ� " & meNO_DATA,"����ȳ�"
			exit sub
		End If

		intRtn = mobjPDCMCONTRACT_BASE.Processrtn_CONFIRM(gstrConfigXml, vntData,confirm)
		If not gDoErrorRtn ("Processrtn_CONFIRM") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox " �����Ͱ� " & strMSG & " �Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   	End With
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
												<td class="TITLE">�⺻ &amp; �ܰ� ��༭ ü��</td>
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
									<TD class="SEARCHLABEL" style="WIDTH: 70px; CURSOR: hand; HEIGHT: 9px" onclick="vbscript:Call gCleanField(txtfrom,txtTo)"
										width="56">���Ⱓ</TD>
									<TD class="SEARCHDATA" style="WIDTH: 250px; HEIGHT: 9px"><INPUT class="INPUT" id="txtFrom" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtFrom"> <IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgFrom">&nbsp; ~&nbsp; <INPUT class="INPUT" id="txtTo" title="���˻� ��������" style="WIDTH: 88px; HEIGHT: 22px" accessKey="DATE"
											type="text" maxLength="10" size="9" name="txtTo"> <IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" height="15" src="../../../images/btnCalEndar.gIF"
											align="absMiddle" border="0" name="imgTo">
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 64px; CURSOR: hand; HEIGHT: 9px">��༭Ȯ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 181px; CURSOR: hand; HEIGHT: 9px"><SELECT id="cmbCONFIRM" style="WIDTH: 120px" name="cmbCONFIRM">
											<OPTION value="" selected>��ü</OPTION>
											<OPTION value="0">��༭ �̽���</OPTION>
											<OPTION value="1">��༭ ����</OPTION>
										</SELECT></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 60px; HEIGHT: 9px" onclick="vbscript:Call gCleanField(txtOUTSNAME, txtOUTSCODE)">����ó</TD>
									<TD class="SEARCHDATA" style="HEIGHT: 9px" colSpan="3"><INPUT class="INPUT_L" id="txtOUTSNAME" title="����ó�� ��ȸ" style="WIDTH: 160px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="32" name="txtOUTSNAME"> <IMG id="ImgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle" border="0" name="ImgCLIENTCODE">
										<INPUT class="INPUT" id="txtOUTSCODE" title="����ó�ڵ���ȸ" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="6" align="left" size="3" name="txtOUTSCODE">
									</TD>
									<td><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
											height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" align="right" border="0"
											name="imgQuery">
									</td>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNO, '')">��༭��ȣ</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCONTRACTNO" title="��༭��ȣ ��ȸ" style="WIDTH: 240px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="34" name="txtCONTRACTNO">
									</TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTNAME1, '')">����</TD>
									<TD class="SEARCHDATA" style="WIDTH: 181px"><INPUT class="INPUT_L" id="txtCONTRACTNAME1" title="����� ��ȸ" style="WIDTH: 180px; HEIGHT: 22px"
											type="text" maxLength="255" align="left" size="30" name="txtCONTRACTNAME1"></TD>
									<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField('', '')">��༭����</TD>
									<TD class="SEARCHDATA" style="WIDTH: 116px"><SELECT id="cmbGBN1" title="��༭ ����" style="WIDTH: 112px" name="cmbGBN1">
											<OPTION value="" selected>��ü</OPTION>
											<OPTION value="ATL">ATL</OPTION>
											<OPTION value="BTL">BTL</OPTION>
										</SELECT>
									</TD>
									<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField('', '')">��༭����</TD>
									<TD class="SEARCHDATA"><SELECT id="cmbGUBUN1" title="��༭ ����" style="WIDTH: 112px" name="cmbGUBUN1">
											<OPTION value="" selected>��ü</OPTION>
											<OPTION value="�⺻">�⺻</OPTION>
											<OPTION value="�ܰ�">�ܰ�</OPTION>
										</SELECT>
									</TD>
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
													<TABLE cellSpacing="0" cellPadding="0" width="180" background="../../../images/back_p.gIF"
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
												<td class="TITLE">�⺻&amp; �ܰ� ��༭ ���� �� ��ȸ</td>
											</tr>
										</table>
									</TD>
									<!--<td>
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
									-->
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<!--<TD><IMG id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" alt="�ű��ڷḦ �����մϴ�."
														src="../../../images/imgNew.gIF" border="0" name="imgREG"></TD> -->
												<td><IMG id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
														height="20" alt="������ ���� �����մϴ�." src="../../../images/imgAgree.gIF" align="absMiddle"
														border="0" name="imgAgree"><IMG id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'" height="20" alt="������ ���� ������� �մϴ�."
														src="../../../images/imgAgreeCanCel.gIF" align="absMiddle" border="0" name="imgAgreeCanCel">
												</td>
												<!--<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
												<TD width="15"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
														name="imgDelete"></TD>
														-->
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
												<TD class="SEARCHDATA" style="WIDTH: 289px" colSpan="3"><INPUT dataFld="CONTRACTNAME" class="INPUT_L" id="txtCONTRACTNAME" title="����" style="WIDTH: 240px; HEIGHT: 21px"
														accessKey=",M" dataSrc="#xmlBind" type="text" size="30" name="txtCONTRACTNAME"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 85px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCONTRACTDAY,'')">�����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 200px"><INPUT dataFld="CONTRACTDAY" class="INPUT" id="txtCONTRACTDAY" title="�����" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtCONTRACTDAY">
													<IMG id="ImgCONTRACTDAY" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'"
														height="15" alt="ImgCONTRACTDAY" src="../../../images/btnCalEndar.gIF" align="absMiddle"
														border="0" name="ImgCONTRACTDAY">&nbsp;&nbsp;��༭����<INPUT dataFld="CONFIRMFLAG" id="chkCONFIRMFLAG" title="��༭����" dataSrc="#xmlBind" type="checkbox"
														value="" name="chkCONFIRMFLAG"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtSTDATE,txtEDDATE)">��� 
													�Ⱓ</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="STDATE" class="INPUT" id="txtSTDATE" title="���Ⱓ ������" style="WIDTH: 88px; HEIGHT: 22px"
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
											</TR>
											<TR>
												<TD class="SEARCHLABEL">��༭����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 100px; HEIGHT: 25px"><SELECT dataFld="GBN" id="cmbGBN" title="��༭����" style="WIDTH: 100px" dataSrc="#xmlBind"
														name="cmbGBN">
														<OPTION value="ATL" selected>ATL</OPTION>
														<OPTION value="BTL">BTL</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="WIDTH: 82px">��༭����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 99px" HEIGHT:style="WIDTH: 102px"><SELECT dataFld="GUBUN" id="cmbGUBUN" title="��༭ ����" style="WIDTH: 100px" dataSrc="#xmlBind"
														name="cmbGUBUN">
														<OPTION value="�⺻" selected>�⺻</OPTION>
														<OPTION value="�ܰ�" id="test">�ܰ�</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" style="HEIGHT: 21px">��»���</TD>
												<TD class="SEARCHDATA" style="HEIGHT: 21px" colSpan="5"><INPUT dataFld="CONFLAG" id="rdCONFLAG" title="����뿪 �ܰ� ��༭" dataSrc="#xmlBind" type="radio"
														CHECKED value="rdCONFLAG" name="rdCONFLAG">����뿪 �⺻ &amp; �ܰ� ��༭
												</TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 53px" onclick="vbscript:Call gCleanField(txtCOMENT, '')">Ư�����</TD>
												<TD class="SEARCHDATA" colSpan="7"><TEXTAREA dataFld="COMENT" id="txtCOMENT" style="WIDTH: 778px" dataSrc="#xmlBind" name="txtCOMENT"
														wrap="hard" cols="10"></TEXTAREA></TD>
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
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="48763">
									<PARAM NAME="_ExtentY" VALUE="11324">
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
