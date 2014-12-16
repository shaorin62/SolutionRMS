<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOPTLIST.aspx.vb" Inherits="SC.SCCOPTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>PT_����</title>
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'�ý��۱��� : SC/ PT����
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOPTLIST.aspx
'��      �� : PT ������ ����
'�Ķ�  ���� : 
'Ư��  ���� :
'----------------------------------------------------------------------------------------
'HISTORY    :1)Ver. Oh Se Hoon
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLES.CSS">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT id="clientEventHandlersVBS" language="vbscript">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjSCCOPTLIST 
Dim mobjSCCOGET
Dim mstrCheck
Dim mcomecalender1, mcomecalender2, mcomecalender3, mcomecalender4
Dim mstrHIDDEN
CONST meTAB = 9

mstrCheck = True
mcomecalender1 = FALSE
mcomecalender2 = FALSE
mcomecalender3 = FALSE
mcomecalender4 = FALSE
mstrHIDDEN = 0
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
'�Է� �ʵ� �����
Sub Set_TBL_HIDDEN()
	With frmThis
		If mstrHIDDEN Then
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableUp' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableUp.gif' align='absmiddle' border='0' name='imgTableUp'>"
			'document.getElementById("SizeOrSdt").innerHTML="������"
			document.getElementById("tblBody").style.display = "inline"
			document.getElementById("tblSheet").style.height = "65%"
		Else
			document.getElementById("spnHIDDEN").innerHTML="<IMG id='imgTableDown' style='CURSOR: hand' alt='�ڷḦ �˻��մϴ�.' src='../../../images/imgTableDown.gif' align='absmiddle' border='0' name='imgTableDown'>"
			document.getElementById("tblBody").style.display = "none"
			document.getElementById("tblSheet").style.height = "95%"
		End If

		If mstrHIDDEN Then
			mstrHIDDEN = 0
		Else
			mstrHIDDEN = 1
		End If
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
'��ȸ��ư
Sub imgQuery_onclick
	If frmThis.txtSTYEARMON.value = "" AND frmThis.txtEDYEARMON.value = "" Then
		gErrorMsgBox "��ȸ����� �Է��Ͻÿ�","��ȸ�ȳ�"
		exit Sub
	End If
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
'�ʱ�ȭ��ư
Sub imgCho_onclick
	InitPageData
End Sub

'�űԹ�ư
Sub imgREG_onclick ()
	Call sprSht_Keydown(meINS_ROW, 0)	
end Sub

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
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht
	end With
	gFlowWait meWAIT_OFF
End Sub


'-----------------------------------------------------------------------------------------
' �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	On error resume next
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData

		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE1_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'-----------------------------------------------------------------------------------------
' �˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME.value))
	    vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			.txtBUSINO.value	 = trim(vntRet(2,0))
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUSINO",frmThis.sprSht.ActiveRow, trim(vntRet(2,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE.value),trim(.txtCLIENTNAME.value), "A")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = trim(vntData(0,1))
					.txtCLIENTNAME.value = trim(vntData(1,1))
					.txtBUSINO.value	 = trim(vntData(2,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUSINO",frmThis.sprSht.ActiveRow, trim(vntData(2,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'����ó �˾�
Sub ImgGREATCODE_onclick
	Call GREATCODE_POP()
End Sub

Sub GREATCODE_POP
	Dim vntRet, vntInParams
	With frmThis 
		vntInParams = array(trim(.txtGREATCODE.value),trim(.txtGREATNAME.value))
		vntRet = gShowModalWindow("../SCCO/SCCOGREATCUSTPOP.aspx",vntInParams , 413,440)
		
		If isArray(vntRet) Then
		    .txtGREATCODE.value = trim(vntRet(0,0))	'Code�� ����
			.txtGREATNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtGREATCODE
		End If
	end With
End Sub

Sub txtGREATNAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetGREATCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtGREATCODE.value,.txtGREATNAME.value)
		
			If not gDoErrorRtn ("GetGREATCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtGREATCODE.value = trim(vntData(0,1))	'Code�� ����
					.txtGREATNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
			
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call GREATCODE_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'���μ� �˾� 
Sub imgDEPT_CD_onclick
	Call DEPT_CD_POP()
End Sub

Sub DEPT_CD_POP
	Dim vntRet, vntInParams
	With frmThis
		vntInParams = array(trim(.txtDEPT_NAME.value))
		vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,440)
		If isArray(vntRet) Then
		    .txtDEPT_CD.value = trim(vntRet(0,0))	'Code�� ����
			.txtDEPT_NAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
			If .sprSht.MaxRows > 0 Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
			End If
			gSetChangeFlag .txtDEPT_CD
		End If
	end With
End Sub

Sub txtDEPT_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjSCCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtDEPT_NAME.value))
			
			If not gDoErrorRtn ("GetCC") Then
				If mlngRowCnt = 1 Then
					.txtDEPT_CD.value = trim(vntData(0,1))
					.txtDEPT_NAME.value = trim(vntData(1,1))
					If .sprSht.MaxRows > 0 Then
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
					End If
				Else
					Call DEPT_CD_POP()
				End If
   			End If
   		end With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


'****************************************************************************************
' SpreadSheet �̺�Ʈ
'****************************************************************************************
'--------------------------------------------------
'��Ʈ Ű�ٿ�
'--------------------------------------------------
Sub sprSht_Keydown(KeyCode, Shift)
	Dim intRtn
	
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	If KeyCode = meINS_ROW Then
		frmThis.sprSht.MaxRows = 0
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"YEARMON",frmThis.sprSht.ActiveRow, gNowDate
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_RESULT",frmThis.sprSht.ActiveRow, ""
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_STATUS",frmThis.sprSht.ActiveRow, "����"
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLASS",frmThis.sprSht.ActiveRow, ""
		
		mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,frmThis.sprSht.MaxRows
		
		frmThis.txtCLIENTNAME1.focus
		frmThis.sprSht.focus
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub sprSht_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strSUM
	Dim intSelCnt, intSelCnt1
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

'****************************************************************************************
' ������ �޷�
'****************************************************************************************
Sub imgCalEndar1_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender1 = true
	gShowPopupCalEndar frmThis.txtPT_DATE1,frmThis.imgCalEndar1,"txtPT_DATE1_onchange()"
	IF frmThis.sprSht.MaxRows <> 0 THEN
		Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PT_DATE1"), frmThis.sprSht.ActiveRow)
	END IF 
	
	mcomecalender1 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalEndar2_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender2 = true
	gShowPopupCalEndar frmThis.txtPT_DATE2,frmThis.imgCalEndar2,"txtPT_DATE2_onchange()"
	IF frmThis.sprSht.MaxRows <> 0 THEN
		Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PT_DATE2"), frmThis.sprSht.ActiveRow)
	end if
	mcomecalender2 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalEndar3_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender3 = true
	gShowPopupCalEndar frmThis.txtPT_DATE3,frmThis.imgCalEndar3,"txtPT_DATE3_onchange()"
	IF frmThis.sprSht.MaxRows <> 0 THEN
		Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"PT_DATE3"), frmThis.sprSht.ActiveRow)
	end if
	mcomecalender3 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalEndar4_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	mcomecalender4 = true
	gShowPopupCalEndar frmThis.txtOT_DATE,frmThis.imgCalEndar4,"txtOT_DATE_onchange()"
	IF frmThis.sprSht.MaxRows <> 0 THEN
		Call sprSht_Change(mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"OT_DATE"), frmThis.sprSht.ActiveRow)
	end if
	mcomecalender4 = false
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub



'****************************************************************************************
' �Է��ʵ� Ű�ٿ� �̺�Ʈ
'****************************************************************************************
Sub txtCLIENTCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtBUSINO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtBUSINO_onkeydown
	Dim strBUSINO
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		
		strBUSINO = frmThis.txtBUSINO.value
		
		if instr(1,strBUSINO,"-") = 0 then
			strBUSINO = mid(strBUSINO,1,3) & "-" & mid(strBUSINO,4,2) & "-" & mid(strBUSINO,7,len(strBUSINO))		
			frmThis.txtBUSINO.value = strBUSINO
		end if
		
		if frmThis.sprSht.MaxRows <> 0 then
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUSINO",frmThis.sprSht.ActiveRow, frmThis.txtBUSINO.value
			mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
		end if 
	
		frmThis.txtGREATNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtGREATCODE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_LIST.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtPT_LIST_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbPT_STATUS.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtOLDCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_BILL.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEX_BILL_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_CONDITION.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEX_CONDITION_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.cmbPT_CLASS.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub cmbPT_CLASS_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtOT_DATE.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtOT_DATE_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEX_INFO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEX_INFO_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtOT_INFO.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtOT_INFO_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_ESTAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtPT_ESTAMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_ACTAMT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtPT_ACTAMT_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_DATE1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub


Sub txtPT_DATE1_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_DATE2.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPT_DATE2_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_DATE3.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPT_DATE3_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_CLIENTNAME1.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPT_CLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_CLIENTNAME2.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPT_CLIENTNAME2_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_CLIENTNAME3.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtPT_CLIENTNAME3_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtETCCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtDEPT_CD_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtEXCLIENTNAME.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

Sub txtEXCLIENTNAME_onkeydown
	If window.event.keyCode = meEnter or window.event.keyCode = meTAB Then
		frmThis.txtPT_TEXT.focus()
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' �Է��ʵ� ü���� �̺�Ʈ
'****************************************************************************************
Sub txtCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtCLIENTCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"CLIENTCODE",frmThis.sprSht.ActiveRow, frmThis.txtCLIENTCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtGREATNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, frmThis.txtGREATNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtGREATCODE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, frmThis.txtGREATCODE.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtBUSINO_onchange
	Dim strBUSINO
	If frmThis.sprSht.ActiveRow >0 Then
		
		strBUSINO = frmThis.txtBUSINO.value
		
		if instr(1,strBUSINO,"-") = 0 then
			strBUSINO = mid(strBUSINO,1,3) & "-" & mid(strBUSINO,4,2) & "-" & mid(strBUSINO,7,len(strBUSINO))		
			frmThis.txtBUSINO.value = strBUSINO
		end if
		
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUSINO",frmThis.sprSht.ActiveRow, frmThis.txtBUSINO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub cmbPT_STATUS_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		IF frmThis.cmbPT_STATUS.value = "����" THEN 
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_STATUS",frmThis.sprSht.ActiveRow, "����"
		elseif frmThis.cmbPT_STATUS.value = "�ܵ�" THEN 
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_STATUS",frmThis.sprSht.ActiveRow, "�ܵ�"
		else
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_STATUS",frmThis.sprSht.ActiveRow, "ANNUAL"
		END IF 
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_LIST_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_LIST",frmThis.sprSht.ActiveRow, frmThis.txtPT_LIST.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEX_BILL_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_BILL",frmThis.sprSht.ActiveRow, frmThis.txtEX_BILL.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtOLDCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OLDCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtOLDCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub cmbPT_CLASS_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLASS",frmThis.sprSht.ActiveRow, frmThis.cmbPT_CLASS.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEX_CONDITION_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_CONDITION",frmThis.sprSht.ActiveRow, frmThis.txtEX_CONDITION.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEX_INFO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_INFO",frmThis.sprSht.ActiveRow, frmThis.txtEX_INFO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtOT_DATE_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EX_INFO",frmThis.sprSht.ActiveRow, frmThis.txtEX_INFO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtOT_DATE_onchange
	Dim strdate 
	Dim strOT_DATE
	strdate = ""
	strOT_DATE =""
	With frmThis
		strdate=.txtOT_DATE.value
	
		If mcomecalender4 Then
			strOT_DATE = strdate
		else
			If len(strdate) = 4 Then
				strOT_DATE = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strOT_DATE = strdate
			elseif len(strdate) = 3 Then
				strOT_DATE = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strOT_DATE = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"OT_DATE",.sprSht.ActiveRow, strOT_DATE
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtOT_INFO_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"OT_INFO",frmThis.sprSht.ActiveRow, frmThis.txtOT_INFO.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_ESTAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_ESTAMT",frmThis.sprSht.ActiveRow, frmThis.txtPT_ESTAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_ACTAMT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_ACTAMT",frmThis.sprSht.ActiveRow, frmThis.txtPT_ACTAMT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_DATE1_onchange
	Dim strdate 
	Dim strPT_DATE1
	strdate = ""
	strPT_DATE1 =""
	With frmThis
		strdate=.txtPT_DATE1.value
	
		If mcomecalender1 Then
			strPT_DATE1 = strdate
		else
			If len(strdate) = 4 Then
				strPT_DATE1 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strPT_DATE1 = strdate
			elseif len(strdate) = 3 Then
				strPT_DATE1 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strPT_DATE1 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"PT_DATE1",.sprSht.ActiveRow, strPT_DATE1
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtPT_DATE2_onchange
	Dim strdate 
	Dim strPT_DATE2
	strdate = ""
	strPT_DATE2 =""
	With frmThis
		strdate=.txtPT_DATE2.value
	
		If mcomecalender2 Then
			strPT_DATE2 = strdate
		else
			If len(strdate) = 4 Then
				strPT_DATE2 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strPT_DATE2 = strdate
			elseif len(strdate) = 3 Then
				strPT_DATE2 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strPT_DATE2 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"PT_DATE2",.sprSht.ActiveRow, strPT_DATE2
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtPT_DATE3_onchange
	Dim strdate 
	Dim strPT_DATE3
	strdate = ""
	strPT_DATE3 =""
	With frmThis
		strdate=.txtPT_DATE3.value
	
		If mcomecalender3 Then
			strPT_DATE3 = strdate
		else
			If len(strdate) = 4 Then
				strPT_DATE3 = Mid(gNowDate2,1,4) & strdate
			elseif len(strdate) = 10 Then
				strPT_DATE3 = strdate
			elseif len(strdate) = 3 Then
				strPT_DATE3 = Mid(gNowDate2,1,4) & "0" & strdate
			else
				strPT_DATE3 = strdate
			End If
		End If
		
		If .sprSht.ActiveRow >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht,"PT_DATE3",.sprSht.ActiveRow, strPT_DATE3
			mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol,.sprSht.ActiveRow
		End If
	End With
	gSetChange
End Sub

Sub txtPT_CLIENTNAME1_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME1",frmThis.sprSht.ActiveRow, frmThis.txtPT_CLIENTNAME1.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_CLIENTNAME2_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME2",frmThis.sprSht.ActiveRow, frmThis.txtPT_CLIENTNAME2.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_CLIENTNAME3_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME3",frmThis.sprSht.ActiveRow, frmThis.txtPT_CLIENTNAME3.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub rMNC_onclick
	
	frmThis.txtETCCLIENTNAME.className	= "NOINPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= True 
	frmThis.txtETCCLIENTNAME.value = ""
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_RESULT",frmThis.sprSht.ActiveRow, "SK�����þ����۴�(��)"
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub rETC_onclick
		frmThis.txtETCCLIENTNAME.className	= "INPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= false 
	
	If frmThis.sprSht.ActiveRow >0 then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_RESULT",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub rDE_onclick
	
	frmThis.txtETCCLIENTNAME.className	= "NOINPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= True 
	frmThis.txtETCCLIENTNAME.value = ""
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_RESULT",frmThis.sprSht.ActiveRow, ""
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub chkATTEND_onclick
	WITH frmThis
		if .chkATTEND.checked = true then
			.txtPT_DATE1.value = "" : .txtPT_DATE2.value = "" : .txtPT_DATE3.value = "" : 
			.txtPT_CLIENTNAME1.value = "" : .txtPT_CLIENTNAME2.value = "" : .txtPT_CLIENTNAME3.value = "" : 
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_DATE1",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_DATE2",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_DATE3",frmThis.sprSht.ActiveRow, ""
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME1",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME2",frmThis.sprSht.ActiveRow, ""
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_CLIENTNAME3",frmThis.sprSht.ActiveRow, ""
			
			
			.txtPT_DATE1.className	= "NOINPUT_L" : .txtPT_DATE1.readOnly	= true : .imgCalEndar1.disabled = true
			.txtPT_DATE2.className	= "NOINPUT_L" : .txtPT_DATE2.readOnly	= true : .imgCalEndar2.disabled = true
			.txtPT_DATE3.className	= "NOINPUT_L" : .txtPT_DATE3.readOnly	= true : .imgCalEndar3.disabled = true
			.txtPT_CLIENTNAME1.className	= "NOINPUT_L" : .txtPT_CLIENTNAME1.readOnly	= true
			.txtPT_CLIENTNAME2.className	= "NOINPUT_L" : .txtPT_CLIENTNAME2.readOnly	= true
			.txtPT_CLIENTNAME3.className	= "NOINPUT_L" : .txtPT_CLIENTNAME3.readOnly	= true
		
		else
			.txtPT_DATE1.className	= "INPUT_L" : .txtPT_DATE1.readOnly	= FALSE : .imgCalEndar1.disabled = FALSE
			.txtPT_DATE2.className	= "INPUT_L" : .txtPT_DATE2.readOnly	= FALSE : .imgCalEndar2.disabled = FALSE
			.txtPT_DATE3.className	= "INPUT_L" : .txtPT_DATE3.readOnly	= FALSE : .imgCalEndar3.disabled = FALSE
			.txtPT_CLIENTNAME1.className	= "INPUT_L" : .txtPT_CLIENTNAME1.readOnly	= FALSE
			.txtPT_CLIENTNAME2.className	= "INPUT_L" : .txtPT_CLIENTNAME2.readOnly	= FALSE
			.txtPT_CLIENTNAME3.className	= "INPUT_L" : .txtPT_CLIENTNAME3.readOnly	= FALSE
		end if
		
	end with
END SUB


Sub txtDEPT_NAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_NAME",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_NAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtDEPT_CD_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"DEPT_CD",frmThis.sprSht.ActiveRow, frmThis.txtDEPT_CD.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtETCCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_RESULT",frmThis.sprSht.ActiveRow, frmThis.txtETCCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtEXCLIENTNAME_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"EXCLIENTNAME",frmThis.sprSht.ActiveRow, frmThis.txtEXCLIENTNAME.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

Sub txtPT_TEXT_onchange
	If frmThis.sprSht.ActiveRow >0 Then
		mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PT_TEXT",frmThis.sprSht.ActiveRow, frmThis.txtPT_TEXT.value
		mobjSCGLSpr.CellChanged frmThis.sprSht, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'���� ����
Sub txtEX_BILL_onblur
	With frmThis
		Call gFormatNumber(.txtEX_BILL,0,True)
	end With
End Sub

'PT ����
Sub txtPT_ESTAMT_onblur
	With frmThis
		Call gFormatNumber(.txtPT_ESTAMT,0,True)
	end With
End Sub

'PT ������ ���
Sub txtPT_ACTAMT_onblur
	With frmThis
		Call gFormatNumber(.txtPT_ACTAMT,0,True)
	end With
End Sub


'-----------------------------------------------------------------------------------------
' õ���� ������ ���ֱ� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
Sub txtEX_BILL_onfocus
	With frmThis
		.txtEX_BILL.value = Replace(.txtEX_BILL.value,",","")
	end With
End Sub

Sub txtPT_ESTAMT_onfocus
	With frmThis
		.txtPT_ESTAMT.value = Replace(.txtPT_ESTAMT.value,",","")
	end With
End Sub

Sub txtPT_ACTAMT_onfocus
	With frmThis
		.txtPT_ACTAMT.value = Replace(.txtPT_ACTAMT.value,",","")
	end With
End Sub



Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
   	Dim strCode, strCodeName
   
	With frmThis
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTCODE") Then	.txtCLIENTCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			'���� �����Ǹ� �ڵ带 �����.
			mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,  _
													  strCode, strCodeName, "A")

				If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,1)
						mobjSCGLSpr.CellChanged .sprSht, Col-1,Row
						.txtCLIENTCODE.value = vntData(0,1)
						.txtCLIENTNAME.value = vntData(1,1)
						.txtBUSINO.value	 = vntData(2,1)
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
						mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
					End If
   				End If
   			End If
		End If
		
		'����ó��
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATCODE")  Then .txtGREATCODE.value = mobjSCGLSpr.GetTextBinding(.sprSht,"GREATCODE",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"GREATNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"GREATCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetGREATCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(strCode),trim(strCodeName))
					
				If not gDoErrorRtn ("GetGREATCUSTCODE") Then
					If mlngRowCnt = 1 Then
						
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntData(1,1))
						
						.txtGREATCODE.value = trim(vntData(0,1))	'Code�� ����
						.txtGREATNAME.value = trim(vntData(1,1))	'�ڵ�� ǥ��
						
						.txtGREATNAME.focus
						.sprSht.focus
						
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME"), Row
						.txtGREATNAME.focus
						.sprSht.focus 
					End If
   				End If
   			End If
		End IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_STATUS") Then 
			 .cmbPT_STATUS.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_STATUS",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_LIST") Then 
			 .txtPT_LIST.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_LIST",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_BILL") Then 
			 .txtEX_BILL.value = mobjSCGLSpr.GetTextBinding( .sprSht,"EX_BILL",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_CLASS") Then 
			 .cmbPT_CLASS.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_CLASS",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_CONDITION") Then 
			 .txtEX_CONDITION.value = mobjSCGLSpr.GetTextBinding( .sprSht,"EX_CONDITION",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"EX_INFO") Then 
			 .txtEX_INFO.value = mobjSCGLSpr.GetTextBinding( .sprSht,"EX_INFO",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OT_DATE") Then 
			 .txtOT_DATE.value = mobjSCGLSpr.GetTextBinding( .sprSht,"OT_DATE",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"OT_INFO") Then 
			 .txtOT_INFO.value = mobjSCGLSpr.GetTextBinding( .sprSht,"OT_INFO",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_ESTAMT") Then 
			 .txtPT_ESTAMT.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_ESTAMT",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_ACTAMT") Then 
			 .txtPT_ACTAMT.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_ACTAMT",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_DATE1") Then 
			 .txtPT_DATE1.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_DATE1",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_DATE2") Then 
			 .txtPT_DATE2.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_DATE2",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_DATE3") Then 
			 .txtPT_DATE3.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_DATE3",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_CLIENTNAME1") Then 
			 .txtPT_CLIENTNAME1.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_CLIENTNAME1",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_CLIENTNAME1") Then 
			 .txtPT_CLIENTNAME1.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_CLIENTNAME1",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_CLIENTNAME2") Then 
			 .txtPT_CLIENTNAME2.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_CLIENTNAME2",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_CLIENTNAME3") Then 
			 .txtPT_CLIENTNAME3.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_CLIENTNAME3",Row) 
		END IF
		
		
		'���μ�
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_CD") Then .txtDEPT_CD.value = mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjSCCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName)

				If not gDoErrorRtn ("GetCC") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.txtDEPT_CD.value = trim(vntData(0,1))
						.txtDEPT_NAME.value = trim(vntData(1,1))
						
						.txtCLIENTNAME1.focus()
						.sprSht.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME"), Row
						.txtCLIENTNAME1.focus()
						.sprSht.focus 
					End If
   				End If
   			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"PT_TEXT") Then 
			 .txtPT_TEXT.value = mobjSCGLSpr.GetTextBinding( .sprSht,"PT_TEXT",Row) 
		END IF
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"BUSINO") Then 
		Dim strBUSINO
			strBUSINO = mobjSCGLSpr.GetTextBinding( .sprSht,"BUSINO",Row) 
			if instr(1,strBUSINO,"-") = 0 then
				strBUSINO = mid(strBUSINO,1,3) & "-" & mid(strBUSINO,4,2) & "-" & mid(strBUSINO,7,len(strBUSINO))		
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"BUSINO",frmThis.sprSht.ActiveRow, strBUSINO
			end if
		END IF
		
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub


Sub mobjSCGLSpr_ClickProc(Col, Row)
	Dim vntRet
	Dim vntInParams
	With frmThis
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"CLIENTNAME") Then			
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)))
			
			vntRet = gShowModalWindow("../SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)		
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)
				.txtCLIENTCODE.value = vntRet(0,0)		
				.txtCLIENTNAME.value = vntRet(1,0)
				.txtBUSINO.value = vntRet(2,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"GREATNAME") Then
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"GREATNAME",Row)))
			
			vntRet = gShowModalWindow("../SCCO/SCCOGREATCUSTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATCODE",frmThis.sprSht.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding frmThis.sprSht,"GREATNAME",frmThis.sprSht.ActiveRow, trim(vntRet(1,0))
				
				.txtGREATCODE.value = trim(vntRet(0,0))	'Code�� ����
				.txtGREATNAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+1,Row
			End If
		End If
	
		'���μ�
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht,"DEPT_NAME") Then			
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../SCCO/SCCODEPTPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_NAME",Row, vntRet(1,0)
				
				.txtDEPT_CD.value = trim(vntRet(0,0))	'Code�� ����
				.txtDEPT_NAME.value = trim(vntRet(1,0))	'�ڵ�� ǥ��
				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End If
		End If
		
		sprShtToFieldBinding Col, Row
	End With
End Sub


Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	
	With frmThis
		If Row > 0 and Col > 1 Then		
			sprShtToFieldBinding Col,Row
			If Col = 4 Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",Row) = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 0
				ELSE
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",Row, 1
				End If 
			End If
		elseif Row = 0 and Col = 1 Then
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

Sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	DIM strYEARMON
	DIM dblSEQ
	
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		else
			strYEARMON =  mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
			dblSEQ =  mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
			
			if dblSEQ = "" then
				gErrorMsgBox "���� ���� ���� �����ʹ� �޷��� Ȯ�� �ϽǼ� �����ϴ�","��ȸ�ȳ�"
				exit sub
			else
				vntInParams = array(strYEARMON, dblSEQ) '<< �޾ƿ��°��
				
				vntRet = gShowModalWindow("SCCOPTLISTPOP.aspx",vntInParams , 813,545)
				gFlowWait meWAIT_ON
				SelectRtn
				gFlowWait meWAIT_OFF
			end if 
			
		End If
	End With
End Sub

'��Ʈ�� �������ѷο��� ������ ��� �ʴ��� ���ε�
Function sprShtToFieldBinding (ByVal Col, ByVal Row)
	With frmThis
		If .sprSht.MaxRows = 0 Then exit function '�׸��� �����Ͱ� ������ ������.
		
		'.txtSTYEARMON.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",Row)
		.txtSEQ.value			=	mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",Row)
		.txtCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTNAME",Row)
		.txtCLIENTCODE.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row)
		.txtGREATNAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"GREATNAME",Row)
		.txtGREATCODE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"GREATCODE",Row)
		.txtBUSINO.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"BUSINO",Row)
		.cmbPT_STATUS.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_STATUS",Row)
		.txtPT_LIST.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_LIST",Row)
		.txtEX_BILL.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_BILL",Row)
		.txtOLDCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"OLDCLIENTNAME",Row)
		.cmbPT_CLASS.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_CLASS",Row)
		.txtEX_CONDITION.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_CONDITION",Row)
		.txtEX_INFO.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"EX_INFO",Row)
		.txtOT_DATE.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"OT_DATE",Row)
		.txtOT_INFO.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"OT_INFO",Row)
		.txtPT_ESTAMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_ESTAMT",Row)
		.txtPT_ACTAMT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_ACTAMT",Row)
		.txtPT_DATE1.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_DATE1",Row)
		.txtPT_DATE2.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_DATE2",Row)
		.txtPT_DATE3.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_DATE3",Row)
		.txtPT_CLIENTNAME1.value=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_CLIENTNAME1",Row)
		.txtPT_CLIENTNAME2.value=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_CLIENTNAME2",Row)
		.txtPT_CLIENTNAME3.value=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_CLIENTNAME3",Row)
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht,"PT_RESULT",Row)  = "SK�����þ����۴�(��)" THEN
			.rMNC.checked = TRUE
			.rETC.checked = FALSE
			.rDE.checked = FALSE
			rMNC_onclick
			frmThis.txtETCCLIENTNAME.className	= "NOINPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= TRUE
		ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"PT_RESULT",Row)  = "" THEN
			.rMNC.checked = FALSE
			.rETC.checked = FALSE
			.rDE.checked = TRUE
			rDE_onclick
			frmThis.txtETCCLIENTNAME.className	= "NOINPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= TRUE 
		ELSE
			.rMNC.checked = FALSE
			.rETC.checked = TRUE
			.rDE.checked = FALSE
			.txtETCCLIENTNAME.value = mobjSCGLSpr.GetTextBinding(.sprSht,"PT_RESULT",Row)
			
			frmThis.txtETCCLIENTNAME.className	= "INPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= false 
		END IF
		
		.txtDEPT_NAME.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_NAME",Row)
		.txtDEPT_CD.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"DEPT_CD",Row)
		.txtEXCLIENTNAME.value	=	mobjSCGLSpr.GetTextBinding(.sprSht,"EXCLIENTNAME",Row)
		.txtPT_TEXT.value		=	mobjSCGLSpr.GetTextBinding(.sprSht,"PT_TEXT",Row)
		
	end with   
	Call gFormatNumber(frmThis.txtEX_BILL,0,True)
	Call gFormatNumber(frmThis.txtPT_ESTAMT,0,True)
	Call gFormatNumber(frmThis.txtPT_ACTAMT,0,True)
End Function

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjSCCOPTLIST		= gCreateRemoteObject("cSCCO.ccSCCOPTLIST")
	set mobjSCCOGET			= gCreateRemoteObject("cSCCO.ccSCCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    Dim strPT_CLASS
    strPT_CLASS = "" & vbTab & "1 ���" & vbTab & "2 ���" & vbTab & "3 ���" & vbTab & "4 ���" & vbTab & "5 ���" & vbTab & "6 ���" & vbTab & "7 ���" & vbTab & "8 ���" & vbTab & "9 ���" & vbTab & "10 ���" & vbTab & "��Ÿ"
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 30, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME  | GREATCODE | GREATNAME | BUSINO | PT_STATUS | PT_LIST | EX_BILL | OLDCLIENTNAME | PT_CLASS | EX_CONDITION | EX_INFO | OT_DATE | OT_INFO | PT_ESTAMT | PT_ACTAMT | PT_DATE1 | PT_DATE2 | PT_DATE3 | PT_CLIENTNAME1 | PT_CLIENTNAME2 | PT_CLIENTNAME3 | PT_RESULT | DEPT_CD | DEPT_NAME | EXCLIENTNAME | PT_TEXT"
		mobjSCGLSpr.SetHeader .sprSht,		 "����|���|����|�������ڵ�|�����ָ�|����ó�ڵ�|����ó��|����ڵ�Ϲ�ȣ|PT_����|PT_ǰ��|�������|����������|�ſ���|��������|��������|0T�Ͻ�|OT����|PT����|PT��������|PT1���Ͻ�|PT2���Ͻ�|PT3���Ͻ�|PT 1��������|PT 2��������|PT 3��������|PT���|���μ��ڵ�|���μ���|CU/���ֻ��|PT_��ȹ��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", " 4|   8|   3|         0|      10|         0|      10|            12|      8|     12|      10|          10|       8|      12|      12|     8|    10|    10|          10|        8|        8|        8|          10|          10|          10|    10|           0|         8|         10|       12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "PT_STATUS", -1, -1, "����" & vbTab & "�ܵ�" & vbTab & "ANNUAL" , 10, 60, False, False
		mobjSCGLSpr.SetCellTypeComboBox2 .sprSht, "PT_CLASS", -1, -1, strPT_CLASS , 10, 60, False, False
		
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "YEARMON | OT_DATE | PT_DATE1 | PT_DATE2 | PT_DATE3 ", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, " SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | BUSINO | PT_LIST | EX_BILL | OLDCLIENTNAME | EX_CONDITION | EX_INFO | OT_INFO | PT_ESTAMT | PT_ACTAMT | PT_CLIENTNAME1 | PT_CLIENTNAME2 | PT_CLIENTNAME3 | PT_RESULT | DEPT_CD | DEPT_NAME | EXCLIENTNAME | PT_TEXT", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SEQ | EX_BILL | PT_ESTAMT | PT_ACTAMT ", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht, True, "SEQ"
		mobjSCGLSpr.ColHidden .sprSht, " CLIENTCODE | GREATCODE", True
		mobjSCGLSpr.SetCellAlign2 .sprSht, " CLIENTNAME | GREATNAME | BUSINO | OLDCLIENTNAME | DEPT_NAME | EXCLIENTNAME",-1,-1,2,2,False

		.sprSht.style.visibility = "visible"
    End With
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjSCCOPTLIST = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	With frmThis
		.sprSht.MaxRows = 0

		.txtSTYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	'���
		.txtEDYEARMON.value  = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)	'���
		
		'���� PT ��� Ÿȸ�� ����� ��ٴ�.
		frmThis.txtETCCLIENTNAME.className	= "NOINPUT_L" : frmThis.txtETCCLIENTNAME.readOnly	= True 

		.rDE.checked		= TRUE
		.chkATTEND.checked	= FALSE
		.cmbPT_STATUS.value = "����"
		.cmbPT_CLASS.value	= ""
		'Sheet�ʱ�ȭ
		.txtSTYEARMON.focus
	
	End With
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strSTYEARMON
	Dim strEDYEARMON
	Dim strCLIENTCODE
	Dim strCLIENTNAME

	strSTYEARMON = "" : strEDYEARMON = "" : strCLIENTCODE = "" :strCLIENTNAME = "" 
		
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		strSTYEARMON	= .txtSTYEARMON.value 
		strEDYEARMON	= .txtEDYEARMON.value 
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strCLIENTNAME	= .txtCLIENTNAME1.value

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)

		vntData = mobjSCCOPTLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt, _
										   strSTYEARMON, strEDYEARMON, _
										   strCLIENTCODE, strCLIENTNAME)

		If not gDoErrorRtn ("SelectRtn") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)

   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				sprShtToFieldBinding 2, 1
   				
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				InitPageData
   			End If
   		End If
   	end With
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
   	Dim lngCol,lngRow
	Dim strDataCHK
	With frmThis
   		'������ Validation
		'If DataValidation =False Then exit Sub
		'On error resume Next

		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht, "YEARMON | CLIENTNAME | PT_LIST ",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ���/������/ǰ�� �� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		End If

		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | GREATCODE | GREATNAME | BUSINO | PT_STATUS | PT_LIST | EX_BILL | OLDCLIENTNAME | PT_CLASS | EX_CONDITION | EX_INFO | OT_DATE | OT_INFO | PT_ESTAMT | PT_ACTAMT | PT_DATE1 | PT_DATE2 | PT_DATE3 | PT_CLIENTNAME1 | PT_CLIENTNAME2 | PT_CLIENTNAME3 | PT_RESULT | DEPT_CD | DEPT_NAME | EXCLIENTNAME | PT_TEXT")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		intRtn = mobjSCCOPTLIST.ProcessRtn(gstrConfigXml,vntData)
		
		If not gDoErrorRtn ("ProcessRtn") Then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","����ȳ�!"
			SelectRtn
   		End If
   	end With
End Sub

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON, dblSEQ
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim lngchkCnt
	
	lngchkCnt = 0
	strSEQFLAG = False
	
	With frmThis
		If gDoErrorRtn ("DeleteRtn") Then exit Sub
		
		for i = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
					lngchkCnt = lngchkCnt +1
			End If
		Next
		
		If lngchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0

		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 Then
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",i)
				
				If dblSEQ = "" Then
					mobjSCGLSpr.DeleteRow .sprSht,i
				else
					intRtn = mobjSCCOPTLIST.DeleteRtn(gstrConfigXml,strYEARMON,dblSEQ )
					
					If not gDoErrorRtn ("DeleteRtn") Then
						mobjSCGLSpr.DeleteRow .sprSht,i
   					End If
   					strSEQFLAG = True
				End If				
   				intCnt = intCnt + 1
   			End If
		Next
		
		If not gDoErrorRtn ("DeleteRtn") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn
		End If
	End With
	err.clear	
End Sub


'��ȣ�� Ŭ�����Ѵ�.
Sub CleanField (objField1, objField2, objField3)
	if isobject(objField1) then 
		objField1.value = ""
	end if
	if isobject(objField2) then 
		objField2.value = ""
	End If
	if isobject(objField3) then 
		objField3.value = ""
	End If
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="100" background="../../../images/back_p.gIF">
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
											<td class="TITLE">PT_������ ����</td>
										</tr>
									</table>
								</TD>
								<TD height="20" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 246px"
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
						<!--Top Define Table Start-->
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblKey" class="SEARCHDATA" cellSpacing="0" cellPadding="0" width="100%" height="48">
							<TR>
								<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtSTYEARMON,'')"
									width="30">�� ��</TD>
								<TD style="WIDTH: 300px; HEIGHT: 24px" class="SEARCHDATA">
									<INPUT accessKey="NUM" style="WIDTH: 78px; HEIGHT: 22px" id="txtSTYEARMON" class="INPUT"
										title="�����ȸ" maxLength="6" size="7" name="txtSTYEARMON">~ <INPUT accessKey="NUM" style="WIDTH: 78px; HEIGHT: 22px" id="txtEDYEARMON" class="INPUT"
										title="�����ȸ" maxLength="6" size="7" name="txtEDYEARMON"> <INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtSEQ" dataSrc="#xmlBind" class="NOINPUT_L"
										title="����" dataFld="SEQ" readOnly maxLength="6" size="3" name="txtSEQ">
								</TD>
								<TD style="WIDTH: 45px; HEIGHT: 24px; CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)">������</TD>
								<TD style="HEIGHT: 24px" class="SEARCHDATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="�ڵ��"
										maxLength="100" align="left" size="16" name="txtCLIENTNAME1"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
									<INPUT style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT_L" title="�ڵ���ȸ"
										maxLength="6" align="left" name="txtCLIENTCODE1"></TD>
							<TR>
								<TD class="SEARCHDATA" colSpan="4"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
										onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ �˻��մϴ�." align="right"
										src="../../../images/imgQuery.gIF">
								</TD>
							</TR>
						</TABLE>
						<TABLE height="25">
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 20px" class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28"> <!--background="../../../images/TitleBG.gIF"-->
							<TR>
								<TD height="20" width="1000" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
										<tr>
											<td class="TITLE" vAlign="middle"><span style="CURSOR: hand" id="spnHIDDEN" onclick="vbscript:Call Set_TBL_HIDDEN ()"><IMG style="CURSOR: hand" id="imgTableUp" border="0" name="imgTableUp" alt="�ڷḦ �˻��մϴ�."
														align="absMiddle" src="../../../images/imgTableUp.gif"></span>
											</td>
										</tr>
									</table>
								</TD>
								<TD height="28" vAlign="top" align="right">
									<!--Common Button Start-->
									<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
										<TR>
											<TD><IMG style="CURSOR: hand" id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgCho"
													alt="�ڷḦ �ʱ�ȭ�մϴ�." src="../../../images/imgCho.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgREG" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgNew.gif'" border="0" name="imgREG"
													alt="�ű��ڷḦ ���� �մϴ�." src="../../../images/imgNew.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" border="0" name="imgSave"
													alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"></TD>
											<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
													alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF"></TD>
											<!--<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
											-->
											<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
													alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 120px" vAlign="top" align="center">
									<TABLE id="tblHidden" class="DATA" border="0" cellSpacing="1" cellPadding="0" width="100%">
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME,txtCLIENTCODE)"
												width="100">�����ָ�</TD>
											<TD style="WIDTH: 220px" class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="�����ָ�" dataFld="CLIENTNAME" maxLength="100" name="txtCLIENTNAME"> <IMG style="CURSOR: hand" id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtCLIENTCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="�������ڵ�" dataFld="CLIENTCODE" maxLength="6" size="3" name="txtCLIENTCODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtBUSINO,'')"
												width="100">����ڹ�ȣ</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtBUSINO" dataSrc="#xmlBind" class="INPUT_L"
													title="����� ��ȣ" dataFld="BUSINO" maxLength="50" name="txtBUSINO"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtGREATNAME,txtGREATCODE)">����ó��</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtGREATNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="����ó��" dataFld="GREATNAME" maxLength="100" name="txtGREATNAME"> <IMG style="CURSOR: hand" id="ImgGREATCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgGREATCODE" align="absMiddle" src="../../../images/imgPopup.gIF">
												<INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtGREATCODE" dataSrc="#xmlBind"
													class="INPUT_L" title="����ó�ڵ�" dataFld="GREATCODE" maxLength="6" size="3" name="txtGREATCODE"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtPT_LIST,'')">PT_ǰ��</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtPT_LIST" dataSrc="#xmlBind" class="INPUT_R"
													title="PT_ǰ��" dataFld="PT_LIST" maxLength="50" name="txtPT_LIST"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(cmbPT_STATUS,'')">PT_����</TD>
											<TD class="DATA"><SELECT style="WIDTH: 120px" id="cmbPT_STATUS" dataSrc="#xmlBind" title="PT_����" dataFld="PT_STATUS"
													name="cmbPT_STATUS">
													<OPTION selected value="����">����</OPTION>
													<OPTION value="�ܵ�">�ܵ�</OPTION>
													<OPTION value="ANNUAL">ANNUAL</OPTION>
												</SELECT>
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtOLDCLIENTNAME,'')">���� 
												�����</TD>
											<TD class="DATA"><INPUT accessKey=",M" style="WIDTH: 199px; HEIGHT: 22px" id="txtOLDCLIENTNAME" dataSrc="#xmlBind"
													class="INPUT_L" title="���� �����" dataFld="OLDCLIENTNAME" maxLength="6" size="3" name="txtOLDCLIENTNAME"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEX_BILL,'')">���� 
												����</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtEX_BILL" dataSrc="#xmlBind" class="INPUT_R"
													title="���� ����" dataFld="EX_BILL" maxLength="50" name="txtEX_BILL"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEX_CONDITION,'')">���� 
												����</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtEX_CONDITION" dataSrc="#xmlBind" class="INPUT_L"
													title="���� ����" dataFld="EX_CONDITION" maxLength="50" name="txtEX_CONDITION"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(cmbPT_CLASS,'')">�ſ� 
												���</TD>
											<TD class="DATA"><SELECT style="WIDTH: 120px" id="cmbPT_CLASS" dataSrc="#xmlBind" title="�ſ� ���" dataFld="PT_CLASS"
													name="cmbPT_CLASS">
													<OPTION selected value=""></OPTION>
													<OPTION value="1 ���">1 ���</OPTION>
													<OPTION value="2 ���">2 ���</OPTION>
													<OPTION value="3 ���">3 ���</OPTION>
													<OPTION value="4 ���">4 ���</OPTION>
													<OPTION value="5 ���">5 ���</OPTION>
													<OPTION value="6 ���">6 ���</OPTION>
													<OPTION value="7 ���">7 ���</OPTION>
													<OPTION value="8 ���">8 ���</OPTION>
													<OPTION value="9 ���">9 ���</OPTION>
													<OPTION value="10 ���">10 ���</OPTION>
													<OPTION value="��Ÿ">��Ÿ</OPTION>
												</SELECT>
											</TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtOT_DATE,'')">O/T 
												�Ͻ�</TD>
											<TD class="DATA"><INPUT accessKey="DATE,M" style="WIDTH: 123px; HEIGHT: 22px" id="txtOT_DATE" dataSrc="#xmlBind"
													class="INPUT" title="O/T �Ͻ�" dataFld="OT_DATE" maxLength="10" size="16" name="txtOT_DATE">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar4" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar4" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEX_INFO,'')">���� 
												����</TD>
											<TD class="DATA" colSpan="3"><INPUT style="WIDTH: 526px; HEIGHT: 22px" id="txtEX_INFO" dataSrc="#xmlBind" class="INPUT_L"
													title="���� ����" dataFld="EX_INFO" maxLength="50" name="txtEX_INFO"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtOT_INFO,'')">O/T 
												����</TD>
											<TD class="DATA" colSpan="3"><INPUT style="WIDTH: 526px; HEIGHT: 22px" id="txtOT_INFO" dataSrc="#xmlBind" class="INPUT_L"
													title="O/T ����" dataFld="OT_INFO" maxLength="50" name="txtOT_INFO"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtPT_ESTAMT,'')">PT 
												����</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtPT_ESTAMT" dataSrc="#xmlBind" class="INPUT_R"
													title="PT ����" dataFld="PT_ESTAMT" maxLength="50" name="txtPT_ESTAMT"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtPT_ACTAMT,'')">PT 
												�� ������</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtPT_ACTAMT" dataSrc="#xmlBind" class="INPUT_R"
													title="PT �� ������" dataFld="PT_ACTAMT" maxLength="50" name="txtPT_ACTAMT"></TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtPT_DATE1,txtPT_DATE2,txtPT_DATE3)">PT 
												�Ͻ�</TD>
											<TD class="DATA" colSpan="3">1��:<INPUT accessKey="DATE,M" style="WIDTH: 95px; HEIGHT: 22px" id="txtPT_DATE1" dataSrc="#xmlBind"
													class="INPUT" title="PT �Ͻ�" dataFld="PT_DATE1" maxLength="10" size="16" name="txtPT_DATE1">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar1" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16">
												2��:<INPUT accessKey="DATE,M" style="WIDTH: 95px; HEIGHT: 22px" id="txtPT_DATE2" dataSrc="#xmlBind"
													class="INPUT" title="PT �Ͻ�" dataFld="PT_DATE2" maxLength="10" size="16" name="txtPT_DATE2">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar2" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar2" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16">
												3��:<INPUT accessKey="DATE,M" style="WIDTH: 95px; HEIGHT: 22px" id="txtPT_DATE3" dataSrc="#xmlBind"
													class="INPUT" title="PT �Ͻ�" dataFld="PT_DATE3" maxLength="10" size="16" name="txtPT_DATE3">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndar3" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndar3" align="absMiddle" src="../../../images/btnCalEndar.gIF" height="16">
												���� <INPUT id="chkATTEND" title="����" type="checkbox" name="chkATTEND" CHECKED>
											</TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call CleanField(txtPT_CLIENTNAME1,txtPT_CLIENTNAME2,txtPT_CLIENTNAME3)">PT 
												������</TD>
											<TD class="DATA" colSpan="3">1��:<INPUT style="WIDTH: 115px; HEIGHT: 22px" id="txtPT_CLIENTNAME1" dataSrc="#xmlBind" class="INPUT_R"
													title="PT ������" dataFld="PT_CLIENTNAME1" maxLength="50" name="txtPT_CLIENTNAME1">
												2��:<INPUT style="WIDTH: 115px; HEIGHT: 22px" id="txtPT_CLIENTNAME2" dataSrc="#xmlBind" class="INPUT_R"
													title="PT ������" dataFld="PT_CLIENTNAME2" maxLength="50" name="txtPT_CLIENTNAME2">
												3��:<INPUT style="WIDTH: 115px; HEIGHT: 22px" id="txtPT_CLIENTNAME3" dataSrc="#xmlBind" class="INPUT_R"
													title="PT ������" dataFld="PT_CLIENTNAME3" maxLength="50" name="txtPT_CLIENTNAME3">
											</TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtETCCLIENTNAME,txtETCCLIENT)">PT 
												���</TD>
											<TD class="DATA" colSpan="3">
												���� <INPUT id="rDE" value="DEALY" CHECKED type="radio" name="chkFLAG">&nbsp;&nbsp;
												M&amp;C <INPUT id="rMNC" value="MNC"  type="radio" name="chkFLAG">&nbsp;&nbsp; 
												Ÿȸ��<INPUT id="rETC" value="ETC" type="radio" name="chkFLAG"> <INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtETCCLIENTNAME" dataSrc="#xmlBind" class="NOINPUT_L"
													title="Ÿȸ��" dataFld="ETCCLIENTNAME" maxLength="100" name="txtETCCLIENTNAME"> 
												</TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtDEPT_CD,txtDEPT_NAME)">M&amp;C 
												������</TD>
											<TD class="DATA"><INPUT style="WIDTH: 123px; HEIGHT: 22px" id="txtDEPT_NAME" dataSrc="#xmlBind" class="INPUT_L"
													title="M&amp;C ������" dataFld="DEPT_NAME" maxLength="100" name="txtDEPT_NAME">
												<IMG style="CURSOR: hand" id="imgDEPT_CD" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="imgDEPT_CD"
													align="absMiddle" src="../../../images/imgPopup.gIF"> <INPUT accessKey=",M" style="WIDTH: 53px; HEIGHT: 22px" id="txtDEPT_CD" dataSrc="#xmlBind"
													class="INPUT_L" title="M&amp;C �������ڵ�" dataFld="DEPT_CD" maxLength="6" size="3" name="txtDEPT_CD"></TD>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtEXCLIENTNAME,'')">CU����/���ֻ�</TD>
											<TD class="DATA"><INPUT style="WIDTH: 199px; HEIGHT: 22px" id="txtEXCLIENTNAME" dataSrc="#xmlBind" class="INPUT_L"
													title="CU����/���ֻ�" dataFld="EXCLIENTNAME" maxLength="100" name="txtEXCLIENTNAME">
											</TD>
										</TR>
										<TR>
											<TD style="CURSOR: hand" class="LABEL" onclick="vbscript:Call gCleanField(txtPT_TEXT,'')">PT��ȹ��/���۹�</TD>
											<TD class="DATA" colSpan="3"><INPUT style="WIDTH: 526px; HEIGHT: 22px" id="txtPT_TEXT" dataSrc="#xmlBind" class="INPUT_L"
													title="PT��ȹ��/���۹�" dataFld="PT_TEXT" maxLength="50" name="txtPT_TEXT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--Input End-->
							<!--BodySplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="BODYSPLIT"></TD>
							</TR>
							<!--BodySplit End-->
						</TABLE>
						<TABLE id="tblSheet" border="0" cellSpacing="0" cellPadding="0" width="100%" height="65%">
							<TR>
								<td style="WIDTH: 100%; HEIGHT: 100%" class="DATA" vAlign="top" align="center">
									<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="31855">
										<PARAM NAME="_ExtentY" VALUE="13387">
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
										<PARAM NAME="MaxCols" VALUE="44">
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
								</td>
							</TR>
							<TR>
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
