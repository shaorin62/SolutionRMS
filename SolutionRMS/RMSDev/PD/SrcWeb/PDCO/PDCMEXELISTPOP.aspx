<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMEXELISTPOP.aspx.vb" Inherits="PD.PDCMEXELISTPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�������� ���</title>
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
'HISTORY    :1) 2003/04/29 By Kwon Hyouk Jin
'			 2) 2003/07/25 By Kim Jung Hoon
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
Const meTAB = 9
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjPDCMEXE, mobjPDCMGET
Dim mstrCheck
Dim mALLCHECK
Dim mstrChk
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

Sub imgNew_onclick
	InitPageData
End Sub

Sub imgDelete_onclick
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "�Ϸ�� ó���� �Ϸ��� ������ �����մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "�˻��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn_ALL
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "�Ϸ�� ó���� �Ϸ��� ������ �����մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	End with
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgAddRow_onclick ()
with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "�Ϸ�� ó���� �Ϸ��� ������ �����մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "�˻��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
End with
call sprSht_Keydown(meINS_ROW, 0)
End Sub
Sub imgDelRow_onclick()
	with frmThis
	If .txtENDDAY.value <> "" Then
		gErrorMsgbox "�Ϸ�� ó���� �Ϸ��� ������ �����մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	End with
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
	mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub
Sub ImgAccInput_onclick()
	Dim vntInParams
	Dim vntRet
	Dim vntData
	Dim strGUBN
	with frmThis
	
	
	If .txtJOBNO.value = "" Then
		gErrorMsgbox "���۹�ȣ ��ȸ�� �Է� ���� �մϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	If .txtENDDAY.value <> "" Then
	strGUBN = "END"
	Else
	strGUBN = ""
	End If
	vntData = mobjPDCMEXE.SelectRtn_ACCEXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		If mlngRowCnt = 0 Then
			ProcessRtn_SUB
		End If
		vntInParams = array(.txtJOBNO.value,strGUBN)
		vntRet = gShowModalWindow("PDCMACCLIST.aspx",vntInParams , 788,540)
		SelectRtn
	End If
	
	
	End with
End Sub

Sub sprSht_Keydown(KeyCode, Shift)
'JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME
Dim strSEQ
Dim strJOBNO
Dim strPREESTNO
Dim strITMECODESEQ
Dim strITEMCODE
Dim strITEMCLASS
Dim strITEMNAME
Dim strADDFLAG
Dim intRtn
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 12 Then
		
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
		DefaultValue
		end if
	Else
	strSEQ = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"SORTSEQ",frmThis.sprSht.ActiveRow) 
	strSEQ = strSEQ+1
	strJOBNO = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"JOBNO",frmThis.sprSht.ActiveRow) 
	strPREESTNO = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"PREESTNO",frmThis.sprSht.ActiveRow) 
	strITMECODESEQ = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCODESEQ",frmThis.sprSht.ActiveRow)  
	strITEMCODE = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCODE",frmThis.sprSht.ActiveRow)  
	strITEMCLASS = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMCLASS",frmThis.sprSht.ActiveRow) 
	strITEMNAME = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ITEMNAME",frmThis.sprSht.ActiveRow) 
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: Call DefaultValue(strSEQ,strJOBNO,strPREESTNO,strITMECODESEQ,strITEMCODE,strITEMCLASS,strITEMNAME)
						
				Case meDEL_ROW: DeleteRtn
		End Select

	End if
End Sub

Sub DefaultValue(ByVal strSEQ,ByVal strJOBNO,ByVal strPREESTNO,ByVal strITMECODESEQ,ByVal strITEMCODE, ByVal strITEMCLASS,ByVal strITEMNAME)
	Dim intCnt
	with frmThis

	
	For intCnt = 1 To .sprSht.MaxRows
		if cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",intCnt)) >= strSEQ Then 
			mobjSCGLSpr.SetTextBinding .sprSht,"SORTSEQ",intCnt,cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",intCnt))+1
		End if
	Next
	mobjSCGLSpr.SetTextBinding .sprSht,"SORTSEQ",.sprSht.ActiveRow, strSEQ
	mobjSCGLSpr.SetTextBinding .sprSht,"JOBNO",.sprSht.ActiveRow, strJOBNO
	mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, strPREESTNO
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODESEQ",.sprSht.ActiveRow, strITMECODESEQ
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",.sprSht.ActiveRow, strITEMCODE
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCLASS",.sprSht.ActiveRow, strITEMCLASS
	mobjSCGLSpr.SetTextBinding .sprSht,"ITEMNAME",.sprSht.ActiveRow, strITEMNAME
	mobjSCGLSpr.SetTextBinding .sprSht,"ADDFLAG",.sprSht.ActiveRow, "A"
	mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",.sprSht.ActiveRow,"�ڵ弱��"
	mobjSCGLSpr.SetTextBinding .sprSht,"INCOMCODE",.sprSht.ActiveRow,"����ҵ�(3,3%)"
	mobjSCGLSpr.SetSheetSortUser  .sprSht,3,1
	.txtCLIENTNAME.focus()
	.sprSht.Focus()	
	End with 
End Sub
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim strUSERID
	Dim intCnt2
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht1.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	For intCnt2 = 1 To frmThis.sprSht1.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TAXNO",intCnt2) <> "" THEN
			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht1,"TRANSNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "���ݰ�꼭��ȣ�� �����ϴ� ������ ������� �� �����ϴ�.","�μ�ȳ�!"
			Exit Sub
		End If
	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�.
		'md_trans_temp���� ����
		intRtn = mobjPDCMEXE.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht1,"TRANSNO",1)
		
		vntData = mobjPDCMEXE.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
	
		strcntsum = 0
		IF not gDoErrorRtn ("Get_CATVTRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum + mlngRowCnt
			
			for i=1 to 3
				strUSERID = ""
				vntDataTemp = mobjPDCMEXE.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
		Params = strUSERID
		Opt = "A"
		
		gShowReportWindow ModuleDir, ReportName, Params, Opt
				
		window.setTimeout "printSetTimeout", 10000
	
	end with
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMCATVTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtTRANSYEARMON.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
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
	Dim strSPONSOR
	
	with frmThis
		strSPONSOR = "Y"
		
		vntInParams = array(.txtTRANSYEARMON.value, .txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELECSPON", strSPONSOR) 
		vntRet = gShowModalWindow("MDCMTRANSCUSTPOP.aspx",vntInParams , 413,425)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = vntRet(1,0)		  ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			IF vntRet(3,0) = "�Ϸ�" THEN
				window.event.keyCode = meEnter
				txtTRANSNO_onkeydown
			ELSE
				.txtTRANSNO.value = ""
			END IF
			gSetChangeFlag .txtCLIENTCODE             ' gSetChangeFlag objectID	 Flag ���� �˸�
		end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strSPONSOR
   		
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			strSPONSOR = "Y"
			
			vntData = mobjMDCMGET.GetTRANSCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON.value, .txtTRANSNO.value,.txtCLIENTNAME1.value,"ALL","trans", "ELECSPON", strSPONSOR)
			
			if not gDoErrorRtn ("GetTRANSCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(0,0)
					.txtCLIENTNAME1.value = vntData(1,0)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' �ŷ�ó��ȣ�˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgTRU_onclick
	Call TRU_POP()
End Sub

Sub txtTRANSNO_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strTRANSYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strTRANSYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "ELECSPON", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)   ' Code�� ����
					.txtTRANSNO.value = vntData(1,0)		' �ڵ�� ǥ��
					.txtCLIENTCODE.value = vntData(2,0)     ' �ڵ�� ǥ��
					.txtCLIENTNAME1.value = vntData(3,0)    ' �ڵ�� ǥ��
					'Call SelectRtn ()
				Else
					Call TRU_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRU_POP
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	with frmThis
		If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
		strTRANSYEARMON = .txtTRANSYEARMON.value
		End If
		
		vntInParams = array(strTRANSYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value,.txtCLIENTNAME1.value, "trans", "ELECSPON") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("MDCMTRANSPOP.aspx",vntInParams , 423,435	)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = vntRet(2,0)  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = vntRet(3,0)  ' �ڵ�� ǥ��
			'Call SelectRtn ()
     	end if
	End with
	gSetChange
End Sub


'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtENDDAY,frmThis.imgCalEndar,"txtENDDAY_onchange()"
		gSetChange
	end with
End Sub

'�Ϸ�����
Sub txtENDDAY_onchange
	gSetChange
End Sub




'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------

Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
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

Sub txtDEMANDAMT_onfocus
	with frmThis
		.txtDEMANDAMT.value = Replace(.txtDEMANDAMT.value,",","")
	end with
End Sub
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
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

Sub txtPAYMENT_onfocus
	with frmThis
		.txtPAYMENT.value = Replace(.txtPAYMENT.value,",","")
	end with
End Sub

Sub txtPAYMENT_onblur
	with frmThis
		call gFormatNumber(.txtPAYMENT,0,true)
	end with
End Sub

Sub txtINCOM_onfocus
	with frmThis
		.txtINCOM.value = Replace(.txtINCOM.value,",","")
	end with
End Sub
Sub txtINCOM_onblur
	with frmThis
		call gFormatNumber(.txtINCOM,0,true)
	end with
End Sub

Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'txtACCAMT
Sub txtACCAMT_onfocus
	with frmThis
		.txtACCAMT.value = Replace(.txtACCAMT.value,",","")
	end with
End Sub
Sub txtACCAMT_onblur
	with frmThis
		call gFormatNumber(.txtACCAMT,0,true)
	end with
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mALLCHECK = FALSE
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			mALLCHECK = TRUE
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht.MaxRows
				sprSht_Change 1, intcnt
			next
		end if
	end with
End Sub

sub sprSht_DblClick (ByVal Col, ByVal Row)
Dim vntInParams
Dim vntRet
Dim strCONTRACTNO
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If Col = 17 AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row) <> "" Then
				strCONTRACTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"CONTRACTNO",Row)	
				vntInParams = array(strCONTRACTNO)
				vntRet = gShowModalWindow("PDCMCONTRACTPOP.aspx",vntInParams , 1060,900)
			End If
		end if
	end with
end sub

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
				IF Col = 13 Then
					strCode = ""
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",.sprSht.ActiveRow)
					vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName)
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntData(1,0)
						mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					Else
						mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
					End If
					.txtCLIENTNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
					.sprSht.Focus	
					mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
				ELSEIF Col = 14 then
					Payment_changevalue
				END IF
	end with
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
'����ó ��ưŬ��
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	with frmThis
	
		IF Col = 12 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("",mobjSCGLSpr.GetTextBinding( .sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				'SUSUAMT_CHANGEVALUE2
				'BUDGET_AMT_SUM
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtCLIENTNAME.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

'���������� �׸��� ���ҽ� ��� �Լ��� �¿���� �Ҷ� ���
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = 13 Then
			vntInParams = array("",mobjSCGLSpr.GetTextBinding(sprSht,"OUTSNAME",Row))
			vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"OUTSNAME",Row, vntRet(1,0)		
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
			End IF
			
			.txtCLIENTNAME.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub

Sub Payment_changevalue
Dim intCnt
Dim lngAMT
Dim lngAMTSUM
Dim lngDEMANDAMT
Dim lngRATE
Dim lngACCAMT
with frmThis
	lngAMT= 0
	lngAMTSUM = 0
	For intCnt = 1 To .sprSht.MaxRows
		lngAMT = CDBL(mobjSCGLSpr.GetTextBinding( .sprSht,"ADJAMT",intCnt))
		lngAMTSUM = lngAMTSUM + lngAMT
	Next
	lngACCAMT = Replace(.txtACCAMT.value,",","")
	.txtPAYMENT.value = lngAMTSUM
	lngDEMANDAMT = Replace(.txtDEMANDAMT.value,",","")
	If lngACCAMT = "" Then
		lngACCAMT = 0
	End If
	If lngAMTSUM = "" Then
		lngAMTSUM = 0
	End If
	If lngDEMANDAMT = "" Then
		lngDEMANDAMT = 0
	End If
	.txtINCOM.value = lngDEMANDAMT - (lngAMTSUM+lngACCAMT)
	If lngDEMANDAMT = 0 Then
	lngRATE = 0
	Else
	
	lngRATE = gRound(((lngDEMANDAMT-(lngAMTSUM+lngACCAMT))/lngDEMANDAMT)*100,2)
	End if
	.txtRATE.value = lngRATE
	txtPAYMENT_onblur
	txtINCOM_onblur
	
 
End with

End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	Dim vntInParam
	Dim intNo,i
	Dim strComboList
	Dim strComboList2
	'����������ü ����	
	set mobjPDCMEXE	= gCreateRemoteObject("cPDCO.ccPDCOEXE")
	set mobjPDCMGET			= gCreateRemoteObject("cPDCO.ccPDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "268px"
	pnlTab1.style.left= "7px"
	

	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
		strComboList =  "�ڵ弱��" & vbTab & "���ݰ�꼭(10%)" & vbTab & "���ݰ�꼭�Ұ���" & vbTab & "���ݰ�꼭������" & vbTab & "��꼭" & vbTab & "INVOICE" & vbTab & "����ҵ�(3,3%)" & vbTab & "��Ÿ�ҵ�(22%)" & vbTab & "��Ÿ�ҵ�(�ʿ���80%)" & vbTab & "�������(���Ѽ���)" & vbTab & "�������"
		strComboList2 =  "������"
		'******************************************************************
		'�ŷ����� ���� �׸���
		'******************************************************************
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 22, 0, 13
		mobjSCGLSpr.AddCellSpan  .sprSht, 11, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht,   "JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|BTN|OUTSNAME|ADJAMT|STD|VOCHNO|CONTRACTNO|VATCODE|INCOMCODE|ADJDAY|ADDFLAG|SEQ"
		mobjSCGLSpr.SetHeader .sprSht,		   "���۹�ȣ|������ȣ|����|�����׸����|�����׸��ڵ�|��з�|�����׸�|����|�ܰ�|�ݾ�|����ó�ڵ�|����ó|���޾�|����|��ǥ��ȣ|��༭��ȣ|�����ڵ�|�ҵ汸���ڵ�|������|���Ա���|��ȣ"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "       0|       0|   4|           0|           0|    10|14      |7   |9   |11  |       8|2|17    |11    |20  |0       |11        |14        |14          |9     |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SORTSEQ|QTY|PRICE|AMT|ADJAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "ITEMCODESEQ|ITEMCLASS|ITEMNAME", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "OUTSNAME|STD", -1, -1, 255
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "ADJDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "SORTSEQ|QTY|PRICE|AMT|ADJDAY|CONTRACTNO"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "OUTSCODE|CONTRACTNO",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "JOBNO|PREESTNO|ITEMCODESEQ|ITEMCODE|ADDFLAG|SEQ|VOCHNO|INCOMCODE", true
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,18,18,-1,-1,strComboList
		mobjSCGLSpr.SetCellTypeComboBox .sprSht,19,19,-1,-1,strComboList2
		
		
		 		
    End With    
	pnlTab1.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
			case 0 : .txtJOBNO1.value = vntInParam(i)
				
				'case 2 : .txtCLIENTNAME1.value = vntInParam(i)			'��ȸ�߰��ʵ�
				'case 3 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				'case 4 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				'case 5 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjPDCMEXE = Nothing
	set mobjPDCMGET = Nothing
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
		'.txtTRANSYEARMON.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'DateClean
		'.txtDEMANDDAY.value = gNowDate
		'.txtPRINTDAY.value  = gNowDate
		'.sprSht.MaxRows = 0	
		'.sprSht1.MaxRows = 0
		
		'.txtDEMANDDAY.readOnly = "FALSE"
		'.txtDEMANDDAY.className = "INPUT"
		'.imgCalDemandday.disabled = FALSE
	
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
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strDELCODE
	Dim intRtnSave
	Dim vntData2
	Dim strPREESTNO
	Dim intHDR
	
	with frmThis
	
	'On error resume next
	
		if .txtJOBNO.value = "" Then
			gErrorMsgBox "��ȸ�� ���۰�����ȣ�� �����ϴ�.","����ȳ�!"
			Exit Sub
			Else
			strCODE = .txtJOBNO.value 
		End If
		
  		'������ Validation
		if DataValidation =false then exit sub
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ|VATCODE|INCOMCODE")
		strMasterData = gXMLGetBindingData (xmlBind)
		if  not IsArray(vntData) then 
			If gXMLIsDataChanged (xmlBind) Then 'XML ������ �� ����Ȱ��� �ִٸ�
			Else
				gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
				exit sub
			End If
		End If
		'ó�� ������ü ȣ��
		
		
			intRtn = mobjPDCMEXE.ProcessRtn(gstrConfigXml,strMasterData,vntData,strCODE)
				
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ" & intRtn & " �� ����" & mePROC_DONE,"����ȳ�" 
			SelectRtn
  		end if
 	end with
End Sub
'****************************************************************************************
' ������ ó�� - ��� ���� ��� ����� �켱���� �ϱ� ���� ����� �켱�� �����ϵ��� ����
'****************************************************************************************
Sub ProcessRtn_SUB ()
   Dim intRtn
  	dim vntData
	Dim strMasterData
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strDELCODE
	Dim intRtnSave
	Dim vntData2
	Dim strPREESTNO
	Dim intHDR
	with frmThis
		'On error resume next
		strCODE = .txtJOBNO.value 
	
  		'������ Validation
		if DataValidation =false then exit sub
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"JOBNO|PREESTNO|SORTSEQ|ITEMCODESEQ|ITEMCODE|ITEMCLASS|ITEMNAME|QTY|PRICE|AMT|OUTSCODE|OUTSNAME|ADJAMT|STD|VOCHNO|ADJDAY|ADDFLAG|SEQ|VATCODE|INCOMCODE")
		strMasterData = gXMLGetBindingData (xmlBind)
		
		intRtn = mobjPDCMEXE.ProcessRtn(gstrConfigXml,strMasterData,vntData,strCODE)	
		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			SelectRtn
  		end if
 	end with
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		for intCnt = 1 to .sprSht.MaxRows
   		'DIVNAME|CLASSNAME|ITEMCODE,ITEMCODENAME
			if mobjSCGLSpr.GetTextBinding(.sprSht,"OUTSCODE",intCnt) <> "" AND mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt) = "�ڵ弱��" Then 
				gErrorMsgBox intCnt & " ��° ���� �����ڵ� �� Ȯ���Ͻʽÿ�","�Է¿���"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	with frmThis
	strCODE = .txtJOBNO1.value 
		if strCODE = "" Or Len(strCODE) <> 7 Then
			gErrorMsgBox "���۹�ȣ��Ȯ���Ͻʽÿ�.","��ȸ�ȳ�!"
			Exit Sub
		End if
	
	IF not SelectRtn_Head (strCODE) Then Exit Sub

	'��Ʈ ��ȸ
	CALL SelectRtn_Detail (strCODE)
	
	txtSUSUAMT_onblur
	txtCOMMITION_onblur
	txtDEMANDAMT_onblur
	txtPAYMENT_onblur
	txtINCOM_onblur
	txtNONCOMMITION_onblur
	txtACCAMT_onblur
	txtESTAMT_onblur
	End with
	
End Sub

Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	Dim strCODENAME
	SelectRtn_Head = false
	strCODENAME = frmThis.txtJOBNAME1.value 
	'on error resume next
	
	'�ʱ�ȭ
	
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMEXE.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE,strCODENAME)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ JOBNO �� ���Ͽ� Ȯ���������� " & meNO_DATA, ""
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
	dim vntData
	Dim intCnt
	Dim strRows
	Dim intCnt2
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMEXE.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'�ʱ� ���·� ����
		

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt2 = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",intCnt2) <> "" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt2,-1,-1,true
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,intCnt2,11,19,true
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt2,17,17,true
					End If
					If mobjSCGLSpr.GetTextBinding(.sprSht,"VATCODE",intCnt2) = "" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"VATCODE",intCnt2,"�ڵ弱��"
					sprSht_Change 18,intCnt2
					End If
					
				Next
				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
		End with
		
	End IF
End Function

Sub PreSearchFiledValue (strTRANSYEARMON, strCLIENTCODE, strCLIENTNAME)
	frmThis.txtTRANSYEARMON.value = strTRANSYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME1.value = strCLIENTNAME
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************


'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2
	dim strYEARMON
	Dim strSEQ
	Dim strJOBNO
	Dim strSORTSEQ
	Dim lngCnt
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
		'PREESTNO,ITEMCODESEQ
		'���õ� �ڷḦ ������ ���� ����
		intRtn2 = 0
		lngCnt = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)) <> "" Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"ADJDAY",vntData(i)) <> "" Then
					gErrorMsgbox "������ �� �ִ� ����Ȯ�� ���� �����ɼ� �����ϴ�.","�����ȳ�!"
					Exit Sub
				End If
				strJOBNO = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",vntData(i))
				strSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i)))
				strSORTSEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"SORTSEQ",vntData(i)))
				intRtn2 = mobjPDCMEXE.DeleteRtn(gstrConfigXml,strJOBNO, strSEQ,strSORTSEQ)
			End IF
			
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			End IF
		next
		If lngCnt <> 0 Then
		gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
		End If
		
   		If intRtn2 = 0 Then
   		Else
			Payment_changevalue
			DelProc
		End If
		mobjSCGLSpr.DeselectBlock .sprSht
		
	End with
	err.clear
End Sub

Sub DelProc
Dim intHDR
Dim strMasterData
Dim strPREESTNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		intHDR = mobjPDCMEXE.ProcessRtn_DELHDR(gstrConfigXml,strMasterData)
				if not gDoErrorRtn ("ProcessRtn_DELHDR") then
					SelectRtn
				End If
	End with
End Sub
'������� ���
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
		vntInParams = array(trim(.txtJOBNO1.value), trim(.txtJOBNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO1.value = vntRet(0,0) and .txtJOBNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtJOBNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtJOBNAME1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			SelectRtn
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO1.value),trim(.txtJOBNAME1.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO1.value = trim(vntData(0,0))
					.txtJOBNAME1.value = trim(vntData(1,0))
					SelectRtn
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
Sub DeleteRtn_ALL
	Dim intRtn
	dIM strJOBNO
	Dim vntInParams
	Dim vntRet
	Dim vntData
	Dim intCnt
	Dim intCntV
	with frmThis
	


	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	vntData = mobjPDCMEXE.SelectRtn_ACCEXIST(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtJOBNO.value)
	IF not gDoErrorRtn ("SelectRtn_Detail") then
		If mlngRowCnt = 0 Then
			gErrorMsgBox "���� �� �����Ͱ� �����ϴ�.","��ü�����ȳ�!"
			Exit Sub	
		End If
	End If

	if .txtENDDAY.value <> "" Then
		gErrorMsgBox "�Ϸ�� ������� �����ɼ� �����ϴ�.","��ü�����ȳ�!"
		Exit Sub
	End if
	intCntV = 0
	For intCnt =1 To .sprSht.MaxRows
		intCntV = mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"ADJDAY",frmThis.sprSht.ActiveRow)
		If intCntV <> "" Then
			gErrorMsgBox "�������� �����ϴ� ���� ��ü���� �ɼ� �����ϴ�.","��ü�����ȳ�!"
			Exit Sub
		End If
	Next
	
	
	intRtn = gYesNoMsgbox("�ڷḦ ��ü �����Ͻðڽ��ϱ�?" & vbcrlf & "��ü�ڷᰡ �����˴ϴ�.","�ڷ���� Ȯ��")
	IF intRtn <> vbYes then exit Sub
	
	strJOBNO = .txtJOBNO.value 
	intRtn = mobjPDCMEXE.DeleteRtn_ALL(gstrConfigXml,strJOBNO)
	if not gDoErrorRtn ("DeleteRtn_ALL") then
					gOkMsgbox "������ �Ǿ����ϴ�.","�����ȳ�"
					SelectRtn
				End If
	End with 
End Sub


-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0" border="0" >
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD >
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;���� ����</td>
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
										<!---->
										<TABLE id="tblButton1" style="WIDTH: 50px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="50" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
												<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" border="0" name="imgDelete"></TD>
												<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose"></TD>
											</TR>
										</TABLE>
										<!---->
									</TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 93px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtJOBNAME1, txtJOBNO1)"
										width="93">JOB��</TD>
									<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtJOBNAME1" title="�����Ƿڸ� ��ȸ����" style="WIDTH: 266px; HEIGHT: 22px"
											type="text" maxLength="100" align="left" size="38" name="txtJOBNAME1"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" alt="�����Ƿڹ�ȣ�� ��ȸ�մϴ�" src="../../../images/imgPopup.gIF" width="23"
											align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO1" title="�����Ƿڹ�ȣ ��ȸ����" style="WIDTH: 65px; HEIGHT: 22px"
											type="text" maxLength="7" align="left" size="3" name="txtJOBNO1"> <INPUT dataFld="JOBNO" id="txtJOBNO" dataSrc="#xmlBind" type="hidden" name="txtJOBNO"><INPUT dataFld="JOBNOINS" id="txtJOBNOINS" dataSrc="#xmlBind" type="hidden" name="txtJOBNOINS"><INPUT dataFld="PREESTNO" id="txtPREESTNO" dataSrc="#xmlBind" type="hidden" name="txtPREESTNO"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE class="DATA" id="tblKey1" cellSpacing="1" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 95px">������Ʈ��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="PROJECTNM" class="NOINPUT_L" id="txtPROJECTNM" title="������Ʈ��" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPROJECTNM"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">JOB��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 148px"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="���۰Ǹ�" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">��ü�ι�</TD>
									<TD class="SEARCHDATA" style="WIDTH: 142px"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="��ü�ι�" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtJOBGUBN"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 103px">��ü�з�</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="��ü�ι�" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20"></TD>
								</TR>
								<TR>
									<TD class="SEARCHLABEL" style="WIDTH: 95px">������</TD>
									<TD class="SEARCHDATA" style="WIDTH: 155px"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="������" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">�����</TD>
									<TD class="SEARCHDATA" style="WIDTH: 148px"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="�����" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCLIENTSUBNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 106px">�귣��</TD>
									<TD class="SEARCHDATA" style="WIDTH: 142px"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="�귣��" style="WIDTH: 152px; HEIGHT: 22px"
											dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUBSEQNAME"></TD>
									<TD class="SEARCHLABEL" style="WIDTH: 103px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtENDDAY, '')">�����</TD>
									<TD class="SEARCHDATA"><INPUT dataFld="ENDDAY" class="NOINPUT" id="txtENDDAY" title="�Ϸ���" style="WIDTH: 152px; HEIGHT: 22px"
											accessKey="date" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtENDDAY"></TD>
								</TR>
							</TABLE>
							<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
								</TR>
							</TABLE>
							<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0"> <!--background="../../../images/TitleBG.gIF"-->
								<TR>
									<TD align="left" height="20">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;���� ����</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
														src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
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
							<TABLE id="tblBody" style="WIDTH: 100%"  cellSpacing="0" cellPadding="0" border="0" align="left">
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" vAlign="middle" align="left">
										<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
											align="left" border="0">
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">Noncommition</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="����������ұݾ�"
														style="WIDTH: 152px; HEIGHT: 22px" dataSrc="#xmlBind" readOnly type="text" size="20" name="txtNONCOMMITION"></TD>
												<TD class="LABEL" style="WIDTH: 106px">Commition</TD>
												<TD class="DATA" style="WIDTH: 148px"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="���������ұݾ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtCOMMITION"></TD>
												<TD class="LABEL" style="WIDTH: 106px">��������</TD>
												<TD class="DATA" style="WIDTH: 142px"><INPUT dataFld="SUSURATE" class="NOINPUTB_R" id="txtSUSURATE" title="��������" style="WIDTH: 128px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="16" name="txtSUSURATE">&nbsp;(%)</TD>
												<TD class="LABEL" style="WIDTH: 103px">
													������</TD>
												<TD class="DATA"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="�������հ�ݾ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtSUSUAMT"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">
													û���ݾ�</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="DEMANDAMT" class="NOINPUTB_R" id="txtDEMANDAMT" title="û���ݾ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtDEMANDAMT"></TD>
												<TD class="LABEL" style="WIDTH: 106px">���ֺ�</TD>
												<TD class="DATA" style="WIDTH: 148px"><INPUT dataFld="PAYMENT" class="NOINPUTB_R" id="txtPAYMENT" title="���ֺ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtPAYMENT"></TD>
												<TD class="LABEL" style="WIDTH: 106px">������</TD>
												<TD class="DATA" style="WIDTH: 142px"><INPUT dataFld="RATE" class="NOINPUTB_R" id="txtRATE" title="������" style="WIDTH: 128px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="16" name="txtRATE">&nbsp;(%)</TD>
												<TD class="LABEL" style="WIDTH: 103px">
													������</TD>
												<TD class="DATA"><INPUT dataFld="INCOM" class="NOINPUTB_R" id="txtINCOM" title="������" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtINCOM"></TD>
											</TR>
											<TR>
												<TD class="LABEL" style="WIDTH: 95px">
													�����ݾ�</TD>
												<TD class="DATA" style="WIDTH: 155px"><INPUT dataFld="ESTAMT" class="NOINPUTB_R" id="txtESTAMT" title="�����ݾ� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtESTAMT"></TD>
												<TD class="LABEL">
													�����</TD>
												<TD class="DATA"><INPUT dataFld="ACCAMT" class="NOINPUTB_R" id="txtACCAMT" title="��� �հ�" style="WIDTH: 152px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" size="20" name="txtACCAMT"></TD>
												<TD></TD>
												<TD></TD>
												<TD vAlign="bottom" align="right" colSpan="2"><IMG id="ImgAccInput" onmouseover="JavaScript:this.src='../../../images/ImgAccInputOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAccInput.gIF'" height="20" alt="���������" src="../../../images/ImgAccInput.gIF"
														align="absMiddle" border="0" name="ImgAccInput">&nbsp;<IMG id="imgAddRow" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'" height="20" alt="�� �� �߰�" src="../../../images/imgRowAdd.gIF"
														align="absMiddle" border="0" name="imgAddRow">&nbsp;<IMG id="imgDelRow" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'" height="20" alt="�� �� ����" src="../../../images/imgRowDel.gIF"
														align="absMiddle" border="0" name="imgDelRow"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								</TABLE>
								<!--Input End-->
								
						</TD>
					</TR>
					<TR>
						<TD style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="center">
							<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="11536">
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
					<!--List End-->
					<!--BodySplit Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
