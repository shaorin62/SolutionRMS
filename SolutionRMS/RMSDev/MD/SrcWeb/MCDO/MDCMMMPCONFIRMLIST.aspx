<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMMMPCONFIRMLIST.aspx.vb" Inherits="MD.MDCMMMPCONFIRMLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���̳�MMP û�� ���ο�û</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOCUSTLIST.aspx
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/08 By KTY
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
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
		
<!--
option explicit 
Dim mlngRowCnt, mlngColCnt
Dim mobjMDCOMMPCONFIRMLIST '�����ڵ�, Ŭ����
Dim mobjMDCOGET
Dim mobjSCCOGET
Dim mstrCheck
Dim mstrFlag
CONST meTAB = 9
mstrCheck = True

'====================================================
' �̺�Ʈ ���ν��� 
'====================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	'EndPage
End Sub

'-----------------------------
' �ݱ�
'-----------------------------
Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'---------------------------------------------------
' ��� ��ư Ŭ�� �̺�Ʈ
'---------------------------------------------------

'-----------------------------------
'��ȸ
'-----------------------------------
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------------
'û�� ��û�� �� ���� ���� 
'-----------------------------------
Sub ImgConfirmRequest_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_Confirm_User
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
' ����
'-----------------------------
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	End With
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcelDTR_onclick ()
	gFlowWait meWAIT_ON
	With frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	End With
	gFlowWait meWAIT_OFF
End Sub

'-----------------------------
'�μ� ��ư Ŭ����
'-----------------------------
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim strCNT
	Dim vntData
	Dim intRtn
	Dim strUSERID
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	If frmThis.sprSht_HDR.MaxRows = 0 Then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	End If

	gFlowWait meWAIT_ON
	With frmThis
		if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONFIRMGBN",.sprSht_HDR.ActiveRow) = "" then
			gErrorMsgBox "���ε��� ���������ʹ� ��� �Ͻ� �� �����ϴ�.","��� �ȳ�!"
			exit sub
		end if 
		
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCOMMPCONFIRMLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMMMPCOMMI.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",.sprSht_HDR.ActiveRow)
		'�ѹ��� �Ѱ��� �����͸� ��� �Ѵ�.(�ӽ� �߰��� ������ �Ѱ��� ������ ���� ��)
		strCNT			= 1
		
		strUSERID = ""
		vntData = mobjMDCOMMPCONFIRMLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, strCNT, strUSERID)
		
		Params = strUSERID
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
		'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
		'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
		window.setTimeout "call printSetTimeout('" & strTRANSYEARMON & "', '" & strTRANSNO & "')", 10000
	End With
	gFlowWait meWAIT_OFF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout(strTRANSYEARMON, strTRANSNO)
	Dim intRtn, intRtn2
	With frmThis
		intRtn = mobjMDCOMMPCONFIRMLIST.DeleteRtn_temp(gstrConfigXml)
	End With
End Sub


'-----------------------------------------------------------------------------------------
' ����ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgEMPNO_onclick
	Call EMP_POP()
End Sub

'���� ������List ��������
Sub EMP_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array("", "", trim(.txtEMPNO.value), trim(.txtEMPNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMEMPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtEMPNO.value = vntRet(0,0) and .txtEMPNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
		
			.txtEMPNO.value = trim(vntRet(0,0))
			.txtEMPNAME.value = trim(vntRet(1,0))
			'.txtMEMO.focus()				' ��Ŀ�� �̵�
			gSetChangeFlag .txtEMPNO		' gSetChangeFlag objectID	 Flag ���� �˸�
			gSetChangeFlag .txtEMPNAME
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtEMPNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetMDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtEMPNO.value, .txtEMPNAME.value,"A","","")
			
			if not gDoErrorRtn ("GetMDEMP") then
				If mlngRowCnt = 1 Then
					.txtEMPNO.value = trim(vntData(0,1))
					.txtEMPNAME.value = trim(vntData(1,1))
					gSetChangeFlag .txtEMPNO
				Else
					Call EMP_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
end sub

'--------------------------------------------------
' SpreadSheet �̺�Ʈ
'--------------------------------------------------
Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
	With frmThis
		mobjSCGLSpr.CellChanged .sprSht_HDR, Col, Row
	End With
End Sub


Sub sprSht_DTL_Change(ByVal Col, ByVal Row)
	With frmThis
		
	End With
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht_DTL, Col, Row
End Sub

'-----------------------------------
'��Ʈ Ŭ��
'-----------------------------------
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	With frmThis		
		If Row > 0 Then
			SelectRtn_DTL Col, Row
		End If
	End With
End Sub

'-----------------------------------
'��Ʈ ����Ŭ��
'-----------------------------------
sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		End If
	End With
End sub

sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	With frmThis
		If Row = 0  Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_DTL, ""
		End If
	End With
End sub

Sub sprSht_HDR_Keyup(KeyCode, Shift)
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
		SelectRtn_DTL frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
End Sub

Sub txtCLIENTNAME_onKeyDown
	if window.event.keyCode <> meEnter then Exit Sub
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub
	
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'------------------------------------------------------------------------------------------------------------
Sub InitPage()
' ������ ȭ�� ������ �� �ʱ�ȭ 
'----------------------------------------------------------------------
	'����������ü ����	
	set mobjMDCOMMPCONFIRMLIST	= gCreateRemoteObject("cMDCO.ccMDCOMMPCONFIRMLIST")
	set mobjMDCOGET				= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjSCCOGET				= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor()
    With frmThis
		'MMP ���ʵ����� ��Ʈ
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR	
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 19, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "����|��������|���|����|�������ڵ�|�����ָ�|��ü���ڵ�|��ü���|���μ��ڵ�|���μ���|û�����|�ű԰��Աݾ�|��������|MMPû���ݾ�|���α���|��û��|��û����|������|��������"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1", " 4|       8|   8|   4|         0|      10|         0|      10|           0|         8|      10|          10|       8|         12|       0|     8|       8|     8|       8"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "AMT | MMP_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
		mobjSCGLSpr.colhidden .sprSht_HDR, "CLIENTCODE | REAL_MED_CODE | DEPT_CD | CONFIRMFLAG",true
		mobjSCGLSpr.SetCellAlign2 .sprSht_HDR, " SEQ | CONFIRMGBN " ,-1,-1,2,2,false
		
		
		'MMP ������ ����ü�� ��Ȳ �� ���� ó�� ��Ʈ
		gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 10, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		 "���|����|�������ڵ�|�����ָ�|��ü����|��ü���հ�|��ü���ݾ�|�д���|MMP�ݾ�|��ü���д��"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", " 8|   4|         8|      12|      10|        12|        12|     8|     12|          12"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "AMTSUM | AMT  | HDR_AMT | DIVAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_DTL, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG  ", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT "
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "SEQ | CLIENTCODE ",-1,-1,2,2,False
		mobjSCGLSpr.CellGroupingEach .sprSht_DTL, "AMTSUM | HDR_AMT "
		
		.sprSht_HDR.style.visibility = "visible"
		.sprSht_DTL.style.visibility = "visible"
	
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
End Sub

Sub EndPage()
	set mobjMDCOMMPCONFIRMLIST = Nothing
	set mobjMDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub

'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis

	'�ʱ� ������ ����
	With frmThis
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
	End With
End Sub

'------------------------------------------
' HDR ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON
   	Dim strCLIENTNAME
   	Dim intCnt
   	
	With frmThis
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_DTL.MaxRows = 0
		
		'���� �ʱ�ȭ
		strYEARMON = "" : strCLIENTNAME = "" 
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value 
		strCLIENTNAME	= .txtCLIENTNAME.value 
		
		vntData = mobjMDCOMMPCONFIRMLIST.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCLIENTNAME)

		If not gDoErrorRtn ("SelectRtn_HDR") Then
			mobjSCGLSpr.SetClipbinding .sprSht_HDR, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			
			If mlngRowCnt > 0 Then
				For intCnt = 1 To .sprSht_HDR.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",intCnt) <> "" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, " CONFIRMGBN | YEARMON | SEQ | CLIENTCODE | CLIENTNAME | REAL_MED_CODE | REAL_MED_NAME | DEPT_CD | DEPT_NAME | DEMANDDAY | AMT | RATE | MMP_AMT | CONFIRMFLAG | REQUEST_USER | REQUEST_DATE | CONFIRM_USER | CONFIRM_DATE"
					ELSE
						mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, false, "CHK "
					END IF 
				Next
				
   				gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
	   			
   				Call SelectRtn_DTL(1,1)
   			else
   				gWriteText lblStatus, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				.sprSht_HDR.MaxRows = 0
				.sprSht_DTL.MaxRows = 0
			end if
   		End If
   	End With
End Sub

'------------------------------------------
' DTL ������ ��ȸ
'------------------------------------------
Sub SelectRtn_DTL(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, intCnt, intCnt2
	Dim strYEARMON
	Dim lngSEQ 
	Dim strCLIENTCODE
	Dim strREQUEST_USER
	Dim intAMT, intSUMAMT, intHDR_AMT
	
	With frmThis
		'sprSht2�ʱ�ȭ
		.sprSht_DTL.MaxRows = 0
		strYEARMON = "" : lngSEQ = 0 : strREQUEST_USER = "" : strCLIENTCODE = ""
		intAMT = 0		: intSUMAMT = 0 : intHDR_AMT = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",Row)
		lngSEQ			= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",Row)
		strCLIENTCODE	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTCODE",Row)
		strREQUEST_USER	= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",Row)
		intHDR_AMT		= mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"MMP_AMT",Row)
		
		
		vntData = mobjMDCOMMPCONFIRMLIST.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, _ 
														lngSEQ,strCLIENTCODE,strREQUEST_USER )
														

		If not gDoErrorRtn ("SelectRtn_DTL") Then
			If mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht_DTL, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				
				If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"REQUEST_USER",Row) <> "" Then
					
					mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT"
				
				ELSE
					For intCnt = 1 To .sprSht_DTL.MaxRows
						'���ε��� ������������ SEQ �� HDR�� Ű�� �ִ´�.
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"SEQ",intCnt , lngSEQ
						
						'DTL ������ �հ�ݾ��� ���Ѵ�.
						intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",intCnt)
						intSUMAMT = intSUMAMT + intAMT
					Next
					
					'�ݺ��ϸ� �ϴ��� ������ ����
					for intCnt2 = 1 To .sprSht_DTL.MaxRows
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"AMTSUM",intCnt2 , intSUMAMT
						mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"HDR_AMT",intCnt2 , intHDR_AMT
					next
					
					'����ü�� �ݾ��� �������� ���Ͽ� HDR �ѱݾ��� �������� ��ü�� ����� �ǽ�
					call DIV_RATE (intSUMAMT,intHDR_AMT)
				END IF
			ELSE
				.sprSht_DTL.MaxRows = 0
			End If	
		End If
   		gWriteText lblStatusDTR, mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
	End With
End Sub

SUB DIV_RATE(ByVal intSUMAMT, ByVal intHDR_AMT)
	Dim i
	Dim intAMT
	Dim intRATE
	Dim intDIVAMT
	With frmThis
		
		for i = 1 To .sprSht_DTL.MaxRows
			intAMT = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"AMT",i)
			
			intRATE = (clng(intAMT) / clng(intSUMAMT)) * 100
			
			intDIVAMT = intHDR_AMT * intRATE / 100
			
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"RATE",i , intRATE
			mobjSCGLSpr.SetTextBinding frmThis.sprSht_DTL,"DIVAMT",i , intDIVAMT
		next
		
	End With
END SUB
	
Sub ProcessRtn_Confirm_User ()
   	Dim vntData
   	Dim VNTDATA_DTL
   	Dim intRtn
	Dim strCLIENTCODE, strCLIENTNAME

	Dim intCnt,intCnt2,intCnt3, chkcnt
	Dim intSaveRtn
	Dim strMsg
	Dim strMstMsg

	'SMS ����
	Dim vntData_info
	Dim strFromUserName
	Dim strFromUserEmail
	Dim strFromUserPhone
	Dim strToUserName
	Dim strToUserEmail
	Dim strToUserPhone
	Dim strAMT
	
	with frmThis
		IF .txtEMPNO.value = "" THEN
			gErrorMsgBox "���ο�û�ڸ� �Է� �Ͻʽÿ�",""
			exit sub
		END IF 
		
		chkcnt=0
		For intCnt = 1 To .sprSht_HDR.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "���ο�û�� �����͸� üũ �Ͻʽÿ�",""
			exit sub
		end if
		
		For intCnt2 = 1 To .sprSht_HDR.MaxRows
			'�׸����� ���۰Ǹ� �� �����´�
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",intCnt2) = 1 Then
				 strMsg = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CLIENTNAME",intCnt2)
				 Exit For
			End If
		Next
	
		If chkcnt = 1 Then
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "...] ���ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ���ο�û���ֽ��ϴ�"
			End If
		Else
			If Len(strMsg) > 10 Then
				strMstMsg = "[ " & MID(strMsg,1,10) & "] ��" & chkcnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			Else
				strMstMsg = "[ " & strMsg & "] ��" & chkcnt-1 & "���ǽ��ο�û���ֽ��ϴ�"
			End If
		End If
		
		intSaveRtn = gYesNoMsgbox("�ش絥���͸� ���ο�û �Ͻðڽ��ϱ�?","���ο�û Ȯ��")
		IF intSaveRtn = vbYes then 
		
			'������ �����Ͽ����Ƿ� SMS �߼�
			'������ ����� ���� ��������
			mlngRowCnt=clng(0) : mlngColCnt=clng(0)
			
			vntData_info = mobjSCCOGET.Get_SENDINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtEMPNO.value),Trim(.txtEMPNAME.value))
			
			'�����»������
			strFromUserName		= vntData_info(0,2)
			strFromUserEmail	= vntData_info(1,2)
			strFromUserPhone	= vntData_info(2,2)

			'�޴»�� ����
			strToUserName		=  vntData_info(0,1)
			strToUserEmail		=  vntData_info(1,1)
			strToUserPhone		=  vntData_info(2,1)
		
			call SMS_SEND(strFromUserName,strFromUserPhone,strToUserPhone,strMstMsg)
			
			'�����÷��� ����
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meINS_TRANS
			mobjSCGLSpr.SetFlag  .sprSht_DTL,meINS_TRANS

			'��Ʈ�� ����� �����͸� �����´�.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | YEARMON | SEQ")
			'�ϴ��� �д�� �ݾ׵� �Բ� ���� �Ѵ�.
			vntData_DTL = mobjSCGLSpr.GetDataRows(.sprSht_DTL,"YEARMON | SEQ | CLIENTCODE | CLIENTNAME | MED_FLAG | AMTSUM | AMT | RATE | HDR_AMT | DIVAMT")
			
			intRtn = mobjMDCOMMPCONFIRMLIST.ProcessRtn_Confirm_User(gstrConfigXml,vntData,vntData_DTL, .txtEMPNO.value)

   			if not gDoErrorRtn ("ProcessRtn_Confirm_User") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
				initpagedata	
				gOkMsgBox "���ο�û�� �Ǿ����ϴ�.","Ȯ��"
				
				If intRtn <> 0  Then
					.txtCLIENTNAME1.value = strCLIENTNAME
					selectRtn
				Else
					initpagedata
				End If
				DateClean
   			end if
   		End If
   	end with
End Sub

-->		
		</script>
		<script language="javascript">
		//SMS �߼�
		function SMS_SEND(strFromUserName , strFromUserPhone, strToUserPhone,strMstMsg){
			frmSMS.location.href = "../../../SC/SrcWeb/SCCO/SMS.asp?MSTMSG="+ strMstMsg + "&FromUserName=" + strFromUserName + "&ToUserPhone=" + strToUserPhone + "&FromUserPhone=" + strFromUserPhone; 
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="70" background="../../../images/back_p.gIF"
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
												<td class="TITLE">���̳� MMP ���ʴ����� ���ο�û</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE id="tblWaitP" style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
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
									<TD align="left" width="100%" height="1">
									</TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 100%; HEIGHT: 4px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="left">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" align="left"
											border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 42px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">���</TD>
												<TD class="SEARCHDATA" width="113" style="WIDTH: 113px"><INPUT class="INPUT" id="txtYEARMON" title="���" style="WIDTH : 100px; HEIGHT : 22px" maxLength="100"
														align="left" size="6" name="txtYEARMON" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 53px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTNAME, '')">�����ָ�</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 200px; HEIGHT: 22px"
														maxLength="100" align="left" size="28" name="txtCLIENTNAME"></TD>
												<TD class="SEARCHDATA" width="50">
													<TABLE cellSpacing="0" cellPadding="2" align="right" border="0">
														<TR>
															<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
																	height="20" alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								<tr>
									<td>
										<table class="DATA" height="10" cellSpacing="0" cellPadding="0" width="100%">
											<TR>
												<TD style="WIDTH: 100%; HEIGHT: 4px"></TD>
											</TR>
										</table>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD style="FONT-SIZE: 12px; FONT-WEIGHT: bold"><span id="title2" onclick="vbscript:Call gCleanField(txtEMPNAME, txtEMPNO)" style="CURSOR: hand">������:</span>
																&nbsp;<INPUT class="INPUT_L" id="txtEMPNAME" title="�����ȸ" style="WIDTH: 62px; HEIGHT: 22px" maxLength="255"
																	align="left" size="5" name="txtEMPNAME"> <IMG id="ImgEMPNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
																	border="0" name="ImgEMPNO" title="���α��ڼ���"> <INPUT class="INPUT" id="txtEMPNO" title="�����ȸ" style="WIDTH: 46px; HEIGHT: 22px" readOnly
																	maxLength="8" align="left" size="2" name="txtEMPNO">&nbsp; <IMG id="ImgConfirmRequest" onmouseover="JavaScript:this.src='../../../images/ImgConfirmRequestOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmRequest.gIF'" height="20" alt="���õ� �����͸� ���ο�û�մϴ�." src="../../../images/ImgConfirmRequest.gIF"
																	align="absMiddle" border="0" name="ImgConfirmRequest"></TD>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
														</TR>
													</TABLE>
													<!--Common Button End-->
												</TD>
											</TR>
										</TABLE>
									</td>
								</tr>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
											ms_positioning="GridLayout">
											<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_HDR" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="7117">
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
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
								</TR>
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="top" align="center">
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20"></TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButtonDTR" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="���ε� �����͸� ����մϴ�." src="../../../images/imgPrint.gIF" border="0" name="imgPrint"></TD>
															<TD><IMG id="imgExcelDTR" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcelDTR"></TD>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"></TD>
								</TR>
								<!--Input End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="center">
										<DIV id="pnlTab2" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden"
											ms_positioning="GridLayout">
											<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="31829">
												<PARAM NAME="_ExtentY" VALUE="7117">
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
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatusDTR" style="WIDTH: 100%"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
				</TBODY>
			</TABLE>
		</FORM>
		<iframe id="frmSMS" style="WIDTH: 500px;DISPLAY: none;HEIGHT: 500px" name="frmSMS"></iframe> <!-- -->
	</body>
</HTML>
