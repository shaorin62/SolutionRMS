<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPROJECTJOBLIST.aspx.vb" Inherits="PD.PDCMPROJECTJOBLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������Ʈ/JOB ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD
'����  ȯ�� : ASP.NET, VB.NET, COM+
'���α׷��� : PDCMPROJECTJOBLIST.aspx
'��      �� : ������Ʈ �� JOB �� ���ÿ� �����Ѵ�.
'�Ķ�  ���� : 
'Ư��  ���� : ���� 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
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
Option explicit
Dim mlngRowCnt, mlngColCnt
Dim mobjPDCOGET , mobjPDCOPONO , mobjPDCOJOBNO, mobjSCCOGET
Dim mstrCheck

CONST meTAB = 9
mstrCheck = true

Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
End Sub

'=============================
' �̺�Ʈ���ν��� 
'=============================
Sub imgQuery_onclick
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn_PROJECT
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_PROJECT
	gFlowWait meWAIT_OFF
End Sub

Sub imgJobDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_DTL
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_PROJECT
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgJobExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExcelExportOption = true 
		mobjSCGLSpr.ExportExcelFile .sprSht_DTL
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgProjectNew_onclick
	call sprSht_PROJECT_Keydown(meINS_ROW, 0)
End Sub

Sub imgDTLNew_onclick
	Dim vntInParams
	Dim vntRet
	Dim strRow, strCol
	'�ű� ���ѱ��
	Dim strPROJECTNO, strPROJECTNM
	Dim strCLIENTNAME
	Dim strSUBSEQNAME
	Dim strGROUPGBN
	Dim strCREDAY
	Dim strCPDEPTNAME
	Dim strCPEMPNAME
	Dim strCLIENTTEAMNAME
	Dim strMEMO
	
	with frmThis
		IF .sprSht_PROJECT.MaxRows > 0  then
			strPROJECTNO	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNO",.sprSht_PROJECT.ActiveRow)
			strPROJECTNM	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNM",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			strSUBSEQNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
			strGROUPGBN		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow)
			
			If strGROUPGBN = "2" Then
				strGROUPGBN = "�׷�"
			Elseif strGROUPGBN = "1" Then
				strGROUPGBN = "��׷�"
			End If
			
			strCREDAY			= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CREDAY",.sprSht_PROJECT.ActiveRow)
			strCPDEPTNAME		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
			strCPEMPNAME		= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
			strCLIENTTEAMNAME	= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
			strMEMO				= mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"MEMO",.sprSht_PROJECT.ActiveRow)
			
			If strPROJECTNO = "" Then
				gErrorMsgBox "������Ʈ�� ������ JOB�� ��� �Ͻʽÿ�.","�Է¾ȳ�"
				Exit Sub
			Else
				vntInParams = array("New", strPROJECTNO, strPROJECTNM, strCLIENTNAME, strSUBSEQNAME, _
									strGROUPGBN, strCREDAY, strCPDEPTNAME, strCPEMPNAME, strCLIENTTEAMNAME, strMEMO)
				
				vntRet = gShowModalWindow("PDCMJOBNONEW.aspx",vntInParams , "1060", "600")
			End If

			sprSht_PROJECT_click 2,.sprSht_PROJECT.ActiveRow
			.txtCLIENTCODE1.focus()
			.sprSht_DTL.Focus
		else
			gErrorMsgBox "JOB����� ������Ʈ�� ��ȸ�ϼ���.","�Է¾ȳ�"
		end if	
	end with
End Sub

Sub imgJobDetail_onclick()
	Dim strJOBNO, strPRONO
	Dim vntInParams
	Dim vntRet
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	strWith =  Screen.width
	strHeight =  Screen.height - 100
	with frmThis
		IF .sprSht_DTL.MaxRows >0  then
			If mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"SEQ",.sprSht_DTL.ActiveRow) = "1" Then
				strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNO",.sprSht_DTL.ActiveRow)
				strPRONO = mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"PROJECTNO",.sprSht_DTL.ActiveRow)
				
				vntInParams = array("Detail",mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNO",.sprSht_DTL.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht_DTL,"JOBNAME",.sprSht_DTL.ActiveRow))
				vntRet = gShowModalWindow("PDCMJOBNODETAIL.aspx",vntInParams , strWith,strHeight)
				strRow = .sprSht_DTL.ActiveRow
				strCol = .sprSht_DTL.ActiveCol
				'���⼭ ���� ���� ���� ȭ�� ȣ��
				.txtCLIENTCODE1.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_DTL.Focus
				
				SelectRtn_DTL(strPRONO)
				mobjSCGLSpr.ActiveCell .sprSht_DTL, strCol, strRow		
			Else
				msgbox "��ǥJOBNO �� �ƴմϴ�.SUBNO �� 1 �� �׸��� �����Ͽ� �ֽʽÿ�"
			End If
		end if	
	end with
End Sub

'=============================
' ��ɹ�ưŬ���̺�Ʈ
'=============================
Sub imgClose_onclick()
    Window_OnUnload()
End Sub

'-----------------------------------------------------------------------------------------
' Project �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'ProjectNO ��ȸ�˾�
Sub ImgPROJECTNO1_onclick
	with frmThis
		'1�� PROJECT ��ȸ   2�� JOBNO��ȸ
		IF .cmbPOPUPTYPE.value = "1" then
			Call PONO_POP()
		else
			Call SEARCHJOB_POP()
		end IF
	
	End with
End Sub
'���� ������List ��������
Sub PONO_POP
	Dim vntRet
	Dim vntInParams
	

	with frmThis
		vntInParams = array(trim(.txtPROJECTNO1.value), trim(.txtPROJECTNM1.value)) '<< �޾ƿ��°��
		vntRet = gShowModalWindow("PDCMPONOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			'.txtCLIENTNAME1.focus()					' ��Ŀ�� �̵�
     	end if
	End with
	gSetChange
End Sub

Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array( trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtPROJECTNO1.value = vntRet(0,0) and .txtPROJECTNM1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtPROJECTNO1.value = trim(vntRet(0,0))  ' Code�� ����
			.txtPROJECTNM1.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub
'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtPROJECTNM1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
		if .cmbPOPUPTYPE.value = "1" Then '������Ʈ �ڵ� ���
			vntData = mobjPDCOGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,1))
					.txtPROJECTNM1.value = trim(vntData(1,1))
				Else
					Call PONO_POP()
				End If
   			end if
		Else
			vntData = mobjPDCOGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		End If
   		
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub




'****************************************************************************************
' �˾� �̺�Ʈ, ������, ��ü��, ��ü��
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
		vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))       ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE1                  ' gSetChangeFlag objectID	 Flag ���� �˸�
		end if
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	if window.event.keyCode = meEnter then
	
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value),"A")
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' �޷�
'****************************************************************************************
'��ȸ��
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub


Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub

Sub cmbPOPUPTYPE_onchange
	with frmThis
		.txtPROJECTNM1.value = ""
		.txtPROJECTNO1.value = ""
	End with
	gSetChange
End Sub
'=============================
'SheetEvent
'=============================
'����Ŭ��
sub sprSht_PROJECT_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_PROJECT, ""
		end if
	end with
end sub

'��Ʈ������ �����
Sub sprSht_PROJECT_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	Dim strDeptCodeName
	Dim strCode
	Dim strCodeName
	Dim vntData
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strCLIENTSUBCODE
	Dim strCLIENTSUBNAME
	Dim strTIMCODE
	Dim strCLIENTTEAMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	
	With frmThis
		'�����
		if Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNAME")  Then
					strCode = ""
					strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
					strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
				
					vntData = mobjPDCOGET.GetPDEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A","",strDeptCodeName)
				
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(2,1)
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(3,1)
						
						mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNO"),frmThis.sprSht_PROJECT.ActiveRow
					Else
						mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
					End If
					.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش� �̰ż�
					.sprSht_PROJECT.Focus	
					If Row <> .sprSht_PROJECT.MaxRows Then
						mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
					Else
						mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
					End IF
		'���μ�
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTNAME")  Then
				strCode = ""
				strDeptCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
				vntData = mobjPDCOGET.GetCC(gstrConfigXml,mlngRowCnt,mlngColCnt,strDeptCodeName)
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(1,0)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTCD"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		'������
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				vntData = mobjSCCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,"A")
				
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(4,1)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTCODE"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTTEAMNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntData = mobjSCCOGET.GetTIMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,strCLIENTCODE,strCLIENTNAME,"",strCodeName)
				
		
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(4,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(5,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(6,1)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"TIMCODE"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		'�귣��
		Elseif  Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQNAME")  Then
				strCode = ""
				strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntData = mobjSCCOGET.Get_BrandInfo(gstrConfigXml,mlngRowCnt,mlngColCnt,"",strCodeName,strCLIENTCODE,strCLIENTNAME)
				
				If mlngRowCnt = 1 Then	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",Row, vntData(0,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",Row, vntData(1,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntData(2,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntData(3,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",Row, vntData(4,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",Row, vntData(5,1)
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntData(8,1)
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntData(9,1)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntData(10,1)
					
					mobjSCGLSpr.CellChanged .sprSht_PROJECT,mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQ"),frmThis.sprSht_PROJECT.ActiveRow
				Else
					mobjSCGLSpr_ClickProc "sprSht_PROJECT", Col, .sprSht_PROJECT.ActiveRow
				End If
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End IF
		End If
	End With
	mobjSCGLSpr.CellChanged frmThis.sprSht_PROJECT, Col, Row
End Sub

'��ư����ó��
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	Dim vntRet, vntInParams
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strCLIENTSUBCODE
	Dim strCLIENTSUBNAME
	Dim strTIMCODE
	Dim strCLIENTTEAMNAME
	Dim strSUBSEQ
	Dim strSUBSEQNAME
	
	
	With frmThis
		'PROJECT �׸���
		If sprSht = "sprSht_PROJECT" Then
			
			'�����
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPEMPNAME") Then
			
				vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTNAME",Row),"",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPEMPNAME",Row))
				
				vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
				
				'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntRet(2,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntRet(3,0)		
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			
			'���μ�
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CPDEPTNAME") Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTNAME",Row))
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			'������
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTNAME") Then
			
				vntInParams = array("",mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CLIENTNAME",Row))
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",Row, vntRet(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",Row, vntRet(1,0)	
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",Row, vntRet(4,0)
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If

			'��
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"CLIENTTEAMNAME") Then
			
				strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTTEAMNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
					
				vntInParams = array("", trim(strCLIENTNAME),"", trim(strCLIENTTEAMNAME) )  '<< �޾ƿ��°��
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(5,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow,  trim(vntRet(6,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			
			'�귣��
			ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"SUBSEQNAME") Then
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow)
				strSUBSEQNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
					
				vntInParams = array("", trim(strSUBSEQNAME),"", trim(strCLIENTNAME))  '<< �޾ƿ��°��
				
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 520,430)
				IF isArray(vntRet) then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(5,0))
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(8,0))
					'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(9,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(10,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, Col,Row
				End IF
				
				.txtFROM.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
				.sprSht_PROJECT.Focus	
				If Row <> .sprSht_PROJECT.MaxRows Then
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row -1
				Else
					mobjSCGLSpr.ActiveCell .sprSht_PROJECT, Col+2, Row
				End If
			End If
		'JOB �׸��� ������
		Elseif sprSht = "sprSht_DTL" Then
		
		
		End If	
	
	End With
End Sub

'����Ŭ��
sub sprSht_DTL_DblClick (ByVal Col, ByVal Row)
	imgJobDetail_onclick
end sub
'������Ʈ ����Ʈ ����Ű ������
Sub sprSht_PROJECT_Keyup(KeyCode, Shift)
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
	
	'Ű�� �����϶� ���ε�
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprSht_PROJECT_Click frmThis.sprSht_PROJECT.ActiveCol,frmThis.sprSht_PROJECT.ActiveRow
	End If
End Sub

'������Ʈ ����Ʈ Ŭ����
Sub sprSht_PROJECT_Click(ByVal Col, ByVal Row)
	Dim intcnt,intCnt2
	Dim strPROJECTNO
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_PROJECT, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
			for intcnt = 1 to .sprSht_PROJECT.MaxRows
				sprSht_PROJECT_Change 1, intcnt
			next
		Else
			'��Ʈ���ε� ������Ʈ-JOB 
			strPROJECTNO = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"PROJECTNO",.sprSht_PROJECT.ActiveRow)
			
			IF strPROJECTNO <> "" Then
				SelectRtn_DTL(strPROJECTNO)
				'JOBNO ����� �ִ� ��� ������Ʈ ��� �� ������Ʈ �� ��, �����Ұ�
				if .sprSht_DTL.MaxRows = 0 Then
					mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT,false,"CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO",Row,Row,false
				Else
					mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT,true,"CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO",Row,Row,false
				End If
			Else
				.sprSht_DTL.MaxRows = 0	
			End If
		end if
	end with
End Sub

'�� �ű�
Sub sprSht_PROJECT_Keydown(KeyCode, Shift)
	Dim intRtn
	Dim vntData
	
	On error resume Next
	
	if KeyCode <> meINS_ROW then exit sub
	
	intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht_PROJECT, cint(KeyCode), cint(Shift), -1, 1)
	
	with frmThis
		'����� ���� ��������
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCCOGET.GetSCEMP(gstrConfigXml,mlngRowCnt,mlngColCnt,gstrUsrID,"","A","","")
		
		if not gDoErrorRtn ("GetSCEMP") then
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CREDAY",.sprSht_PROJECT.ActiveRow,gNowDate
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow,gstrUsrID
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow,vntData(1,1)
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow,vntData(2,1)
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow,vntData(3,1)
			Else
				gErrorMsgBox "����� ������ ���� ���Ͽ����ϴ�." & vbcrlf & "��α��� �Ͽ� �ֽʽÿ�.","�Է¾ȳ�" 
			End If
		End If
	End with
End Sub

'--------------------------------------------------
'��Ʈ ��ưŬ��
'--------------------------------------------------
Sub sprSht_PROJECT_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	Dim strMEDFLAG
	Dim strDel
	Dim strCLIENTSUBCODE , strCLIENTSUBNAME , strCLIENTCODE , strCLIENTNAME,strTIMCODE,strCLIENTTEAMNAME
	Dim strSUBSEQ , strSUBSEQNM
	Dim strCPDEPTCD , strCPDEPTNAME
	Dim strCPEMPNO , strCPEMPNAME
	
	with frmThis

		'������
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CLIENT") Then
		
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CLIENT") then exit Sub
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME))
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTPOP.aspx",vntInParams , 413,435)
		
			if isArray(vntRet) then
				if strCLIENTCODE = vntRet(0,0) and strCLIENTNAME = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
				mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
				mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
			end if
			.txtFrom.focus()
			.sprSht_PROJECT.focus()	
			gSetChange
     	'��
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_TEAM") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_TEAM") then exit Sub
			strTIMCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTTEAMNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow)
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
			
			
			vntInParams = array("", trim(strCLIENTNAME),"", trim(strCLIENTTEAMNAME) ) '<< �޾ƿ��°��
			
			vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOTIMPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if strTIMCODE = vntRet(0,0) and strCLIENTTEAMNAME = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
				if .sprSht_PROJECT.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(1,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow,  trim(vntRet(5,0))
					
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow,  trim(vntRet(6,0))
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
				end if
				.txtFrom.focus()
				.sprSht_PROJECT.focus()					' ��Ŀ�� �̵�
				gSetChange 
     		end if
     	'�귣��
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_BRAND") Then
		
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_BRAND") then exit Sub
				
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow)
				strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow)
				strSUBSEQ = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow)
				strSUBSEQNM = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow)
			
				
				vntInParams = array("", trim(strSUBSEQNM),"", trim(strCLIENTNAME)) '<< �޾ƿ��°��
		
				vntRet = gShowModalWindow("../../../SC/SrcWeb/SCCO/SCCOCUSTSEQPOP.aspx",vntInParams , 520,430)
				if isArray(vntRet) then
					if strSUBSEQ = vntRet(0,0) and strSUBSEQNM = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit

					if .sprSht_PROJECT.ActiveRow >0 Then
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQ",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"SUBSEQNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"TIMCODE",.sprSht_PROJECT.ActiveRow, trim(vntRet(4,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CLIENTTEAMNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(5,0))
								'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(8,0))
								'mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(9,0))
								mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",.sprSht_PROJECT.ActiveRow, trim(vntRet(10,0))
						
								mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
					end if
					.txtFrom.focus()
					.sprSht_PROJECT.focus()
					gSetChange	
     			end if
     	'���μ�
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPDEPT") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPDEPT") then exit Sub
				
				strCPDEPTCD = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow)
				strCPDEPTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
				
				vntInParams = array(trim(strCPDEPTNAME))
				
				vntRet = gShowModalWindow("PDCMDEPTPOP.aspx",vntInParams , 413,440)
				if isArray(vntRet) then
			
					if .sprSht_PROJECT.ActiveRow >0 Then	
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
						mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
						
						mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
					end if
					.txtFrom.focus()
					.sprSht_PROJECT.focus()
					gSetChange	
				end if
		'�����
		ElseIf Col = mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPEMP") Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht_PROJECT,"BTN_CPEMP") then exit Sub
		
			strCPDEPTCD = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow)
			strCPDEPTNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow)
			strCPEMPNO = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow)
			strCPEMPNAME = mobjSCGLSpr.GetTextBinding( .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow)
			
			vntInParams = array("", trim(strCPDEPTNAME), "", trim(strCPEMPNAME)) '<< �޾ƿ��°��
		
			vntRet = gShowModalWindow("PDCMEMPPOP.aspx",vntInParams , 413,435)
			if isArray(vntRet) then
				if strCPEMPNO = vntRet(0,0) and strCPEMPNAME = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
				
				if .sprSht_PROJECT.ActiveRow >0 Then
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTCD",.sprSht_PROJECT.ActiveRow, trim(vntRet(2,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPDEPTNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(3,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNO",.sprSht_PROJECT.ActiveRow, trim(vntRet(0,0))
					mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"CPEMPNAME",.sprSht_PROJECT.ActiveRow, trim(vntRet(1,0))
					
					mobjSCGLSpr.CellChanged .sprSht_PROJECT, .sprSht_PROJECT.ActiveCol,.sprSht_PROJECT.ActiveRow
				end if
				.txtFrom.focus()
				.sprSht_PROJECT.focus()
				gSetChange
     		end if
     	END IF	
	End with
End Sub

Sub InitPage()
    '����������ü ����	
    set mobjPDCOPONO	= gCreateRemoteObject("cPDCO.ccPDCOPONO")
    set mobjPDCOJOBNO	= gCreateRemoteObject("cPDCO.ccPDCOJOBNO")
    set mobjPDCOGET		= gCreateRemoteObject("cPDCO.ccPDCOGET")
    set mobjSCCOGET		= gCreateRemoteObject("cSCCO.ccSCCOGET")
	
   '���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	gSetSheetDefaultColor() 
	with frmThis
		'������Ʈ ����Ʈ ��Ʈ����
		gSetSheetColor mobjSCGLSpr, .sprSht_PROJECT
		mobjSCGLSpr.SpreadLayout .sprSht_PROJECT, 21, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,7, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,9, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,11, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.AddCellSpan  .sprSht_PROJECT,13, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_PROJECT, "CHK | PROJECTNO | CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | CLIENTTEAMNAME | BTN_TEAM | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO"
		mobjSCGLSpr.SetHeader .sprSht_PROJECT,		"����|������Ʈ�ڵ�|�����|������Ʈ��|������|��|�귣��|���μ�|�����|�׷챸��|���|�������ڵ�|���ڵ�|�귣���ڵ�|�μ��ڵ�|���"
		mobjSCGLSpr.SetColWidth .sprSht_PROJECT, "-1","4 |          10|     8|        25|  20|2|18|2|18|2|    15|2|  10|2|      10|  25|         0|     10|  0       |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht_PROJECT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_PROJECT, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_PROJECT, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CLIENT"

		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_BRAND"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CPDEPT"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_CPEMP"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_PROJECT,"..", "BTN_TEAM"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_PROJECT, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_PROJECT, true, "TIMCODE | PROJECTNO | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP"
		mobjSCGLSpr.ColHidden .sprSht_PROJECT, "CLIENTCODE | CPDEPTCD | CPEMPNO | SUBSEQ | TIMCODE", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_PROJECT, "PROJECTNM | CLIENTCODE | CLIENTNAME | SUBSEQ | SUBSEQNAME | GROUPGBN | CREDAY | CPDEPTCD | CPDEPTNAME | CPEMPNO | CPEMPNAME | MEMO | CLIENTTEAMNAME",-1,-1,0,2,false
        mobjSCGLSpr.SetCellAlign2 .sprSht_PROJECT, "PROJECTNO | CPEMPNAME",-1,-1,2,2,false '���
        
   
        '���� Ȯ���̵ǰ� û����û ���ε� JOB LIST
        gSetSheetColor mobjSCGLSpr, .sprSht_DTL
		mobjSCGLSpr.SpreadLayout .sprSht_DTL, 24, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_DTL, "CHK | CREDAY | JOBNAME | PREESTNO | JOBNO | SEQ | JOBGUBN | CREPART | BUDGETAMT | ENDFLAG | CREGUBN | JOBBASE | EMPNO | EMPNAME | DEPTCD | DEPTNAME | EXCLIENTCODE | EXCLIENTNAME | BIGO | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | PROJECTNO | RANKJOB"
		mobjSCGLSpr.SetHeader .sprSht_DTL,		   "����|�Ƿ���|JOB��|Ȯ��������ȣ|JOBNO|SUBNO|��ü�ι�|��ü�з�|����|����|�ű�|����|������ڵ�|�����|������ڵ�|�����|���ۻ��ڵ�|���ۻ�|���|���ǿ�|û����|����|������Ʈ��ȣ|�׷�����"
		mobjSCGLSpr.SetColWidth .sprSht_DTL, "-1", "  4|      8|   25|           0|    8|    5|       8|      10|  10|   6|   6|   6|         0|     6|         0|    15|         0|    15|  17|    8 |    8 |   8  |0           |10"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_DTL, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_DTL, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_DTL, "BUDGETAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_DTL, "", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht_DTL, true, "CREDAY | JOBNAME | PREESTNO | JOBNO | SEQ | JOBGUBN | CREPART | BUDGETAMT | ENDFLAG | CREGUBN | JOBBASE | EMPNO | EMPNAME | DEPTCD | DEPTNAME | EXCLIENTCODE | EXCLIENTNAME | BIGO | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | PROJECTNO | RANKJOB"
		mobjSCGLSpr.ColHidden .sprSht_DTL, "PREESTNO | EMPNO | DEPTCD | EXCLIENTCODE | PROJECTNO | RANKJOB", true 
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "JOBNAME|BIGO",-1,-1,0,2,false ' ����
		mobjSCGLSpr.SetCellAlign2 .sprSht_DTL, "CHK | CREDAY | JOBNO | ENDFLAG | CREGUBN | JOBBASE | EMPNAME | DEPTNAME | EXCLIENTNAME | AGREEMONTH | TRANSYEARMON | CLOSINGMONTH | JOBGUBN | CREPART | SEQ",-1,-1,2,2,false '���
				
        .cmbPOPUPTYPE.value=1	
        
        .sprSht_PROJECT.style.visibility = "visible"
        .sprSht_DTL.style.visibility = "visible"
	end with
	InitPageData
end Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		
		.sprSht_PROJECT.MaxRows = 0
		.sprSht_DTL.maxRows = 0
		DateClean
		
		call COMBO_TYPE()
	End with
	'���ο� XML ���ε��� ����
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub


Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		'.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub

sub COMBO_TYPE()
   	Dim vntGROUPGUBN
    With frmThis   
		On error resume next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntGROUPGUBN = mobjPDCOPONO.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt,"PONOGUBN")  'JOB���� ȣ��

		if not gDoErrorRtn ("COMBO_TYPE") then 

			mobjSCGLSpr.SetCellTypeComboBox2 .sprSht_PROJECT, "GROUPGBN",,,vntGROUPGUBN,,60 
			mobjSCGLSpr.SetTextBinding .sprSht_PROJECT,"GROUPGBN",-1, "1"
			mobjSCGLSpr.TypeComboBox = True 
			 gLoadComboBox .cmbGROUPGUBN, vntGROUPGUBN, False
   		end if    
   	end with     	
End Sub	

Sub EndPage()
	set mobjPDCOPONO = Nothing
	set mobjPDCOGET = Nothing
	set mobjPDCOJOBNO = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_PROJECT.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		'���ݰ�꼭 �Ϸ���ȸ
		vntData = mobjPDCOPONO.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtPROJECTNM1.value),Trim(.txtPROJECTNO1.value),Trim(.txtCLIENTNAME1.value),Trim(.txtCLIENTCODE1.value),"AA",Trim(.cmbPOPUPTYPE.value))
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_PROJECT,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht_PROJECT,meCLS_FLAG
			gWriteText lblstatus_hdr, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
			If mlngRowCnt = 0 Then
				.sprSht_PROJECT.MaxRows = 0	
				.sprSht_DTL.maxRows= 0
			else
				Call sprSht_PROJECT_Click(1,1)
			End If

		End If		
	END WITH
End Sub

Sub SelectRtn_DTL (ByVal strPONO)
	Dim vntData
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
		vntData = mobjPDCOJOBNO.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt, strPONO)
		
		If not gDoErrorRtn ("SelectRtn_DTL") then
			If mlngRowCnt > 0 Then
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht_DTL,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				mobjSCGLSpr.SetFlag  frmThis.sprSht_DTL,meCLS_FLAG
				
				gWriteText lblstatus_dtl, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				If mlngRowCnt < 1 Then  '��ȸ�Ȱ� ������
					frmThis.sprSht_DTL.MaxRows = 0   '�ο츦 0���� �ϰ�
				Else     '��ȸ�Ȱ� ������
					For intCnt = 1 To .sprSht_DTL.MaxRows '��ȸ�� ������ ó������ ������ ���鼭
						'JOB�� �÷� ����
						If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"RANKJOB",intCnt) Mod 2 = "0" Then
							mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
						Else
							mobjSCGLSpr.SetCellShadow .sprSht_DTL, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
						End If
						'�Ƿ��� ��� CHK Lock Ǯ��
						if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ENDFLAG",intCnt)  = "�Ƿ�" Then
							mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,false,"CHK",intCnt,intCnt,false
						Else
							mobjSCGLSpr.SetCellsLock2 .sprSht_DTL,true,"CHK",intCnt,intCnt,false
						End If
					Next
			  End If
			ELSE
				.sprSht_DTL.MaxRows = 0
			END IF
		END IF
	End with
End SUB

'------------------------------------------
' ������ ó��
'------------------------------------------
Sub ProcessRtn_PROJECT ()
    Dim intRtn
  	Dim vntData
  	Dim vntData1
	Dim intRtnSave
	Dim strPROJECTNO
	Dim intCnt
	Dim intEDITCODE
	Dim strPROJECTLIST
	Dim strDataCHK
	Dim lngCol, lngRow
	
	with frmThis
		If .sprSht_PROJECT.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF
   		
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_PROJECT, "PROJECTNM | CLIENTCODE | TIMCODE | SUBSEQ | CPDEPTCD | CPEMPNO | GROUPGBN",lngCol, lngRow, False) 

		If strDataCHK = False Then
			gErrorMsgBox lngRow & " ���� ������Ʈ��/������/��/�귣��/���μ�/�����/�׷챸���� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub		 
		End If
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_PROJECT,"CREDAY | PROJECTNM | CLIENTNAME | BTN_CLIENT | SUBSEQNAME | BTN_BRAND | CPDEPTNAME | BTN_CPDEPT | CPEMPNAME | BTN_CPEMP | GROUPGBN | MEMO | PROJECTNO | CLIENTCODE | SUBSEQ | CPDEPTCD | CPEMPNO | TIMCODE")
		
		if  not IsArray(vntData) then 
			gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
			exit sub
		End If
		
		'if PROJECT_DataValidation =false then exit sub
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strPROJECTNO = ""
		strPROJECTLIST = ""
		
		intRtn = mobjPDCOPONO.ProcessRtnSheet_Insert(gstrConfigXml,vntData, strPROJECTNO, strPROJECTLIST)
		
		if not gDoErrorRtn ("ProcessRtnSheet_Insert") then
			mobjSCGLSpr.SetFlag  .sprSht_PROJECT,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ" & intRtn & " �� ����" & mePROC_DONE,"����ȳ�" 
			SelectRtn

			For intCnt = 1 To .sprSht_PROJECT.MaxRows 
				If strPROJECTNO = mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",intCnt) Then
					intEDITCODE = intCnt 
					Exit For
				End If
			Next
			
			.txtFROM.focus()
			.sprSht_PROJECT.focus()
			mobjSCGLSpr.ActiveCell .sprSht_PROJECT, 2,intEDITCODE
			sprSht_PROJECT_Click 2,intEDITCODE
				
  		end if
  		
  		vntData1 = mobjPDCOPONO.SelectRtn_PROJECTLIST(gstrConfigXml,mlngRowCnt,mlngColCnt,strPROJECTLIST)
  		
  		Dim strIF_CNT : strIF_CNT = 0
		Dim strIF_USER : strIF_USER = "68300"
		Dim strITEMLIST : strITEMLIST = ""
		Dim strHSEQ : strHSEQ = 1		
		Dim IF_GUBUN : IF_GUBUN = "RMS_0011"
		Dim intCol, intRow, i
		
		intCol = ubound(vntData1, 1)
		intRow = ubound(vntData1, 2)
		
		
		For i = 1 To intRow
			strIF_CNT = strIF_CNT + 1
			
			if strIF_CNT = "1" then
				strITEMLIST = strITEMLIST + cstr(strHSEQ) + "|" + _
								vntData1(0,i) + "|" + _
								vntData1(1,i) + "|" + _
								vntData1(2,i) + "|" + _
								vntData1(3,i) + "|" + _
								vntData1(4,i) + "|" + _
								vntData1(5,i) + "|" + _
								vntData1(6,i) + "|" + _
								vntData1(7,i) + "|" + _
								vntData1(8,i)
			else
				strITEMLIST = strITEMLIST + ":" + cstr(strHSEQ) + "|" + _
								vntData1(0,i) + "|" + _
								vntData1(1,i) + "|" + _
								vntData1(2,i) + "|" + _
								vntData1(3,i) + "|" + _
								vntData1(4,i) + "|" + _
								vntData1(5,i) + "|" + _
								vntData1(6,i) + "|" + _
								vntData1(7,i) + "|" + _
								vntData1(8,i)
			end if
		
			strHSEQ = strHSEQ+1
		Next
		
		
		Call Set_WebServer (strIF_CNT, IF_GUBUN, strIF_USER, strITEMLIST)
		
 	end with
End Sub

Function PROJECT_DataValidation ()
	PROJECT_DataValidation = false
   	Dim intCnt
   	Dim intCntChk
   	Dim intChk
   	
	On error resume next
	with frmThis
		intChk= 0
		For intCntChk = 1 To .sprSht_PROJECT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",intCnt) = "" Then
			Else
				intChk = intChk +1
			End If
		Next
		If intChk = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
  			Exit Function
		End If
		
  		for intCnt = 1 to .sprSht_PROJECT.MaxRows
  			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",intCnt) = "" Then
  				'�ʼ� �׸� üũ
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNM",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� ������Ʈ���� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CLIENTCODE",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� �����ִ� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"TIMCODE",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� ���� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"SUBSEQ",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� �귣��� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPDEPTCD",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� ���μ��� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CPEMPNO",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� ������� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
	  			
  				If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"GROUPGBN",intCnt) = "" Then
  					gErrorMsgBox intCnt & " ���� �׷챸���� �ʼ� �Է� �����Դϴ�.","����ȳ�!"
  					Exit Function
  				End If
  			End If
		next
	End with

	PROJECT_DataValidation = true
End Function


'�ڷ����
Sub DeleteRtn_PROJECT ()
	Dim intSelCnt, intRtn, i , intCnt,intCnt2
	Dim vntData
	Dim strPROJECTNO , strCODE
	Dim intDelCount
	Dim intColSum
	
	with frmThis
	

		'����� üũ�Ȱ͸� ����ɼ� �ֵ���
		intColSum = 0
  		for intCnt2 = 1 to .sprSht_PROJECT.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�����ȳ�"
			exit Sub
		End If

		'JOB�� ��ϵǾ� �ִ��� Ȯ��
		If .sprSht_DTL.MaxRows <> 0  Then
			gErrorMsgBox "��ϵ�JOBNO �� �ֽ��ϴ�.","�����ȳ�"
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		

		intDelCount = 0
		'�������� ; ���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_PROJECT.MaxRows to 1 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"CHK",i) = 1 then
				strPROJECTNO = mobjSCGLSpr.GetTextBinding(.sprSht_PROJECT,"PROJECTNO",i)
				'�ڷ� ����
				intRtn = mobjPDCOPONO.DeleteRtn(gstrConfigXml,strPROJECTNO)
				
				IF not gDoErrorRtn ("DeleteRtn") then
					mobjSCGLSpr.DeleteRow .sprSht_PROJECT,i
   				End IF
   			
			End If
   			intDelCount = intDelCount + 1
   			gWriteText lblstatus_hdr, "������ �ڷῡ ���ؼ� " & intDelCount & " ���� ����" & mePROC_DONE	
   		next
			
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht_PROJECT
		SelectRtn
	End with
	err.clear
End Sub


'�ڷ����
Sub DeleteRtn_DTL ()
    Dim intSelCnt, intRtn, i , intCnt,intCnt2
	Dim vntData
	Dim strCODE
	Dim intDelCount
	Dim intColSum
	Dim strENDFLAG
	with frmThis
	
		'����� üũ�Ȱ͸� ����ɼ� �ֵ���
		intColSum = 0
  		for intCnt2 = 1 to .sprSht_DTL.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",intCnt2) = 1  Then 
					intColSum = intColSum + 1
			End if
		next
		
		If intColSum = 0 Then 
			gErrorMsgBox "���õ� �����Ͱ� �����ϴ�.","�����ȳ�"
			exit Sub
		End If
			
		for i = .sprSht_DTL.MaxRows to 1 step -1
			strCODE = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 then
				strENDFLAG = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"ENDFLAG",i)
				If strENDFLAG <> "�Ƿ�" Then
					gErrorMsgBox "[" & i & "��] ��������°� �Ƿڰ� �ƴѰ��� �����ϽǼ� �����ϴ�.","�����ȳ�!"
					Exit Sub
				ELSE
					mlngRowCnt=clng(0) : mlngColCnt=clng(0)
					
					strCODE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"JOBNO",i)
					
					vntData = mobjPDCOJOBNO.GetJOBNOSELECT(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
					If mlngRowCnt <> 0 Then
						gErrorMsgBox "�ش� JOBNO �� �������� �Ǵ� �������곻���� Ȯ���Ͻʽÿ�","ó���ȳ�"
						Exit Sub
					End If
				End If
			End If
   		next
			
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_DTL.MaxRows to 1 step -1
			'Insert Transaction�� �ƴ� ��� ���� ������ü ȣ��
			strCODE = ""
			If mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"CHK",i) = 1 then
				strCODE = mobjSCGLSpr.GetTextBinding(.sprSht_DTL,"JOBNO",i)
				
				intRtn = mobjPDCOJOBNO.DeleteRtn(gstrConfigXml,strCODE)
			End IF
		next	
		'���� ���� ����
		
		mobjSCGLSpr.DeselectBlock .sprSht_DTL
		'SelectRtn
		sprSht_PROJECT_click 2,.sprSht_PROJECT.ActiveRow
		.txtCLIENTCODE1.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.sprSht_DTL.Focus
	End with
	err.clear
	
End Sub

-->
		</SCRIPT>
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
			//EAI ���� ���� �Ǽ� ���̻� ������ ����.
			//window.frames[0].document.forms[0].submit();
		}
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							height="28">
							<TR>
								<td style="WIDTH: 400px" height="28" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="110" background="../../../images/back_p.gIF">
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
											<td id="tblTitleName" class="TITLE">������Ʈ/JOB ����</td>
										</tr>
									</table>
								</td>
								<TD height="28" vAlign="middle" width="640" align="right">
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 350px"
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
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
							<TR>
								<TD style="HEIGHT: 10px" class="TOPSPLIT"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="left">
									<TABLE id="tblKey0" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
										align="left">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call DateClean()" width="80">�����</TD>
											<TD class="SEARCHDATA" width="230"><INPUT accessKey="DATE" style="WIDTH: 80px; HEIGHT: 22px" id="txtFROM" class="INPUT" title="�Ⱓ�˻�(FROM)"
													maxLength="10" size="6" name="txtFROM">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarFROM1" align="absMiddle" src="../../../images/btnCalEndar.gIF"
													height="15">&nbsp;~ <INPUT accessKey="DATE" style="WIDTH: 80px; HEIGHT: 22px" id="txtTO" class="INPUT" title="�Ⱓ�˻�(TO)"
													maxLength="10" size="7" name="txtTO">&nbsp;<IMG style="CURSOR: hand" id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/btnCalEndarOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/btnCalEndar.gIF'" border="0" name="imgCalEndarTO1" align="absMiddle" src="../../../images/btnCalEndar.gIF"
													height="15"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80"><FONT face="����">������</FONT></TD>
											<TD class="SEARCHDATA" width="260"><FONT face="����"><FONT face="����"><INPUT style="WIDTH: 179px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="�ڵ��"
															maxLength="100" size="24" name="txtCLIENTNAME1"></FONT> <IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF">
													<INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT" title="�ڵ��Է�"
														maxLength="6" size="4" name="txtCLIENTCODE1"></FONT></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtPROJECTNO1, txtPROJECTNM1)"
												width="80"><SELECT style="WIDTH: 88px" id="cmbPOPUPTYPE" title="������Ʈ,JOBNO����" name="cmbPOPUPTYPE">
													<OPTION selected value="1">PROJECT</OPTION>
													<OPTION value="2">JOBNO</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA"><FONT face="����"><INPUT style="WIDTH: 142px; HEIGHT: 22px" id="txtPROJECTNM1" class="INPUT_L" title="�ڵ��"
														maxLength="100" size="18" name="txtPROJECTNM1"> <IMG style="CURSOR: hand" id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgPROJECTNO1" align="absMiddle" src="../../../images/imgPopup.gIF">
													<INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtPROJECTNO1" class="INPUT" title="�ڵ�" maxLength="7"
														size="4" name="txtPROJECTNO1"></FONT></TD>
											<td class="SEARCHDATA" width="53"><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery" alt="�ڷḦ �˻��մϴ�." align="right"
													src="../../../images/imgQuery.gIF" height="20"></td>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<tr>
								<td>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="13">
										<TR>
											<TD style="WIDTH: 1040px; HEIGHT: 25px" class="TOPSPLIT"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="20" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="97" background="../../../images/back_p.gIF">
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
														<td class="TITLE">������Ʈ ����Ʈ</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgProjectNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" border="0" name="imgProjectNew"
																alt="�ű��ڷḦ �ۼ��մϴ�." src="../../../images/imgNew.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgProjectSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" border="0" name="imgProjectSave"
																alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF" width="54" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgProjectDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" border="0" name="imgProjectDelete"
																alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" height="20"></td>
														<TD><IMG style="CURSOR: hand" id="imgProjectExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgProjectExcel"
																alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							<!--BodySplit Start-->
							<TR>
								<TD style="WIDTH: 1040px" class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 40%" vAlign="top" align="left">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: visible" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 95%" id="sprSht_PROJECT" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="4709">
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
										<!--/DIV--></DIV>
								</TD>
							</TR>
							<TR>
								<TD id="lblstatus_hdr" class="BODYSPLIT"></TD>
							</TR>
							<TR>
								<TD>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="20" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="67" background="../../../images/back_p.gIF">
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
														<td class="TITLE">JOB ����Ʈ</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" id="tblButton1" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgDTLNew" onmouseover="JavaScript:this.src='../../../images/imgNewOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgNew.gIF'" border="0" name="imgDTLNew"
																alt="�ű��ڷḦ �ۼ��մϴ�." src="../../../images/imgNew.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgJobDetail" onmouseover="JavaScript:this.src='../../../images/imgDetailOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgDetail.gif'" border="0" name="imgJobDetail"
																alt="�ڷḦ �󼼺����մϴ�." src="../../../images/imgDetail.gIF" height="20"></TD>
														<td><IMG style="CURSOR: hand" id="imgJobDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'" border="0" name="imgJobDelete"
																alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" height="20"></td>
														<TD><IMG style="CURSOR: hand" id="imgJobExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgJobExcel"
																alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD style="WIDTH: 1040px" class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 55%" vAlign="top" align="left">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 95%; VISIBILITY: visible" id="pnlTab2"
										ms_positioning="GridLayout">
										<!--DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 1038px; POSITION: relative" ms_positioning="GridLayout"-->
										<OBJECT style="WIDTH: 100%; HEIGHT: 93.13%" id="sprSht_DTL" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="6217">
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
										<!--/DIV--></DIV>
								</TD>
							</TR>
							<!--Bottom Split Start-->
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<TR>
					<TD id="lblStatus_dtl" class="BOTTOMSPLIT"></TD>
				</TR>
				<!--Top TR End--></TABLE>
			<!--Main End--></form>
		</TR></TBODY></TABLE> <iframe id="frmSapCon" style="WIDTH: 100%; DISPLAY: none; HEIGHT: 300px" src="../../../PD/WebService/PROJECTWEBSERVICE.aspx"
			name="frmSapCon"></iframe><!--style="DISPLAY: none"-->
	</body>
</HTML>
