<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMDEMANDCONFIRM_USER.aspx.vb" Inherits="PD.PDCMDEMANDCONFIRM_USER" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>û����û</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : RMS/PD/PDCO
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMDEMAND.aspx
'��      �� : SpreadSheet�� �̿��� û����û/JOB����/��ȸ �� ����� ������.
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/08/10 By KimTH
'****************************************************************************************
-->
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- StyleSheet ���� --><LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">	
	
<!--
option explicit
Dim mlngRowCnt, mlngColCnt			'�������� �ο�� �÷� ��ȯ
Dim mobjPDCODEMAND					'û����û �� Control Class
Dim mobjPDCOGET						'���۰��� Control Class
Dim mobjSCCOGET						'��ü���� Control Class
Dim mstrCheck						'��ü ���� �� ���� ������
Dim mstrSelect						'��ȸ���� (������ �̷���ȸ Or ���� �Է´�� ��ȸ)
Dim mlngRowChk						'�ϴܱ׸��� ����߸����̼� ���
Dim mstrDEPTCD						'�α��λ���ںμ�


Dim mlngTaxRowCnt
Dim mlngTaxColCnt
Const meTab = 9
mstrCheck = True					'��ü������ ���� ����	
mstrSelect = false					'��ȸ���� Default Value: �Է´�� ��ȸ

'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgDivDemand_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_HDR
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' ��ɹ�ư
'=========================================================================================
Sub imgQuery_onclick
	with frmThis
		
	End with
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

'�μ� - �ش���� ����
Sub imgPrint_onclick ()
	
End Sub	
Sub imgAgree_onclick
	Data_Confirm("3")
	SelectRtn
	
End Sub

Sub imgAgreeCanCel_onclick
	Data_Confirm("2")
	SelectRtn
End Sub

Sub imgBackProc_onclick
	Data_Confirm("0")
	SelectRtn
End Sub

Sub Chk_False
	Dim intCnt
	with frmThis
		If .sprSht.MaxRows <> 0 Then
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt, "0"	
		Next
		End If
	End with
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowDelUp_onclick
	gFlowWait meWAIT_ON
	DeleteRtnProc
	gFlowWait meWAIT_OFF

End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub imgSave_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
'=========================================================================================
' SpreadSheet �̺�Ʈ 
'=========================================================================================
Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim intcnt
	Dim intRtn
	Dim dblChk
	Dim dblChkSum
	Dim vntData
	Dim intRtnChk
	'mlngRowChk
	
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
			if mstrCheck = True then 
				mstrCheck = False
			elseif mstrCheck = False then 
				mstrCheck = True
			end if
		end if
		'��üŬ���� �������� ���� �ݿ�
		For intCnt = 1 To .sprSht.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, intCnt
		Next
		
		
	end with	
End Sub


Sub sprSht_Change(ByVal Col, ByVal Row)	
	
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
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
		'�ʵ� To ���ε� ����� ����
	End If
		
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
		Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") Then
				strCOLUMN = "DIVAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") Then
				strCOLUMN = "ADJAMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
				strCOLUMN = "CHARGE"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT")) _
				Or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE"))  Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"DIVAMT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"ADJAMT") _
			Or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"CHARGE") Then
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
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM
	Dim lngEXECnt,IntEXEAMT,IntEXEAMTSUM
	Dim lngChCnt,IntChAMT,IntChAMTSUM
	
	With frmThis
		IntAMTSUM = 0
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0	
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"DIVAMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtDIVAMT.value = 0
		else
			.txtDIVAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtDIVAMT,0,True)
		End If
		
		IntEXEAMTSUM = 0
		For lngEXECnt = 1 To .sprSht.MaxRows
			IntEXEAMT = 0	
			IntEXEAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"ADJAMT", lngEXECnt)
			IntEXEAMTSUM = IntEXEAMTSUM + IntEXEAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtADJAMT.value = 0
		else
			.txtADJAMT.value = IntEXEAMTSUM
			Call gFormatNumber(frmThis.txtADJAMT,0,True)
		End If
		
		IntChAMTSUM = 0
		For lngChCnt = 1 To .sprSht.MaxRows
			IntChAMT = 0	
			IntChAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"CHARGE", lngChCnt)
			IntChAMTSUM = IntChAMTSUM + IntChAMT
		Next
		If .sprSht.MaxRows = 0 Then
			.txtCHARGE.value = 0
		else
			.txtCHARGE.value = IntChAMTSUM
			Call gFormatNumber(frmThis.txtCHARGE,0,True)
		End If
	End With
End Sub




sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strSUBNO,strPREESTNO,strPRIJOBNAME,strPROJECTNM,strJOBNAME
	Dim strRow, strCol
	Dim strWith
	Dim strHeight
	Dim strCLIENTCODE,strCLIENTNAME,strCLIENTSUBCODE,strCLIENTSUBNAME,strTIMCODE,strTIMNAME,strSUBSEQ,strSUBSEQNAME,strJOBGUBN,strJOBGUBNNAME,lngCOMMITIONVALUE
	Dim vntInParams
	Dim vntRet
	With frmThis
		If Row = 0 and Col >1 Then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
		
			strWith =  Screen.width
			strHeight =  Screen.height - 100
			
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNO",.sprSht.ActiveRow)
			strSUBNO= mobjSCGLSpr.GetTextBinding( .sprSht,"SEQ",.sprSht.ActiveRow)
			strJOBNAME	= mobjSCGLSpr.GetTextBinding( .sprSht,"JOBNAME",.sprSht.ActiveRow)	
			strPREESTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PREESTNO",.sprSht.ActiveRow)		
			strPRIJOBNAME = mobjSCGLSpr.GetTextBinding( .sprSht,"PRIJOBNAME",.sprSht.ActiveRow)	
			strPROJECTNM = mobjSCGLSpr.GetTextBinding( .sprSht,"PROJECTNM",.sprSht.ActiveRow) 
			strCLIENTNAME = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",.sprSht.ActiveRow) 
			strJOBGUBNNAME  = mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBNNAME",.sprSht.ActiveRow) 
			strCLIENTCODE = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTCODE",.sprSht.ActiveRow)	 
			strTIMCODE =  mobjSCGLSpr.GetTextBinding( .sprSht,"TIMCODE",.sprSht.ActiveRow)	
			strSUBSEQ =  mobjSCGLSpr.GetTextBinding( .sprSht,"SUBSEQ",.sprSht.ActiveRow)
			strJOBGUBN =  mobjSCGLSpr.GetTextBinding( .sprSht,"JOBGUBN",.sprSht.ActiveRow)	
			
			vntInParams = array(strJOBNO,strSUBNO,strJOBNAME,strPREESTNO,strPRIJOBNAME,strPROJECTNM,strCLIENTNAME,strJOBGUBNNAME,strCLIENTCODE,strTIMCODE,strSUBSEQ,strJOBGUBN)
			vntRet = gShowModalWindow("PDCMJOBMST.aspx",vntInParams , strWith,strHeight)
			
		End If
	End With
end sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)

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
	
	'Ű�� �����϶� ���ε�
	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		sprSht_Click frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================

' ������ ȭ�� ������ �� �ʱ�ȭ 
Sub InitPage()
	'����������ü ����	
	set mobjPDCODEMAND	= gCreateRemoteObject("cPDCO.ccPDCODEMAND")
	set mobjPDCOGET	= gCreateRemoteObject("cPDCO.ccPDCOGET")
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'=========================================================================================
		'û����û SHEET 'CHK|YEARMON|JOBNO|SEQ|PREESTNO
		'=========================================================================================
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 30, 0
		mobjSCGLSpr.SpreadDataField .sprSht,  "CHK|DATAYEARMON|YEARMON|JOBNAME|JOBNO|SEQ|PREESTNO|CLIENTNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|DEPTNAME|EMPNAME|DEMANDPERSON|MANAGERNAME|RANKDIV|USENO|MANAGER|CHARGEHISTORY|DELCHK|CLIENTCODE|TIMCODE|SUBSEQ|JOBGUBN|PRIJOBNAME|PROJECTNM|JOBGUBNNAME"
		mobjSCGLSpr.SetHeader .sprSht,		  "����|���ο�û��|û����û��|JOB��|JOBno.|SUBno.|������ȣ|�����ָ�|�����ݾ�|û���ݾ�|����|û������|����|û�����|���μ�|�����|��û��|������|������|��û�ڻ��|�����ڻ��|��ûSEQ|�ݷ�����|�������ڵ�|���ڵ�|�귣���ڵ�|��ü����|��ǥ�Ÿ�|������Ʈ��|��ü���и�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   4|10        |10        |20   |8     |6     |0       |22      |11      |11      |11  |10      |10  |10      |12      |8     |8     |8     |0     |0         |0         |10     |10      |0         |0     |0         |0       |0       |0         |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"	
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "DIVAMT|ADJAMT|CHARGE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "DATAYEARMON|YEARMON|JOBNAME|JOBNO|SEQ|PREESTNO|CLIENTNAME|DIVAMT|ADJAMT|CHARGE|DEMANDFLAG|MEMO|TAXCODE|DEPTNAME|EMPNAME|DEMANDPERSON|MANAGERNAME|MANAGER|CHARGEHISTORY|DELCHK|CLIENTCODE|TIMCODE|SUBSEQ|JOBGUBN|PRIJOBNAME|PROJECTNM|JOBGUBNNAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DATAYEARMON|YEARMON|JOBNO|SEQ|DEMANDFLAG|TAXCODE|EMPNAME|DEMANDPERSON|MANAGERNAME|MEMO",-1,-1,2,2,false '���
		mobjSCGLSpr.SetCellAlign2 .sprSht, "JOBNAME|CLIENTNAME|DEPTNAME",-1,-1,0,2,false '����
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|RANKDIV|USENO|MANAGER", true
		.sprSht.style.visibility = "visible"
		
		.rdT.style.display = "none"
		.rdF.style.display = "none"
		.imgBackProc.style.display = "none"
    End With
	
	InitPageData	
	'SelectRtn
End Sub

Sub EndPage()
	set mobjPDCODEMAND = Nothing
	set mobjPDCOGET = Nothing
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


' ȭ���� �ʱ���� ������ ����

Sub InitPageData
	'��� ������ Ŭ����
	Dim vntData
	
	gClearAllObject frmThis
	'�ʱ� ������ ����
	with frmThis
		.sprSht.maxrows = 0
		.txtYEARMON.value  = MID(gNowDate,1,4) & MID(gNowDate,6,2) '���� �̰����� ��ó �ӽ÷� �׽�Ʈ�� ���� �Ͽ���
		'.txtYEARMON.value = "200910"

	vntData = mobjPDCODEMAND.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
		if mlngRowCnt > 0 Then
		mstrDEPTCD = vntData(0,1)
		end if
   	end if	
	
	rdChecked
	End with
	'���ο� XML ���ε��� ����
	'gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

Sub rdChecked
	with frmThis
		If .rdT.checked = True Then
			.imgAgreeCanCel.style.display = "none"
			.imgAgree.style.display = "inline"
			'.imgBackProc.style.display = "inline"
		Else
			.imgAgree.style.display = "none"
			.imgBackProc.style.display = "none"
			.imgAgreeCanCel.style.display = "inline"
		End If
	End with

End Sub
Sub rdT_onclick
	rdChecked
	SelectRtn
End Sub
Sub rdF_onclick
	rdChecked
	SelectRtn
End Sub

' �׸����޺�
'�ڵ��޺� ����
Sub Get_COMBO_PVALUE (ByVal blnRow)		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",blnRow,blnRow,vntData_TaxCode,,80,,true
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub	
'��ܱ׸��� �޺�
Sub Get_COMBO_UPVALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_UPVALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_UPVALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "DEMANDFLAG",,,vntData_Demand,,80	
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht, "TAXCODE",,,vntData_TaxCode,,80						
			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		
'�ϴܱ׸����޺�
Sub Get_COMBO_VALUE ()		
	Dim vntData_Demand, vntData_TaxCode	
   	Dim i, strCols	
   	Dim intCnt	
   		
	With frmThis	
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData_Demand = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_DEMAND")
		vntData_TaxCode = mobjPDCODEMAND.Get_COMBO_VALUE(gstrConfigXml,mlngRowCnt,mlngColCnt,"PD_TAXCODE")
						

		If not gDoErrorRtn ("Get_COMBO_VALUE") Then 					
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "DEMANDFLAG",,,vntData_Demand,,80		
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht1, "TAXCODE",,,vntData_TaxCode,,80

			mobjSCGLSpr.TypeComboBox = True 						
   		End If    					
   	End With						
End Sub		

'****************************************************************************************
' ������ ó�� 
'****************************************************************************************
'���߰�
Sub imgRowAdd_onclick ()
	with frmThis
		If mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI03" Or mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDFLAG",.sprSht.ActiveRow) = "DI04" Then 
			call sprSht1_Keydown(meINS_ROW, 0)
			mlngRowChk = .sprSht.ActiveRow
		Else 
			gErrorMsgBox "��ܼ��õ� ������ û�������� ���ҳ��� ����� �ƴմϴ�." & vbcrlf & "û�������� Ȯ���Ͻʽÿ�.","���߰�ó���ȳ�"
		End If
	End with
End Sub


Sub sprSht1_Keydown(KeyCode, Shift)

	Dim intRtn
	
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	
	'if KeyCode = meCR  Or KeyCode = meTab Then
	'	if frmThis.sprSht1.ActiveRow = frmThis.sprSht1.MaxRows and frmThis.sprSht1.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht1,"SAVEFLAG")  Then ' ���� frmThis.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(frmThis.sprSht,"DETAIL_BTN")
	'		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(13), cint(Shift), -1, 1)
	'		DefaultValue
	'	End If
	'Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht1, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
				Case meINS_ROW: DefaultValue
						
				Case meDEL_ROW: DeleteRtn
		End Select

	'End If
End Sub

'�űԵ��޵� ���� ����
Sub DefaultValue
	
	
End Sub
'��ȸ
Sub SelectRtn
	Dim vntData
	Dim intCnt
	Dim strGbn

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		mlngTaxRowCnt=clng(0)
		mlngTaxColCnt=clng(0)
		
		If .rdT.checked = True Then
			strGbn = "2"
		Else
			strGbn = "3"
		End If
		
		vntData = mobjPDCODEMAND.SelectRtn_DEMANDPRECONFIRM(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,strGbn)
		
		if not gDoErrorRtn ("SelectRtn") then
			if mlngRowCnt > 0 Then
				mobjSCGLSpr.SetClipbinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				For intCnt = 1 To .sprSht.MaxRows 
					'JOB�� �÷� ����
					If mobjSCGLSpr.GetTextBinding(.sprSht,"RANKDIV",intCnt) Mod 2 = "0" Then
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HF4EDE3, &H000000,False
					Else
						mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&HFFFFFF, &H000000,False
					End If
				Next
				
				Chk_False
   			Else
   				.sprSht.MaxRows = 0
   			end If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		window.setTimeout "AMT_SUM",1	
		.txtSELECTAMT.value = 0
   	end with
	
End Sub


' ����
Sub ProcessRtn ()
	
End Sub

Sub Data_Confirm(byVal strConfirmFlag)
	Dim vntData
	Dim intRtn
	Dim intCnt
	Dim strMSG
	Dim intSaveRtn
	Dim intCnt3
	Dim intCnt4
	Dim strJOBNAME
	Dim strSEQ
	with frmThis
		If strConfirmFlag = "3" Then
			For intCnt3 = 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt3) = "1" Then
					mobjSCGLSpr.SetTextBinding .sprSht,"MANAGER",intCnt3,gstrUsrID
				End If
			Next
		End If
		'�ݷ� �ϰ�� �ݷ� �Ұ��� �׸��� PD_DEMANDRETURN_FUN ���� ���� Y �� ǥ��Ǹ�,�ŷ����� �Ǵ� ���ݰ�꼭, ��ǥ ������ �ؾ� �ݷ��� ���� �ϴ�.
		If strConfirmFlag = "0" Then
			For intCnt4= 1 To .sprSht.MaxRows
				If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt4) = "1" Then 
					If mobjSCGLSpr.GetTextBinding(.sprSht,"DELCHK",intCnt4) = "Y" Then
						strJOBNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"JOBNO",intCnt4)
						strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",intCnt4)
						gErrorMsgBox "JOBNO [" & strJOBNAME & "-" & strSEQ & "] ������ �� �ݷ������ �ƴմϴ�." & vbcrlf & "���õ� �����̿� �� �� �ŷ����� �� Ȯ�� �Ͻʽÿ�.","�ݷ�ó���ȳ�!"
						Exit Sub
					End If
				End If
			Next
		End If
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|YEARMON|JOBNO|SEQ|PREESTNO|USENO|MANAGER")
		
		
		
		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� �����Ͱ� �����ϴ�.","����ó�� �ȳ�"
		End If
		if  not IsArray(vntData)  then
			gErrorMsgBox "����� " & meNO_DATA,"����ó�� �ȳ�"
			Exit Sub
		End If
		
		select case strConfirmFlag
			case "2": strMSG = "�������"
			case "3": strMSG = "����"
			case "0": strMSG = "�ݷ�"
		end select
		
		intSaveRtn = gYesNoMsgbox("�ش絥���͸� " & strMSG & "�Ͻðڽ��ϱ�?","û����û Ȯ��")
		
		IF intSaveRtn <> vbYes then exit Sub
		
		
		intRtn = mobjPDCODEMAND.Data_Confirm(gstrConfigXml,vntData,strConfirmFlag)
		'������ sms �߼�
		Call SMS_SEND()
		
		if not gDoErrorRtn ("Data_Confirm") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "�ڷᰡ " & strMSG & mePROC_DONE,"ó���ȳ�" 		
		End If
	End with

End Sub

'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
   	Dim intCnt
   	Dim dblSumAmt
   	Dim dblAMT
	'On error resume next
	with frmThis
  	
		
   		IF not gDataValidation(frmThis) then exit Function
   		
   		dblSumAmt = 0
		
   		for intCnt = 1 to .sprSht1.MaxRows
   			'Sheet �ʼ� �Է»���
   			
			if mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTCODE",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"JOBNAME",intCnt) = "" Or _
			mobjSCGLSpr.GetTextBinding(.sprSht1,"YEARMON",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� ���� ���Կ��� �� Ȯ���Ͻʽÿ�","�������"
				Exit Function
			End if
			dblAMT = 0
			dblAMT = mobjSCGLSpr.getTextBinding(.sprSht1,"DIVAMT",intCnt) 
			dblSumAmt = dblSumAmt + dblAMT
			'�ݾ� ��������
		next
   		If mobjSCGLSpr.getTextBinding(.sprSht,"DIVAMT",.sprSht.ActiveRow) < dblSumAmt Then
   			gErrorMsgBox "���Ҵ��ݾ��� ���� �����ݾ� �� �ʰ��Ҽ� �����ϴ�","�������"
   			Exit Function
   		End If
   	End with
   	
	DataValidation = true
End Function



-->
		</script>
		<script language="javascript">
		//SMS �߼�
		function SMS_SEND(){
			frmSMS.location.href = "PD_SMS.asp"; 
		}
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" border="0">
				<TR valign="top">
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<TD align="left" width="400" height="20">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
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
											<td class="TITLE">û����û����&nbsp;</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 326px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<!--Wait Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" title="û����û �� �� ���� �մϴ�." style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="80">���ο�û ��</TD>
											<TD class="SEARCHDATA" width="90"><INPUT class="INPUT" id="txtYEARMON" title="��Ͽ�" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
													type="text" maxLength="6" onchange="vbscript:Call gYearmonCheck(txtYEARMON)" size="9" name="txtYEARMON"></TD>
											<TD class="SEARCHDATA">&nbsp;<INPUT id="rdT" title="��û������ȸ" type="radio" CHECKED value="rdT" name="rdGBN">
												&nbsp; <INPUT id="rdF" title="���γ�����ȸ" type="radio" value="rdF" name="rdGBN"></TD>
											<td align="right" ><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery">
												<IMG id="imgAgree" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
													height="20" alt="������ ���� �����մϴ�." src="../../../images/imgAgree.gIF" align="absMiddle"
													border="0" name="imgAgree"> <IMG id="imgBackProc" onmouseover="JavaScript:this.src='../../../images/imgBackProcOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgBackProc.gIF'" height="20" alt="������ ���� �ݷ��մϴ�."
													src="../../../images/imgBackProc.gIF" align="absMiddle" border="0" name="imgBackProc">
												<IMG id="imgAgreeCanCel" onmouseover="JavaScript:this.src='../../../images/imgAgreeCanCelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgreeCanCel.gIF'"
													height="20" alt="������ ���� ������� �մϴ�." src="../../../images/imgAgreeCanCel.gIF" align="absMiddle"
													border="0" name="imgAgreeCanCel"> <IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF"
													align="absMiddle" border="0" name="imgExcel">
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR valign="top">
					<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 10px"><FONT face="����"></FONT></TD>
				</TR>
				<tr>
					<td>
						<table cellSpacing="0" cellPadding="0" width="100%" border="0">
							<tr>
								<td class="TITLE">�� �� : <INPUT class="NOINPUTB_R" id="txtDIVAMT" title="�����ݾ��հ�" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtDIVAMT"> <INPUT class="NOINPUTB_R" id="txtADJAMT" title="û���ݾ��հ�" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtADJAMT">&nbsp;<INPUT class="NOINPUTB_R" id="txtCHARGE" title="�ܾ��հ�" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtCHARGE">&nbsp;<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="HEIGHT: 22px" accessKey="NUM"
										readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
								</td>
							</tr>
						</table>
					</td>
				</tr>
				<TR>
					<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
						<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
							VIEWASTEXT>
							<PARAM NAME="_Version" VALUE="393216">
							<PARAM NAME="_ExtentX" VALUE="40323">
							<PARAM NAME="_ExtentY" VALUE="16325">
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
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblstatus"></TD>
				</TR>
			</TABLE>
			</TD></TR></TBODY></TABLE></form>
		<iframe id="frmSMS" style="DISPLAY: none;WIDTH: 0px;HEIGHT: 0px" name="frmSMS"></iframe> <!--DISPLAY: none; -->
	</body>
</HTML>
