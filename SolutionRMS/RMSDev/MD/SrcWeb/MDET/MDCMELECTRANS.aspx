<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMELECTRANS.aspx.vb" Inherits="MD.MDCMELECTRANS" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����Ź �ŷ���ǥ ��ü����</title>
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
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECTRANS, mobjMDCMGET
Dim mstrCheck
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
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub
Sub imgSaveProc_onclick()
	gFlowWait meWAIT_ON
	ProcessRtn_BatchProc
	gFlowWait meWAIT_OFF
End Sub
Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'----------------------------
'����Ź ��Ȳ TAB BUTTON CLICK
'----------------------------
Sub btnTab1_onclick
	
	frmThis.btnTab1.style.backgroundImage = meURL_TABON
	
	pnltab1.style.visibility = "visible" 
	
	mobjSCGLCtl.DoEventQueue
End Sub

'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------
Sub imgCalDemandday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgCalPrintday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalPrintday,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û�����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' �˻����� ����� MON ������ �����ֱ� ����
'-----------------------------------------------------------------------------------------
Sub txtYEARMON_onblur
	With frmThis
		If .txtYEARMON.value <> "" AND Len(.txtYEARMON.value) = 6 Then DateClean
	End With
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim intAMT,intADJAMT,intBALANCE,intCalCul	
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row  
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
	'����������ü ����	
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECTRANS	= gCreateRemoteObject("cMDET.ccMDETELECTRANS")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "126px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 30, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht,   "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetHeader .sprSht,		   "YEARMON|SEQ|������|MEDNAME| ��ü��|INPUT_MEDFLAG|��ü����|PROGRAM|ADLOCALFLAG|WEEKDAY|����ݾ�|�ΰ���|��|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|�����|GFLAG|SUBSEQ|�귣���|CLIENTSUBCODE|����θ�"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "	  0|  0|    20|      0|     20|            0|       8|      0|          0|      0|      10|    10|10|0         |0     |0    |0  |0         |0           |0         |0      |0             |0        |19    |0    |0     |13      |0            |12"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " CLIENTNAME|REAL_MED_NAME|ATTR02|BRANDNAME|CLIENTSUBNAME ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, " INPUT_MEDNAME", -1, -1, 2
		mobjSCGLSpr.ColHidden .sprSht, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE|TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true 'GFLAG �տ� TRANSRANK �߰�
		
		
		'�հ� ǥ�� �׸��� �⺻ȭ�� ����
		gSetSheetColor mobjSCGLSpr, .sprSht_TRANSSUM
		mobjSCGLSpr.SpreadLayout .sprSht_TRANSSUM, 30, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprSht_TRANSSUM, "YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|MATTERCODE"
		mobjSCGLSpr.SetText .sprSht_TRANSSUM, 3, 1, "           ��       ��"
	    mobjSCGLSpr.SetScrollBar .sprSht_TRANSSUM, 0
	    mobjSCGLSpr.SetBackColor .sprSht_TRANSSUM,"1|3",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeFloat2 .sprSht_TRANSSUM, "AMT|VAT|SUMATMVAT", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprSht_TRANSSUM, "YEARMON|SEQ|MEDNAME|INPUT_MEDFLAG|PROGRAM|ADLOCALFLAG|WEEKDAY|COMMISSION|DEPTCD|PRICE|CNT|ROLLSTDATE|TRU_TAX_FLAG|CLIENTCODE|MEDCODE|REAL_MED_CODE |TRANSRANK|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE", true
		
		mobjSCGLSpr.SetRowHeight .sprSht_TRANSSUM, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht, .sprSht_TRANSSUM
    End With    
    
	pnlTab1.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
	.txtYEARMON.value =  Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtYEARMON.value = vntInParam(i)	
	'			case 1 : mstrFields = vntInParam(i)
	'			case 2 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
	'			case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
	'			case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
	'		end select
	'	next
	end with
	DateClean
	'SelectRtn
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECTRANS = Nothing
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
	.txtPRINTDAY.value  = gNowDate
	.sprSht.MaxRows = 0	

	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	
	strDATE = MID(frmThis.txtYEARMON.value,1,4) & "-" & MID(frmThis.txtYEARMON.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtDEMANDDAY.value = date2
	End With
End Sub
Sub ProcessRtn_BatchProc
	
		Dim intSaveChkRtn	
		Dim intRtn
   		Dim vntData
		Dim strMasterData
		Dim strTRANSYEARMON
		Dim intTRANSNO
		Dim intRANKTRANS
		Dim intCnt,bsdiv
		Dim intColFlag
		Dim strDESCRIPTION
		intSaveChkRtn = gYesNoMsgbox("�����ֺ� �ŷ������� ��ü ���� �Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intSaveChkRtn <> vbYes then exit Sub
		
		
		
		If SelectRtn_Proc = False Then Exit Sub	
		with frmThis
			strDESCRIPTION = ""
			'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
			If .txtDEMANDDAY.value = "" Then
				msgbox "û������ �ʼ� �Է� ���� �Դϴ�."
				Exit Sub
			End If
			
			If .txtPRINTDAY.value = "" Then
				msgbox "�������� �ʼ� �Է� ���� �Դϴ�."
				Exit Sub
			End If

			'�����÷��� ����
			mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
			gXMLSetFlag xmlBind, meINS_TRANS

			'�׷� �ִ밪 ����
			intColFlag = 0
			For intCnt = 1 To .sprSht.MaxRows
			'�ִ밪
				bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
				IF intColFlag < bsdiv THEN
					intColFlag = bsdiv
				END IF
			Next
			
   			'������ Validation
   			If .sprSht.MaxRows = 0 Then
   				msgbox "���׸� �� �����ϴ�."
   				Exit Sub
   			End If
			if DataValidation =false then exit sub
			'On error resume next
			'��Ʈ�� ����� �����͸� �����´�.
			vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
			
			'������ �����͸� ���� �´�.
			strMasterData = gXMLGetBindingData (xmlBind)
			
			'ó�� ������ü ȣ��
			intTRANSNO = 0
			strTRANSYEARMON = .txtYEARMON.value
			
			intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

			if not gDoErrorRtn ("ProcessRtn") then
				'��� �÷��� Ŭ����
				
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				InitPageData
				gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
				gEndPage
   			end if
   		end with

End Sub
'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
	Dim intSaveChkRtn
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strTRANSYEARMON
	Dim intTRANSNO
	Dim intRANKTRANS
	Dim intCnt,bsdiv
	Dim intColFlag
	Dim strDESCRIPTION
	
	intSaveChkRtn = gYesNoMsgbox("����κ� �ŷ������� ��ü ���� �Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
	IF intSaveChkRtn <> vbYes then exit Sub
	with frmThis
		strDESCRIPTION = ""
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .txtDEMANDDAY.value = "" Then
			msgbox "û������ �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If
		
		If .txtPRINTDAY.value = "" Then
			msgbox "�������� �ʼ� �Է� ���� �Դϴ�."
			Exit Sub
		End If

		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

		'�׷� �ִ밪 ����
		intColFlag = 0
		For intCnt = 1 To .sprSht.MaxRows
		'�ִ밪
			bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSRANK",intCnt))
			IF intColFlag < bsdiv THEN
				intColFlag = bsdiv
			END IF
		Next
		
   		'������ Validation
   		If .sprSht.MaxRows = 0 Then
   			msgbox "���׸� �� �����ϴ�."
   			Exit Sub
   		End If
		if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON | SEQ | CLIENTNAME | MEDNAME | REAL_MED_NAME  | INPUT_MEDFLAG| INPUT_MEDNAME | PROGRAM |ADLOCALFLAG |WEEKDAY | AMT | VAT | SUMATMVAT |COMMISSION | DEPTCD | PRICE | CNT | ROLLSTDATE | TRU_TAX_FLAG | CLIENTCODE | MEDCODE | REAL_MED_CODE | TRANSRANK|ATTR02|GFLAG|SUBSEQ|CLIENTSUBCODE|MATTERCODE")
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		intTRANSNO = 0
		strTRANSYEARMON = .txtYEARMON.value
		
		intRtn = mobjMDCMELECTRANS.ProcessRtn(gstrConfigXml,strMasterData,vntData,intTRANSNO,strTRANSYEARMON,intColFlag)

		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			InitPageData
			gOkMsgBox "�ŷ������� �����Ǿ����ϴ�.","Ȯ��"
			gEndPage
   		end if
   	end with
End Sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = false
	Dim vntData
   	Dim i, strCols,intCnt
   	Dim intColSum
   	
	'On error resume next
	with frmThis
'		intColSum = 0
' 		for intCnt = 1 to .sprSht.MaxRows
'			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1  Then 
'				intColSum = intColSum + 1
'			End if
'		next
'		If intColSum = 0 Then exit Function
  	End with
	DataValidation = true
End Function

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
'-----------------------------------------------------------------------------------------
' �ŷ����� ���� ��ȸ[�����Է���ȸ]
'-----------------------------------------------------------------------------------------
Sub SelectRtn ()
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	'����Ʃ�� �ʿ� ����
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
	'On error resume next
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtYEARMON.value
		
		vntPreData = mobjMDCMELECTRANS.SelectRtn_PreCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
			if not gDoErrorRtn ("SelectRtn_PreCnt") then
				lngCnt = vntPreData(0,0)
				If lngCnt < 100 Then
					lngCnt = 1
				Else
					lngCnt = int(lngCnt/100)
					lngCnt = lngCnt+1
				End If
			End if
		
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		vntDataConfirm = mobjMDCMELECTRANS.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		
		If IngCOMMITRowCnt = 0 Then
			gErrorMsgBox strYEARMON & "���� ����ó������ �ʾҽ��ϴ�.",""
			EXIT SUB	
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		'For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			end if
   			'strST = strST + 100
   			'strED = strED + 100
   		'Next
   		'for intDelCnt = .sprSht.MaxRows to 1 step -1				
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intDelCnt) = "" Then
		'		mobjSCGLSpr.DeleteRow .sprSht,intDelCnt
		'	End If		
		'next
   		
   		
   		AMT_SUM
   		PreSearchFiledValue strYEARMON	
   		gWriteText lblStatus, "����Ź " & mlngRowCnt & " �� �� �ڷᰡ �˻�" & mePROC_DONE
   	end with
End Sub
Function SelectRtn_Proc ()
SelectRtn_Proc = False
	Dim vntData, vntDataConfirm
	Dim strYEARMON
	Dim strPRINTDAY
   	Dim i, strCols
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	'����Ʃ�� �ʿ� ����
    Dim strST
   	Dim strED
   	Dim intSQLCnt
   	Dim intDelCnt
   	Dim vntPreData
   	Dim lngCnt
	'On error resume next
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit Function
		End If 
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		strST = 1
		strED = 100
		lngCnt = 0
		strYEARMON	= .txtYEARMON.value
		
		vntPreData = mobjMDCMELECTRANS.SelectRtn_PreCnt(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
			if not gDoErrorRtn ("SelectRtn_PreCnt") then
				lngCnt = vntPreData(0,0)
				If lngCnt < 100 Then
					lngCnt = 1
				Else
					lngCnt = int(lngCnt/100)
					lngCnt = lngCnt+1
				End If
			End if
		
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		vntDataConfirm = mobjMDCMELECTRANS.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		
		If IngCOMMITRowCnt = 0 Then
			gErrorMsgBox strYEARMON & "���� ����ó������ �ʾҽ��ϴ�.",""
			EXIT Function	
		End If
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		'For intSQLCnt = 1 To lngCnt
			vntData = mobjMDCMELECTRANS.SelectRtn_Proc(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strST,strED)

			if not gDoErrorRtn ("SelectRtn") then
					mobjSCGLSpr.SetClip .sprSht, vntData, 1, strST, mlngColCnt, mlngRowCnt, True
					mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
   			end if
   		'	strST = strST + 100
   		'	strED = strED + 100
   		'Next
   		'for intDelCnt = .sprSht.MaxRows to 1 step -1				
		'	If mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",intDelCnt) = "" Then
		'		mobjSCGLSpr.DeleteRow .sprSht,intDelCnt
		'	End If		
		'next
   		
   		
   		AMT_SUM
   		'PreSearchFiledValue strYEARMON	
   		'gWriteText lblStatus, "����Ź " & mlngRowCnt & " �� �� �ڷᰡ �˻�" & mePROC_DONE
   	end with
SelectRtn_Proc = True
End Function

Sub PreSearchFiledValue (strYEARMON)
	frmThis.txtYEARMON.value = strYEARMON
End Sub

'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntVAT, IntVATSUM, IntSUMATMVAT, IntSUMATMVATSUM
	With frmThis
		IntAMTSUM = 0
		IntVATSUM = 0
		IntSUMATMVATSUM = 0
		
		'����Ź �׸��� �հ�׸��� ���ֱ�
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = 0
			IntVAT = 0
			IntSUMATMVAT = 0
			
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"VAT", lngCnt)
			IntSUMATMVAT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMATMVAT", lngCnt)
			
			IntAMTSUM = IntAMTSUM + IntAMT
			IntVATSUM = IntVATSUM + IntVAT
			IntSUMATMVATSUM = IntSUMATMVATSUM + IntSUMATMVAT
		Next
		if .sprSht.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"AMT",1, IntAMTSUM
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"VAT",1, IntVATSUM
			mobjSCGLSpr.SetTextBinding .sprSht_TRANSSUM,"SUMATMVAT",1, IntSUMATMVATSUM
		end if
	End With
End Sub


'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprSht_TRANSSUM	
	End with
end sub

'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprSht_TRANSSUM, NewTop, NewLeft
End Sub

'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON, dblSEQ

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
		
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			'Insert Transaction�� �ƴ� ��� ���� ������ü ȣ��
			IF mobjSCGLSpr.GetFlagMode(.sprSht,vntData(i)) <> meINS_TRANS then
				dblSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SEQ",vntData(i))
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",vntData(i))
			
				intRtn = mobjMDCMELECTRANS.DeleteRtn(gstrConfigXml,dblSEQ, strYEARMON)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
   			End IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intSelCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
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
			<TABLE id="tblForm" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="793" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD>
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD style="WIDTH: 400px" align="left" width="427" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">
													&nbsp;����Ź �ŷ����� ����</td>
											</tr>
										</table>
									</TD>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 282px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
							<TABLE id="tblBody" style="WIDTH: 1040px" cellSpacing="0" cellPadding="0" width="1040"
								border="0"> <!--TopSplit Start->
									<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="top" align="center">
										<TABLE class="DATA" id="tblDATA1" style="WIDTH: 1040px" cellSpacing="1" cellPadding="0"
											width="1040" border="0">
											<TR>
												<TD class="SEARCHLABEL" title="�����մϴ�." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')">���</TD>
												<TD class="SEARCHDATA" width="180"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 89px; HEIGHT: 22px" type="text"
														maxLength="6" size="9" name="txtYEARMON" accessKey="MON"></TD>
												<TD class="SEARCHLABEL" title="�����մϴ�." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY,'')">û������</TD>
												<TD class="SEARCHDATA" width="180"><INPUT dataFld="DEMANDDAY" class="INPUT" id="txtDEMANDDAY" title="û������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="9" name="txtDEMANDDAY"><IMG id="imgCalDemandday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalDemandday"></TD>
												<TD class="SEARCHLABEL" title="�����մϴ�." style="WIDTH: 80px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')">��������</TD>
												<TD class="SEARCHDATA"><INPUT dataFld="PRINTDAY" class="INPUT" id="txtPRINTDAY" title="��������" style="WIDTH: 88px; HEIGHT: 22px"
														accessKey="date,M" dataSrc="#xmlBind" type="text" maxLength="10" size="12" name="txtPRINTDAY"><IMG id="imgCalPrintday" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalPrintday">
												</TD>
												<TD class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
							</TABLE>
						</TD>
					<!--BodySplit Start-->
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px;HEIGHT: 25px"></TD>
					</TR>
					<TR>
						<TD class="KEYFRAME" vAlign="middle" align="center">
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
												<td class="TITLE">
													&nbsp;�ŷ����� ��ü����</td>
											</tr>
										</table>
									</TD>
									<TD vAlign="middle" align="right" height="20">
										<!--Common Button Start-->
										<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
											<TR>
												<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/ImgTRANSALLSUBOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTRANSALLSUB.gIF'"
														height="20" alt="�ش���� ����κ� �ŷ����� ��ü�� �����մϴ�. " src="../../../images/ImgTRANSALLSUB.gIF"
														border="0" name="imgSave"></TD>
												<TD><IMG id="imgSaveProc" onmouseover="JavaScript:this.src='../../../images/ImgTRANSALLOn.gif'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTRANSALL.gIF'"
														height="20" alt="�ش���� �����ֺ� �ŷ����� ��ü�� �����մϴ�." src="../../../images/ImgTRANSALL.gIF"
														border="0" name="imgSaveProc"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gIF'"
														height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
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
					<!--BodySplit End-->
					<!--List Start-->
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 714px" vAlign="top" align="center">
							<DIV id="pnlTab1" style="POSITION: relative; vWIDTH: 100%" ms_positioning="GridLayout">
								<OBJECT id="sprSht" style="WIDTH: 1040px; HEIGHT: 690px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="18256">
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
								<OBJECT id="sprSht_TRANSSUM" style="WIDTH: 1040px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="635">
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
					<!--List End-->
					<!--Bottom Split Start-->
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"></TD>
					</TR>
					<!--Bottom Split End--></TBODY></TABLE>
			<!--Input Define Table End--> </TD></TR> 
			<!--Top TR End--> </TBODY></TABLE> 
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
