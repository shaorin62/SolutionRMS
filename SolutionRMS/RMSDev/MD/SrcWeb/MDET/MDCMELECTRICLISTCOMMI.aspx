<%@ Page CodeBehind="MDCMELECTRICLISTCOMMI.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRICLISTCOMMI" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ������ ����ó��</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/ǥ�ػ���/�������彬Ʈ
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : SpreadSheet�� �̿��� ��ȸ/�Է�/����/����/�μ� ó�� ǥ�� ����
'�Ķ�  ���� : 
'Ư��  ���� : ǥ�ػ����� ���� ���� ����
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/04/15 By KimKS
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
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDCMELECTRICLISTCOMMI 
Dim mobjMDCMGET
Dim mstrCheck
mstrCheck = True
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
	if frmThis.txtYEARMON.value = "" then
		gErrorMsgBox "����� �Է��Ͻÿ�",""
		exit Sub
	end if
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgSetting_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub ImgConfirmCancel_onclick ()
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht1
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub


'----------------------------
'������ ���� TAB BUTTON CLICK
'----------------------------
Sub btnTab2_onclick
	pnltab2.style.visibility = "visible"
	mobjSCGLCtl.DoEventQueue
End Sub


'****************************************************************************************
' ��Ʈ ����Ŭ�� �̺�Ʈ
'****************************************************************************************
sub sprSht1_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		end if
	end with
end sub


'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	frmThis.imgSetting.style.visibility = "hidden"
	frmThis.ImgConfirmCancel.style.visibility = "hidden"
	'����������ü ����	
	set mobjMDCMELECTRICLISTCOMMI = gCreateRemoteObject("cMDET.ccMDETELECTRICLISTCOMMI")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab2.style.position = "absolute"
	pnlTab2.style.top = "102px"
	pnlTab2.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
   
    '*********************************
    '�������Ʈ
    '*********************************
    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
	With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 13, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht1,   "YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET"
		mobjSCGLSpr.SetHeader .sprSht1,		   "���|��ü��|������|INPUT_MEDFLAG|��ü����|����ݾ�|��������|������|�������ڵ�|��ü���ڵ�|�μ��ڵ�|TRANSRANK"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "  0|    30|    38|            0|      13|      13|      13|    15|0         |0         |0       |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, " YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU", -1, -1, 0
		'mobjSCGLSpr.SetCellTypeStatic2 .sprSht1, "AMT|SUSURATE|SUSU", -1, -1, 1
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht1, "AMT|SUSU|SUSURATE", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"AMT|SUSURATE"		
		mobjSCGLSpr.ColHidden .sprSht1, "CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET", true
		mobjSCGLSpr.CellGroupingEach .sprSht1, "YEARMON|REAL_MED_NAME|CLIENTNAME"
		
		gSetSheetColor mobjSCGLSpr, .sprShtSum
		mobjSCGLSpr.SpreadLayout .sprShtSum, 13, 1, 0,0,1,1,1,false,true,true,1
		mobjSCGLSpr.SpreadDataField .sprShtSum, "YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET"
		mobjSCGLSpr.SetText .sprShtSum, 1, 1, "        ��      ��"
	    mobjSCGLSpr.SetScrollBar .sprShtSum, 0
	    mobjSCGLSpr.SetBackColor .sprShtSum,"1|1",rgb(205,219,215),false
	    mobjSCGLSpr.SetCellTypeStatic2 .sprShtSum,  "AMT|SUSURATE|SUSU ", -1, -1, 0
	    mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "AMT|SUSU|SUSURATE", -1, -1, 0
		mobjSCGLSpr.ColHidden .sprShtSum, "CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET", true
		
		mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "13"	  
	    mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum
    End With
    
    pnlTab2.style.visibility = "visible"

	'ȭ�� �ʱⰪ ����
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : mstrFields = vntInParam(i)
				case 2 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	
End Sub
sub sprSht1_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht1, .sprShtSum	
	End with
end sub
Sub sprSht1_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub
Sub EndPage()
	set mobjMDCMELECTRICLISTCOMMI = Nothing
	set mobjMDCMGET = Nothing
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
		.txtYEARMON.value =  Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'Sheet�ʱ�ȭ
		.sprSht1.MaxRows = 0
		
		.txtYEARMON.focus
		
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData, vntData1,vntDataPre
	Dim strYEARMON
	Dim strCLIENTCODE
	Dim strCLIENTNAME
	Dim strREAL_MED_CODE
	Dim strREAL_MED_NAME
	Dim strGFLAG
	Dim strINPUT_MEDFLAG
	Dim vntDataConfirm
	Dim strCONFIRM
   	Dim i, strCols
   	Dim IngsusuColCnt, IngsusuRowCnt
   	Dim IngCOMMITColCnt, IngCOMMITRowCnt
   	Dim strSEARCHGBN
   	Dim strSETENDFLAG
   	Dim lngCnt
	'on error resume next
	with frmThis
		If .txtYEARMON.value = "" Then 
			gErrorMsgBox "����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		.sprSht1.MaxRows = 0
		IngCOMMITColCnt=clng(0)
		IngCOMMITRowCnt=clng(0)
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		IngsusuColCnt=clng(0)
		IngsusuRowCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		
		strCLIENTNAME	= .txtCLIENTNAME.value
		
		
		strGFLAG = .cmbGROUP.value
		'����,�̿�����ȸ�� �ŷ�ó�� �������� ������ش�.
		If strGFLAG <> "A" Then
			.txtCLIENTNAME.value = ""
		End IF

		vntDataConfirm = mobjMDCMELECTRICLISTCOMMI.SelectRtn_CONFIRM(gstrConfigXml,IngCOMMITRowCnt,IngCOMMITColCnt, strYEARMON)
		'Ȯ����
		If IngCOMMITRowCnt > 0 Then
			strSETENDFLAG = "T"
			.ImgConfirmCancel.style.visibility = "visible"
			.imgSetting.style.visibility = "hidden"
			.btnTab2.value = "��������ȸ"
			mobjSCGLSpr.SetCellsLock2 .sprSht1,true,"SUSU"
			'2���� ��ȸ
			vntData1 = mobjMDCMELECTRICLISTCOMMI.SelectRtn_ENDSUSU(gstrConfigXml,IngsusuRowCnt,IngsusuColCnt,strYEARMON)
			
		Else
		'��Ȯ����
			strSETENDFLAG = "F"
			.imgSetting.style.visibility = "visible"
			.ImgConfirmCancel.style.visibility = "hidden"
			mobjSCGLSpr.SetCellsLock2 .sprSht1,false,"SUSU"
			.btnTab2.value = "���������"
			'2���� ��ȸ
 			vntData1 = mobjMDCMELECTRICLISTCOMMI.SelectRtn_SUSU(gstrConfigXml,IngsusuRowCnt,IngsusuColCnt,strYEARMON)
		End If
		
		if IngsusuRowCnt > 0 then
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData1,1,1,IngsusuColCnt,IngsusuRowCnt,TRUE)
			mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			AMT_SUM
			If strSETENDFLAG = "F" Then
			gWriteText lblStatus, "�̻������� " & mlngRowCnt & " ��,��������� " & IngsusuRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
			Else
			gWriteText lblStatus, "�̻������� " & mlngRowCnt & " ��,��������ȸ " & IngsusuRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
			End IF
		else
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   		end if
   		
   		REAL_MED_CODE_AMT_SUM
	end with   	
End Sub

Sub PreSearchFiledValue (strYEARMON, strCLIENTCODE, strCLIENTNAME, strREAL_MED_CODE, strREAL_MED_NAME)
	frmThis.txtYEARMON.value = strYEARMON
	frmThis.txtCLIENTCODE.value = strCLIENTCODE
	frmThis.txtCLIENTNAME.value = strCLIENTNAME
	frmThis.txtREAL_MED_CODE.value = strREAL_MED_CODE
	frmThis.txtREAL_MED_NAME.value = strREAL_MED_NAME
End Sub

'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntVAT, IntVATSUM, IntSUMATMVAT, IntSUMATMVATSUM
	Dim IntAMT1, IntSUSU, IntSUSUVAT, IntSUMSUSUVAT, IntAMT1SUM, IntSUSUSUM, IntSUSUVATSUM, IntSUMSUSUVATSUM
	With frmThis
		IntAMTSUM = 0
		
		IntAMT1SUM = 0
		IntSUSUSUM = 0
		IntSUSUVATSUM = 0
		IntSUMSUSUVATSUM = 0
		
		'������ �׸��� �հ�׸��� ���ֱ�
		For lngCnt = 1 To .sprSht1.MaxRows
			IntAMT1 = 0
			IntSUSU = 0
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht1,"CLIENTNAME", lngCnt) <> "" THEN
				IntAMT1 = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				IntSUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntAMT1SUM = IntAMT1SUM + IntAMT1
				IntSUSUSUM = IntSUSUSUM + IntSUSU
			END IF
		Next
		
		if .sprSht1.MaxRows >0 Then
			mobjSCGLSpr.SetTextBinding .sprShtSum,"AMT",1, IntAMT1SUM
			mobjSCGLSpr.SetTextBinding .sprShtSum,"SUSU",1, IntSUSUSUM
			
		end if
	End With
End Sub


'��Ʈ�� �ݾ��� �ջ��� ���� �հ��ƮM�� �ѷ��ش�.
Sub REAL_MED_CODE_AMT_SUM
	Dim lngCnt
	Dim lntB1AMT, lntB1SUSU, IntB1AMTSUM, IntB1SUSUSUM
	Dim lntB2AMT, lntB2SUSU, IntB2AMTSUM, IntB2SUSUSUM
	Dim lntB3AMT, lntB3SUSU, IntB3AMTSUM, IntB3SUSUSUM
	Dim lntB4AMT, lntB4SUSU, IntB4AMTSUM, IntB4SUSUSUM
	Dim lntB6AMT, lntB6SUSU, IntB6AMTSUM, IntB6SUSUSUM
	Dim lntB7AMT, lntB7SUSU, IntB7AMTSUM, IntB7SUSUSUM

	With frmThis
		IntB1AMTSUM = 0
		IntB1SUSUSUM = 0
		
		IntB2AMTSUM = 0
		IntB2SUSUSUM = 0
		
		IntB3AMTSUM = 0
		IntB3SUSUSUM = 0
		
		IntB4AMTSUM = 0
		IntB4SUSUSUM = 0
		
		IntB6AMTSUM = 0
		IntB6SUSUSUM = 0
		
		IntB7AMTSUM = 0
		IntB7SUSUSUM = 0
		
		'������ �׸��� �հ�׸��� ���ֱ�
		For lngCnt = 1 To .sprSht1.MaxRows
			lntB1AMT = 0
			lntB1SUSU = 0
			lntB2AMT = 0
			lntB2SUSU = 0
			lntB3AMT = 0
			lntB3SUSU = 0
			lntB4AMT = 0
			lntB4SUSU = 0
			lntB6AMT = 0
			lntB6SUSU = 0
			lntB7AMT = 0
			lntB7SUSU = 0
			
			IF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00140" THEN
				lntB1AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB1SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB1AMTSUM = IntB1AMTSUM  + lntB1AMT
				IntB1SUSUSUM = IntB1SUSUSUM + lntB1SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00144" THEN
				lntB2AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB2SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB2AMTSUM = IntB2AMTSUM  + lntB2AMT
				IntB2SUSUSUM = IntB2SUSUSUM + lntB2SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00142" THEN
				lntB3AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB3SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB3AMTSUM = IntB3AMTSUM  + lntB3AMT
				IntB3SUSUSUM = IntB3SUSUSUM + lntB3SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00143" THEN
				lntB4AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB4SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB4AMTSUM = IntB4AMTSUM  + lntB4AMT
				IntB4SUSUSUM = IntB4SUSUSUM + lntB4SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00141" THEN
				lntB6AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB6SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB6AMTSUM = IntB6AMTSUM  + lntB6AMT
				IntB6SUSUSUM = IntB6SUSUSUM + lntB6SUSU
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht1,"REAL_MED_CODE", lngCnt) = "B00145" THEN
				lntB7AMT = mobjSCGLSpr.GetTextBinding(.sprSht1,"AMT", lngCnt)
				lntB7SUSU = mobjSCGLSpr.GetTextBinding(.sprSht1,"SUSU", lngCnt)
				
				IntB7AMTSUM = IntB7AMTSUM  + lntB7AMT
				IntB7SUSUSUM = IntB7SUSUSUM + lntB7SUSU
			
			END IF
		Next
		
		if .sprSht1.MaxRows >0 Then
			.txtB1AMT.value = IntB1AMTSUM
			.txtB1SUSU.value = IntB1SUSUSUM
			
			.txtB2AMT.value = IntB2AMTSUM
			.txtB2SUSU.value = IntB2SUSUSUM
			
			.txtB3AMT.value = IntB3AMTSUM
			.txtB3SUSU.value = IntB3SUSUSUM
			
			.txtB4AMT.value = IntB4AMTSUM
			.txtB4SUSU.value = IntB4SUSUSUM
			
			.txtB6AMT.value = IntB6AMTSUM
			.txtB6SUSU.value = IntB6SUSUSUM
			
			.txtB7AMT.value = IntB7AMTSUM
			.txtB7SUSU.value = IntB7SUSUSUM
			
		end if
	End With
End Sub


'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
   	Dim vntData
	Dim strMasterData
	Dim strYEARMON
	Dim intCnt
	with frmThis
		'�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht1,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht1.MaxRows = 0 Then
   			gErrorMsgBox "���׸��� �����ϴ�.","Ȯ������"
   			Exit Sub
   		End If
		
		'������ �����͸� ���� �´�.
		strMasterData = gXMLGetBindingData (xmlBind)
		
		'ó�� ������ü ȣ��
		strYEARMON = .txtYEARMON.value
		for intCnt = 1 to .sprSht1.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht1,"SAVESET",intCnt, "T"	
			Call sprSht1_Change (13,intCnt)
		next
		vntData = mobjSCGLSpr.GetDataRows(.sprSht1,"YEARMON|REAL_MED_NAME|CLIENTNAME|INPUT_MEDFLAG|INPUT_MEDNAME|AMT|SUSURATE|SUSU|CLIENTCODE|REAL_MED_CODE|DEPTCD|TRANSRANK|SAVESET")
		intRtn = mobjMDCMELECTRICLISTCOMMI.ProcessRtn(gstrConfigXml, strMasterData,vntData,strYEARMON)

		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht1,meCLS_FLAG
			'InitPageData
			gOkMsgBox "����ó�� �Ǿ����ϴ�.","Ȯ��"
			SelectRtn
   		end if
   	end with
End Sub
Sub sprSht1_change(ByVal Col,ByVal Row)
AMT_SUM
mobjSCGLSpr.CellChanged frmThis.sprSht1, Col,Row
End Sub	
'****************************************************************************************
' ��ü ������ �� ��Ʈ�� ����
'****************************************************************************************
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i
	dim strYEARMON

	with frmThis
		intSelCnt = 0
		
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		If .sprSht1.MaxRows = 0 Then
   			gErrorMsgBox "���׸��� �����ϴ�.","Ȯ����ҿ���"
   			Exit Sub
   		End If
   		
   		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
   		strYEARMON = .txtYEARMON.value
   		intRtn = mobjMDCMELECTRICLISTCOMMI.SelectRtn_CANCEL(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON)
   		if mlngRowCnt > 0 then
   			gErrorMsgBox "�ŷ������� ������ �����ʹ� Ȯ����Ұ� �ȵ˴ϴ�.","Ȯ����ҿ���"
   			Exit Sub
   		end if
		
		intRtn = gYesNoMsgbox("Ȯ����� �Ͻðڽ��ϱ�?","Ȯ����� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		
		'���õ� �ڷḦ ������ ���� ����
		strYEARMON = .txtYEARMON.value
	
		intRtn = mobjMDCMELECTRICLISTCOMMI.DeleteRtn(gstrConfigXml,strYEARMON)
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gOkMsgBox  strYEARMON & " �� �ڷᰡ Ȯ����� �Ǿ����ϴ�.","Ȯ��"
			SelectRtn
   		End IF
	End with
	err.clear	
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<P dir="ltr" style="MARGIN-RIGHT: 0px">
				<TABLE id="tblForm" style="WIDTH: 1040px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
					width="1040" border="0">
					<TR>
						<TD>
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
								border="0">
								<TR>
									<td align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gif" width="14"></td>
												<td align="left" height="4"></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;������&nbsp;������ ����ó��</td>
											</tr>
										</table>
									</td>
									<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
										<!--Wait Button Start-->
										<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
											cellSpacing="1" cellPadding="1" width="75%" border="0">
											<TR>
												<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
														border="0" name="imgWaiting">
												</TD>
											</TR>
										</TABLE>
										<!--Wait Button End-->
										<!--Common Button Start-->
										<TABLE id="tblButton" style="WIDTH: 108px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
											width="108" border="0">
											<TR>
												<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
														height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
														name="imgQuery"></TD>
												<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
														height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
														name="imgExcel"></TD>
												<TD><!--<IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
														height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
														name="imgClose">--></TD>
											</TR>
										</TABLE>
										<!--Common Button End--></TD>
								</TR>
							</TABLE>
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 60px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">�� 
													��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 181px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" style="WIDTH: 64px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="5" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 74px; CURSOR: hand">��ȸ����</TD>
												<TD class="SEARCHDATA" style="WIDTH: 174px"><SELECT class="INPUT" id="cmbGROUP" title="�׷챸��" style="WIDTH: 99px" name="cmbGROUP">
														<OPTION value="A" selected>��ü</OPTION>
														<OPTION value="G">������</OPTION>
														<OPTION value="N">�̿���</OPTION>
													</SELECT></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 77px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">�����ָ�</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 384px; HEIGHT: 22px"
														type="text" maxLength="100" size="58" name="txtCLIENTNAME"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px">
							<TABLE id="tblTab" style="WIDTH: 1040px; HEIGHT: 5px" cellSpacing="0" cellPadding="0" width="787"
								border="0">
								<TR>
									<TD style="HEIGHT: 26px" align="left" width="100%"><INPUT class="BTNTAB" id="btnTab2" style="BACKGROUND-IMAGE: url(../../../images/imgTabOn.gIF)"
											type="button" size="20" value="���������" name="btnTab2">
									</TD>
									<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgSettingOn.gIF'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgSetting.gIF'"
											height="20" alt="Ȯ���մϴ�." src="../../../images/imgSetting.gIF" width="54" align="right"
											border="0" name="imgSetting"></TD>
									<TD><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgConfirmCancelOn.gif'"
											style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgConfirmCancel.gif'"
											height="20" alt="Ȯ������մϴ�." src="../../../images/ImgConfirmCancel.gIF" border="0"
											name="ImgConfirmCancel"></TD>
								</TR>
								<TR class="TABBAR">
									<TD colSpan="3"></TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD class="LISTFRAME" style="WIDTH: 1040px; HEIGHT: 600px" vAlign="top" align="center">
							<DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative" ms_positioning="GridLayout">
								<OBJECT id="sprSht1" style="WIDTH: 1040px; HEIGHT: 576px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
									VIEWASTEXT>
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="27517">
									<PARAM NAME="_ExtentY" VALUE="15240">
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
								<OBJECT id="sprShtSum" style="WIDTH: 1040px; HEIGHT: 24px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
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
					<TR>
						<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
					</TR>
					<TR>
						<TD>
							<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
										<TABLE class="DATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="LABEL" width="90">�������ݾ�</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB1AMT" title="�ѱ���۱�����纻�����ݾ�" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB1AMT"></TD>
												<TD class="LABEL" width="90">���������</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB1SUSU" title="�ѱ���۱�����纻��������Ѿ�" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB1SUSU"></TD>
												<TD class="LABEL" width="90">�λ����ݾ�</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB2AMT" title="�ѱ���۱������λ��������ݾ�" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB2AMT"></TD>
												<TD class="LABEL" width="90">�λ������</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB2SUSU" title="�ѱ���۱������λ�����������Ѿ�" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB2SUSU"></TD>
											</TR>
											<TR>
												<TD class="LABEL" width="90">�뱸����ݾ�</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB3AMT" title="�ѱ���۱������뱸�������ݾ�" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB3AMT"></TD>
												<TD class="LABEL" width="90">�뱸������</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB3SUSU" title="�ѱ���۱������뱸����������Ѿ�" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB3SUSU"></TD>
												<TD class="LABEL" width="90">��������ݾ�</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB4AMT" title="�ѱ���۱����������������ݾ�" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB4AMT"></TD>
												<TD class="LABEL" width="90">����������</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB4SUSU" title="�ѱ���۱�������������������Ѿ�" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB4SUSU"></TD>
											</TR>
											<TR>
												<TD class="LABEL" width="90">���ִ���ݾ�</TD>
												<TD class="DATA" width="107"><INPUT class="INPUT" id="txtB6AMT" title="�ѱ���۱�����籤���������ݾ�" style="WIDTH: 103px; HEIGHT: 22px"
														type="text" size="11" name="txtB6AMT"></TD>
												<TD class="LABEL" width="90">���ּ�����</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB6SUSU" title="�ѱ���۱�����籤������������Ѿ�" style="WIDTH: 110px; HEIGHT: 22px"
														type="text" size="13" name="txtB6SUSU"></TD>
												<TD class="LABEL" width="90">���ϴ���ݾ�</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB7AMT" title="�ѱ���۱�����������������ݾ�" style="WIDTH: 108px; HEIGHT: 22px"
														type="text" size="12" name="txtB7AMT"></TD>
												<TD class="LABEL" width="90">���ϼ�����</TD>
												<TD class="DATA" width="108"><INPUT class="INPUT" id="txtB7SUSU" title="�ѱ���۱��������������������Ѿ�" style="WIDTH: 106px; HEIGHT: 22px"
														type="text" size="12" name="txtB7SUSU"></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
							</TABLE>
						</TD>
					</TR>
					<TR>
						<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
					</TR>
				</TABLE>
			</P>
		</form>
	</body>
</HTML>
