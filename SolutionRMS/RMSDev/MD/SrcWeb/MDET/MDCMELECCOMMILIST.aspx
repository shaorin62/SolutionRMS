<%@ Page CodeBehind="MDCMELECCOMMILIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECCOMMILIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ������ �ŷ��� ��ȸ</title> 
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
Dim mobjMDCOGET 
Dim mobjMDCMELECCOMMILIST
Dim mstrCheck
Dim mALLCHECK
mALLCHECK = TRUE
mstrCheck = True

CONST meTAB = 9

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

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
		with frmThis
			mobjSCGLSpr.ExcelExportOption = true 
			mobjSCGLSpr.ExportExcelFile .sprSht
		end with
	gFlowWait meWAIT_OFF
End Sub
Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub
'��� �μ��ư Ŭ���� �̺�Ʈ
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j,k
	Dim datacnt
	Dim strTRANSYEARMON
	Dim strTRANSNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim strUSERID
	
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
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCMELECCOMMILIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECCOMMI_NEW.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_ELECCOMMI_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					
					datacnt = strcntsum
					for k=1 to 2
						strUSERID = ""
						vntDataTemp = mobjMDCMELECCOMMILIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
					next
				End IF
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
End Sub

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMELECCOMMILIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgREAL_MED_CODE_onclick
	Call REAL_MED_CODE_POP()
End Sub

'���� ������List ��������
Sub REAL_MED_CODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtREAL_MED_CODE.value), trim(.txtREAL_MED_NAME.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMREAL_MEDPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtREAL_MED_CODE.value = vntRet(0,0) and .txtREAL_MED_NAME.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtREAL_MED_CODE.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtREAL_MED_NAME.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SELECTRTN
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtREAL_MED_NAME_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtREAL_MED_CODE.value),trim(.txtREAL_MED_NAME.value), "B")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtREAL_MED_CODE.value = trim(vntData(0,1))
					.txtREAL_MED_NAME.value = trim(vntData(1,1))
					SELECTRTN
				Else
					Call REAL_MED_CODE_POP()
				End If
   			End If
   		End With
   		
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' ��Ʈ ����Ŭ�� �̺�Ʈ
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

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row

End Sub
sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	DIM strTRANSYEARMON
	DIM strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		else
			strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",Row)
			strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",Row)
			
			vntInParams = array(strTRANSYEARMON, strTRANSNO) '<< �޾ƿ��°��
			vntRet = gShowModalWindow("MDCMELECCOMMIGUNLIST.aspx",vntInParams , 813,545)
			if isArray(vntRet) then
     		end if
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
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		'sprShtToFieldBinding frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow
	End If
	
	With frmThis
		If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or  _
			.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") OR .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") Then
			strSUM = 0
			intSelCnt = 0
			intSelCnt1 = 0
			strCOLUMN = ""
			
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") Then
				strCOLUMN = "AMT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") Then
				strCOLUMN = "VAT"
			ELSEIF .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") Then
				strCOLUMN = "SUMAMTVAT"
			End If
			
			vntData_col = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt1)

			FOR i = 0 TO intSelCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT")) OR (vntData_col(i) = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT")) Then
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
			If .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"AMT") or  _
				.sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"VAT") or .sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht,"SUMAMTVAT") Then
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



'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'����������ü ����	
	set mobjMDCOGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECCOMMILIST = gCreateRemoteObject("cMDET.ccMDETELECCOMMILIST")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 1, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TRANSYEARMON | TRANSNO | CLIENTCODE |CLIENTNAME| MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME| AMT | VAT | SUMAMTVAT | DEMANDDAY | PRINTDAY | TAXYN "
		mobjSCGLSpr.SetHeader .sprSht,		"����|TRANSYEARMON|TRANSNO|CLIENTCODE|������|MEDCODE|��ü��|REAL_MED_CODE|û����|�������Ѿ�|�ΰ���|��|û����|������"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "5|		  	 0|		 0|		    0|	   0|      0|     0|	        0|    31|        16|    16|18|    18|    18"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT | VAT |SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MEDNAME|REAL_MED_NAME", -1, -1, 20
		mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO |CLIENTNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|MEDNAME|TAXYN", true
		.sprSht.style.visibility = "visible"
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
    End With
    
    pnlTab1.style.visibility = "visible"


	'ȭ�� �ʱⰪ ����
	
	InitPageData
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	'�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'WITH frmThis
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtYEARMON.value = vntInParam(i)	
	'			case 1 : .txtREAL_MED_CODE.value = vntInParam(i)
	'			case 2 : .txtREAL_MED_NAME.value = vntInParam(i)		'���� ������� �͸�
	'			case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
	'			case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
	'		end select
	'	next
	'end with
	
	SelectRtn		
End Sub

Sub EndPage()
	set mobjMDCOGET = Nothing
	set mobjMDCMELECCOMMILIST = Nothing
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
		
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtYEARMON.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim i, strCols
	on error resume next
	with frmThis
		'�ʱ�ȭ
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		strYEARMON			= .txtYEARMON.value
		strREAL_MED_CODE	= .txtREAL_MED_CODE.value
		
		vntData = mobjMDCMELECCOMMILIST.Get_ELECCOMMI_ALLLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,"", strREAL_MED_CODE)
		
		IF not gDoErrorRtn ("Get_ELECTRANS_ALLLIST") then
			'��ȸ�� �����͸� ���ε�
			IF mlngRowCnt > 0 THEN
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			ELSE
				INITPAGEDATA
			END IF
			AMT_SUM
			gWriteText lblstatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
		End IF

   	end with
End Sub

'****************************************************************************************
'��Ʈ�� �ݾ��� �ջ��� ���� �հ��Ʈ�� �ѷ��ش�.
'****************************************************************************************
Sub AMT_SUM
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
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
	End With
End Sub


'****************************************************************************************
' �ŷ����� ����
'****************************************************************************************
Sub ImgDeleteAll_onclick()
	gFlowWait meWAIT_ON
	DeleteAll
	gFlowWait meWAIT_OFF
End Sub

Sub DeleteAll
	Dim intCnt
	Dim strVOCHCnt
	Dim strVOCHSumCnt
	Dim intRtn
	Dim vntData
	Dim strSUMRTN
	Dim intCnt2
	Dim intDelRtn
	Dim intDelete
	with frmThis
		intDelete = gYesNoMsgbox("�ش�� ������ŷ����� ��ü�� �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intDelete <> vbYes then exit Sub
		
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "���� �����Ͻ� �����͸� ��ȸ�Ͻʽÿ�.","��ü�����ȳ�!"
			Exit Sub
		End If
		
		'ó�� ������ü ȣ��
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strSUMRTN = 0
		intRtn = 0
		
		For intCnt2 = 1 to .sprSht.MaxRows		
			strSUMRTN = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYN", intCnt2)
			IF strSUMRTN = "Y" Then
				Exit For
			end If
		Next
		
		If strSUMRTN = "Y" Then
			gErrorMsgBox "���ݰ�꼭�� �ۼ��Ǿ��ֽ��ϴ�." & vbcrlf & "���ݰ�꼭 �� �����Ͻð� ��ü������ �Ͻʽÿ�.","��ü�����ȳ�!"
			Exit Sub
		Else
		    intDelRtn = mobjMDCMELECCOMMILIST.Delete_TRANS(gstrConfigXml,.txtYEARMON.value) 
		    if not gDoErrorRtn ("Delete_TRANS") then
				'��� �÷��� Ŭ����
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				if intDelRtn > 1 Then
				gErrorMsgBox intDelRtn & " �� �� �����Ǿ����ϴ�.","�����ȳ�"
				End If
				SelectRtn
   			end if
			
		End If
		
	End With
End Sub
-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left">
													<TABLE cellSpacing="0" cellPadding="0" width="125" background="../../../images/back_p.gIF"
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
											<td class="TITLE">������ �ŷ��� ����</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<!--Wait Button Start-->
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 336px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" height="95%" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 100%"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 100%" vAlign="middle" align="center">
									<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">�� 
												��</TD>
											<TD class="SEARCHDATA" style="WIDTH: 380px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" accessKey="MON" type="text" maxLength="6"
													size="10" name="txtYEARMON"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 90px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtREAL_MED_CODE, txtREAL_MED_NAME)">û����&nbsp;
											</TD>
											<TD class="SEARCHDATA" ><INPUT class="INPUT_L" id="txtREAL_MED_NAME" title="�ڵ��" style="WIDTH: 184px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="25" name="txtREAL_MED_NAME"> <IMG id="ImgREAL_MED_CODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" src="../../../images/imgPopup.gIF" align="absMiddle"
													border="0" name="ImgREAL_MED_CODE"> <INPUT class="INPUT_L" id="txtREAL_MED_CODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtREAL_MED_CODE"></TD>
											<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" border="0" align="right" name="imgQuery"></td>
										</TR>
									</TABLE>
								</TD>
							</tr>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 25px"></TD>
							</TR>	
							<tr>
								<td>
									<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" width="400" height="20">
												<table height="100%" cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td class="TITLE" vAlign="absmiddle">�������հ� : <INPUT class="NOINPUTB_R" id="txtSUMAMT" title="�հ�ݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																accessKey="NUM" readOnly type="text" maxLength="100" size="13" name="txtSUMAMT">
															<INPUT class="NOINPUTB_R" id="txtSELECTAMT" title="���ñݾ�" style="WIDTH: 120px; HEIGHT: 20px"
																readOnly type="text" maxLength="100" size="16" name="txtSELECTAMT">
														</td>
													</tr>
												</table>
											</TD>
											<TD vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<!--TD><IMG id="ImgDeleteAll" onmouseover="JavaScript:this.src='../../../images/ImgDeleteAllOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgDeleteAll.gif'"
																height="20" alt="��ü�ڷḦ �����մϴ�." src="../../../images/ImgDeleteAll.gIF" border="0"
																name="ImgDeleteAll"></TD-->
														<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
																name="imgPrint"></TD>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											VIEWASTEXT>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="16140">
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
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 100%"></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></form>
	</body>
</HTML>
