<%@ Page CodeBehind="MDCMELECTRANSCONFLIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMELECTRANSCONFLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����Ź �ŷ��� ��ȸ</title> 
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
Dim mobjMDCMGET 
Dim mobjMDCMELECTRANSLIST
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
Sub imgSetting_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_ConfirmOK
	gFlowWait meWAIT_OFF
End Sub
Sub ImgConfirmCancel_onclick
	gFlowWait meWAIT_ON
	ProcessRtn_ConfirmCancel
	gFlowWait meWAIT_OFF
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
		mobjSCGLSpr.ExportExcelFile .sprSht
	
	end with
	gFlowWait meWAIT_OFF
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
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMELECTRANS_NEW.rpt"
		
		for i=1 to .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
				mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
				strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",i)
				strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",i)
				vntData = mobjMDCMELECTRANSLIST.Get_ELETRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
				
				strcntsum = 0
				IF not gDoErrorRtn ("Get_ELETRANS_CNT") then
					for j=1 to mlngRowCnt
						strcnt = 0
						strcnt = vntData(0,j)
						strcntsum =  strcntsum + strcnt
					next
					datacnt = strcntsum + mlngRowCnt
					strUSERID = ""
					vntDataTemp = mobjMDCMELECTRANSLIST.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
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
		intRtn = mobjMDCMELECTRANSLIST.DeleteRtn_temp(gstrConfigXml)
	end with
end sub

Sub imgClose_onclick ()
	Window_OnUnload
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
	dim vntRet
	Dim vntInParams
	with frmThis
	
	vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME.value)
		
	vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
		
	if isArray(vntRet) then
		if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
		.txtCLIENTCODE.value = vntRet(0,0)		        ' Code�� ����
		.txtCLIENTNAME.value = vntRet(1,0)             ' �ڵ�� ǥ��
		gSetChangeFlag .txtCLIENTCODE                  ' gSetChangeFlag objectID	 Flag ���� �˸�
    end if
			
	End with
	
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value,.txtCLIENTNAME.value)
		
			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE.value = vntData(0,0)
					.txtCLIENTNAME.value = vntData(1,0)
				Else
					Call CLIENTCODE_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'****************************************************************************************
' ��Ʈ ����Ŭ�� �̺�Ʈ
'****************************************************************************************

Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
	
		If Row = 0 and Col = 1  then 'AND mobjSCGLSpr.GetTextBinding( .sprSht,"CONFIRMFLAG",Row) = "N"
			If mobjSCGLSpr.GetTextBinding( .sprSht,"CHK",intCnt) <> "" Then
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1,,, , , , , , mstrCheck
				
				if mstrCheck = True then 
					mstrCheck = False
				elseif mstrCheck = False then 
					mstrCheck = True
				end if
				
				for intcnt = 1 to .sprSht.MaxRows
					sprSht_Change 1, intcnt
					
				next
				For intCnt = 1 To .sprSht.MaxRows
					If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
						'����ƽ
						mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
						mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
					'Else
						'üũ
					'	mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
					End If			
				Next
			End IF
		end if
	
	end with
End Sub  	
'------------------------------------------
' ���� �������
'------------------------------------------
Sub ProcessRtn_ConfirmOK
	Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG 
	
	strFLAG = "CONFIRM"
	
	with frmThis
   		
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
   			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		
		If lngCHKSUM = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		'���⼭ ���� ����
		'if DataValidation =false then exit sub
	    '������ Validation End
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG|TAXNO")
		
		intRtn = mobjMDCMELECTRANSLIST.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
		
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ ����" & mePROC_DONE
			'gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
   		end if
   	end with
End Sub
'------------------------------------------
' ������� �������
'------------------------------------------
Sub ProcessRtn_ConfirmCancel

    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	Dim lngCHK,lngCHKSUM
	Dim strFLAG
	strFLAG = "CANCEL"
	with frmThis
   		'������ Validation Start
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		end if
		
   		lngCHK = 0
   		lngCHKSUM = 0
   		For intCnt = 1 to .sprSht.MaxRows
			 IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = "1" THEN
				lngCHK = mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt)
				lngCHKSUM = lngCHKSUM + lngCHK
			END IF
		Next
		If lngCHKSUM = 0 Then
			gErrorMsgBox "������ �����͸� ���� �Ͻʽÿ�.","����ȳ�!"
			Exit Sub
		End If
		
		
		'if DataValidation =false then exit sub
	    '������ Validation End
		On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK|TRANSYEARMON|TRANSNO|CONFIRMFLAG|TAXNO")
		
		intRtn = mobjMDCMELECTRANSLIST.ProcessRtn_Confirm_OK(gstrConfigXml,vntData,strFLAG)
	
		if not gDoErrorRtn ("ProcessRtn_Confirm_OK") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			msgbox lngCHKSUM & " ���� �ڷᰡ �������" & mePROC_DONE
			'gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
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
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		elseif Row = 0 and Col =1 then
		else
			strTRANSYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",Row)
			strTRANSNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",Row)
			
			vntInParams = array(strTRANSYEARMON, strTRANSNO) '<< �޾ƿ��°��
			vntRet = gShowModalWindow("MDCMELECTRANSGUNLIST.aspx",vntInParams , 813,545)
			if isArray(vntRet) then
     		end if
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
	
	'����������ü ����	
	set mobjMDCMGET	= gCreateRemoteObject("cMDCO.ccMDCOGET")
	set mobjMDCMELECTRANSLIST = gCreateRemoteObject("cMDET.ccMDETELECTRANSLIST")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 14, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK|CONFIRMFLAG|TAXNO|TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|AMT |VAT|SUMAMTVAT"
		mobjSCGLSpr.SetHeader .sprSht,		"����|����|��꼭��ȣ|��� | ��ȣ | CLIENTCODE | ������ | MEDCODE | ��ü�� | REAL_MED_CODE | ��ü��|����ݾ� |�ΰ���|��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "4|   6|        10|    8|	 6|		      0|	  25|        0|      20|	          0|     19|       13|    12|13"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT |VAT|SUMAMTVAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "TRANSYEARMON | TRANSNO | CLIENTCODE | CLIENTNAME | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME|CONFIRMFLAG", -1, -1, 20
		'mobjSCGLSpr.SetCellsLock2 .sprSht, true, "AMT |VAT|SUMAMTVAT" 
		mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|CLIENTCODE|MEDCODE|REAL_MED_CODE", true
		.sprSht.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData
	
	'vntInParam = window.dialogArguments
	'intNo = ubound(vntInParam)
	''�⺻�� ����
	'mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	'WITH frmThis
	'	for i = 0 to intNo
	'		select case i
	'			case 0 : .txtYEARMON.value = vntInParam(i)	
	'			case 1 : .txtCLIENTCODE.value = vntInParam(i)
	'			case 2 : .txtCLIENTNAME.value = vntInParam(i)		'���� ������� �͸�
	'			case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
	'			case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
	'		end select
	'	next
	'end with
	SelectRtn	
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMELECTRANSLIST = Nothing
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
	Dim strYEARMON, strCUSTCODE
   	Dim i, strCols
   	Dim intCnt
	on error resume next
	with frmThis
		strYEARMON	= .txtYEARMON.value
		strCUSTCODE	= .txtCLIENTCODE.value

		'�ʱ�ȭ
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		vntData = mobjMDCMELECTRANSLIST.Get_ELECTRANS_ALLLIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,"", strCUSTCODE)
		
		IF not gDoErrorRtn ("Get_ELECTRANS_ALLLIST") then
			'��ȸ�� �����͸� ���ε�
			
			

			
			mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			For intCnt = 1 To .sprSht.MaxRows
				If  mobjSCGLSpr.GetTextBinding( .sprSht,"TAXNO",intCnt) <> "" Then
					'����ƽ
					mobjSCGLSpr.SetCellTypeStatic .sprSht, 1,1, intCnt, intCnt,0,2
					mobjSCGLSpr.SetTextBinding .sprSht,"CHK",intCnt," "
				Else
					'üũ
					mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1,1,intCnt,intCnt,,0,1,2,2,false
				End If			
			Next
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If 	mlngRowCnt < 1 Then
			.sprSht.MaxRows= 0
			End If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE	
		End IF
	end with
End Sub

Sub PreSearchFiledValue (strCUSTCODE, strCUSTNAME)
	frmThis.txtYEARMON.value = strCUSTCODE
	frmThis.txtCLIENTCODE.value = strCUSTNAME		
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm"  width="100%" height="100%" cellSpacing="0" cellPadding="0"  border="0">
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
											<td class="TITLE">&nbsp;����Ź �ŷ����� ����</td>
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
									<TABLE id="tblButton" style="WIDTH: 100px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="100" border="0">
										<TR>
											<TD></TD>
											<TD></TD>
											<TD><!--<IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
													height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
													name="imgExcel">--></TD>
											<TD><!--<IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose">--></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height = "95%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" vAlign="middle" align="center"><FONT face="����">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="WIDTH: 95px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON, '')">�� 
													��</TD>
												<TD class="SEARCHDATA" style="WIDTH: 375px"><INPUT class="INPUT" id="txtYEARMON" title="�����ȸ" accessKey="NUM" type="text" maxLength="6"
														size="10" name="txtYEARMON"></TD>
												<TD class="SEARCHLABEL" style="WIDTH: 95px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME)">������
												</TD>
												<TD class="SEARCHDATA" width="400"><INPUT class="INPUT_L" id="txtCLIENTNAME" title="�ڵ��" style="WIDTH: 312px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="46" name="txtCLIENTNAME"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
												<TD class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
														src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></TD>
											</TR>
										</TABLE>
										<!-- �߰� �����ν���-->
										<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px;HEIGHT: 25px"><FONT face="����"></FONT></TD>
											</TR>
										</TABLE>
										<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0"> <!--background="../../../images/TitleBG.gIF"-->
											<TR>
												<TD align="left" width="400" height="20">
													<table cellSpacing="0" cellPadding="0" width="100%" border="0">
														<tr>
															<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
															<td align="left" height="4"><FONT face="����"></FONT></td>
														</tr>
														<tr>
															<td class="TITLE">&nbsp;�ŷ����� ��ȸ �� ����</td>
														</tr>
													</table>
												</TD>
												<TD vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSetting" onmouseover="JavaScript:this.src='../../../images/imgAgreeOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAgree.gIF'"
																	height="20" alt="�ڷḦ����ó���մϴ�." src="../../../images/imgAgree.gIF" border="0" name="imgSetting">
															</TD>
															<td><IMG id="ImgConfirmCancel" onmouseover="JavaScript:this.src='../../../images/ImgAgreeCancelOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgAgreeCancel.gIF'"
																	height="20" alt="����ó���� ����մϴ�." src="../../../images/ImgAgreeCancel.gif" border="0"
																	name="ImgConfirmCancel">
															</td>
															<td><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
																	height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
																	name="imgPrint">
															</td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel">
															</td>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
										<!--���̺��� �������°��� �����ش�-->
										<TABLE cellSpacing="0" cellPadding="0" width="1040" border="0">
											<TR>
												<TD align="left" width="100%" height="1"></TD>
											</TR>
										</TABLE>
										<!-- �߰� �����γ�-->
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 100%; HEIGHT: 3px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
								<DIV id="pnlTab1" style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%" ms_positioning="GridLayout">
									<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 100%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5 >
									<PARAM NAME="_Version" VALUE="393216">
									<PARAM NAME="_ExtentX" VALUE="32438">
									<PARAM NAME="_ExtentY" VALUE="21352">
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
								</div>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" height = "3"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
