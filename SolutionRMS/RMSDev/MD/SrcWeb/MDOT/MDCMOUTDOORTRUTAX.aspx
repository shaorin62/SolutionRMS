<%@ Page CodeBehind="MDCMOUTDOORTRUTAX.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMOUTDOORTRUTAX" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� ����Ź ���ݰ�꼭 ����</title> 
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
Dim mobjMDCMOUTDOORTAX , mobjMDCMGET
Dim mstrCheck
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
' ���� ��ư Ŭ�� �̺�Ʈ
'-----------------------------------
Sub imgQuery_onclick
	if frmThis.txtTRANSYEARMON1.value = "" then
	    gErrorMsgBox "��� �Է��Ͻÿ�",""
		exit Sub
	end if
	If LEN(frmThis.txtTRANSYEARMON1.value) <> 6 Then
		 gErrorMsgBox "����� 6�ڸ� �Դϴ�",""
		exit Sub
	End If
	
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgPrint_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "�μ�� �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i,j
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim vntData
	Dim vntDataTemp
	Dim strcnt, strcntsum
	Dim intRtn
	Dim intCount
	Dim VATFLAG
	Dim FLAG
	Dim strUSERID
	
	IF frmThis.sprSht.MaxRows = 0 then
		gFlowWait meWAIT_ON
		with frmThis		
			ModuleDir = "MD"
			ReportName = "TAXNO_BLACK.rpt"
						
			IF .cmbFLAG.value = "receipt" THEN
				FLAG = "Y"
			ELSE
				FLAG = "N"
			END IF
						
			Params = FLAG
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
		end with
		gFlowWait meWAIT_OFF
	else
	
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
			'�μ��ư�� Ŭ���ϱ� ���� md_tax_temp���̺��� ������ �����Ѵ�
			'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
			'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
			'md_trans_temp���� ����
			intRtn = mobjMDCMOUTDOORTAX.DeleteRtn_TEMP(gstrConfigXml)
			'md_trans_temp���� ��
			
			ModuleDir = "MD"
			ReportName = "TAX_OUTDOORNEW.rpt"
			
			for i=1 to .sprSht.MaxRows
				IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = "1" THEN
					mlngRowCnt=clng(0): mlngColCnt=clng(0)
			
					strTAXYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
					strTAXNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
					IF mobjSCGLSpr.GetTextBinding(.sprSht,"VAT",i) = 0 THEN
						VATFLAG = "N"
					ELSE
						VATFLAG = "Y"
					END IF
					IF .cmbFLAG.value = "receipt" THEN
						FLAG = "Y"
					ELSE
						FLAG = "N"
					END IF
					strUSERID = ""
					
					vntDataTemp = mobjMDCMOUTDOORTAX.ProcessRtn_TEMP(gstrConfigXml,strTAXYEARMON, strTAXNO, VATFLAG, FLAG, i, strUSERID)
				END IF
			next
			Params = "MD_TAXOUTDOOR_TEMP" & ":" & strUSERID
			Opt = "A"
			gShowReportWindow ModuleDir, ReportName, Params, Opt
			
			'10���Ŀ� printSetTimeout ����� ȣ���Ͽ� temp���̺��� �����Ѵ�.
			'���ȭ���� �ߴ� �ӵ����� �����ϴ� �ӵ��� ���� �ؿ��� �ٷ� ������ �ȵǱ⶧���� �ð��� ���Ƿ� ��..
			window.setTimeout "printSetTimeout", 10000
		end with
		gFlowWait meWAIT_OFF
	END IF
End Sub	

'����� �Ϸ���� md_trans_temp(��������� ���� �ӽ����̺�)�� �����
Sub printSetTimeout()
	Dim intRtn
	with frmThis
		intRtn = mobjMDCMOUTDOORTAX.DeleteRtn_TEMP(gstrConfigXml)
	end with
end sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub ImgTaxCre_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "���ݰ�꼭 ������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdF.checked <> true then
		gErrorMsgBox "���ݰ�꼭������ �̿Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "������ �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

'��������
Sub imgModifyDelete_onclick ()
	If frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "���������� �����Ͱ� �����ϴ�.","ó���ȳ�!"
		Exit Sub
	End If
	
	If frmThis.rdT.checked <> true then
		gErrorMsgBox "���������� �Ϸ�����϶� �����մϴ�..","ó���ȳ�!"
		Exit Sub
	end if
	
	gFlowWait meWAIT_ON
	DeleteRtn_Modify
	gFlowWait meWAIT_OFF
End Sub



Sub imgClose_onclick ()
	Window_OnUnload
End Sub

Sub ImgTAX_onclick ()
	TAXPOP
End Sub

Sub btnCOMMISSION_onclick ()
	Dim intCnt
	Dim intRtn
	Dim strDEMANDDAY
	Dim strTAXYEARMON
	With frmThis
		If rdT.checked = True OR .rdA.checked = TRUE Then
			gErrorMsgBox "û���� ������ �̿Ϸ���� ���� ����˴ϴ�.","ó���ȳ�!"
			Exit Sub
		End If
		
		strDEMANDDAY = .txtDEMANDDAY.value
		strTAXYEARMON = MID(.txtDEMANDDAY.value,1,4) & MID(.txtDEMANDDAY.value,6,2)
		intRtn = gYesNoMsgbox("���õ� �׸��� û������ ���� �Ͻðڽ��ϱ�?","���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
			
		For intCnt = 1 To .sprSht.MaxRows
			If  mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				mobjSCGLSpr.setTextBinding .sprSht,"DEMANDDAY",intCnt,strDEMANDDAY
				mobjSCGLSpr.setTextBinding .sprSht,"TAXYEARMON",intCnt,strTAXYEARMON
			End If
		Next
	End With
End Sub

Sub TAXPOP
	Dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis
		vntInParams = array(.txtTRANSYEARMON1.value, .txtTRANSNO1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "trans","OUTDOOR") '<< �޾ƿ��°��
		If .rdF.checked = True Then
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		Else
		vntRet = gShowModalWindow("../MDCO/MDCMTAXPOP.aspx",vntInParams , 413,435)
		End If
		if isArray(vntRet) then
			.txtTRANSYEARMON1.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO1.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtCLIENTCODE1.value = vntRet(2,0)  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = vntRet(3,0)  ' �ڵ�� ǥ��
			DateClean
			selectRtn
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' �������ڵ��˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE_onclick
	Call CLIENTCODE_POP()
End Sub

'û���� ��ȸ���� ����
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = MID(frmThis.txtTRANSYEARMON1.value,1,4) & "-" & MID(frmThis.txtTRANSYEARMON1.value,5,2)
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(.txtTRANSYEARMON1.value, .txtCLIENTCODE1.value, .txtCLIENTNAME1.value, "trans", "OUTDOOR") 
		vntRet = gShowModalWindow("../MDCO/MDCMTAXCUSTPOP.aspx",vntInParams , 413,445)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(1,0) and .txtCLIENTNAME1.value = vntRet(2,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = vntRet(1,0)		  ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(2,0)       ' �ڵ�� ǥ��
			IF vntRet(3,0) = "�Ϸ�" THEN
				'window.event.keyCode = meEnter
				'txtTRANSNO1_onkeydown
			ELSE
				.txtTRANSNO1.value = ""
			END IF
			gSetChangeFlag .txtCLIENTCODE1             ' gSetChangeFlag objectID	 Flag ���� �˸�
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
			
			vntData = mobjMDCMGET.GetTAXCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtTRANSYEARMON1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value,"trans","OUTDOOR")
											  'gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtCUSTNO.value,.txtCUSTNAME.value, mtranscommiflag, mtransTblflag
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = vntData(1,0)
					.txtCLIENTNAME1.value = vntData(2,0)
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' �ŷ��������˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'1���� ã����� ��� ��ȸ
Sub txtTRANSNO1_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
   		Dim strYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON1.value <> "" Or Len(.txtTRANSYEARMON1.value) = 6 Then
				strYEARMON = .txtTRANSYEARMON1.value
			End If
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, .txtTRANSNO1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "trans", "OUTDOOR", "0")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON1.value = vntData(0,0)  ' Code�� ����
					.txtTRANSNO1.value = vntData(1,0)  ' �ڵ�� ǥ��
					.txtCLIENTCODE1.value = vntData(2,0)  ' �ڵ�� ǥ��
					.txtCLIENTNAME1.value = vntData(3,0)  ' �ڵ�� ǥ��
				Else
					Call TRANSPOP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

Sub TRANSPOP
	dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis

	If .txtTRANSYEARMON1.value <> "" Or Len(.txtTRANSYEARMON1.value) = 6 Then
	strYEARMON = .txtTRANSYEARMON1.value
	End If
	'msgbox strYEARMON
		vntInParams = array(strYEARMON, .txtTRANSNO1.value,.txtCLIENTCODE1.value,.txtCLIENTNAME1.value, "trans","OUTDOOR") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			.txtTRANSYEARMON1.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO1.value = vntRet(1,0)  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' �޷�
'-----------------------------------------------------------------------------------------
Sub ImgPRINTDAY_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.ImgPRINTDAY,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub


Sub txtTRANSYEARMON1_onblur
	With frmThis
	If .txtTRANSYEARMON1.value <> "" AND Len(.txtTRANSYEARMON1.value) = 6 Then DateClean
	End With
End Sub

Sub imgCalDemandday_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.imgCalDemandday,"txtDEMANDDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'û����
Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
End Sub

Sub imgFROM_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtFROM,frmThis.imgFROM,"txtFROM_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub imgTO_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtTO,frmThis.imgTO,"txtTO_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub img1_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtDEMANDDAY,frmThis.img1,"txtDEMANDDAY_onchange()"
	'gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

Sub txtFROM_onchange
	gSetChange
End Sub

Sub txtTO_onchange
	gSetChange
End Sub

Sub txtDEMANDDAY_onchange
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht, 1, 1, , , "", , , , , mstrCheck
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
	Dim vntInParams
	Dim strMEDFLAG
	strMEDFLAG = "O"
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		Else
			If .rdT.checked = True Then
				vntInParams = array(mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON", Row),mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO", Row),strMEDFLAG) '<< �޾ƿ��°��
				gShowModalWindow "MDCMOUTDOORTRUTAXDTL.aspx",vntInParams , 813,565
				SelectRtn
			End IF
		end if	
	end with
end sub

Sub cmbGUBUN_onchange
	with frmThis
		If .cmbGUBUN.value = "taxdiv" Then
			selectRtn
		Elseif  .cmbGUBUN.value = "taxgroup" Then
			selectRtn
		End If
	End with
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'****************************************************************************************
' ������ ȭ�� ������ �� �ʱ�ȭ 
'****************************************************************************************
Sub InitPage()
	'����������ü ����	
	set mobjMDCMOUTDOORTAX	= gCreateRemoteObject("cMDOT.ccMDOTOUTDOORTRUTAX")
	set mobjMDCMGET			= gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 38, 0, 1, 2
		mobjSCGLSpr.SpreadDataField .sprSht, "CHK | TAXYEARMON | TAXNO | TAXMANAGE | TRANSYEARMON | TRANSNO | SEQ | DEMANDDAY | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | VAT | SUMAMT | COMMI_RATE | COMMISSION | SUMM | DEPT_NAME | PRINTDAY  | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPTCODE | MEDFLAG | VOCHNO | RANKTRANS | PROGRAM | SPONSOR| CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetHeader .sprSht,		   "����|���|��ȣ|������ȣ|���|��ȣ|����|û����|�����ָ�|�����ֻ���ڵ�Ϲ�ȣ|��ü���|��ü�����ڵ�Ϲ�ȣ|��ü��|����ݾ�|�ΰ�����|�հ�ݾ�|��������|������|����|�μ���|������|�������ڵ�|������AC�ڵ�|��ü���ڵ�|��ü��AC�ڵ�|MEDCODE|�μ��ڵ�|���豸��|��ǥ��ȣ|����|�����|��������|CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "   4|	5|   4|	     11|   5|   4|	 4|     8|	    19|                  17|	  19|                  17|     9|       9|       9|	     9|        7|     9|  30|    10| 	8|         0|           0|         0|           0|      0|       0|       0|      10|   0|    30|       0|0|0|0|0|0|0"
			
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "DEMANDDAY|PRINTDAY"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT|SUMAMT|COMMISSION", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "COMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "AMT|VAT|SUMAMT|COMMI_RATE|COMMISSION"
		mobjSCGLSpr.SetCellsLock2 .sprSht,true, "TAXYEARMON|TAXNO|TRANSYEARMON|TRANSNO|CLIENTNAME|REAL_MED_NAME|MEDNAME|AMT|VAT|SUMAMT|DEPT_NAME|CLIENTCODE|CLIENTACCODE|CLIENTBISNO|REAL_MED_CODE|REAL_MED_ACCODE|REAL_MED_BISNO|MEDCODE|DEPTCODE|MEDFLAG|SEQ|VOCHNO|RANKTRANS|PROGRAM"
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MEDNAME|REAL_MED_NAME|PROGRAM|TAXMANAGE", -1, -1, 20
		mobjSCGLSpr.ColHidden .sprSht, "TAXNO|CLIENTCODE|CLIENTACCODE|REAL_MED_CODE|REAL_MED_ACCODE|MEDCODE|DEPTCODE|MEDFLAG|VOCHNO|TAXYEARMON|RANKTRANS|SPONSOR|CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2", true
		.sprSht.style.visibility = "visible"
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDCMOUTDOORTAX = Nothing
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
		.txtTRANSYEARMON1.value = Mid(gNowDate,1,4)  & Mid(gNowDate,6,2)
		'Sheet�ʱ�ȭ
		DateClean
		.sprSht.MaxRows = 0
		
		.txtCLIENTNAME1.focus
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'�Ϸ�üũ
Sub rdT_onclick
	SelectRtn
End Sub

'�̿Ϸ�üũ
Sub rdF_onclick
	SelectRtn
End Sub

'��üüũ
Sub rdA_onclick
	SelectRtn
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCUSTCODE, strTRANSNO,strGUBUN
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	Dim strGROUP
   
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		If .txtTRANSYEARMON1.value = "" Then
			gErrorMsgBox "����� �Է��Ͻʽÿ�","��ȸ�ȳ�!"
			Exit Sub
		End If	
		If Len(.txtTRANSYEARMON1.value) <> 6 Then
			gErrorMsgBox "����� ������ �ƴմϴ�.","��ȸ�ȳ�!"
			Exit Sub
		End If
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		strYEARMON	= .txtTRANSYEARMON1.value
		strTRANSNO	= .txtTRANSNO1.value
		strTAXNO = .txtTRANSNO1.value
		strCUSTCODE	= .txtCLIENTCODE1.value
		strGUBUN = .cmbGUBUN.value
		'���ݰ�꼭 �Ϸ���ȸ
		If .rdT.checked = True Then
			vntData = mobjMDCMOUTDOORTAX.Get_OUTDOOR_TAX(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strCUSTCODE,strFROM,strTO,strTAXNO)
			If not gDoErrorRtn ("Get_OUTDOOR_TAX") then
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				mobjSCGLSpr.ColHidden .sprSht, "TAXMANAGE|VOCHNO", False
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|TAXYEARMON|TAXNO|SEQ|MEDNAME|PROGRAM|COMMI_RATE|COMMISSION", True
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
			End If
		'�̿Ϸ� �ŷ������� ������ ��ȸ
		ElseIf .rdF.checked = True Then
			vntData = mobjMDCMOUTDOORTAX.Get_OUTDOOR_TAXBUILD(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE,strGUBUN)
			If not gDoErrorRtn ("Get_OUTDOOR_TAXBUILD") then
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
				mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON|TAXNO|TAXMANAGE|VOCHNO", True
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|SEQ|MEDNAME|PROGRAM|COMMI_RATE|COMMISSION", False
				Layout_change
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
			End If
		ElseIf .rdA.checked = True Then			
			vntData = mobjMDCMOUTDOORTAX.Get_OUTDOOR_TAXALL(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE, strGUBUN)
			If not gDoErrorRtn ("Get_OUTDOOR_TAXALL") then
				'�ʱ� ���·� ����
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
				mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CHK ", -1, -1, 100
				mobjSCGLSpr.ColHidden .sprSht, "TAXYEARMON|TAXNO|TAXMANAGE|VOCHNO", True
				mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO|SEQ|MEDNAME|PROGRAM|COMMI_RATE|COMMISSION", False
				'��ȸ�� �����͸� ���ε�
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				
				Layout_change
				gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE
				mobjSCGLSpr.ActiveCell .sprSht, 2, 1
			End If
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub

Sub Layout_change ()
	Dim intCnt
	with frmThis
	For intCnt = 1 To .sprSht.MaxRows 
		If mobjSCGLSpr.GetTextBinding(.sprSht,"SPONSOR",intCnt) = "Y" Then
		mobjSCGLSpr.SetCellShadow .sprSht, -1, -1, intCnt, intCnt,&H99CCFF, &H000000,False
		End If
	Next 
	End With
End Sub

Sub PreSearchFiledValue (strCUSTCODE, strCUSTNAME)
	frmThis.txtTRANSYEARMON1.value = strCUSTCODE
	frmThis.txtCLIENTCODE1.value = strCUSTNAME		
End Sub

'****************************************************************************************
' ������ ó��
'****************************************************************************************
Sub ProcessRtn ()
   	Dim intRtn
    Dim intRtn2
   	dim vntData, vntData1
	Dim strMasterData
	Dim strTAXYEARMON
	Dim intTAXNO
	Dim strTAXSET
	Dim strSUMM
	Dim intCnt
	Dim strDEMANDDAY,strPRINTDAY
	Dim chkcnt
	Dim intCnt2
	Dim intColFlag
	Dim intMaxCnt
	Dim bsdiv
	Dim strVALIDATION
	with frmThis
		'�������� xml ���� ó���Ҽ� �����Ƿ� �ݵ�� ����üũ �ʿ�
		If .rdT.checked = True Then
			gErrorMsgBox "�̿Ϸ� ���¿��� ������ �����մϴ�.","����ȳ�!"
			Exit Sub
		End If
		intRtn2 = gYesNoMsgbox("û������ Ȯ���ϼ̽��ϱ�?","Ȯ��")
		IF intRtn2 <> vbYes then exit Sub
			
		For intCnt = 1 To .sprSht.MaxRows
		strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
		strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
			If strDEMANDDAY  = "" Then
				gErrorMsgBox "û������ �ʼ� �Դϴ�.","����ȳ�!"
				Exit Sub
			End If
			If  strPRINTDAY = "" Then
				gErrorMsgBox "û������ �ʼ� �Դϴ�.","����ȳ�!"
				Exit Sub
			End If
		Next
		'üũ ���� ��� ���� �ȵǵ���
		chkcnt = 0
		For intCnt2 = 1 To .sprSht.MaxRows
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
				chkcnt = chkcnt + 1
			END IF
		next
		if chkcnt = 0 then
			gErrorMsgBox "���ݰ�꼭�� ������ �����͸� üũ �Ͻʽÿ�","����ȳ�!"
			exit sub
		end if
		 '�����÷��� ����
		mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
		gXMLSetFlag xmlBind, meINS_TRANS

   		If .sprSht.MaxRows = 0 Then
   			gErrorMsgBox "���׸� �� �����ϴ�.",""
   			Exit Sub
   		End If
		'if DataValidation =false then exit sub
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"CHK | TAXYEARMON | TAXNO | TAXMANAGE | TRANSYEARMON | TRANSNO | SEQ | DEMANDDAY | CLIENTNAME | CLIENTBISNO | REAL_MED_NAME | REAL_MED_BISNO | MEDNAME | AMT | VAT | SUMAMT | SUMM | DEPT_NAME | PRINTDAY  | CLIENTCODE | CLIENTACCODE | REAL_MED_CODE | REAL_MED_ACCODE | MEDCODE | DEPTCODE | MEDFLAG | VOCHNO | RANKTRANS | PROGRAM | SPONSOR| CLIENTOWNER| CLIENTADDR1| CLIENTADDR2| REAL_MEDOWNER| REAL_MEDADDR1| REAL_MEDADDR2")
		
		'������ �����͸� ���� �´�.
		'
		
		'ó�� ������ü ȣ��
		intTAXNO = 0
		If .cmbGUBUN.value = "taxdiv" Then
		intRtn = mobjMDCMOUTDOORTAX.ProcessRtn_Div(gstrConfigXml,vntData, intTAXNO)
		Else
		
			'validation
			If Not TaxGroup(strVALIDATION) Then 
				gErrorMsgBox strVALIDATION & vbCrlf & "������ [������,��ü��] [û����,�ۼ���,����] �� ���� �Ͽ��� �մϴ�.","����ȳ�!"
				Exit Sub
			Else
				'�ִ밪
				intColFlag = 0
				For intMaxCnt = 1 To .sprSht.MaxRows
					If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intMaxCnt) = 1 Then
						bsdiv = cint(mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intMaxCnt))
						IF intColFlag < bsdiv THEN
							intColFlag = bsdiv
							'rowflag = lngCnt
						END IF
					End IF
				Next
				'�ƽ����� �߰��Ͽ� ������
				intRtn = mobjMDCMOUTDOORTAX.ProcessRtn_Group(gstrConfigXml,vntData, intTAXNO,intColFlag)
			End IF
		End If

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox "���强��","����ȳ�!"
			.rdT.checked = True
			.txtTRANSNO1.value = ""
			selectRtn
   		end if
   	end with
End Sub


Function TaxGroup(ByRef strVALIDATION)
	Dim intCnt
	Dim strCLIENTCODE '������ ����� ��Ϲ�ȣ
	Dim strREAL_MED_CODE 'û���� ����� ��Ϲ�ȣ
	Dim strDEMANDDAY
	Dim strPRINTDAY
	Dim strSUMM
	Dim strStartRank
	TaxGroup = False
	with frmThis
		strStartRank = "0"
		strCLIENTCODE = ""
		strREAL_MED_CODE = ""
		strDEMANDDAY = ""
		strPRINTDAY = ""
		strSUMM = ""
		For intCnt = 1 To .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt) = 1 Then
				
				If strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt) Then
					If strCLIENTCODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBISNO",intCnt) Then
						Exit Function
					End If
					If strREAL_MED_CODE <> mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt) Then
						Exit Function
					End If 
					If strDEMANDDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt) Then
						strVALIDATION = "û����Ȯ�� �ŷ���������ȣ " & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If 
					If strPRINTDAY <> mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt) Then
						strVALIDATION = "������Ȯ�� �ŷ���������ȣ" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If 
					If strSUMM <> mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt) Then
						strVALIDATION = "����Ȯ�� �ŷ���������ȣ" & mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",intCnt) & " ��"
						Exit Function
					End If 
				End If
				
				strStartRank = mobjSCGLSpr.GetTextBinding(.sprSht,"RANKTRANS",intCnt)
				strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTBISNO",intCnt)
				strREAL_MED_CODE = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_MED_BISNO",intCnt)
				strDEMANDDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"DEMANDDAY",intCnt)
				strPRINTDAY = mobjSCGLSpr.GetTextBinding(.sprSht,"PRINTDAY",intCnt)
				strSUMM = mobjSCGLSpr.GetTextBinding(.sprSht,"SUMM",intCnt)
				
			End If
		Next
	End With
	TaxGroup = True
End Function

Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strDESCRIPTION
	with frmThis
	strDESCRIPTION = ""
		For intCnt2 = 1 To .sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
			If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "��ǥ�� �����ϴ� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
				Exit Sub
			End If
		END IF
		Next
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
			
				intRtn = mobjMDCMOUTDOORTAX.DeleteRtn_TruTax(gstrConfigXml,strTAXYEARMON, strTAXNO)
				IF not gDoErrorRtn ("DeleteRtn_TruTax") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"�����ȳ�!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn_TruTax") then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ������ ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub


Sub DeleteRtn_Modify ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strTAXYEARMON
	Dim strTAXNO
	Dim strDESCRIPTION
	Dim strMEDFLAG
	with frmThis
	strDESCRIPTION = ""
		For intCnt2 = 1 To .sprSht.MaxRows
		if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",intCnt2) = 1 THEN
			If mobjSCGLSpr.GetTextBinding(.sprSht,"VOCHNO",intCnt2) <> "" THEN
				gErrorMsgBox mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "��ǥ�� �����ϴ� ������ ������ ���� �ʽ��ϴ�.","�����ȳ�!"
				Exit Sub
			End If
		END IF
		Next
		IF gDoErrorRtn ("DeleteRtn") then exit Sub
		
		intRtn = gYesNoMsgbox("�ڷḦ ���������Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht,"CHK",i) = 1 THEN
				strTAXNO = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXNO",i)
				strTAXYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"TAXYEARMON",i)
				strMEDFLAG = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAG",i)
				
				intRtn = mobjMDCMOUTDOORTAX.ProcessRtn_MODIFYTEMP(gstrConfigXml,strTAXYEARMON,strTAXNO, strMEDFLAG)
				
				if not gDoErrorRtn ("ProcessRtn_MODIFYTEMP") then
					gErrorMsgBox "�������强��","����ȳ�!"
   				end if
			
				intRtn = mobjMDCMOUTDOORTAX.DeleteRtn_TruTax(gstrConfigXml,strTAXYEARMON, strTAXNO)
				IF not gDoErrorRtn ("DeleteRtn_TruTax") then
					If strDESCRIPTION <> "" Then
						gErrorMsgBox strDESCRIPTION,"�����ȳ�!"
						Exit Sub
					End If
					mobjSCGLSpr.DeleteRow .sprSht,i
   				End IF
   				intCnt = intCnt + 1
   			END IF
		next
		
		IF not gDoErrorRtn ("DeleteRtn_TruTax") then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
		'���� ������ ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn
	End with
	err.clear	
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 1040px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
				width="684" border="0">
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
											<td class="TITLE">
												&nbsp;����Ź ���ݰ�꼭 ����</td>
										</tr>
									</table>
								</td>
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
									<!--Wait Button End-->
								</TD>
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
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, txtTRANSNO1)"
												width="80">��ȣ</TD>
											<TD class="SEARCHDATA" width="200"><INPUT class="INPUT" id="txtTRANSYEARMON1" title="�ŷ��������" style="WIDTH: 72px; HEIGHT: 22px"
													accessKey="MON" type="text" maxLength="6" size="6" name="txtTRANSYEARMON1"><IMG id="ImgTAX" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTAX"><INPUT class="INPUT" id="txtTRANSNO1" title="�ŷ�������ȣ" style="WIDTH: 56px; HEIGHT: 22px"
													type="text" maxLength="6" size="4" name="txtTRANSNO1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80">������
											</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 203px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="28" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="ImgCLIENTCODE"><INPUT class="INPUT" id="txtCLIENTCODE1" title="�ڵ���ȸ" style="WIDTH: 65px; HEIGHT: 22px"
													type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtFROM, txtTO)"
												width="80">û����
											</TD>
											<TD class="SEARCHDATA" width="300"><INPUT class="INPUT" id="txtFROM" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="6" size="6" name="txtFROM"><IMG id="imgFrom" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgFrom">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="û������" style="WIDTH: 72px; HEIGHT: 22px" accessKey="date"
													type="text" maxLength="6" size="6" name="txtTO"><IMG id="imgTo" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgTo">
											</TD>
											<td class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" border="0" name="imgQuery" align="absMiddle"></td>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON1, txtTRANSNO1)"
												width="80">����</TD>
											<TD class="SEARCHDATA" colspan="6">&nbsp;&nbsp;&nbsp;&nbsp;<INPUT id="rdT" type="radio" name="rdGBN" title="�Ϸ᳻����ȸ">&nbsp;�Ϸ�&nbsp;&nbsp;&nbsp; 
												&nbsp; <INPUT id="rdF" type="radio" name="rdGBN" title="�̿Ϸ� ������ȸ" CHECKED>&nbsp;�̿Ϸ�&nbsp;&nbsp;&nbsp; 
												&nbsp;<INPUT id="rdA" title="��ü ������ȸ" type="radio" value="on" name="rdGBN">&nbsp;��ü</TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
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
														<td class="TITLE">
															&nbsp;���ݰ�꼭 ����</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<td><IMG id="ImgTaxCre" onmouseover="JavaScript:this.src='../../../images/ImgTaxCreOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgTaxCre.gif'"
																height="20" alt="���õǾ��� ��Ŀ� ���� ���ݰ�꼭�� �ۼ��մϴ�." src="../../../images/ImgTaxCre.gif"
																align="absMiddle" border="0" name="ImgTaxCre"></td>
														<td><IMG id="imgModifyDelete" onmouseover="JavaScript:this.src='../../../images/imgModifyDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgModifyDelete.gif'"
																height="20" alt="�ڷḦ ����/�����մϴ�." src="../../../images/imgModifyDelete.gIF" border="0"
																name="imgModifyDelete"></td>
														<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
																height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" width="54" border="0"
																name="imgDelete"></TD>
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
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
									<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 1040px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="left" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 67px">����/�ջ�</TD>
											<TD class="DATA" style="WIDTH: 189px"><SELECT id="cmbGUBUN" title="��ü����" style="WIDTH: 96px" name="cmbGUBUN">
													<OPTION value="taxdiv" selected>���ҹ���</OPTION>
													<OPTION value="taxgroup">�ջ����</OPTION>
												</SELECT>
											</TD>
											<TD class="LABEL" style="WIDTH: 73px; CURSOR: hand" onclick="vbscript:Call gCleanField(txtDEMANDDAY, '')">û���� 
												����</TD>
											<TD class="DATA" style="WIDTH: 294px"><INPUT class="INPUT" id="txtDEMANDDAY" title="û������" style="WIDTH: 112px; HEIGHT: 22px"
													accessKey="date" type="text" maxLength="8" size="13" name="txtDEMANDDAY"><IMG id="Img1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="Img1"><IMG id="btnCOMMISSION" onmouseover="JavaScript:this.src='../../../images/imgAppOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgApp.gIF'" height="20" alt="�ش� �����Ϸ� �������� ���� ���׸��� Setting �մϴ�"
													src="../../../images/imgApp.gIF" width="54" align="absMiddle" border="0" name="btnCOMMISSION">
											</TD>
											<TD class="LABEL" style="WIDTH: 71px">��꼭����</TD>
											<TD class="DATA" style="WIDTH: 113px">
												<SELECT id="chkPRINT" title="��¹�����" style="WIDTH: 112px" name="chkPRINT">
													<OPTION value="1" selected>�Ϲݿ�</OPTION>
													<OPTION value="0">���޹޴��ڿ�</OPTION>
												</SELECT>
											</TD>
											<TD class="LABEL" style="WIDTH: 83px">û��/��������</TD>
											<td class="DATA">
												<SELECT id="cmbFLAG" title="����/û������" style="WIDTH: 120px" name="cmbFLAG">
													<OPTION value="receipt" selected>û��</OPTION>
													<OPTION value="demand">����</OPTION>
												</SELECT>
											</td>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
				</TR>
				<TR>
					<TD class="BODYSPLIT"></TD>
				</TR>
				<TR>
					<TD align="center">
						<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="LISTFRAME" style="HEIGHT: 101px" height="101">
									<OBJECT id="sprSht" style="WIDTH: 1040px; HEIGHT: 416px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="27517">
										<PARAM NAME="_ExtentY" VALUE="11007">
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
						</TABLE>
					</TD>
				</TR>
				<TR>
					<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
				</TR>
			</TABLE>
			</TD></TR></TBODY></TABLE></form>
	</body>
</HTML>