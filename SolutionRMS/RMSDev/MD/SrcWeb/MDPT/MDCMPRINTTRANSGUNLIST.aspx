<%@ Page CodeBehind="MDCMPRINTTRANSGUNLIST.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="MD.MDCMPRINTTRANSGUNLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�μ� ����Ź �ŷ��� ��ȸ</title> 
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
Dim mobjMDCMPRINTTRANS 

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
	if frmThis.txtTRANSYEARMON.value = "" or frmThis.txtTRANSNO.value = "" then
		gErrorMsgBox "����� �ŷ��� ��ȣ�� �Է��Ͻÿ�",""
		exit Sub
	end if
	gFlowWait meWAIT_ON
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
	Dim strUSERID
	Dim intCnt2
	
	'üũ�� �����Ͱ� ���ٸ� �޽����� �Ѹ��� Sub�� ������
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
		Exit Sub
	end if
	
	For intCnt2 = 1 To frmThis.sprSht.MaxRows
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TAXYEARMON",intCnt2) <> "" OR mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TAXNO",intCnt2) <> "" THEN
			gErrorMsgBox mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TRANSYEARMON",intCnt2) & "-" & mobjSCGLSpr.GetTextBinding(frmThis.sprSht,"TRANSNO",intCnt2) & " �� ���Ͽ�" &vbcrlf & "���ݰ�꼭��ȣ�� �����ϴ� ������ ������� �� �����ϴ�.","�μ�ȳ�!"
			Exit Sub
		End If
	Next
	
	gFlowWait meWAIT_ON
	with frmThis
		'�μ��ư�� Ŭ���ϱ� ���� md_trans_temp���̺� ������ �����Ѵ�
		'�μ��Ŀ� temp���̺��� �����ϰ� �Ǹ� ũ����Ż ����Ʈ�� �Ķ���� ���� �Ѿ������
		'�����Ͱ� �����ǹǷ� �Ķ���Ͱ� �Ѿ�� �ʴ´�. by kty
		'md_trans_temp���� ����
		intRtn = mobjMDCMPRINTTRANS.DeleteRtn_temp(gstrConfigXml)
		'md_trans_temp���� ��
		
		ModuleDir = "MD"
		ReportName = "MDCMPRINTTRANS_NEW.rpt"
		
		mlngRowCnt=clng(0): mlngColCnt=clng(0)

		strTRANSYEARMON	= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSYEARMON",1)
		strTRANSNO		= mobjSCGLSpr.GetTextBinding(.sprSht,"TRANSNO",1)
		
		vntData = mobjMDCMPRINTTRANS.Get_PRINTTRANS_CNT(gstrConfigXml,mlngRowCnt,mlngColCnt, strTRANSYEARMON,strTRANSNO)
		
		strcntsum = 0
		IF not gDoErrorRtn ("Get_PRINTTRANS_CNT") then
			for j=1 to mlngRowCnt
				strcnt = 0
				strcnt = vntData(0,j)
				strcntsum =  strcntsum + strcnt
			next
			datacnt = strcntsum 
			for i=1 to 1
				strUSERID = ""
				vntDataTemp = mobjMDCMPRINTTRANS.ProcessRtn_TEMP(gstrConfigXml,strTRANSYEARMON, strTRANSNO, datacnt, strUSERID)
			next
		End IF
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
		intRtn = mobjMDCMPRINTTRANS.DeleteRtn_temp(gstrConfigXml)
	end with
end sub


Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
' �޷�
'-----------------------------------------------------------------------------------------
Sub imgPRINTDAY_onclick
	'CalEndar�� ȭ�鿡 ǥ��
	gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgPRINTDAY,"txtPRINTDAY_onchange()"
	gXMLDataChanged xmlBind           ' gXMLDataChanged  xmlBindID
End Sub

'������
Sub txtPRINTDAY_onchange
	gSetChange
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
		vntInParams = array(.txtCLIENTCODE.value, .txtCLIENTNAME1.value)
			
		vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
			
		if isArray(vntRet) then
			if .txtCLIENTCODE.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE.value = vntRet(0,0)		        ' Code�� ����
			.txtCLIENTNAME1.value = vntRet(1,0)             ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE                  ' gSetChangeFlag objectID	 Flag ���� �˸�
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
			
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtCLIENTCODE.value,.txtCLIENTNAME1.value)
		
			if not gDoErrorRtn ("GetCUSTNO") then
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
' ������ڵ��˾� ��ư[�Է¿�]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTSUBCODE_onclick
	Call CLIENTSUBCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTSUBCODE_POP
	Dim vntRet
	Dim vntInParams

	with frmThis
		vntInParams = array(trim(.txtCLIENTSUBCODE.value), trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("../MDCO/MDCMHIGHCUSTGROUPPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtCLIENTSUBCODE.value = vntRet(0,0) and .txtCLIENTSUBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTSUBCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtCLIENTSUBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = trim(vntRet(5,0))
			.txtCLIENTNAME1.value = trim(vntRet(6,0))
			gSetChangeFlag .txtCLIENTSUBCODE		' gSetChangeFlag objectID	 Flag ���� �˸�
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTSUBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCMGET.GetCUSTNO_HIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTSUBCODE.value),trim(.txtCLIENTSUBNAME.value), trim(.txtCLIENTCODE.value), trim(.txtCLIENTNAME1.value))
			
			if not gDoErrorRtn ("GetCUSTNO_HIGHCUSTCODE") then
				If mlngRowCnt = 1 Then
					.txtCLIENTSUBCODE.value = trim(vntData(0,0))
					.txtCLIENTSUBNAME.value = trim(vntData(1,0))
					.txtCLIENTCODE.value = trim(vntData(5,0))
					.txtCLIENTNAME.value = trim(vntData(6,0))
				Else
					Call CLIENTSUBCODE_POP()
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
   		Dim strYEARMON
		On error resume next
		with frmThis
			If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
				strYEARMON = .txtTRANSYEARMON.value
			End If
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			
			vntData = mobjMDCMGET.GetTRANSNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "PRINT")
			if not gDoErrorRtn ("GetTRANSNO") then
				If mlngRowCnt = 1 Then
					.txtTRANSYEARMON.value = vntData(0,0)  ' Code�� ����
					.txtTRANSNO.value = vntData(1,0)  ' �ڵ�� ǥ��
					.txtCLIENTCODE.value = vntData(2,0)  ' �ڵ�� ǥ��
					.txtCLIENTNAME1.value = vntData(3,0)  ' �ڵ�� ǥ��
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
	dim vntRet
	Dim vntInParams
	Dim strYEARMON
	with frmThis

	If .txtTRANSYEARMON.value <> "" Or Len(.txtTRANSYEARMON.value) = 6 Then
	strYEARMON = .txtTRANSYEARMON.value
	End If
		vntInParams = array(strYEARMON, .txtTRANSNO.value,.txtCLIENTCODE.value, .txtCLIENTNAME1.value, "trans", "PRINT") '<< �޾ƿ��°��
		vntRet = gShowModalWindow("../MDCO/MDCMTRANSPOP.aspx",vntInParams , 413,425)
		if isArray(vntRet) then
			'if .txtTRANSYEARMON.value = vntRet(0,0) and .txtTRANSNO.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtTRANSYEARMON.value = vntRet(0,0)  ' Code�� ����
			.txtTRANSNO.value = vntRet(1,0)  ' �ڵ�� ǥ��
			.txtCLIENTCODE.value = vntRet(2,0)  ' �ڵ�� ǥ��
			.txtCLIENTNAME1.value = vntRet(3,0)  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
Sub sprSht_Click(ByVal Col, ByVal Row)
	dim intcnt
	with frmThis
		if Row > 0 and Col > 1 then		
			'sprShtToFieldBinding Col,Row			
		end if
	end with
End Sub  

sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
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
	set mobjMDCMPRINTTRANS = gCreateRemoteObject("cMDPT.ccMDPTPRINTTRANS")
	set mobjMDCMGET			  = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mobjSCGLCtl.DoEventQueue

    'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
        gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 26, 0, 0, 0
		mobjSCGLSpr.SpreadDataField .sprSht, "TRANSYEARMON|TRANSNO|CLIENTCODE|CLIENTNAME|MEDCODE|MEDNAME| REAL_MED_CODE | REAL_MED_NAME|SUBSEQ|BRANDNAME|CLIENTSUBCODE|CLIENTSUBNAME|PROGRAM_NAME|DEPT_CD|DEMANDDAY|PRINTDAY|AMT| VAT| TAXYEARMON|TAXNO|MED_FLAG"
		mobjSCGLSpr.SetHeader .sprSht,		 "TRANSYEARMON|TRANSNO|CLIENTCODE|CLIENTNAME|MEDCODE|��ü��|REAL_MED_CODE|��ü��|�귣���ڵ�|�귣���|������ڵ�|����θ�|�����|DEPT_CD|DEMANDDAY|��������|���ް���|������|�ΰ���|���ݰ�꼭���|���ݰ�꼭��ȣ|��ü����"
		mobjSCGLSpr.SetColWidth .sprSht, "-1", "         0|	     0|		    0|		   0|      0|    10|	        0|    10|         0|      10|         0|      10|    10|      0|        0|       8|       9|       9|     9|             0|             0|      0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "PRINTDAY", -1, -1, 10
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, " AMT| VAT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeStatic2 .sprSht, "MEDNAME|REAL_MED_NAME|PROGRAM_NAME|MED_FLAG|BRANDNAME|CLIENTSUBNAME", -1, -1, 20
		mobjSCGLSpr.ColHidden .sprSht, "TRANSYEARMON|TRANSNO | CLIENTCODE|CLIENTNAME|MEDCODE|REAL_MED_CODE|DEPT_CD|DEMANDDAY|SUBSEQ|CLIENTSUBCODE|TAXYEARMON|TAXNO ", true
		.sprSht.style.visibility = "visible"
		
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtTRANSYEARMON.value = vntInParam(i)	
				case 1 : .txtTRANSNO.value = vntInParam(i)
				case 2 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 3 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 4 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	SelectRtn
End Sub

Sub EndPage()
	set mobjMDCMGET = Nothing
	set mobjMDCMPRINTTRANS = Nothing
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
		'.txtTRANSYEARMON.value = "200712"
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		.txtCLIENTNAME1.focus
		
	End with

	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strCUSTCODE, strTRANSNO, strCLIENTSUBCODE
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON	= .txtTRANSYEARMON.value
		strTRANSNO	= .txtTRANSNO.value
		strCUSTCODE	= .txtCLIENTCODE.value
		strCLIENTSUBCODE = .txtCLIENTSUBCODE.value
		
		IF not SelectRtn_HDR (strYEARMON, strTRANSNO, strCUSTCODE) Then Exit Sub

		'��Ʈ ��ȸ
		Call SelectRtn_DTL (strYEARMON, strTRANSNO, strCUSTCODE, strCLIENTSUBCODE)
				
	END WITH
	
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub

Function SelectRtn_HDR (ByVal strYEARMON, ByVal strTRANSNO, ByVal strCUSTCODE)
	dim vntData
	on error resume next

	'�ʱ�ȭ
	SelectRtn_HDR = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjMDCMPRINTTRANS.Get_PRINTTRANS_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_HDR") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ �ŷ�����ȣ�� ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			gWriteText "", "������ �ŷ�����ȣ�� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			
			txtAMT_onblur
			txtVAT_onblur
			txtSUMAMTVAT_onblur
			SelectRtn_HDR = True
		End IF
	End IF
End Function

Function SelectRtn_DTL (ByVal strYEARMON,ByVal strTRANSNO, ByVal strCUSTCODE, ByVal strCLIENTSUBCODE)
	dim vntData
	on error resume next

	'�ʱ�ȭ
	SelectRtn_DTL = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	
	vntData = mobjMDCMPRINTTRANS.Get_PRINTTRANS_LIST(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON,strTRANSNO, strCUSTCODE, strCLIENTSUBCODE)
	
	IF not gDoErrorRtn ("Get_PRINTTRANS_LIST") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,TRUE)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_DTL = True
		gWriteText "", "������ �ŷ�����ȣ���� �󼼳����� ���Ͽ�" & mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
	End IF
End Function

Sub PreSearchFiledValue (strCUSTCODE, strCUSTNAME)
	frmThis.txtTRANSYEARMON.value = strCUSTCODE
	frmThis.txtCLIENTCODE.value = strCUSTNAME		
End Sub

'-----------------------------------------------------------------------------------------
' õ���� ������ ǥ�� ( �ܰ�, �ݾ�, ������)
'-----------------------------------------------------------------------------------------
'�ܰ�
Sub txtAMT_onblur
	with frmThis
		call gFormatNumber(.txtAMT,0,true)
	end with
End Sub

Sub txtVAT_onblur
	with frmThis
		call gFormatNumber(.txtVAT,0,true)
	end with
End Sub

Sub txtSUMAMTVAT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMTVAT,0,true)
	end with
End Sub

-->
		</script>
		<XML id="xmlBind"></XML>
	</HEAD>
	<body class="base">
		<form id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 684px; HEIGHT: 403px" cellSpacing="0" cellPadding="0"
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
												&nbsp;����Ź �ŷ����� �󼼰˻�</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 375px" vAlign="middle" align="right" height="28">
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
									<TABLE id="tblButton" style="WIDTH: 203px; HEIGHT: 20px" cellSpacing="0" cellPadding="2"
										width="203" border="0">
										<TR>
											<TD><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'"
													height="20" alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" width="54" border="0"
													name="imgQuery"></TD>
											<TD><IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'"
													height="20" alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF" width="54" border="0"
													name="imgPrint"><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" height="20" alt="�ڷḦ ������ �޽��ϴ�."
													src="../../../images/imgExcel.gIF" width="54" border="0" name="imgExcel"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
									<!--Common Button End--></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<!--TopSplit Start-->
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 794px"><FONT face="����"></FONT></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 791px" vAlign="middle" align="center"><FONT face="����">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtTRANSYEARMON, txtTRANSNO)">
													�ŷ�����ȣ</TD>
												<TD class="SEARCHDATA" width="315"><INPUT class="INPUT" id="txtTRANSYEARMON" title="�ŷ������" style="WIDTH: 72px; HEIGHT: 22px"
														accessKey="NUM" type="text" maxLength="6" size="6" name="txtTRANSYEARMON"><IMG id="ImgTRU" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'" style="CURSOR: hand"
														onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle" border="0" name="ImgTRU"><INPUT class="INPUT" id="txtTRANSNO" title="�ŷ�����ȣ" style="WIDTH: 69px; HEIGHT: 22px" type="text"
														maxLength="6" size="5" name="txtTRANSNO" accessKey="NUM"></TD>
												<TD class="SEARCHLABEL" width="80" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE, txtCLIENTNAME1)">������
												</TD>
												<TD class="SEARCHDATA" width="316"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 227px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="32" name="txtCLIENTNAME1"><IMG id="ImgCLIENTCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtCLIENTCODE"></TD>
											</TR>
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTSUBCODE, txtCLIENTSUBNAME)">�����
												</TD>
												<TD class="SEARCHDATA"><INPUT class="INPUT_L" id="txtCLIENTSUBNAME" title="�ڵ��" style="WIDTH: 222px; HEIGHT: 22px"
														type="text" maxLength="100" align="left" size="31" name="txtCLIENTSUBNAME"><IMG id="ImgCLIENTSUBCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE"><INPUT class="INPUT_L" id="txtCLIENTSUBCODE" title="�ڵ���ȸ" style="WIDTH: 64px; HEIGHT: 22px"
														type="text" maxLength="6" align="left" size="5" name="txtCLIENTSUBCODE"></TD>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPRINTDAY,'')"><FONT face="����">��������</FONT></TD>
												<TD class="SEARCHDATA">
													<INPUT class="INPUT" id="txtPRINTDAY" title="���μ���" style="WIDTH: 82px; HEIGHT: 22px" accessKey="DATE"
														type="text" maxLength="100" size="9" name="txtPRINTDAY"><IMG id="ImgPRINTDAY" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23"
														align="absMiddle" border="0" name="ImgPRINTDAY">
												</TD>
											</TR>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 791px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblDATA" style="WIDTH: 791px; HEIGHT: 6px" cellSpacing="1" cellPadding="0"
										align="right" border="0">
										<TR>
											<TD class="LABEL" width="90"><FONT face="����">������</FONT></TD>
											<TD class="DATA" width="173"></FONT><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="�����ָ�" style="WIDTH: 172px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" align="left" size="22" name="txtCLIENTNAME">
											</TD>
											<TD class="LABEL" width="90"><FONT face="����">���μ�</FONT></TD>
											<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="DEPT_NAME" class="NOINPUT_L" id="txtDEPT_NAME" title="�귣���ڵ�" style="WIDTH: 172px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="22" name="txtDEPT_NAME"></FONT>
											</TD>
											<TD class="LABEL" width="90"><FONT face="����">û������</FONT></TD>
											<TD class="DATA" width="173"><FONT face="����"><INPUT dataFld="DEMANDDAY" class="NOINPUT" id="txtDEMANDDAY" title="�귣���" style="WIDTH: 104px; HEIGHT: 22px"
														accessKey="DATE" dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="12" name="txtDEMANDDAY"></FONT></TD>
										</TR>
										<TR>
											<TD class="LABEL"><FONT face="����">����ݾ�</FONT></TD>
											<TD class="DATA"><FONT face="����"><INPUT dataFld="AMT" class="NOINPUT_R" id="txtAMT" title="����ݾ�" style="WIDTH: 135px; HEIGHT: 22px"
														dataSrc="#xmlBind" readOnly type="text" maxLength="20" size="16" name="txtAMT"></FONT>
											</TD>
											<TD class="LABEL"><FONT face="����">�ΰ���</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT dataFld="VAT" class="NOINPUT_R" id="txtVAT" title="�ΰ���" style="WIDTH: 135px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" maxLength="100" size="17" name="txtVAT"></TD>
											<TD class="LABEL"><FONT face="����">��</FONT></TD>
											<TD class="DATA"></FONT></FONT><INPUT class="NOINPUT_R" id="txtSUMAMTVAT" title="��" style="WIDTH: 154px; HEIGHT: 22px"
													type="text" maxLength="100" size="20" name="txtSUMAMTVAT" readonly dataFld="SUMAMTVAT" dataSrc="#xmlBind"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 791px; HEIGHT: 3px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD align="center">
									<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LISTFRAME" style="HEIGHT: 101px" height="101">
												<OBJECT id="sprSht" style="WIDTH: 790px; HEIGHT: 346px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
													<PARAM NAME="_Version" VALUE="393216">
													<PARAM NAME="_ExtentX" VALUE="20902">
													<PARAM NAME="_ExtentY" VALUE="9155">
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
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
