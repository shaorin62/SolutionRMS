<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCLOUD.aspx.vb" Inherits="MD.MDCMCLOUD" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>CGV Ŭ���� û�����</title>
		<meta name="vs_snapToGrid" content="False">
		<META content="text/html; charset=ks_c_5601-1987" http-equiv="Content-Type">
		<!--
'****************************************************************************************
'�ý��۱��� : ����Ź�ŷ����� ��� ȭ��(MDCMPRINTTRANS1.aspx)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SheetSample.aspx
'��      �� : ����Ź�ŷ����� �Է�/���� ó��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/28 By Kim Tae Yub
'			 2) 
'****************************************************************************************
-->
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<meta name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<LINK rel="STYLESHEET" type="text/css" href="../../Etc/STYLEs.CSS">
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script id="clientEventHandlersVBS" language="vbscript">
<!--
option explicit
Dim mlngRowCnt, mlngColCnt
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Dim mobjMDOTCLOUD, mobjMDCOGET
Dim mstrCheck, mstrCheck1
Dim mProcess_CHK '��ܿ� �ٸ� �׸��� �����Ͱ� ������ �Ǿ����� üũ

CONST meTAB = 9
mstrCheck=True
mstrCheck1=True

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

'�ʱ�ȭ��ư
Sub imgCho_onclick
	InitPageData
End Sub

sub ImgAddRow_OUT_onclick ()
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "����� ��� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
	
		IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
			gErrorMsgBox "�ش� �����ʹ� �ŷ������� ����� ������ �Դϴ�. �߰� �Ͻ� �� �����ϴ�.","����ȳ�"
			Exit Sub
		end if
		
		call sprSht_OUT_Keydown(meINS_ROW, 0)
		.txtCLIENTCODE1.focus
		.sprSht_OUT.focus
	End With 
End sub

sub ImgAddRow_CGV_onclick ()
	With frmThis
		If .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "����� ��� ������ ������ �߰��� �� �����ϴ�.","����ȳ�"
			Exit Sub
		End If
		
		IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
			gErrorMsgBox "�ش� �����ʹ� �ŷ������� ����� ������ �Դϴ�. �߰� �Ͻ� �� �����ϴ�.","����ȳ�"
			Exit Sub
		end if
		
		call sprSht_CGV_Keydown(meINS_ROW, 0)
		.txtCLIENTCODE1.focus
		.sprSht_CGV.focus
	End With 
End sub

Sub ImgSave_AMT_onclick
	If frmThis.sprSht_HDR.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_AMT
	gFlowWait meWAIT_OFF
End Sub

Sub ImgSave_OUT_onclick
	If frmThis.sprSht_OUT.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_OUT
	gFlowWait meWAIT_OFF
End Sub

Sub ImgSave_CGV_onclick
	If frmThis.sprSht_CGV.MaxRows = 0 Then
   		gErrorMsgBox "���׸� �� �����ϴ�.",""
   		Exit Sub
   	End If
   	
	gFlowWait meWAIT_ON
	ProcessRtn_CGV
	gFlowWait meWAIT_OFF
End Sub

'-----------�μ�-----------
Sub imgPrint_onclick ()
	Dim ModuleDir 	    '����� ����
	Dim ReportName      '����Ʈ �̸�
	Dim Params		    '�Ķ����(VARCHAR2)
	Dim Opt             '�̸����� "A" : �̸�����, "B" : ���
	Dim i
	Dim chkcnt
	Dim strCONT_CODE
	
	Dim strYEARMON, strCLIENTCODE, strTITLE, strCONT_TYPE
	
	
	Dim Con1, Con2, Con3, Con4, Con5
	
	with frmThis
		Con1 = "" : Con2 = "" : Con3 = "" : Con4 = "" : Con5 = ""
		
		if frmThis.sprSht_HDR.MaxRows = 0 then
			gErrorMsgBox "�μ��� �����Ͱ� �����ϴ�.",""
			Exit Sub
		end if
		
		ModuleDir = "MD"
		
		ReportName = "MDCMCGVCLOUD.rpt"
		
		strYEARMON		 = .txtYEARMON.value
		strCLIENTCODE	 = .txtCLIENTCODE1.value
		strTITLE		 = .txtTITLE1.value
		strCONT_TYPE	 = .cmbCONT_TYPE.value
		
		If strYEARMON		<> ""	Then Con1  = " AND (YEARMON = '" & strYEARMON & "') "
		If strCLIENTCODE	<> ""	Then Con2  = " AND (CLIENTCODE = '" & strCLIENTCODE & "') "
		If strTITLE			<> ""	Then Con3  = " AND (CONT_NAME = '" & strTITLE & "') " 
		
		If strCONT_TYPE <> "" Then 
			If strCONT_TYPE = "B" Then '�������
				Con4 = " AND (B.CONT_TYPE = '01')"
			Else
				Con4 = " AND (B.CONT_TYPE = '02')"
			End If
		End If
		
		chkcnt=0
		For i=1 To .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = "1" Then
				if chkcnt = 0 then
					strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)
				else
					strCONT_CODE = strCONT_CODE & "," & mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)  
				end if 
				chkcnt = chkcnt +1
			End If
			
		Next
		
		if chkcnt <> 0 then
			Con5 = " AND ( CONT_CODE IN (" & strCONT_CODE &"))"
		end if 
		
		Params = Con1 & ":" & Con2 & ":" & Con3 & ":" & Con4 & ":" & Con5
		Opt = "A"
		gShowReportWindow ModuleDir, ReportName, Params, Opt
	end with  
End Sub	

Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgDelete_OUT_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_OUT
	gFlowWait meWAIT_OFF
End Sub
	
Sub imgDelete_CGV_onclick
	gFlowWait meWAIT_ON
	DeleteRtn_CGV
	gFlowWait meWAIT_OFF
End Sub

Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportMerge = true
		mobjSCGLSpr.ExcelExportOption = true
		mobjSCGLSpr.ExportExcelFile .sprSht_HDR
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgClose_onclick ()
	Window_OnUnload
End Sub

'-----------------------------------------------------------------------------------------
'��ȸ�� �̺�Ʈ
'-----------------------------------------------------------------------------------------
'�������˾���ư
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE_POP
	Dim vntRet
	Dim vntInParams
	With frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
	    vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,435)
		If isArray(vntRet) Then
			If .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) Then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))	    ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			SelectRtn
		End If
	End With
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtCLIENTNAME1_onkeydown
	If window.event.keyCode = meEnter Then
		Dim vntData
   		Dim i, strCols
		'On error resume Next
		With frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjMDCOGET.GetHIGHCUSTCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value), "D")
			
			If not gDoErrorRtn ("GetHIGHCUSTCODE") Then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,1))
					.txtCLIENTNAME1.value = trim(vntData(1,1))
					SelectRtn
				Else
					Call CLIENTCODE_POP()
				End If
   			End If
   		End With
		window.event.returnValue = False
		window.event.cancelBubble = True
	End If
End Sub

'****************************************************************************************
' ��Ʈ Ŭ�� �̺�Ʈ
'****************************************************************************************
Sub sprSht_HDR_Click(ByVal Col, ByVal Row)
	with frmThis
		if  Row > 0 AND Col > 1 then
			SelectRtn_OUT Col, Row
			SelectRtn_CGV Col, Row
		end if
	end with
End Sub

Sub sprSht_OUT_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_OUT, 1, 1, , , "", , , , , mstrCheck1
			if mstrCheck1 = True then 
				mstrCheck1 = False
			elseif mstrCheck1 = False then 
				mstrCheck1 = True
			end if
			for intcnt = 1 to .sprSht_OUT.MaxRows
				sprSht_OUT_Change 1, intcnt
			next
		end if
	end with
End Sub  

Sub sprSht_CGV_Click(ByVal Col, ByVal Row)
	Dim intcnt
	with frmThis
		if Row = 0 and Col = 1 then
			mobjSCGLSpr.SetCellTypeCheckBox .sprSht_CGV, 1, 1, , , "", , , , , mstrCheck1
			if mstrCheck1 = True then 
				mstrCheck1 = False
			elseif mstrCheck1 = False then 
				mstrCheck1 = True
			end if
			for intcnt = 1 to .sprSht_CGV.MaxRows
				sprSht_CGV_Change 1, intcnt
			next
		end if
	end with
End Sub  

'��� ��Ʈ ����Ŭ��
sub sprSht_HDR_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht_HDR, ""
		end if
	end with
end sub

'----------------------------------------------------------
' [��Ʈ Ű��]
'----------------------------------------------------------
Sub sprSht_HDR_Keyup(KeyCode, Shift)
	Dim intRtn
	Dim strCOLUMN
	Dim i, j
	Dim vntData_col, vntData_row
	
	If KeyCode = 229 Then Exit Sub
	
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	If KeyCode = 17 or KeyCode = 33 or KeyCode = 34 or KeyCode = 35 or KeyCode = 36 or KeyCode = 38 or KeyCode = 40 Then
		SelectRtn_OUT frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		SelectRtn_CGV frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
	End If
	
	With frmThis
		KeyUp_SumAmt .sprSht_HDR
	End With
End Sub

Sub sprSht_OUT_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt_OUT .sprSht_OUT
	End With
End Sub

Sub sprSht_CGV_Keyup(KeyCode, Shift)
	If KeyCode = 229 Then Exit Sub
	If KeyCode <> meCR and KeyCode <> meTab _
		and KeyCode <> 37 and KeyCode <> 38 and KeyCode <> 39 and KeyCode <> 40 _
		and KeyCode <> 17 and KeyCode <> 33 and KeyCode <> 34 and KeyCode <> 35 _
		and KeyCode <> 36 and KeyCode <> 38 and KeyCode <> 40 Then Exit Sub

	With frmThis
		KeyUp_SumAmt_CGV .sprSht_CGV
	End With
End Sub

'---------------------------------------------
'��Ʈ ���콺 ��
'---------------------------------------------
'û���� 
Sub sprSht_HDR_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt .sprSht_HDR
	end with
End Sub

'����� ��Ʈ
Sub sprSht_OUT_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt_OUT .sprSht_OUT
	end with
End Sub

'CGV ��Ʈ 
Sub sprSht_CGV_Mouseup(KeyCode, Shift, X,Y)
	with frmThis
		MouseUp_SumAmt_CGV .sprSht_CGV
	end with
End Sub

'-----------------------------------------------
'���������Ʈ ü���� �̺�Ʈ
'-----------------------------------------------
Sub sprSht_HDR_Change(ByVal Col, ByVal Row)
   Dim intAMT		'���û����
   Dim intMC_AMT	'���û����
   Dim intTIM_AMT	'�系����
   Dim intEX_AMT	'�������
   Dim intCGV_AMT	'CGV ����
   Dim intTIM_RATE	'�系������
   Dim intEX_RATE	'���ۼ�����
   Dim intCGV_RATE	'CGV ������
   
	With frmThis
		
		'�����ְ� ������ �Ŀ�ĳ��Ʈ�ϰ�� �ڵ� ������� �ʴ´�. _�����ƾ� ��û ����..._20120224
		if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CLIENTCODE",Row) <> "A00220" then
		
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"AMT") or Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"TIM_RATE") or _
			Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"EX_RATE") or Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"CGV_RATE")  Then 
				AMT_CALCUL Col,Row
				CGV_AMT_CAL 
			end if

			If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"TIM_AMT") or Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"EX_AMT")  Then 
				TIM_EX_CALCUL Col,Row
			end if
			
			If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"CGV_AMT") Then 
				CGV_CALCUL Col,Row
				CGV_AMT_CAL
			end if

		'	If Col = mobjSCGLSpr.CnvtDataField(.sprSht_HDR,"MC_AMT") Then 
		'		AMT_CALCUL Col,Row
		'	end if
		
			if Row = 1 then
				mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow 
			else
				mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow -1
			end if 
			
			'SelectRtn_OUT .sprSht_HDR.ActiveCol,.sprSht_HDR.ActiveRow
   			'SelectRtn_CGV .sprSht_HDR.ActiveCol,.sprSht_HDR.ActiveRow
   		end if 
   		mobjSCGLSpr.CellChanged .sprSht_HDR, Col, Row
	End With
End Sub

'-----------------------------
'�ݾ� ��� ���� ���ν���
'-----------------------------
'��û���� AMT �� �ְų� ������� �����ϸ� ���п� ���� �ڵ� ����Ѵ�.
sub AMT_CALCUL (Col,Row)
   Dim intAMT		'���û����
   Dim intMC_AMT	'���û����
   Dim intTIM_AMT	'�系����
   Dim intEX_AMT	'�������
   Dim intCGV_AMT	'CGV ����
   Dim intTIM_RATE	'�系������
   Dim intEX_RATE	'���ۼ�����
   Dim intCGV_RATE	'CGV ������
	
	With frmThis
	
		if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "CLIENT" then
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			intTIM_RATE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_RATE",ROW)
			intCGV_RATE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CGV_RATE",ROW)
			
			intTIM_AMT = (intAMT * intTIM_RATE /100)
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"TIM_AMT",Row, intTIM_AMT '�系���ͱ�
			
			intCGV_AMT = (intAMT - intTIM_AMT) * intCGV_RATE /100
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_AMT",Row, intCGV_AMT 'CGV���ͱ�
			
			intMC_AMT = (intAMT - intTIM_AMT - intCGV_AMT)
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"MC_AMT",Row, intMC_AMT 'MC���ͱ�
			
			
		elseif mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "EXCLIENT" then
		
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			intEX_RATE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EX_RATE",ROW)
			intCGV_RATE = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CGV_RATE",ROW)
			
			intEX_AMT = (intAMT * intEX_RATE /100)
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"EX_AMT",Row, intEX_AMT '������ͱ�
			
			intCGV_AMT = (intAMT - intEX_AMT) * intCGV_RATE /100
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_AMT",Row, intCGV_AMT 'CGV���ͱ�
			
			intMC_AMT = (intAMT - intEX_AMT - intCGV_AMT)
			
			mobjSCGLSpr.SetTextBinding .sprSht_HDR,"MC_AMT",Row, intMC_AMT 'MC���ͱ�
		
		end if
	end with
end sub

'���п� ���� �系 ���ͱ��̳� ��������� �����ϸ� ������� �ڵ� ����Ѵ�.
sub TIM_EX_CALCUL (Col,Row)
   Dim intAMT		'���û����
   Dim intMC_AMT	'���û����
   Dim intTIM_AMT	'�系����
   Dim intEX_AMT	'�������
   Dim intCGV_AMT	'CGV ����
   Dim intTIM_RATE	'�系������
   Dim intEX_RATE	'���ۼ�����
   Dim intCGV_RATE	'CGV ������
	
	With frmThis
	
		if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "CLIENT" then
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			if intAMT <> 0 then
				intTIM_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",ROW)
				
				intTIM_RATE = (intTIM_AMT / intAMT) * 100
				
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"TIM_RATE",Row, intTIM_RATE '���μ�����
				AMT_CALCUL Col,Row
			else
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"TIM_AMT",Row, 0 '���μ���
				AMT_CALCUL Col,Row
			end if
			
		elseif mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "EXCLIENT" then
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			if intAMT <> 0 then
				intEX_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EX_AMT",ROW) 
				
				intEX_RATE = (intEX_AMT / intAMT) * 100 
				
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"EX_RATE",Row, intEX_RATE '���������
				AMT_CALCUL Col,Row
			else
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"EX_AMT",Row, 0 '�������
				AMT_CALCUL Col,Row
			end if 
		
		end if
	end with
end sub

'CGV �ݾ��� �����ϸ� CGV ������� �ڵ� ����Ѵ�.
sub CGV_CALCUL (Col,Row)
   Dim intAMT		'���û����
   Dim intMC_AMT	'���û����
   Dim intTIM_AMT	'�系����
   Dim intEX_AMT	'�������
   Dim intCGV_AMT	'CGV ����
   Dim intTIM_RATE	'�系������
   Dim intEX_RATE	'���ۼ�����
   Dim intCGV_RATE	'CGV ������
	
	With frmThis
		if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "CLIENT" then
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			if intAMT <> 0 then
				intTIM_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",ROW)
				intCGV_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CGV_AMT",ROW)
				
				intCGV_RATE = (100/ (intAMT - intTIM_AMT)) * intCGV_AMT
			
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_RATE",Row, intCGV_RATE 'CGV������
				AMT_CALCUL Col,Row
			else 
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_AMT",Row, 0 'CGV����
				AMT_CALCUL Col,Row
			end if 
			
		elseif mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",ROW) = "EXCLIENT" then
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"AMT",ROW)
			if intAMT <> 0 then
				intEX_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EX_AMT",ROW) 
				intCGV_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CGV_AMT",ROW)
				
				intCGV_RATE = (100/ (intAMT - intEX_AMT)) * intCGV_AMT
			
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_RATE",Row, intCGV_RATE 'CGV������
				AMT_CALCUL Col,Row
			else
				mobjSCGLSpr.SetTextBinding .sprSht_HDR,"CGV_AMT",Row, 0 'CGV����
				AMT_CALCUL Col,Row
			end if
		
		end if
	end with
end sub



Sub sprSht_OUT_Change(ByVal Col, ByVal Row)
   	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	
   	Dim intTIM_AMT
   	Dim intEX_AMT
   	Dim intRATE
   	Dim intSettingAMT
		
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		'���μ� ��ȸ
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"DEPT_NAME") Then 
			strCode		= mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEPT_CD",Row)
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"DEPT_NAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",Row, ""
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"DEPT_CD",Row) = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCC_CGV(gstrConfigXml,mlngRowCnt,mlngColCnt, strCodeName, strCodeName)

				If not gDoErrorRtn ("GetCC_CGV") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",Row, trim(vntData(0,1))
						mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_NAME",Row, trim(vntData(1,1))
						
						.txtCLIENTNAME1.focus()
						.sprSht_OUT.focus
					Else
						mobjSCGLSpr_ClickProc mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"DEPT_NAME"), Row, .sprSht_OUT
						.txtCLIENTNAME1.focus()
						.sprSht_OUT.focus 
					End If
   				End If
   			End If
		End If
	
		'�系������ �ݾ׺���
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"RATE") Then 
			intTIM_AMT = 0
			intEX_AMT = 0
			intRATE = 0
			intTIM_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow)
			intEX_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EX_AMT",.sprSht_HDR.ActiveRow)
			intRATE = mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow)
			
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"GUBUN",.sprSht_HDR.ActiveRow) = "CLIENT" then	
				intSettingAMT = (intTIM_AMT * intRATE) / 100
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",Row, intSettingAMT	
			else
				intSettingAMT = (intEX_AMT * intRATE) / 100
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",Row, intSettingAMT	
			end if 

		end if 
		mobjSCGLSpr.CellChanged .sprSht_OUT, Col, Row
		
	End With
End Sub

Sub sprSht_CGV_Change(ByVal Col, ByVal Row)
   	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
		
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"MEDNAME") Then 
			strCode		= ""
			strCodeName = TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"MEDNAME",Row))
			mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDCODE",Row, ""
			If strCode = "" AND strCodeName <> "" Then			
				vntData = mobjMDCOGET.GetCGVCODE(gstrConfigXml,mlngRowCnt,mlngColCnt, strCode, strCodeName)

				If not gDoErrorRtn ("GetCGVCODE") Then
					If mlngRowCnt = 1 Then
						mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDCODE",Row, vntData(0,1)
						mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDNAME",Row, vntData(1,1)
						mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_CODE",Row, vntData(3,1)
						mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_NAME",Row, vntData(4,1)
						
						.txtCLIENTCODE1.focus
						.sprSht_CGV.focus
					Else
						mobjSCGLSpr_ClickProc2 mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"MEDNAME"), Row, .sprSht_CGV
						.txtCLIENTCODE1.focus
						.sprSht_CGV.focus 
					End If
  				End If
 			End If
		End If
	
		mobjSCGLSpr.CellChanged .sprSht_CGV, Col, Row
	End With
End Sub

Sub mobjSCGLSpr_ClickProc(Col, Row, sprSht)
	Dim vntRet
	Dim vntInParams
	With frmThis

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"DEPT_NAME") Then	
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"DEPT_CD",Row)),TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"DEPT_NAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTCGVPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_NAME",Row, vntRet(1,0)
				
				mobjSCGLSpr.CellChanged .sprSht_OUT, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_OUT, Col+2,Row
			End If
		End If
		
		.txtCLIENTNAME1.focus
		.sprSht_OUT.Focus
	End With
End Sub


Sub mobjSCGLSpr_ClickProc2(Col, Row, sprSht)
	Dim vntRet
	Dim vntInParams
	With frmThis

		If Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"MEDNAME") Then	
			vntInParams = array("", TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"MEDNAME",Row)))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCGVPOP.aspx",vntInParams , 413,440)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_CODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_NAME",Row, vntRet(4,0)
				
				mobjSCGLSpr.CellChanged .sprSht_CGV, Col,Row
				mobjSCGLSpr.ActiveCell .sprSht_CGV, Col+2,Row
			End If
		End If
		
		.txtCLIENTNAME1.focus
		.sprSht_CGV.Focus
	End With
End Sub


'-------------------------------------------
'�������� ��Ʈ ��ư Ŭ��
'-------------------------------------------
Sub sprSht_OUT_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"BTN_DEPT") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"DEPT_CD",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"DEPT_NAME",Row)))
			vntRet = gShowModalWindow("../MDCO/MDCMDEPTCGVPOP.aspx",vntInParams , 413,435)
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_NAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht_OUT, Col,Row
			End If
			.txtCLIENTCODE1.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht_CGV.Focus
			mobjSCGLSpr.ActiveCell .sprSht_OUT, Col+2, Row
		End If
	End with
End Sub

Sub sprSht_CGV_ButtonClicked (Col,Row,ButtonDown)
	Dim vntRet, vntInParams
	with frmThis
		IF Col = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"BTN_MED") Then
			vntInParams = array(TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"MEDCODE",Row)), TRIM(mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"MEDNAME",Row)))
								
			vntRet = gShowModalWindow("../MDCO/MDCMCGVPOP.aspx",vntInParams , 413,435)
			
			If isArray(vntRet) Then
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"MEDNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_CODE",Row, vntRet(3,0)
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"REAL_MED_NAME",Row, vntRet(4,0)
				mobjSCGLSpr.CellChanged .sprSht_CGV, Col,Row
			End If
			.txtCLIENTCODE1.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht_CGV.Focus
			mobjSCGLSpr.ActiveCell .sprSht_CGV, Col+2, Row
		end if
	End with
End Sub


'------------------------------------
'��Ʈ Ű�ٿ� �̺�Ʈ
'------------------------------------
Sub sprSht_HDR_Keydown(KeyCode, Shift)

	with frmThis
		if keycode = meENTER then
			mobjSCGLSpr.ActiveCell .sprSht_HDR, .sprSht_HDR.ActiveCol +1, .sprSht_HDR.ActiveRow-1
		end if
	end with
end sub

Sub sprSht_OUT_Keydown(KeyCode, Shift)
	Dim intRtn
	With frmThis
		
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
		if .sprSht_OUT.ActiveRow = .sprSht_OUT.MaxRows and .sprSht_OUT.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_OUT,"MEMO") Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_OUT, cint(13), cint(Shift), -1, 1)
			if mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow-1) <> "" and .sprSht_OUT.MaxRows > 1 then
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow-1)
			else
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
			end if 
			
								
			if .sprSht_OUT.maxRows = 0 then 
				intRtn = mobjSCGLSpr.InsDelRow(.sprSht_OUT, meINS_ROW, 0, -1, 1)

				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"CONT_CODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
				
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"GUBUN",ActiveRow) = "CLIENT" then	
					mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow)
				else
					mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EX_AMT",.sprSht_HDR.ActiveRow)
				end if 
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow, 100
			else 
				intRtn = mobjSCGLSpr.InsDelRow(.sprSht_OUT, meINS_ROW, 0, -1, 1)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"CONT_CODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, 0
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow, 0
			end if 
		
		End if
	ElseIf KeyCode = meINS_ROW  Then
		if .sprSht_OUT.maxRows = 0 then 
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_OUT, meINS_ROW, 0, -1, 1)
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"CONT_CODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
			
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"GUBUN",.sprSht_HDR.ActiveRow) = "CLIENT" then	
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow)
			else
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_CD",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTCODE",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"DEPT_NAME",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EXCLIENTNAME",.sprSht_HDR.ActiveRow)
				mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"EX_AMT",.sprSht_HDR.ActiveRow)
			end if 
			
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow, 100
		else 
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_OUT, meINS_ROW, 0, -1, 1)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"YEARMON",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"CONT_CODE",.sprSht_OUT.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"AMT",.sprSht_OUT.ActiveRow, 0
			mobjSCGLSpr.SetTextBinding .sprSht_OUT,"RATE",.sprSht_OUT.ActiveRow, 0
		end if 
		
	End if
	End With
End Sub

Sub sprSht_CGV_Keydown(KeyCode, Shift)
	Dim intRtn
	With frmThis
	If KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab Then Exit Sub
	
	if KeyCode = meCR  Or KeyCode = meTab Then
		if .sprSht_CGV.ActiveRow = .sprSht_CGV.MaxRows and .sprSht_CGV.ActiveCol = mobjSCGLSpr.CnvtDataField(.sprSht_CGV,"MEMO") Then
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_CGV, cint(13), cint(Shift), -1, 1)
			if mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"YEARMON",.sprSht_CGV.ActiveRow-1) <> "" and .sprSht_CGV.MaxRows > 1 then
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"YEARMON",.sprSht_CGV.ActiveRow, mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"YEARMON",.sprSht_CGV.ActiveRow-1)
			else
				mobjSCGLSpr.SetTextBinding .sprSht_CGV,"YEARMON",.sprSht_CGV.ActiveRow, Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)
			end if 
			
			intRtn = mobjSCGLSpr.InsDelRow(.sprSht_CGV, meINS_ROW, 0, -1, 1)	
			mobjSCGLSpr.SetTextBinding .sprSht_CGV,"CONT_CODE",.sprSht_CGV.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
			CGV_AMT_CAL
		End if
	ElseIf KeyCode = meINS_ROW  Then
		intRtn = mobjSCGLSpr.InsDelRow(.sprSht_CGV, meINS_ROW, 0, -1, 1)
		mobjSCGLSpr.SetTextBinding .sprSht_CGV,"CONT_CODE",.sprSht_CGV.ActiveRow, mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
		CGV_AMT_CAL
	End if
	End With
End Sub


'-----------------------------------
'��Ʈ���� Ű�� �ݾ� �ջ� �̺�Ʈ
'-----------------------------------
SUB KeyUp_SumAmt (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"TOTAL_AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"TAX_AMT") or _
							  mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"TIM_AMT") or _ 
							  mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT ") or mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT") or _
							  mobjSCGLSpr.CnvtDataField(sprSht,"CGV_AMT") Then
		
			strSUM = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"TOTAL_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"TAX_AMT")) or _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"TIM_AMT")) or _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT")) or (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT")) or _
										    (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"CGV_AMT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

SUB KeyUp_SumAmt_OUT (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")  Then
		
			strSUM = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and  (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"AMT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

SUB KeyUp_SumAmt_CGV (sprsht)
	Dim intRtn
	Dim strSUM
	Dim intColCnt, intRowCnt
	Dim i, j
	Dim vntData_col, vntData_row
	
	with frmThis
		If sprSht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"OUT_AMT")  Then
		
			strSUM = 0

			vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprSht,intColCnt, False)
			vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprSht,intRowCnt)

			FOR i = 0 TO intColCnt -1
				If vntData_col(i) <> "" and  (vntData_col(i) = mobjSCGLSpr.CnvtDataField(sprSht,"OUT_AMT")) Then
				
					FOR j = 0 TO intRowCnt -1
						If vntData_row(j) <> "" Then
							strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
						End If
					Next
				End If
			Next

			.txtSELECTAMT.value = strSUM
			Call gFormatNumber(.txtSELECTAMT,0,True)
		else
			.txtSELECTAMT.value = 0
		End If
	end with
END SUB

'-----------------------------------
'��Ʈ���� ���콺�� �ݾ��ջ� �̺�Ʈ
'-----------------------------------
sub MouseUp_SumAmt(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"TOTAL_AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"TAX_AMT") or _
								  mobjSCGLSpr.CnvtDataField(sprSht,"AMT") or mobjSCGLSpr.CnvtDataField(sprSht,"TIM_AMT") or _ 
								  mobjSCGLSpr.CnvtDataField(sprSht,"EX_AMT ") or mobjSCGLSpr.CnvtDataField(sprSht,"MC_AMT") or _
								  mobjSCGLSpr.CnvtDataField(sprSht,"CGV_AMT")  Then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

sub MouseUp_SumAmt_OUT(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"AMT") Then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
			end if
		end if 
	end with
end sub

sub MouseUp_SumAmt_CGV(sprSht)
Dim intRtn
Dim strSUM
Dim intColCnt, intRowCnt
Dim i,j
Dim vntData_col, vntData_row

	with frmThis
		strSUM = 0
		intColCnt = 0
		intRowCnt = 0
		
		if sprSht.MaxRows > 0  then
			if sprsht.ActiveCol = mobjSCGLSpr.CnvtDataField(sprSht,"OUT_AMT") Then
				
				vntData_col = mobjSCGLSpr.GetSelectedItemNo(sprsht,intColCnt,false)
				vntData_row = mobjSCGLSpr.GetSelectedItemNo(sprsht,intRowCnt)
					
				for i = 0 to intColCnt -1
					if vntData_col(i) <> "" then
						FOR j = 0 TO intRowCnt -1
							If vntData_row(j) <> "" Then
								if typename(mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))) = "String" then
									exit sub
								end if 
								strSUM = strSUM + mobjSCGLSpr.GetTextBinding(sprSht,vntData_col(i),vntData_row(j))
								
							End If
						Next
					end if 
				next

				.txtSELECTAMT.value = strSUM
				Call gFormatNumber(.txtSELECTAMT,0,True)
			else
				.txtSELECTAMT.value = 0
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
	set mobjMDOTCLOUD = gCreateRemoteObject("cMDOT.ccMDOTCLOUD")
	set mobjMDCOGET	  = gCreateRemoteObject("cMDCO.ccMDCOGET")

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	mobjSCGLCtl.DoEventQueue
	
	'Sheet �⺻Color ����
    gSetSheetDefaultColor() 
    
    With frmThis
		'-------------------------------------
		'��� �׸���
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_HDR
		mobjSCGLSpr.SpreadLayout .sprSht_HDR, 25, 0, 0, 0,0
		mobjSCGLSpr.SpreadDataField .sprSht_HDR, "CHK | YEARMON | SEQ | GUBUN | CONT_TYPE | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | TBRDSTDATE | TBRDEDDATE | TOTAL_AMT | TAX_AMT | AMT | TIM_RATE | TIM_AMT | EXCLIENTCODE | EXCLIENTNAME | EX_RATE | EX_AMT | CGV_RATE | MC_AMT | CGV_AMT | MEMO | COMMI_TRANS_NO"
		mobjSCGLSpr.SetHeader .sprSht_HDR,		 "����|���|û�����|����|��������|�������ڵ�|�����ָ�|����ڵ�|����|������|������|�Ѱ��ݾ�|��û����|���û����|�系���͹����|�系���ͱݾ�|������ڵ�|������|�����������|���������|CGV�����|MC�ݾ�|CGV�ݾ�|���|�ŷ�������ȣ"
		mobjSCGLSpr.SetColWidth .sprSht_HDR, "-1","  4|   5|       0|   7|       7|         0|      15|	      8|    15|     8|     8|        10|      10|        10|             5|          10|         0|      15|           5|        10|        5|    10|     10|  15|             0"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_HDR, "0", "15"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_HDR, "CHK"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "SEQ | TOTAL_AMT | TAX_AMT | AMT | TIM_AMT | EX_AMT | MC_AMT | CGV_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_HDR, "TIM_RATE | EX_RATE | CGV_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht_HDR, "TBRDSTDATE | TBRDEDDATE", -1, -1, 10
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_HDR, "YEARMON | GUBUN | CONT_TYPE | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | EXCLIENTCODE | EXCLIENTNAME | MEMO | COMMI_TRANS_NO", -1, -1, 200
		mobjSCGLSpr.SetCellsLock2 .sprSht_HDR, true, "SEQ | GUBUN | CONT_TYPE | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | TBRDSTDATE | TBRDEDDATE | TOTAL_AMT | TAX_AMT | EXCLIENTCODE | EXCLIENTNAME | COMMI_TRANS_NO"
		
		'-------------------------------------
		'�系���� �׸���
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_OUT
		mobjSCGLSpr.SpreadLayout .sprSht_OUT, 10, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_OUT, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_OUT, "CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | BTN_DEPT | DEPT_NAME | RATE | AMT | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_OUT,		 "����|���|����|����ڵ�|�μ��ڵ�|�μ���|����|�ݾ�|���"
		mobjSCGLSpr.SetColWidth .sprSht_OUT, "-1", " 4|   6|   4|       8|       8|2|  10|  10|    12|10"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_OUT, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_OUT, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_OUT,"..", "BTN_DEPT"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_OUT, " YEARMON | CONT_CODE | DEPT_CD | DEPT_NAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, " SEQ | AMT ", -1, -1, 0
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_OUT, "RATE", -1, -1, 2
		mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, True, "SEQ | CONT_CODE "
		mobjSCGLSpr.SetCellAlign2 .sprSht_OUT, "CHK | YEARMON",-1,-1,2,2,False
		mobjSCGLSpr.ColHidden .sprSht_OUT, "CONT_CODE", True
		
		'-------------------------------------
		'CGV ���� �׸���
		'-------------------------------------
		gSetSheetColor mobjSCGLSpr, .sprSht_CGV
		mobjSCGLSpr.SpreadLayout .sprSht_CGV, 11, 0, 0, 0,0
		mobjSCGLSpr.AddCellSpan  .sprSht_CGV, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht_CGV, "CHK | YEARMON | SEQ | CONT_CODE | MEDCODE | BTN_MED | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | MEMO"
		mobjSCGLSpr.SetHeader .sprSht_CGV,		 "����|���|����|����ڵ�|�����ڵ�|������|û�����ڵ�|û������|�ݾ�|���"
		mobjSCGLSpr.SetColWidth .sprSht_CGV, "-1", " 4|   6|   4|       8|       5|2|   8|        10|      10|   8|  10"
		mobjSCGLSpr.SetRowHeight .sprSht_CGV, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht_CGV, "0", "20"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht_CGV, "CHK"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht_CGV,"..", "BTN_MED"
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht_CGV, " YEARMON | CONT_CODE | MEDCODE | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | MEMO", -1, -1, 100
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht_CGV, "SEQ | OUT_AMT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht_CGV, True, "SEQ | CONT_CODE | REAL_MED_CODE | REAL_MED_NAME"
		mobjSCGLSpr.SetCellAlign2 .sprSht_CGV, "CHK | YEARMON",-1,-1,2,2,False
		mobjSCGLSpr.ColHidden .sprSht_CGV, "CONT_CODE | REAL_MED_CODE ", True
		
		.sprSht_HDR.style.visibility = "visible"
		.sprSht_OUT.style.visibility = "visible"
		.sprSht_CGV.style.visibility = "visible"
    End With

	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub EndPage()
	set mobjMDOTCLOUD = Nothing
	set mobjMDCOGET = Nothing
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
		.txtYEARMON.value = Mid(gNowDate2,1,4)  & Mid(gNowDate2,6,2)

		.sprSht_HDR.MaxRows = 0	
		.sprSht_OUT.MaxRows = 0
		.sprSht_CGV.MaxRows = 0
		
		CALL COMBO_TYPE ()
	End with
End Sub

'-----------------------------------------------------------------------------------------
' COMBO TYPE ����
'-----------------------------------------------------------------------------------------
Sub COMBO_TYPE()
	Dim vntData, vntData_SEARCH, vntData_GUBUN
	
	With frmThis   
		On error resume Next
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
       	
       	vntData_SEARCH  = mobjMDOTCLOUD.GetDataType_SEARCH(gstrConfigXml, mlngRowCnt, mlngColCnt, "CLOUD_CONTTYPE")
       	vntData			= mobjMDOTCLOUD.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt, "CLOUD_CONTTYPE")
       	vntData_GUBUN	= mobjMDOTCLOUD.GetDataType(gstrConfigXml, mlngRowCnt, mlngColCnt, "CLOUD_GUBUN")
       	
		If not gDoErrorRtn ("GetDataTypeChange") Then 
			gLoadComboBox .cmbCONT_TYPE, vntData_SEARCH, False
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht_HDR, "CONT_TYPE",,,vntData,,60 
			mobjSCGLSpr.SetCellTypeComboBox2 .sprsht_HDR, "GUBUN",,,vntData_GUBUN,,60 
			mobjSCGLSpr.TypeComboBox = True 
   		End If  
   		gSetChange
   	end With   
End Sub

'****************************************************************************************
' ������ ��ȸ
'****************************************************************************************
Sub SelectRtn ()
	Dim vntData, vntData2
	Dim strYEARMON
	Dim strCLIENTCODE, strTITLE, strCONT_TYPE
   	Dim i, strCols
   	Dim intCnt , strRows
    
	with frmThis
		intCnt = 1
		
		If .txtYEARMON.value = "" Then
			gErrorMsgBox "��ȸ�� ����� �ݵ�� �־�� �մϴ�.",""
			Exit SUb
		End If 
		
		'Sheet�ʱ�ȭ
		.sprSht_HDR.MaxRows = 0
		.sprSht_OUT.MaxRows = 0
		.sprSht_CGV.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON		= .txtYEARMON.value
		strCLIENTCODE	= .txtCLIENTCODE1.value
		strTITLE		= .txtTITLE1.value
		strCONT_TYPE	= .cmbCONT_TYPE.value
		
		vntData = mobjMDOTCLOUD.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt, _
											  strYEARMON, strCLIENTCODE, strTITLE, strCONT_TYPE)
													
		If not gDoErrorRtn ("SelectRtn_HDR") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_HDR,vntData,1,1,mlngColCnt,mlngRowCnt,True)
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				
   				for i =1 to .sprSht_HDR.maxRows
   					'�����ְ� ������ �Ŀ�ĳ��Ʈ�� �ƴҰ�츸 ���.()�Ŀ�ĳ��Ʈ �����Է�
   					if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CLIENTCODE",i) <> "A00220" then
   						if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"GUBUN",i) = "CLIENT" then
   							mobjSCGLSpr.SetCellsLock2 frmThis.sprSht_HDR,True,i,19,20,True '������ �����ָ� �����ݾ� ���
   						else
   								mobjSCGLSpr.SetCellsLock2 frmThis.sprSht_HDR,True,i,15,16,True '������ ������ �系���� ���
	   						 
   						end if 
   					end if
   					
   					If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",i) <> "" THEN  '�ŷ������� ������ ��� �����
   						If intCnt = 1 Then
							strRows = i
						Else
							strRows = strRows & "|" & i
						End If
						intCnt = intCnt + 1
   					END IF 
   				next
   				
   				mobjSCGLSpr.SetCellsLock2 .sprSht_HDR,True,strRows,2,25,True
   				
   				AMT_SUM (.sprSht_HDR)
   				SelectRtn_OUT 1,1
   				SelectRtn_CGV 1,1
   			else
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				.sprSht_HDR.MaxRows = 0
   			End If
   		End If
   	end with
End Sub


'----------�系����---------
Sub SelectRtn_OUT (Col, Row)
	Dim vntData
	Dim strCONT_CODE, strYEARMON
   	Dim i, strCols
    Dim intCnt, strRows
    
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_OUT.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",Row)
				
		vntData = mobjMDOTCLOUD.SelectRtn_OUT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
																							
		If not gDoErrorRtn ("SelectRtn_OUT") Then
			If mlngRowCnt >0 Then
				Call mobjSCGLSpr.SetClipBinding (.sprSht_OUT,vntData,1,1,mlngColCnt,mlngRowCnt,True)

   				gWriteText lblStatus_OUT, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   				
   				'����� û�� ������ �ŷ������� ������ ��� �׸��带 ��װ� �ƴѰ�� Ǭ��
   				IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
   					mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, True, " YEARMON | SEQ | CONT_CODE | DEPT_CD | BTN_DEPT | DEPT_NAME | RATE | AMT | MEMO "	
   				else 
   					mobjSCGLSpr.SetCellsLock2 .sprSht_OUT, False, " YEARMON | SEQ | CONT_CODE | DEPT_CD | BTN_DEPT | DEPT_NAME | RATE | AMT | MEMO "	
   				END IF 
   				
   			End If
   		End If
   	end with
End Sub

'----------CGV����---------
Sub SelectRtn_CGV (Col, Row)
	Dim vntData
	Dim strCONT_CODE, strYEARMON
    Dim intCnt
    
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht_CGV.MaxRows = 0

		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0) : mlngColCnt=clng(0)
		
		
		strYEARMON = .txtYEARMON.value
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",Row)
		
		'����Ǿ��ִ� cgv������ ��ȸ�Ѵ�.
		intCnt = mobjMDOTCLOUD.SelectRtn_CGVCHK(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
		IF not gDoErrorRtn ("SelectRtn_CGVCHK") then
			If mlngRowCnt > 0 Then
				'����� ���� �ִٸ� ����� ���� �Ѹ���.
				vntData = mobjMDOTCLOUD.SelectRtn_CGV(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
				If not gDoErrorRtn ("SelectRtn_CGV") Then
					If mlngRowCnt >0 Then
						Call mobjSCGLSpr.SetClipBinding (.sprSht_CGV,vntData,1,1,mlngColCnt,mlngRowCnt,True)
						gWriteText lblStatus_CGV, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   					End If
   				End If
   			'����� ���� ���ٸ� ��ü ������ �Ѹ���.
   			else
   				vntData = mobjMDOTCLOUD.SelectRtn_CGVEmpty(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON, strCONT_CODE)
				If not gDoErrorRtn ("SelectRtn_CGVEmpty") Then
					If mlngRowCnt >0 Then
						Call mobjSCGLSpr.SetClipBinding (.sprSht_CGV,vntData,1,1,mlngColCnt,mlngRowCnt,True)
						
						CGV_AMT_CAL '��� �����͸� �Ѹ��ٸ� ����� cgv �ݾ��� ������ �������� ����Ѵ�.
						
						gWriteText lblStatus_CGV, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   					End If
   				End If
			end if 
			
			'����� û�� ������ �ŷ������� ������ ��� �׸��带 ��װ� �ƴѰ�� Ǭ��
   			IF mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" THEN
   				mobjSCGLSpr.SetCellsLock2 .sprSht_CGV, True, " YEARMON | SEQ | CONT_CODE | MEDCODE | BTN_MED | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | MEMO"	
   			else 
   				mobjSCGLSpr.SetCellsLock2 .sprSht_CGV, False, " YEARMON | SEQ | CONT_CODE | MEDCODE | BTN_MED | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | MEMO"	
   			END IF 
   			
			
		end if
		
   	end with
End Sub

'���� ���� �ݾ��� ����Ѵ�.
sub CGV_AMT_CAL
	Dim intAMT
	Dim intSUMAMT, intSUMDTLAMT
	Dim i , j, intCNT
	
	with frmThis
		intSUMAMT = 0
		intSUMDTLAMT = 0
		intCNT = 0
		intSUMAMT = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CGV_AMT",.sprSht_HDR.ActiveRow)
		
		intCNT = .sprSht_CGV.MaxRows
		
		intAMT = INT(intSUMAMT / intCNT)
	
		for i = 1 to .sprSht_CGV.MaxRows
			mobjSCGLSpr.SetTextBinding .sprSht_CGV,"OUT_AMT",i, intAMT
			intSUMDTLAMT = intSUMDTLAMT + intAMT
		next
		
		IF intSUMAMT <> intSUMDTLAMT THEN
			mobjSCGLSpr.SetTextBinding .sprSht_CGV,"OUT_AMT",1, mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"OUT_AMT",1) + (intSUMAMT - intSUMDTLAMT)
		END IF
		
	end with
end sub

'-----------------------------------
'��ȸ�� �ݾ� �ջ�� ���� �ջ�
'-----------------------------------
Sub AMT_SUM (sprSht)
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	With frmThis
		IntAMTSUM = 0
		
		For lngCnt = 1 To sprSht.MaxRows
			IntAMT = 0
			IntAMT = mobjSCGLSpr.GetTextBinding(sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		If sprSht.MaxRows = 0 Then
			.txtSUMAMT.value = 0
		else
			.txtSUMAMT.value = IntAMTSUM
			Call gFormatNumber(frmThis.txtSUMAMT,0,True)
		End If
	End With
End Sub

'**********************************************
'----------------------����--------------------
'**********************************************
'-------------��� û�� ������ ����-----------
Sub ProcessRtn_AMT ()
	Dim intRtn
   	Dim vntData
   	Dim intCnt
	Dim strDataCHK
	Dim lngCol, lngRow
	Dim strCONT_CODE

	With frmThis
	
		'---------------------DATAVALIDATION------------------------
		if .sprSht_HDR.MaxRows = 0 Then
			gErrorMsgBox "�����Ͻ� �����Ͱ� �����ϴ�.","����ȳ�!"
			Exit Sub
		end if
		
		'��� �ʼ��Է»���
		strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_HDR, "YEARMON",lngCol, lngRow, False) 
		If strDataCHK = False Then
			gErrorMsgBox "��� û�൥������" & lngRow & " ���� ��� �� �ʼ� �Է»����Դϴ�.","����ȳ�"
			Exit Sub
		End If
		
		'�� ����簡 �ο찡 �ִ°�� �ʼ� �Է»���
		if .sprSht_OUT.MaxRows <> 0 then
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_OUT, "YEARMON | DEPT_CD",lngCol, lngRow, False) 

			If strDataCHK = False Then
				gErrorMsgBox  "�系���� �׸���" & lngRow & " ���� ���/������ �ʼ� �Է»����Դϴ�.","����ȳ�"
				Exit Sub
			End If
			if DataValidation = false then exit sub 	
		end if 
		
		'CGV �׸��� �����Ͱ� �ִ°�� �ʼ� �Է»���
		if .sprSht_CGV.MaxRows <> 0 then
			strDataCHK = mobjSCGLSpr.DataValidation(.sprSht_CGV, "YEARMON | MEDCODE",lngCol, lngRow, False) 
		
			If strDataCHK = False Then
				gErrorMsgBox "CGV �׸���" & lngRow & " ���� ���/������� �ʼ� �Է»����Դϴ�.","����ȳ�"
				Exit Sub
			end if 
			if DataValidation_CGV = false then exit sub 	
		end if 
		
		if mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"GUBUN",.sprSht_HDR.ActiveRow) = "EXCLIENT" or mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow) <> 0 then
			if .sprSht_OUT.MaxRows = 0 then
				gErrorMsgBox "�系�����̰ų� �ܺδ����� ��� 1)�� �׸���(�系/�����)�� �߰��ϼž��մϴ�.","����ȳ�"
				exit sub
			end if 
		end if

		'---------------------------����---------------------------				
		strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",.sprSht_HDR.ActiveRow)
	
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_HDR,"CHK | YEARMON | SEQ | GUBUN | CONT_TYPE | CLIENTCODE | CLIENTNAME | CONT_CODE | CONT_NAME | TBRDSTDATE | TBRDEDDATE | TOTAL_AMT | TAX_AMT | AMT | TIM_RATE | TIM_AMT | EXCLIENTCODE | EXCLIENTNAME | EX_RATE | EX_AMT | CGV_RATE | MC_AMT | CGV_AMT | MEMO | COMMI_TRANS_NO")
		'ó�� ������ü ȣ��
		
		If isArray(vntData) Then
			intRtn = mobjMDOTCLOUD.ProcessRtn_AMT(gstrConfigXml, vntData)
		end if 
		
		'---------------------------
		'���� �ϴ��� �� ����� ����
		'---------------------------
		For intCnt = 1 to .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"YEARMON",intCnt) <> "" Then
				mobjSCGLSpr.CellChanged frmThis.sprSht_OUT, 1, intCnt
			End If
		Next
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_OUT,"CHK | YEARMON | SEQ | CONT_CODE | DEPT_CD | BTN_DEPT | DEPT_NAME | RATE | AMT | MEMO")
				
		'ó�� ������ü ȣ��
		If isArray(vntData) Then
			intRtn = mobjMDOTCLOUD.ProcessRtn_OUT(gstrConfigXml, vntData)
		end if 
		
		'---------------------------
		'���� �ϴ��� CGV ���� ����
		'---------------------------
		For intCnt = 1 to .sprSht_CGV.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"YEARMON",intCnt) <> "" Then
				mobjSCGLSpr.CellChanged frmThis.sprSht_CGV, 1, intCnt
			End If
		Next
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht_CGV,"CHK | YEARMON | SEQ | CONT_CODE | MEDCODE | BTN_MED | MEDNAME | REAL_MED_CODE | REAL_MED_NAME | OUT_AMT | MEMO")
				
		'ó�� ������ü ȣ��
		If isArray(vntData) Then
			intRtn = mobjMDOTCLOUD.ProcessRtn_CGV(gstrConfigXml, vntData)
		end if 
		'---------------------------------------------------------------------
		if not gDoErrorRtn ("ProcessRtn_AMT") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht_HDR,meCLS_FLAG
			gErrorMsgBox "�ڷᰡ ����" & mePROC_DONE,"����ȳ�" 
			SelectRtn
   		end if
   	end with
end sub

'****************************************************************************************
' ������ ó���� ���� ����Ÿ ����
'****************************************************************************************
Function DataValidation ()
	DataValidation = False
   	Dim i
   	Dim intAMT
   	Dim intSUMAMT
   	Dim intTIM_AMT
   	
   	intAMT = 0
	intTIM_AMT = 0
	intSUMAMT = 0 
	'On error resume Next
	With frmThis
		
		IF mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow) <> 0 THEN
			intTIM_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"TIM_AMT",.sprSht_HDR.ActiveRow)
		ELSE
			intTIM_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"EX_AMT",.sprSht_HDR.ActiveRow)
		END IF
		
		for i =1 to .sprSht_OUT.MaxRows
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_OUT,"AMT",i)
			intSUMAMT = intSUMAMT + intAMT
		next	
		if intSUMAMT > intTIM_AMT then
			gErrorMsgBox "�ϴ��� �հ�ݾ��� ����� �ݾ��� �ʰ��Ҽ� �����ϴ�.","����ȳ�"
			exit Function
		end if 
		
		if intSUMAMT < intTIM_AMT then
			gErrorMsgBox "�ϴ��� �հ�ݾ��� ����� �ݾ׺��� ���� �� �����ϴ�..","����ȳ�"
			exit Function
		end if
   		
   	End With
	DataValidation = True
End Function

Function DataValidation_CGV ()
	DataValidation_CGV = False
   	Dim i
   	Dim intAMT
   	Dim intSUMAMT
   	Dim intCGV_AMT
   	
   	intAMT = 0
	intCGV_AMT = 0
	
	'On error resume Next
	With frmThis
		
		intCGV_AMT = mobjSCGLSpr.GetTextBinding( .sprSht_HDR,"CGV_AMT",.sprSht_HDR.ActiveRow)
		
		intSUMAMT = 0	
		for i =1 to .sprSht_CGV.MaxRows
			intAMT = mobjSCGLSpr.GetTextBinding( .sprSht_CGV,"OUT_AMT",i)
			intSUMAMT = intSUMAMT + intAMT
		next	
		
		if intSUMAMT > intCGV_AMT then
			gErrorMsgBox "�ϴ��� CGV ���ұݾ��� ����� CGV �ݾ��� �ʰ��� �� �����ϴ�..","����ȳ�"
			exit Function
		end if 
		
		if intSUMAMT < intCGV_AMT then
			gErrorMsgBox "�ϴ��� CGV ���ұݾ��� ����� CGV �ݾ׺��� ���� �� �����ϴ�.","����ȳ�"
			exit Function
		end if
   		
   	End With
	DataValidation_CGV = True
End Function



'--------------û�೻����� -------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON,strCONT_CODE
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	With frmThis
		
		for i = 1 to .sprSht_HDR.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
				
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",i) <> "" then
					gErrorMsgBox i & "��° �ο��� �����ʹ� �ŷ������� ����� ���� �Դϴ�. ���� �Ͻ� �� �����ϴ�.","�����ȳ�!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		

		
		intRtn = gYesNoMsgbox("û�� ������ �����Ͻø� �ϴ��� �系�����̳� CGV �ݾ��� ��� ���� �˴ϴ�." & vbcrlf & " �ڷḦ �����Ͻðڽ��ϱ�? ","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_HDR.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"YEARMON",i)
				strCONT_CODE = mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"CONT_CODE",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_HDR,i
				else
					intRtn = mobjMDOTCLOUD.DeleteRtn_AMT(gstrConfigXml, strYEARMON, strSEQ, strCONT_CODE)
					
					IF not gDoErrorRtn ("DeleteRtn_AMT") then
						mobjSCGLSpr.DeleteRow .sprSht_HDR,i
   					End IF
				end if				
   				intCnt = intCnt + 1
   			END IF
		next
		
		If not gDoErrorRtn ("DeleteRtn_AMT") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   			
   		End If
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht_OUT
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		SelectRtn

	End With
	err.clear	
End Sub



'--------------�系���� ����-------------
Sub DeleteRtn_OUT ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	strSEQFLAG = False
	With frmThis
	
		for i = 1 to .sprSht_OUT.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" then
					gErrorMsgBox "�ش� �����ʹ� �ŷ������� ����� ���� �Դϴ�. ���� �Ͻ� �� �����ϴ�.","�����ȳ�!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_OUT.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_OUT,"YEARMON",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_OUT,i
				else
					intRtn = mobjMDOTCLOUD.DeleteRtn_OUT(gstrConfigXml, strYEARMON, strSEQ)
					
					IF not gDoErrorRtn ("DeleteRtn_OUT") then
						mobjSCGLSpr.DeleteRow .sprSht_OUT,i
   					End IF
   					
   					strSEQFLAG = TRUE
				end if				
   				intCnt = intCnt + 1
   			END IF
		next
		
		If not gDoErrorRtn ("DeleteRtn_OUT") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht_OUT
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn_OUT frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		End If
	End With
	err.clear	
End Sub

'--------------CGV���� ����-------------
Sub DeleteRtn_CGV ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim strYEARMON
	Dim strSEQFLAG '���������Ϳ��� �÷�
	Dim strSEQ	
	Dim intchkCnt
	
	intchkCnt = 0
	strSEQFLAG = False
	With frmThis
	
		for i = 1 to .sprSht_CGV.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 Then		
				intchkCnt = intchkCnt + 1
	
				if mobjSCGLSpr.GetTextBinding(.sprSht_HDR,"COMMI_TRANS_NO",.sprSht_HDR.ActiveRow) <> "" then
					gErrorMsgBox "�ش� �����ʹ� �ŷ������� ����� ���� �Դϴ�. ���� �Ͻ� �� �����ϴ�.","�����ȳ�!"
					EXIT Sub
				end if
			END IF
		NEXT
		
		If intchkCnt = 0 Then
			gErrorMsgBox "������ �����͸� üũ�� �ּ���.","�����ȳ�!"
			EXIT Sub
		End If
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		If intRtn <> vbYes Then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		for i = .sprSht_CGV.MaxRows to 1 step -1
			if mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"CHK",i) = 1 THEN
				strSEQ = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"SEQ",i)
				strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht_CGV,"YEARMON",i)
				
				if strSEQ = "" then
					mobjSCGLSpr.DeleteRow .sprSht_CGV,i
				else
					intRtn = mobjMDOTCLOUD.DeleteRtn_CGV(gstrConfigXml, strYEARMON, strSEQ)
					
					IF not gDoErrorRtn ("DeleteRtn_CGV") then
						mobjSCGLSpr.DeleteRow .sprSht_CGV,i
   					End IF
   					
   					strSEQFLAG = TRUE
				end if				
   				intCnt = intCnt + 1
   			END IF
		next
		
		If not gDoErrorRtn ("DeleteRtn_CGV") Then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End If
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht_CGV
		'�������� �� �����ͻ����� ��ȸ�� ���¿��, �� ������ ������ ����ȸ
		If strSEQFLAG Then
			SelectRtn_CGV frmThis.sprSht_HDR.ActiveCol,frmThis.sprSht_HDR.ActiveRow
		End If
		CGV_AMT_CAL
	End With
	err.clear	
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" border="0" cellSpacing="0" cellPadding="0" width="100%" height="100%">
				<!--Top TR Start-->
				<TR>
					<TD>
						<!--Top Define Table Start-->
						<TABLE id="tblTitle" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							height="28">
							<TR>
								<TD height="28" width="400" align="left">
									<table border="0" cellSpacing="0" cellPadding="0" width="100%">
										<tr>
											<td align="left">
												<TABLE border="0" cellSpacing="0" cellPadding="0" width="138" background="../../../images/back_p.gIF">
													<TR>
														<TD height="2" width="100%" align="left"></TD>
													</TR>
												</TABLE>
											</td>
										</tr>
										<tr>
											<td class="TITLE">CGV Ŭ���� û�����</td>
										</tr>
									</table>
								</TD>
								<TD style="WIDTH: 640px" height="28" vAlign="middle" align="right">
									<!--Wait Button Start-->
									<TABLE style="Z-INDEX: 200; POSITION: absolute; WIDTH: 65px; HEIGHT: 23px; VISIBILITY: hidden; TOP: 0px; LEFT: 336px"
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
						<TABLE border="0" cellSpacing="0" cellPadding="0" width="1040" background="../../../images/TitleBG.gIF">
							<TR>
								<TD height="1" width="100%" align="left"></TD>
							</TR>
						</TABLE>
						<!--Top Define Table End-->
						<!--Input Define Table End-->
						<TABLE id="tblBody" border="0" cellSpacing="0" cellPadding="0" width="100%" height="93%"> <!--TopSplit Start->
								<!--TopSplit Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 4px" class="TOPSPLIT"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD style="WIDTH: 100%" class="KEYFRAME" vAlign="top" align="left">
									<TABLE id="tblKey" class="SEARCHDATA" border="0" cellSpacing="1" cellPadding="0" width="100%"
										align="left">
										<TR>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtYEARMON, '')"
												width="50">���</TD>
											<TD class="SEARCHDATA" width="100"><INPUT accessKey="NUM" style="WIDTH: 80px; HEIGHT: 22px" id="txtYEARMON" class="INPUT"
													title="���" maxLength="10" size="6" name="txtYEARMON"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtCLIENTNAME1, txtCLIENTCODE1)"
												width="70">������</TD>
											<TD class="SEARCHDATA" width="250"><INPUT style="WIDTH: 174px; HEIGHT: 22px" id="txtCLIENTNAME1" class="INPUT_L" title="�����ָ�"
													maxLength="100" size="23" name="txtCLIENTNAME1"><IMG style="CURSOR: hand" id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" border="0" name="ImgCLIENTCODE1" align="absMiddle" src="../../../images/imgPopup.gIF"><INPUT style="WIDTH: 56px; HEIGHT: 22px" id="txtCLIENTCODE1" class="INPUT" title="�ڵ��Է�"
													maxLength="6" size="4" name="txtCLIENTCODE1"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" onclick="vbscript:Call gCleanField(txtTITLE1, '')"
												width="70">����</TD>
											<TD class="SEARCHDATA" width="220"><INPUT style="WIDTH: 216px; HEIGHT: 22px" id="txtTITLE1" class="INPUT_L" title="����" maxLength="100"
													size="30" name="txtTITLE1"></TD>
											<TD style="CURSOR: hand" class="SEARCHLABEL" width="60">��������</TD>
											<TD class="SEARCHDATA"><SELECT style="WIDTH: 96px" id="cmbCONT_TYPE" title="��������" name="cmbCONT_TYPE">
													<OPTION selected value="">��ü</OPTION>
													<OPTION value="B">�������</OPTION>
													<OPTION value="C">��������</OPTION>
												</SELECT></TD>
											<TD class="SEARCHDATA" width="50">
												<TABLE border="0" cellSpacing="0" cellPadding="2" align="right">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" border="0" name="imgQuery"
																alt="�ڷḦ ��ȸ�մϴ�." src="../../../images/imgQuery.gIF" height="20"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
									</TABLE>
								</TD>
							<tr>
								<td>
									<table class="DATA" cellSpacing="0" cellPadding="0" width="100%" height="10">
										<TR>
										</TR>
									</table>
									<TABLE border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<td style="WIDTH: 1000px" class="DATA">���û�����հ�:<INPUT accessKey="NUM" style="WIDTH: 120px; HEIGHT: 20px" id="txtSUMAMT" class="NOINPUTB_R"
													title="�հ�ݾ�" readOnly maxLength="100" size="13" name="txtSUMAMT">�հ�<INPUT style="WIDTH: 120px; HEIGHT: 20px" id="txtSELECTAMT" class="NOINPUTB_R" title="���ñݾ�"
													readOnly maxLength="100" size="16" name="txtSELECTAMT">
											</td>
											<TD height="20" width="400" align="left"></TD>
											<TD height="20" vAlign="middle" align="right">
												<!--Common Button Start-->
												<TABLE style="HEIGHT: 20px" border="0" cellSpacing="0" cellPadding="2">
													<TR>
														<TD><IMG style="CURSOR: hand" id="imgCho" onmouseover="JavaScript:this.src='../../../images/imgChoOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgCho.gif'" border="0" name="imgCho"
																alt="ȭ���� �ʱ�ȭ �մϴ�." src="../../../images/imgCho.gif"></TD>
														<TD><IMG style="CURSOR: hand" id="ImgSave_AMT" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gIF'"
																onmouseout="JavaScript:this.src='../../../images/ImgSave.gIF'" border="0" name="ImgSave"
																alt="û�� �ڷḦ �����մϴ�.." src="../../../images/ImgSave.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete"
																alt="û�� ������ �����մϴ�..." src="../../../images/imgDelete.gIF" height="20"></TD>
														<TD><IMG style="CURSOR: hand" id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" border="0" name="imgPrint"
																alt="�ڷḦ �μ��մϴ�." src="../../../images/imgPrint.gIF"></TD>
														<TD><IMG style="CURSOR: hand" id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel"
																alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 3px" class="BODYSPLIT"></TD>
							</TR>
							<!--Input End-->
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 30%" class="LISTFRAME" vAlign="top" align="center">
									<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab1"
										ms_positioning="GridLayout">
										<OBJECT style="WIDTH: 100%; HEIGHT: 100%" id="sprSht_HDR" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											DESIGNTIMEDRAGDROP="213">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="31829">
											<PARAM NAME="_ExtentY" VALUE="3598">
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
								<TD style="WIDTH: 100%" id="lblStatus" class="BOTTOMSPLIT"></TD>
							</TR>
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 3px" class="BODYSPLIT"></TD>
							</TR>
							<!--Input End-->
							<TR>
								<TD>
									<TABLE id="tblTitle3" border="0" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										height="28"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD height="28" width="50%" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="65" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="1" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td class="TITLE">1) �系/�����&nbsp;</td>
														<TD height="22" vAlign="middle" align="right">
															<TABLE style="HEIGHT: 20px" id="tblButton_OUT" border="0" cellSpacing="0" cellPadding="2">
																<TR>
																	<TD><IMG style="CURSOR: hand" id="ImgAddRow_OUT" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" border="0" name="imgAddRow_OUT"
																			alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54"></TD>
																	<!--		<TD><IMG id="ImgSave_OUT" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gIF'"
																			style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgSave.gIF'"
																			height="20" alt="�系������ �����մϴ�..." src="../../../images/ImgSave.gIF" border="0"
																			name="ImgSave_OUT"></TD>
															-->
																	<TD><IMG style="CURSOR: hand" id="imgDelete_OUT" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete_OUT"
																			alt="�系�����ڷḦ �����մϴ�..." src="../../../images/imgDelete.gIF" height="20"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgExcel_OUT" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel_OUT"
																			alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
																</TR>
															</TABLE>
														</TD>
													</tr>
												</table>
											</TD>
											<TD height="28" width="50%" align="left">
												<table border="0" cellSpacing="0" cellPadding="0" width="100%">
													<tr>
														<td align="left">
															<TABLE border="0" cellSpacing="0" cellPadding="0" width="40" background="../../../images/back_p.gIF">
																<TR>
																	<TD height="1" width="100%" align="left"></TD>
																</TR>
															</TABLE>
														</td>
													</tr>
													<tr>
														<td class="TITLE">2) CGV&nbsp;</td>
														<TD height="22" vAlign="middle" align="right">
															<TABLE style="HEIGHT: 20px" id="tblButton_CGV" border="0" cellSpacing="0" cellPadding="2">
																<TR>
																	<TD><IMG style="CURSOR: hand" id="ImgAddRow_CGV" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" border="0" name="imgAddRow_CGV"
																			alt="�� �� �߰�" src="../../../images/imgAddRow.gif" width="54"></TD>
																	<!--	<TD><IMG id="ImgSave_CGV" onmouseover="JavaScript:this.src='../../../images/ImgSaveOn.gIF'"
																			style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/ImgSave.gIF'"
																			height="20" alt="CGV �ڷḦ �����մϴ�.." src="../../../images/ImgSave.gIF" border="0"
																			name="ImgSave_CGV"></TD>
																-->
																	<TD><IMG style="CURSOR: hand" id="imgDelete_CGV" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'" border="0" name="imgDelete_CGV"
																			alt="CGV �ڷḦ �����մϴ�." src="../../../images/imgDelete.gIF" height="20"></TD>
																	<TD><IMG style="CURSOR: hand" id="imgExcel_CGV" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																			onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'" border="0" name="imgExcel_CGV"
																			alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" height="20"></TD>
																</TR>
															</TABLE>
														</TD>
													</tr>
												</table>
											</TD>
										</TR>
										<TR>
											<TD style="WIDTH: 218px; HEIGHT: 4px" colSpan="2"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<!--List Start-->
							<TR>
								<TD style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="left">
									<TABLE border="0" cellSpacing="1" cellPadding="0" width="100%" align="left" height="98%">
										<TR>
											<td style="WIDTH: 50%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab2"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_OUT" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="100%"
														height="100%">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="15875">
														<PARAM NAME="_ExtentY" VALUE="7699">
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
											</td>
											<td style="WIDTH: 50%; HEIGHT: 100%" vAlign="top" align="left">
												<DIV style="POSITION: relative; WIDTH: 100%; HEIGHT: 100%; VISIBILITY: hidden" id="pnlTab3"
													ms_positioning="GridLayout">
													<OBJECT id="sprSht_CGV" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" width="100%"
														height="100%">
														<PARAM NAME="_Version" VALUE="393216">
														<PARAM NAME="_ExtentX" VALUE="15875">
														<PARAM NAME="_ExtentY" VALUE="7699">
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
											</td>
										</TR>
										<TR>
											<TD id="lblStatus_OUT" class="BOTTOMSPLIT"></TD>
											<TD id="lblStatus_CGV" class="BOTTOMSPLIT"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			</TR></TABLE></FORM>
	</body>
</HTML>
