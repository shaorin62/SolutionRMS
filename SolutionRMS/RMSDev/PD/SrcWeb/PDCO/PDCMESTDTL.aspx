<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMESTDTL.aspx.vb" Inherits="PD.PDCMESTDTL" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>���� ��������</title>
		<meta content="False" name="vs_showGrid">
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : ������������ ȭ��(PDCMESTDTL)
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPREESTDTL.aspx
'��      �� : ������ ���� ��� �� Ȯ��
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/11/16 By Tae Ho Kim
'			 2) 
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
Dim mobjPDCMPREESTDTL '�����ڵ�, Ŭ����
Dim mstrPROCESS
Dim mstrPROCESS2 '��ȸ�����̸� true �űԻ����̸� false
Dim mstrCheck
Dim mobjMDLOGIN
Dim mobjMDCMEMP
Dim mobjPDCMGET
CONST meTAB = 9
mstrPROCESS = TRUE
mstrPROCESS2 = TRUE
mstrCheck = True

'=============================
' �̺�Ʈ ���ν��� 
'=============================
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
	DataClean
	call sprSht_Keydown(meINS_ROW, 0)
End Sub

'Sub imgDelete_onclick
'	gFlowWait meWAIT_ON
'	DeleteRtn
'	gFlowWait meWAIT_OFF
'End Sub

Sub imgSave_onclick ()
	with frmThis
		
		if frmThis.txtENDFLAG.value = "T" Then
		
			gErrorMsgBox "�ŷ������� �ۼ��Ǿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		End If
		'if frmThis.txtENDFLAGEXE.value = "T" Then
		'	gErrorMsgBox "���ֺ� ���⳻���� �ۼ��Ǿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
		'	Exit Sub
		'End If
			gFlowWait meWAIT_ON
			if .txtPREESTNO.value = "" Then
				ProcessRtn
			Else
				ProcessRtn_OLD
			End If
			gFlowWait meWAIT_OFF
		
		
	End with
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

Sub imgRowDel_onclick
	
	if frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "�ŷ������� �ۼ��Ǿ� ������� �Ұ��� �մϴ�.","�����ȳ�!"
		Exit Sub
	End If
	'if frmThis.txtENDFLAGEXE.value = "T" Then
	'	gErrorMsgBox "���ֺ� ���⳻���� �ۼ��Ǿ� ������� �Ұ��� �մϴ�.","����ȳ�!"
	'	Exit Sub
	'End If
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub

Sub imgRowAdd_onclick ()
	if frmThis.txtENDFLAG.value = "T" Then
		gErrorMsgBox "�ŷ������� �ۼ��Ǿ� ���߰��� �Ұ��� �մϴ�.","����ȳ�!"
		Exit Sub
	End If
	'if frmThis.txtENDFLAGEXE.value = "T" Then
	'	gErrorMsgBox "���ֺ� ���⳻���� �ۼ��Ǿ�  ���߰��� �Ұ��� �մϴ�.","����ȳ�!"
	'	Exit Sub
	'End If
	call sprSht_Keydown(meINS_ROW, 0)
	
End Sub

Sub ImgExeList_onclick
Dim strJOBNO	
Dim vntInParams
Dim vntRet
	with frmThis
		strJOBNO = Trim(.txtJOBNO.value)
		vntInParams = array(strJOBNO)
		vntRet = gShowModalWindow("PDCMEXELISTPOP.aspx",vntInParams , 1060,780)
		SelectRtn
	End with
End Sub
'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'�Է¿�
Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtPRINTDAY,frmThis.imgCalEndar,"txtPRINTDAY_onchange()"
		gSetChange
	end with
End Sub

Sub txtPRINTDAY_onchange
	gSetChange
End Sub
Sub imgCalEndarAGREE_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtAGREEYEARMON,frmThis.imgCalEndarAGREE,"txtAGREEYEARMON_onchange()"
		gSetChange
	end with
End Sub
Sub txtAGREEYEARMON_onchange
	gSetChange
End Sub
'-----------------------------------
' SpreadSheet �̺�Ʈ
'-----------------------------------
'���������� ��� �׸��� ���� ���� �ɶ� �߻� �ϴ� �̺�Ʈ �Դϴ�.
Sub sprSht_Change(ByVal Col, ByVal Row)
	Dim vntData
	Dim i, strCols
	Dim strCode, strCodeName
	Dim strQTY, strPRICE, strAMT
	Dim lngPrice
	Dim lngVALUE
	Dim lngVALUE1
	Dim lngVALUE2

	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		IF Col = 7 Then
			strCode = ""
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",.sprSht.ActiveRow)
			vntData = mobjPDCMGET.GetITEMCODE(gstrConfigXml,mlngRowCnt,mlngColCnt,"0","",strCodeName)
			If mlngRowCnt = 1 Then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntData(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntData(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntData(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntData(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntData(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntData(4,0)			
				mobjSCGLSpr.CellChanged .sprSht, .sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
			Else
				mobjSCGLSpr_ClickProc .sprSht, Col, .sprSht.ActiveRow
			End If
			.txtSUSURATE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+1, Row
		'��������	
		ElseIf  Col = 11 Then
   			strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   			strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   			strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   			If strPRICE <> "" And strAMT = "" Then
   				lngVALUE = strQTY * strAMT
   				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE
   			ElseIf strPRICE = "" And strAMT <> "" Then
   				lngVALUE1 = gRound(strAMT/strQTY,0)
   				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngVALUE1
   			ElseIf strPRICE <> "" And strAMT <> "" Then
   				lngVALUE2 = strQTY * strPRICE
   				mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row, lngVALUE2
   			End IF
   			Call SUSUAMT_CHANGEVALUE(Row)
   			BUDGET_AMT_SUM
   		'�ܰ� ����
   		ElseIf Col = 12 Then
   			strQTY		= mobjSCGLSpr.GetTextBinding( .sprSht,"QTY",.sprSht.ActiveRow)
			strPRICE   = mobjSCGLSpr.GetTextBinding( .sprSht,"PRICE",.sprSht.ActiveRow)
			strAMT = strQTY * strPRICE
			mobjSCGLSpr.SetTextBinding .sprSht,"AMT",.sprSht.ActiveRow, strAMT	
			Call SUSUAMT_CHANGEVALUE(Row)
			BUDGET_AMT_SUM
		'�ݾ׷���	
   		ElseIf  Col = 13 Then
   			strQTY = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row)
   			strPRICE = mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row)
   			strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
   			If strAMT = 0 Then
   				mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row, strAMT
   				mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, strAMT
   			Else 
   				If strQTY <> 0  Then
   					lngPrice = gRound(strAMT/strQTY,0)
   					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row, lngPrice
   				End IF
   			End IF
   			Call SUSUAMT_CHANGEVALUE(Row)
   			BUDGET_AMT_SUM
   		Elseif Col = 10 Then
   			Call SUSUAMT_CHANGEVALUE2
   			BUDGET_AMT_SUM
		END IF
	end with
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub
Sub SUSUAMT_CHANGEVALUE(ByVal Row)
Dim strAMT,strCOMMIFLAG
Dim strSUSURATE
	with frmThis
		strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row)
		strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",Row)
		strSUSURATE = .txtSUSURATE.value
		if strCOMMIFLAG = "1" Then
			if strSUSURATE = "" then
				strSUSURATE = 0
			end if
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, gRound((strAMT * strSUSURATE /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",.sprSht.ActiveRow, 0	
		End if
	End with
End SUb
Sub SUSUAMT_CHANGEVALUE2
Dim intCnt
Dim strAMT,strCOMMIFLAG
Dim strSUSURATE
	with frmThis
	
	For intCnt = 1 to .sprSht.MaxRows
		strAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
		strCOMMIFLAG  = mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG",intCnt)
		strSUSURATE = .txtSUSURATE.value
		if strCOMMIFLAG = "1" Then
			if strSUSURATE = "" then
				strSUSURATE = 0
			end if
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, gRound((strAMT * strSUSURATE /100),0)
		Else
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSUAMT",intCnt, 0	
		End if
	Next
	
	End with
End Sub
Sub txtSUSURATE_onchange
	with frmThis
		SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		gSetChangeFlag .txtSUSURATE  
	End with
End Sub

Sub BUDGET_AMT_SUM
	'���հ� ����
	Dim lngCnt, IntAMT, IntAMTSUM, IntPRICE, IntPRICESUM
	Dim lngSUSU
	'������ ��� ����
	Dim intCnt,intSUSU,intSUSUSUM 
	'commition ��� ����
	Dim intCnt1,intCOM,intCOMSUM 
	'noncommition ��꺯��
	Dim intCnt2,intNON,intNONSUM 
	
	with frmThis
	
		IntAMTSUM = 0
		IntPRICESUM = 0
		intSUSU = 0
		intSUSUSUM = 0
		intCOM = 0
		intCOMSUM = 0
		intNON = 0
		intNONSUM = 0
		For intCnt = 1 To .sprSht.MaxRows
		
			intSUSU = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSUAMT", intCnt)
			intSUSUSUM = intSUSUSUM + intSUSU
			
		Next
		.txtSUSUAMT.value = intSUSUSUM
		
		
		For lngCnt = 1 To .sprSht.MaxRows
			IntAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
			IntAMTSUM = IntAMTSUM + IntAMT
		Next
		IntAMTSUM = IntAMTSUM + intSUSUSUM
		.txtSUMAMT.value = IntAMTSUM
		
		For intCnt1 = 1 To .sprSht.MaxRows
			if mobjSCGLSpr.GetTextBinding(.sprSht,"COMMIFLAG", intCnt1) = "1" Then
				
				intCOM = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intCOMSUM = intCOMSUM + intCOM
			Else
				
				intNON = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", intCnt1)
				intNONSUM = intNONSUM + intNON
			end if
		Next
		.txtCOMMITION.value = intCOMSUM
		.txtNONCOMMITION.value = intNONSUM
		
		txtSUSUAMT_onblur
		txtCOMMITION_onblur
		txtSUMAMT_onblur
		txtNONCOMMITION_onblur
		
	End With
End Sub
'���������� ���� ���� Ŭ�� �� �߻�
sub sprSht_DblClick (ByVal Col, ByVal Row)
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

'���������� �׸��� ���ҽ� ��� �Լ��� �¿���� �Ҷ� ���
Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
	dim vntRet, vntInParams
	With frmThis
		IF Col = 7 Then
			vntInParams = array(mobjSCGLSpr.GetTextBinding( sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			'ITEMCODE,DIVNAME,CLASSNAME,ITEMNAME
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)				
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
			End IF
			
			.txtSUSURATE.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus	
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End With
End Sub
'�������� �� ��ư�� Ŭ�� �Ͽ����� �߻� �ϴ� �̺�Ʈ
Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	with frmThis
	
		IF Col = 6 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht,"ITEMCODENAME",Row))
			vntRet = gShowModalWindow("PDCMITEMPOP.aspx",vntInParams , 413,435)
			
			IF isArray(vntRet) then
			
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"DIVNAME",Row, vntRet(1,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLASSNAME",Row, vntRet(2,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"ITEMCODENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"FAKENAME",Row, vntRet(3,0)	
				mobjSCGLSpr.SetTextBinding .sprSht,"COMMIFLAG",Row, vntRet(4,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				SUSUAMT_CHANGEVALUE2
				BUDGET_AMT_SUM
				'GetRealMedCode mobjSCGLSpr.GetTextBinding( .sprSht,"MEDCODE",Row), mobjSCGLSpr.GetTextBinding( .sprSht,"MEDNAME",Row)
			End IF
			.txtSUSURATE.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht.Focus
			mobjSCGLSpr.ActiveCell .sprSht, Col+2, Row
		end if
	End with
End Sub

'=============================
' UI���� ���ν��� 
'=============================
Sub txtSUSUAMT_onfocus
	with frmThis
		.txtSUSUAMT.value = Replace(.txtSUSUAMT.value,",","")
	end with
End Sub
Sub txtSUSUAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUSUAMT,0,true)
	end with
End Sub
Sub txtCOMMITION_onfocus
	with frmThis
		.txtCOMMITION.value = Replace(.txtCOMMITION.value,",","")
	end with
End Sub
Sub txtCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtCOMMITION,0,true)
	end with
End Sub
Sub txtSUMAMT_onfocus
	with frmThis
		.txtSUMAMT.value = Replace(.txtSUMAMT.value,",","")
	end with
End Sub
Sub txtSUMAMT_onblur
	with frmThis
		call gFormatNumber(.txtSUMAMT,0,true)
	end with
End Sub
Sub txtNONCOMMITION_onfocus
	with frmThis
		.txtNONCOMMITION.value = Replace(.txtNONCOMMITION.value,",","")
	end with
End Sub
Sub txtNONCOMMITION_onblur
	with frmThis
		call gFormatNumber(.txtNONCOMMITION,0,true)
	end with
End Sub
'-----------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------	
Sub InitPage()
	'����������ü ����	
	dim vntInParam
	dim intNo,i
	
	set mobjPDCMPREESTDTL	= gCreateRemoteObject("cPDCO.ccPDCOPREESTDLT")
	set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
	
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "260px"
	pnlTab1.style.left= "8px"
	
	mobjSCGLCtl.DoEventQueue
	
	vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : frmThis.txtPREESTNO.value = vntInParam(i)	
				case 1 : frmThis.txtJOBNO.value = vntInParam(i)
			end select
		next
    'Sheet �⺻Color ����
	gSetSheetDefaultColor()
	With frmThis

		
		
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 0
		mobjSCGLSpr.AddCellSpan  .sprSht, 5, SPREAD_HEADER, 2, 1
		mobjSCGLSpr.SpreadDataField .sprSht, "PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|BTN|ITEMCODENAME|FAKENAME|STD|COMMIFLAG|QTY|PRICE|AMT|SUSUAMT|GBN"
		mobjSCGLSpr.SetHeader .sprSht,		  "��������ȣ|����|��з�|�ߺз�|�����׸��ڵ�|�����׸��|������|����|Ŀ�̼�|����|�ܰ�|�ݾ�|������ݾ�|���屸��"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|   0|     8|    12|        8 |2|        15|12    |  20|     6|  12|  13|13  |10         |0"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
		mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "COMMIFLAG"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "QTY|PRICE|AMT|SUSUAMT", -1, -1, 0
		mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "ITEMCODENAME|STD|FAKENAME", -1, -1, 255
		'mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "DIVNAME|CLASSNAME|ITEMCODE"
		mobjSCGLSpr.ColHidden .sprSht, "PREESTNO|ITEMCODESEQ|ITEMCODESEQ|GBN", true 'SUSUAMT
		mobjSCGLSpr.SetCellAlign2 .sprSht, "DIVNAME|CLASSNAME|FAKENAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht, "ITEMCODE|ITEMCODESEQ",-1,-1,2,2,false
	
	    .sprSht.style.visibility  = "visible"
		.sprSht.MaxRows = 0

	InitPageData	
	SelectRtn
	If .txtENDFLAG.value = "T" Then
	Else
		.txtPREESTNAME.value = .txtJOBNAME.value 
	End If
	
	End With
End Sub



Sub EndPage()
	set mobjPDCMPREESTDTL = Nothing
	set mobjPDCMGET = Nothing
	gEndPage
End Sub
'-----------------------------
' Ȯ�� �� Ȯ����� ó��
'-----------------------------	
Sub imgSetting_onclick
Dim intRtnConfirm
Dim intRtn
	intRtnConfirm = gYesNoMsgbox("�ڷḦ Ȯ�� �Ͻðڽ��ϱ�?","�ڷ�Ȯ�� Ȯ��")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_Confirm(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_Confirm") then
				gErrorMsgBox " �ڷᰡ Ȯ�� �Ǿ����ϴ�.","Ȯ���ȳ�" 
			End If
			ESTCONFIRM_Search
	End with
End Sub


Sub ImgConfirmCancel_onclick
Dim intRtnConfirm
Dim intRtn
	intRtnConfirm = gYesNoMsgbox("�ڷḦ Ȯ����� �Ͻðڽ��ϱ�?","�ڷ�Ȯ����� Ȯ��")
	IF intRtnConfirm <> vbYes then exit Sub
	with frmThis
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_ConfirmCancel(gstrConfigXml,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
			if not gDoErrorRtn ("ProcessRtn_ConfirmCancel") then
				gErrorMsgBox " �ڷᰡ Ȯ����� �Ǿ����ϴ�.","Ȯ����Ҿȳ�" 
			End If
			ESTCONFIRM_Search
	End with
End Sub
Sub txtPREESTNAME_onchange
	gSetChange
End Sub

'-----------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------	
Sub InitPageData
	'��� ������ Ŭ����
	'gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		
		.txtPRINTDAY.value = gNowDate
		.sprSht.MaxRows = 0
		ESTCONFIRM_Search2
	End with
	'DataNewClean
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"
End Sub

'û��Ȯ�� ��ȸ
Sub ESTCONFIRM_Search
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
				.imgSetting.disabled = true
				.ImgConfirmCancel.disabled = false
				
			Else
				.imgSetting.disabled = false
				.ImgConfirmCancel.disabled = true
				
			End if
   		end if
	end with
End Sub
Sub ESTCONFIRM_Search2
	Dim intRtn
	Dim vntData
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		intRtn = mobjPDCMPREESTDTL.SelectRtn_Confirm(gstrConfigXml,mlngRowCnt,mlngColCnt,Trim(.txtPREESTNO.value),Trim(.txtJOBNO.value))
		If not gDoErrorRtn ("SelectRtn_Confirm") then
			If mlngRowCnt > 0 Then
			
				.imgExeList.style.visibility = "visible"
			Else
			
				.imgExeList.style.visibility = "hidden"
			End if
   		end if
	end with
End Sub
Sub DataNewClean
	with frmThis
	.txtCREDAY.value = ""
	.cmbGROUPGBN.selectedIndex  = -1
	End with
End Sub
Sub DateClean
	Dim date1
	Dim date2
	Dim strDATE
	strDATE = gNowDate
	date1 = Mid(strDATE,1,7)  & "-01"
	date2 = DateAdd("d", -1, DateAdd("m", 1, date1))

	with frmThis
		.txtFROM.value = date1
		.txtTO.value = date2
	End With
End Sub
'------------------------------------------
' ������ ��ȣ�� �ִ� ��� ����ó��
'------------------------------------------
Sub ProcessRtn_OLD ()
    Dim intRtn
  	Dim vntData
	Dim strMasterData
	Dim strJOBYEARMON 
	Dim strJOBCUST
	Dim strJOBSEQ
	Dim strCODE
	Dim strSEQFlag
	Dim strGROUPGBN
	Dim strDELCODE
	Dim intRtnSave
	Dim vntData2
	Dim strPREESTNO
	Dim intHDR
	with frmThis
	'On error resume next
  		'������ Validation
		if DataValidation =false then exit sub
		If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "�������� �Է��Ͻʽÿ�.","����ȳ�"
			Exit Sub
		End If
		strPREESTNO = .txtPREESTNO.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|FAKENAME|SUSUAMT")
		'ó�� ������ü ȣ��
		strMasterData = gXMLGetBindingData (xmlBind)
		
		
		if  not IsArray(vntData) then 
				If gXMLIsDataChanged (xmlBind) Then 
					intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
					if not gDoErrorRtn ("ProcessRtn_HDR") then
						mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
						gErrorMsgBox " �ڷᰡ" & intHDR & " �� ����" & mePROC_DONE,"����ȳ�" 
						SelectRtn
					End If
				Else
					gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
				End If
			Exit Sub
		End If
		
		If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
		End IF

			intRtn = mobjPDCMPREESTDTL.ProcessRtn(gstrConfigXml,vntData,strPREESTNO)
				
		if not gDoErrorRtn ("ProcessRtn") then
			intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
			if not gDoErrorRtn ("ProcessRtn_HDR") then
				mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
				gErrorMsgBox " �ڷᰡ" & intHDR & " �� ����" & mePROC_DONE,"����ȳ�" 
				SelectRtn
			End If
  		end if
 	end with
End Sub
'------------------------------------------
' ��� Ȯ�� ���� ����� (��������ȣ�� ���� ���)
'------------------------------------------
Sub ProcessRtn
Dim intRtn
Dim strMasterData
Dim strPREESTNO
Dim intCnt
Dim intRtnDtl
Dim vntData
Dim strAGREEYEARMON
Dim strJOBNO
Dim intSearchRtn

strMasterData = gXMLGetBindingData (xmlBind)
if DataValidation =false then exit sub
strPREESTNO = ""
	with frmThis
	If .txtPREESTNAME.value = "" Then
			gErrorMsgBox "�������� �Է��Ͻʽÿ�.","����ȳ�"
			Exit Sub
	End If
	If .sprSht.MaxRows = 0 Then
			gErrorMsgBox "������ ������ ���� ���� �ʽ��ϴ�.","����ȳ�"
			Exit Sub
	End IF
	vntData = mobjSCGLSpr.GetDataRows(.sprSht,"PREESTNO|ITEMCODESEQ|DIVNAME|CLASSNAME|ITEMCODE|ITEMCODENAME|STD|COMMIFLAG|QTY|PRICE|AMT|FAKENAME|SUSUAMT")
	if  not IsArray(vntData) then 
		gErrorMsgBox "����� " & meNO_DATA,"����ȳ�"
		exit sub
	End If
	If .txtAGREEYEARMON.value = "" then
		gErrorMsgBox "����Ȯ������ �����Ͻʽÿ�.","����ȳ�"
		exit sub
	End If
	strAGREEYEARMON = MID(.txtAGREEYEARMON.value,1,4) & MID(.txtAGREEYEARMON.value,6,2) & MID(.txtAGREEYEARMON.value,9,2)
	intRtn = mobjPDCMPREESTDTL.ProcessRtn_HDRLESS(gstrConfigXml,strMasterData,vntData,strPREESTNO,strAGREEYEARMON)
		if not gDoErrorRtn ("ProcessRtn_HDRLESS") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gErrorMsgBox " �ڷᰡ" & intRtn & " �� ����" & mePROC_DONE,"����ȳ�" 
			strJOBNO = .txtJOBNO.value
			intSearchRtn =  mobjPDCMPREESTDTL.SelectRtn_PREESTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBNO)
			
			.txtPREESTNO.value = intSearchRtn(0,1)
			SelectRtn
		End If
	End with

End Sub


Sub DelProc
Dim intHDR
Dim strMasterData
Dim strPREESTNO
	strMasterData = gXMLGetBindingData (xmlBind)
	with frmThis
		strPREESTNO = .txtPREESTNO.value
		intHDR = mobjPDCMPREESTDTL.ProcessRtn_PREESTHDR(gstrConfigXml,strMasterData,strPREESTNO)
				if not gDoErrorRtn ("ProcessRtn_PREESTHDR") then
					'SelectRtn
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
	'On error resume next
	with frmThis
  	
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻� TBRDSTDATE|TBRDEDDATE
   		IF not gDataValidation(frmThis) then exit Function
   		for intCnt = 1 to .sprSht.MaxRows
   		'DIVNAME|CLASSNAME|ITEMCODE,ITEMCODENAME
			if mobjSCGLSpr.GetTextBinding(.sprSht,"DIVNAME",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"CLASSNAME",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODE",intCnt) = "" Or mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODENAME",intCnt) = "" Then 
				gErrorMsgBox intCnt & " ��° ���� �����׸� ���� �� Ȯ���Ͻʽÿ�","�Է¿���"
				Exit Function
			End if
		next
   	
   	End with
	DataValidation = true
End Function

'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim strCODE
	Dim strJOBCODE
	With frmThis
		strCODE = .txtPREESTNO.value
		strJOBCODE = .txtJOBNO.value
		IF strCODE = ""  THEN 
			
			
			IF not SelectRtn_HeadLess (strJOBCODE) Then Exit Sub
			
		Else
			IF not SelectRtn_Head (strCODE) Then Exit Sub
			'��Ʈ ��ȸ
			CALL SelectRtn_Detail (strCODE)
			txtSUSUAMT_onblur
			txtCOMMITION_onblur
			txtSUMAMT_onblur
			txtNONCOMMITION_onblur
		End If
		
		If .txtENDFLAG.value = "T" Then
			.txtPREESTNAME.className = "NOINPUT_L"
			.txtPREESTNAME.readOnly = true
		Else
			.txtPREESTNAME.className = "INPUT_L"
			.txtPREESTNAME.readOnly = false
		End If
		'If .txtAGREEYEARMON.value = "" Then
		'	.imgExeList.style.visibility = "hidden"
		'Else
		'	.imgExeList.style.visibility = "visible"
		'End If
	End With
	
	
End Sub
'���������� ���� ��� ��ȸ
Function SelectRtn_HeadLess (ByVal strJOBCODE)
	Dim vntData
	'on error resume next

	'�ʱ�ȭ
	SelectRtn_HeadLess = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDRLESS(gstrConfigXml,mlngRowCnt,mlngColCnt,strJOBCODE)
	
	IF not gDoErrorRtn ("SelectRtn_HeadLess") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ ������ ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_HeadLess = True
		End IF
	End IF
End Function
'���������� ���� ��� ��ȸ
Function SelectRtn_Head (ByVal strCODE)
	Dim vntData
	'on error resume next

	'�ʱ�ȭ
	SelectRtn_Head = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)
	
	vntData = mobjPDCMPREESTDTL.SelectRtn_HDR(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)
	
	IF not gDoErrorRtn ("SelectRtn_Head") then
		IF mlngRowCnt<=0 then
			gErrorMsgBox "������ �������� ���Ͽ�" & meNO_DATA, ""
			exit Function
		else
			'��ȸ�� �����͸� ���ε�
			call gXMLDataBinding (frmThis,xmlBind,"#xmlBind",vntData)
			SelectRtn_Head = True
		End IF
	End IF
End Function


'���� ���̺� ��ȸ
Function SelectRtn_Detail (ByVal strCODE)
	dim vntData
	Dim intCnt
	Dim strRows
	
	'on error resume next
	'�ʱ�ȭ
	SelectRtn_Detail = false
	mlngRowCnt=clng(0): mlngColCnt=clng(0)

	vntData = mobjPDCMPREESTDTL.SelectRtn_DTL(gstrConfigXml,mlngRowCnt,mlngColCnt,strCODE)

	IF not gDoErrorRtn ("SelectRtn_Detail") then
		'��ȸ�� �����͸� ���ε�
		call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,true)
		'�ʱ� ���·� ����
		mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG

		SelectRtn_Detail = True
		with frmThis
			IF mlngRowCnt > 0 THEN
				For intCnt = 1 To .sprSht.MaxRows 
					if mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",intCnt) = "T" Then
						mobjSCGLSpr.SetCellsLock2 .sprSht,true,intCnt,6,7,true
					Else
						mobjSCGLSpr.SetCellsLock2 .sprSht,false,intCnt,6,7,true
					End If
				Next
				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
			ELSE
				.sprSht.MaxRows = 0
			END IF
		End with
	End IF
End Function



'****************************************************************************************
'���� �˻�� ��� ���´�.
'****************************************************************************************
Sub PreSearchFiledValue (strTBRDSTDATE,strTBRDEDDATE, strCAMPAIGN_CODE, strCAMPAIGN_NAME, strCLIENTCODE, strCLIENTNAME)
	With frmThis
		.txtTBRDSTDATE1.value = strTBRDSTDATE
		.txtTBRDEDDATE1.value = strTBRDEDDATE
		.txtCAMPAIGN_CODE1.value = strCAMPAIGN_CODE
		.txtCAMPAIGN_NAME1.value = strCAMPAIGN_NAME
		.txtCLIENTCODE1.value = strCLIENTCODE
		.txtCLIENTNAME1.value = strCLIENTNAME
	End With
End Sub







Sub DataClean
	with frmThis
		.txtPROJECTNM.value = ""
		.txtPROJECTNO.value = ""
		.txtCLIENTCODE.value = ""
		.txtCLIENTNAME.value = ""
		.txtCLIENTSUBCODE.value = ""
		.txtCLIENTSUBNAME.value = ""
		.txtSUBSEQ.value = ""
		.txtSUBSEQNAME.value = ""
		.txtCPDEPTCD.value = ""
		.txtCPDEPTNAME.value = ""
		.txtCPEMPNO.value = ""
		.txtCPEMPNAME.value = ""
		.txtMEMO.value = ""
		.cmbGROUPGBN.value = "1"
		.txtCREDAY.value = gNowDate
		.sprSht.MaxRows = 0
	End With
End Sub
Sub sprSht_Keydown(KeyCode, Shift)
	if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR and KeyCode <> meTab then exit sub
	Dim intRtn
	if KeyCode = meCR  Or KeyCode = meTab Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 14 AND frmThis.txtENDFLAG.value <> "T" Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(13), cint(Shift), -1, 1)
		DefaultValue
		end if
	Else
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		if intRtn = meINS_ROW then
			'DefaultValue
			mobjSCGLSpr.SetTextBinding frmThis.sprSht,"PREESTNO",frmThis.sprSht.ActiveRow, frmThis.txtPREESTNO.value 
		elseif intRtn = meDEL_ROW then
			DeleteRtn
		end if
'		Select Case intRtn
'			Case meINS_ROW: DefaultValue
'			Case meDEL_ROW: DeleteRtn
'		End Select
	End if
End Sub
Sub DefaultValue
	with frmThis
	mobjSCGLSpr.SetTextBinding .sprSht,"PREESTNO",.sprSht.ActiveRow, .txtPREESTNO.value 
	End With
End Sub
'ProjectNO ��ȸ�˾�
Sub ImgPROJECTNO1_onclick
	Call PONO_POP()
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
			.txtCLIENTNAME1.focus()					' ��Ŀ�� �̵�
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
			
			vntData = mobjPDCMGET.GetPONO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtPROJECTNO1.value),trim(.txtPROJECTNM1.value))
			
			if not gDoErrorRtn ("txtPROJECTNM1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtPROJECTNO1.value = trim(vntData(0,0))
					.txtPROJECTNM1.value = trim(vntData(1,0))
				Else
					Call PONO_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub
'�ڷ����
Sub DeleteRtn ()
	Dim vntData
	Dim intSelCnt, intRtn, i,intRtn2,lngCnt
	dim strYEARMON
	Dim strSEQ
	Dim strPREESTNO
	Dim strITEMCODESEQ
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
		'PREESTNO,ITEMCODESEQ
		'���õ� �ڷḦ ������ ���� ����
		lngCnt =0
		intRtn2 = 0
		for i = intSelCnt-1 to 0 step -1
			If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)) <> ""  Then
				If mobjSCGLSpr.GetTextBinding(.sprSht,"GBN",vntData(i)) = "T"  Then
					gErrorMsgBox "������ �ڷ��� �������� ó���� �Ǿ��־� ������ �Ұ��� �մϴ�.","�����ȳ�"
				Exit Sub
				End iF
				strPREESTNO = mobjSCGLSpr.GetTextBinding(.sprSht,"PREESTNO",vntData(i))
				strITEMCODESEQ = cdbl(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMCODESEQ",vntData(i)))
				intRtn2 = mobjPDCMPREESTDTL.DeleteRtn(gstrConfigXml,strPREESTNO, strITEMCODESEQ)
			End IF
			IF not gDoErrorRtn ("DeleteRtn") then
				lngCnt = lngCnt +1
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				gWriteText "", "[" & strITEMCODESEQ & "] �ڷᰡ �����Ǿ����ϴ�."
   			End IF
		next
		'�������
		Call SUSUAMT_CHANGEVALUE2
		BUDGET_AMT_SUM
		'����Ǿ��ִ� ���� ������ DB �� ������� ���� ���� 
		If intRtn2 = 0 Then
   		Else
			DelProc
		End If
		'1���̶� �������� �ִٸ� �޼��� ���
		If lngCnt <> 0 Then
			gOkMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
		End If
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
	End with
	err.clear
End Sub

-->
		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<TABLE id="tblForm" style="WIDTH: 100%" height="100%" cellSpacing="0" cellPadding="0">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
							border="0">
							<TR>
								<td align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE">&nbsp;û�� ��������</td>
										</tr>
									</table>
								</td>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 280px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
								</TD>
								<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
									<!--Common Button Start-->
									<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
										<TR>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" style="WIDTH: 100%; HEIGHT: 100%" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<!--TopSplit End-->
							<!--Input Start-->
							<TR>
								<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">JOB��</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="JOBNAME" class="NOINPUT_L" id="txtJOBNAME" title="���۰Ǹ�" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtJOBNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">��ü�ι�</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="JOBGUBN" class="NOINPUT_L" id="txtJOBGUBN" title="��ü�ι�" style="WIDTH: 255px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtJOBGUBN"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">��ü�з�</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="CREPART" class="NOINPUT_L" id="txtCREPART" title="��ü�з�" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtCREPART"></TD>
										</TR>
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">������</TD>
											<TD class="SEARCHDATA" width="230"><INPUT dataFld="CLIENTNAME" class="NOINPUT_L" id="txtCLIENTNAME" title="������" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtCLIENTNAME"></TD>
											<TD class="SEARCHLABEL" style="WIDTH: 92px; CURSOR: hand" width="92">�����</TD>
											<TD class="SEARCHDATA" width="260"><INPUT dataFld="CLIENTSUBNAME" class="NOINPUT_L" id="txtCLIENTSUBNAME" title="�����" style="WIDTH: 256px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCLIENTSUBNAME"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" width="80">�귣��</TD>
											<TD class="SEARCHDATA"><INPUT dataFld="SUBSEQNAME" class="NOINPUT_L" id="txtSUBSEQNAME" title="�귣��" style="WIDTH: 272px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUBSEQNAME"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 72px" vAlign="top" align="center">
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;&nbsp;û�� ���� �ۼ�</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgExeList" onmouseover="JavaScript:this.src='../../../images/imgExeListOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExeList.gIF'"
																height="20" alt="���ֺ����� ���α׷��� ȣ���մϴ�." src="../../../images/imgExeList.gIF" border="0"
																name="imgExeList"></TD>
														<TD><IMG id="imgRowAdd" onmouseover="JavaScript:this.src='../../../images/imgRowAddOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowAdd.gIF'"
																height="20" alt="�ڷ��Է��� ���� �����߰��մϴ�." src="../../../images/imgRowAdd.gIF" border="0"
																name="imgRowAdd"></TD>
														<TD><IMG id="imgRowDel" onmouseover="JavaScript:this.src='../../../images/imgRowDelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgRowDel.gIF'"
																height="20" alt="������ ���������մϴ�." src="../../../images/imgRowDel.gIF" border="0" name="imgRowDel"></TD>
														<td><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
																onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�."
																src="../../../images/imgSave.gIF" width="54" border="0" name="imgSave"></td>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gIF'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" width="54" border="0"
																name="imgExcel"></TD>
													</TR>
												</TABLE>
											</TD>
										</TR>
										<tr height="5">
											<td></td>
										</tr>
									</TABLE>
									<TABLE class="DATA" id="tblDATA" cellSpacing="1" cellPadding="0" width="1040" border="0"
										align="LEFT">
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">���� �ڵ�</TD>
											<TD class="DATA" width="230"><INPUT dataFld="PREESTNO" class="NOINPUT_L" id="txtPREESTNO" title="�������ڵ�" style="WIDTH: 224px; HEIGHT: 22px"
													dataSrc="#xmlBind" readOnly type="text" size="32" name="txtPREESTNO"></TD>
											<TD class="LABEL" style="WIDTH: 92px; CURSOR: hand" width="92">������</TD>
											<TD class="DATA" width="260"><INPUT dataFld="PREESTNAME" class="NOINPUT_L" id="txtPREESTNAME" title="��������" style="WIDTH: 255px; HEIGHT: 22px"
													accessKey="M" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtPREESTNAME"></TD>
											<TD class="LABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtAGREEYEARMON,'')">����Ȯ����</TD>
											<TD class="DATA"><INPUT dataFld="AGREEYEARMON" class="INPUT" id="txtAGREEYEARMON" title="����������" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" dataSrc="#xmlBind" type="text" maxLength="10" size="10" name="txtAGREEYEARMON"><IMG id="imgCalEndarAGREE" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0" name="imgCalEndarAGREE"><INPUT dataFld="ENDFLAG" id="txtENDFLAG" style="WIDTH: 40px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtENDFLAG"><INPUT dataFld="ENDFLAGEXE" id="txtENDFLAGEXE" style="WIDTH: 40px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtENDFLAGEXE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand" width="80">���ۼ�����</TD>
											<TD class="DATA" width="230"><INPUT dataFld="SUSUAMT" class="NOINPUTB_R" id="txtSUSUAMT" title="���ۼ�����" style="WIDTH: 224px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="32" name="txtSUSUAMT">
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand" align="right" width="94">Commition</TD>
											<TD class="DATA" width="260"><INPUT dataFld="COMMITION" class="NOINPUTB_R" id="txtCOMMITION" title="commition ��" style="WIDTH: 256px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37" name="txtCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand" align="right" width="80">�հ�</TD>
											<TD class="DATA"><INPUT dataFld="SUMAMT" class="NOINPUTB_R" id="txtSUMAMT" title="���հ�ݾ�" style="WIDTH: 272px; HEIGHT: 22px"
													accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="40" name="txtSUMAMT"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" onclick="vbscript:Call gCleanField(txtSUSURATE, '')">��������</TD>
											<TD class="DATA"><INPUT dataFld="SUSURATE" class="INPUT_R" id="txtSUSURATE" style="WIDTH: 200px; HEIGHT: 22px"
													accessKey=",NUM,M" dataSrc="#xmlBind" type="text" size="28" name="txtSUSURATE">&nbsp;(%)
											</TD>
											<TD class="LABEL" style="WIDTH: 94px; CURSOR: hand; HEIGHT: 25px">Non Commition</TD>
											<TD class="DATA"><INPUT dataFld="NONCOMMITION" class="NOINPUTB_R" id="txtNONCOMMITION" title="noncommition ��"
													style="WIDTH: 256px; HEIGHT: 22px" accessKey=",NUM" dataSrc="#xmlBind" readOnly type="text" size="37"
													name="txtNONCOMMITION"></TD>
											<TD class="LABEL" style="CURSOR: hand; HEIGHT: 25px" width="80">������ ���</TD>
											<TD class="DATA"><INPUT class="INPUT" id="txtPRINTDAY" title="������������" style="WIDTH: 96px; HEIGHT: 22px"
													accessKey="DATE,M" type="text" maxLength="10" size="10" name="txtPRINTDAY"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndar">&nbsp;&nbsp; <IMG id="imgPrint" onmouseover="JavaScript:this.src='../../../images/imgPrintOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPrint.gif'" height="20" alt="������ �� �μ��մϴ�." src="../../../images/imgPrint.gIF"
													width="54" align="absMiddle" border="0" name="imgPrint">&nbsp;<INPUT dataFld="JOBNO" id="txtJOBNO" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtJOBNO"><INPUT dataFld="CREDAY" id="txtCREDAY" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCREDAY"><INPUT dataFld="CLIENTSUBCODE" id="txtCLIENTSUBCODE" style="WIDTH: 16px; HEIGHT: 21px"
													dataSrc="#xmlBind" type="hidden" size="1" name="txtCLIENTSUBCODE"><INPUT dataFld="CLIENTCODE" id="txtCLIENTCODE" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtCLIENTCODE"><INPUT dataFld="SUBSEQ" id="txtSUBSEQ" style="WIDTH: 16px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="hidden" size="1" name="txtSUBSEQ"></TD>
										</TR>
										<TR>
											<TD class="LABEL">���</TD>
											<TD class="DATA" colSpan="5"><INPUT dataFld="MEMO" id="txtMEMO" style="WIDTH: 950px; HEIGHT: 21px" dataSrc="#xmlBind"
													type="text" maxLength="255" size="152" name="txtMEMO"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px"></TD>
							</TR>
							<TR>
								<TD class="DATAFRAME" style="WIDTH: 100%; HEIGHT: 98%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: hidden; POSITION: relative;HEIGHT:95%; vWIDTH: 100%"
										ms_positioning="GridLayout">
										<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="27464">
											<PARAM NAME="_ExtentY" VALUE="11721">
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
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</FORM>
	</body>
</HTML>
