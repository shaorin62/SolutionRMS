<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMCHARGEEXMAIN01.aspx.vb" Inherits="PD.PDCMCHARGEEXMAIN01" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�ý��� ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/�ý��۰���/EXCEL���δ�
'����  ȯ�� : ASP.NET, VB.NET, COM+
'���α׷��� : SCEXMAIN0.aspx
'��      �� : �������̺� EXCELUPLOAD
'�Ķ�  ���� : 
'Ư��  ���� : ���� 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/07/03 By ParkJS(������)
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- SpreadSheet/Control ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<SCRIPT language="vbscript" id="clientEventHandlersVBS">
<!--
    Option explicit
    Dim mlngRowCnt, mlngColCnt
    Dim mobjccPDDCCHARGEEXCOM , mobjPDCMGET
    Dim mInsOKFlag 'Insert Flag 
    Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode '�˾�����
'=============================
' �̺�Ʈ���ν��� 
'=============================
Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
End Sub

Sub InitPage()
	dim vntInParam
	dim intNo,i
	
    '����������ü ����	
    Set mobjccPDDCCHARGEEXCOM = gCreateRemoteObject("cPDCO.ccPDDCCHARGEEXCOM")
    set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")

   '���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

   'InsOKFlag �� false ������ �����Ѵ�.
	mInsOKFlag   =  false
	
	gSetSheetDefaultColor
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* �ʱ�ȭ�� �Դϴ�. "& vbcrlf & vbcrlf &"* ����: JOBNO�� �����Ͽ� �ֽð�, �ݵ�� ó����ư�� �����ʽÿ�."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "70"
		
	end with
	pnlTab1.style.visibility = "visible" 
	'ȭ�� �ʱⰪ ����
	InitPageData
	
	vntInParam = window.dialogArguments
	intNo = ubound(vntInParam)
	'�⺻�� ����
	mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
	WITH frmThis
		for i = 0 to intNo
			select case i
				case 0 : .txtJOBNO.value = vntInParam(i)	
				case 1 : .txtJOBNAME.value = vntInParam(i)
				case 2 : .txtOUTSCODE.value = vntInParam(i)		'���� ������� �͸�
				case 3 : .txtOUTSNAME.value = vntInParam(i)		'�ڵ� ��� ����
				case 4 : mstrFields = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with
	Call imgFind_onclick
end Sub

'****************************************************************************************
' ȭ���� �ʱ���� ������ ����
'****************************************************************************************
Sub InitPageData
	'��� ������ Ŭ����
	gClearAllObject frmThis
	
	'�ʱ� ������ ����
	with frmThis
		.sprSht.MaxRows = 0
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
End Sub

Sub EndPage()
	set mobjccPDDCCHARGEEXCOM = Nothing
	'PopUp Window �϶� mInsOKFlag �� �Ѱ��ش�.
	If gIsPopupWindow then
 	  window.returnvalue = mInsOKFlag
	End if
	gEndPage
End Sub

'=============================
' ��ɹ�ưŬ���̺�Ʈ
'=============================
Sub imgFind_onclick
    Dim vntRet, vntInParams, dblTAB_ID		
	gFlowWait meWAIT_ON
	makePageData
	gFlowWait meWAIT_OFF
	
	'�߰��κ�
	Dim i, RowNum, intRows
	RowNum = 101
	
	mobjSCGLSpr.SetMaxRows frmThis.sprSht, RowNum
	gOKMsgbox "�����͸� �Է��� �غ� �Ǿ����ϴ�. Excel Data�� �ٿ��־� �ֽʽÿ�.", ""
				
	mobjSCGLSpr.ActiveCell frmThis.sprSht, 1,1
	frmThis.sprSht.focus()
End Sub

Sub imgSave_onclick()
	if frmThis.sprSht.MaxRows = 0 then
		gErrorMsgBox "������ �����Ͱ� �����ϴ�.","����ȳ�"
		Exit Sub
	end if
	
    gFlowWait(meWAIT_ON)
    ProcessRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgDelete_onclick
    gFlowWait(meWAIT_ON)
    DeleteRtn()
    gFlowWait(meWAIT_OFF)
End Sub

Sub imgClose_onclick()
    Window_OnUnload()
End Sub


Sub imgCalEndar_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtBUDGETDATE,frmThis.imgCalEndar,"txtBUDGETDATE_onchange()"
		gSetChange 
	end with
End Sub

Sub txtBUDGETDATE_onchange
	gSetChange
End Sub

'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgJOBNO_onclick
	Call SEARCHJOB_POP()
End Sub

'���� ������List ��������
Sub SEARCHJOB_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtJOBNO.value), trim(.txtJOBNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMJOBNOPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtJOBNO.value = vntRet(0,0) and .txtJOBNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtJOBNO.value = trim(vntRet(0,0))  ' Code�� ����
			.txtJOBNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtJOBNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		'On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetJOBNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtJOBNO.value),trim(.txtJOBNAME.value))
			if not gDoErrorRtn ("txtJOBNAME_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtJOBNO.value = trim(vntData(0,0))
					.txtJOBNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'-----------------------------------------------------------------------------------------
' ����ó ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub imgOUTSCODE_onclick
	Call SEARCHOUT_POP()
End Sub

'���� ������List ��������
Sub SEARCHOUT_POP
	Dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtOUTSCODE.value), trim(.txtOUTSNAME.value)) '<< �޾ƿ��°��
		
		vntRet = gShowModalWindow("PDCMEXECUSTPOP.aspx",vntInParams , 413,435)
		if isArray(vntRet) then
			if .txtOUTSCODE.value = vntRet(0,0) and .txtOUTSNAME.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtOUTSCODE.value = trim(vntRet(0,0))  ' Code�� ����
			.txtOUTSNAME.value = trim(vntRet(1,0))  ' �ڵ�� ǥ��
     	end if
	End with
	gSetChange
End Sub

'�Ѱ��� ã����� ���� �̺�Ʈ�ν� �ش簪�� �ѷ���
Sub txtOUTSNAME_onkeydown
	if window.event.keyCode = meEnter then
		Dim vntData
   		Dim i, strCols
		On error resume next
		with frmThis
			'Long Type�� ByRef ������ �ʱ�ȭ
			mlngRowCnt=clng(0)
			mlngColCnt=clng(0)
			vntData = mobjPDCMGET.GetEXECUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtOUTSCODE.value),trim(.txtOUTSNAME.value))
			if not gDoErrorRtn ("GetEXECUSTNO") then
				If mlngRowCnt = 1 Then
					.txtOUTSCODE.value = trim(vntData(0,0))
					.txtOUTSNAME.value = trim(vntData(1,0))
				Else
					Call SEARCHJOB_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub

'=============================
'SheetEvent
'=============================
Sub sprSht_Change(ByVal Col, ByVal Row)
   mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row

End Sub

Sub sprSht_KeyDown(KeyCode, Shift)
	mobjSCGLSpr.SetFlag frmThis.sprSht, meINS_FLAG
	'IF KeyCode = 86 THEN
	'	CALL TEST(KeyCode, frmThis.sprSht.ActiveCol,frmThis.sprSht.ActiveRow)
	'END IF
End Sub

Sub sprSht_KeyUp(KeyCode, shift)
	If KeyCode = 86 Then
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,1,100) <> "" then
			gErrorMsgbox "�ϰ����Խ� �ѹ��� ���԰����� �����ʹ� 100���Դϴ�. �ٽ� �÷��ֽʽÿ�.",""
			mobjSCGLSpr.ClearText frmThis.sprSht , -1, -1, -1, -1 
			exit sub
		End If
	end if
end Sub

Sub ProcessRtn ()
	Dim intRtn   'Return ��
   	Dim vntData  'Insert �� ������
   	Dim vntData2
   	Dim strMasterData
   	Dim intCnt
   	Dim lngAMT
   	Dim lngCOMMI_RATE
   	Dim strCOMMISSION
   	Dim strYEARMON
   	Dim strJOBNO
   	dIM strOUTSCODE
   	dim strREVSEQ
   	'������ Validation
   	with frmThis
   		If trim(.txtJOBNO.value) = "" or trim(.txtOUTSCODE.value) = "" or .txtBUDGETDATE.value = ""   Then
			gErrorMsgBox "�����Ƿڹ�ȣ�� ����ó�ڵ�, �������ڴ� �ʼ� �Դϴ�.",""
			exit sub
		End If
		
		'���� Rows ����ó��
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMNAME",intCnt) = ""  then 
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			else
				CALL SetTrim (intCnt) ' ���鹮�ڿ� ����
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,0
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",intCnt,0
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"QTY",intCnt,0
				End If
				
			End If
		Next
		
		
		'==================��������
		'if DataValidation =false then exit sub
		'Exit SUb
		'==================��������
		 For intCnt = 1 To .sprSht.MaxRows
            lngAMT =  mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)  
            if lngAMT = "" or lngAMT ="0" then
				if  mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) <> "" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt) <> "" _
					and mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) <> "0" AND  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt)<> "0" THEN
					lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",intCnt) *  mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",intCnt)
					 mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,lngAMT
				else
					 mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,0
				END IF 
            end if
         Next
 	
 		strMasterData = gXMLGetBindingData (xmlBind)
 		
		strJOBNO = .txtJOBNO.value
		strOUTSCODE =.txtOUTSCODE.value
		strREVSEQ = 0
		'On error resume next
		'����� �����͸� �����´�.
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, "ITEMNAME|STD|QTY|PRICE|AMT|BIGO|ATTR01")
 	    if  not IsArray(vntData) then 
		    gErrorMsgBox "����� " & meNO_DATA,"�������"
		    exit sub
        end if
  	    Dim STime, ETime
  	   
  	    STime = Time
			intRtn = mobjccPDDCCHARGEEXCOM.ProcessRtn(gstrConfigXML, strMasterData, vntData, strJOBNO, strOUTSCODE, strREVSEQ, replace(.txtBUDGETDATE.value,"-",""), .txtMEMO.value)
		ETime = Time

		if not gDoErrorRtn ("ProcessRtn") then
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
	   	    mobjSCGLSpr.SetMaxRows frmThis.sprSht, 0 
	   	    gOKMsgbox "�����͸� ���������� UPLOAD �Ͽ����ϴ�.", "" 

	   	    mInsOKFlag = true
   		end if

   	end with
End Sub

Sub SetTrim (Row) 
	With frmThis
		mobjSCGLSpr.SetTextBinding .sprSht,"ITEMNAME",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"ITEMNAME",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"STD",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"STD",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"QTY",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"QTY",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PRICE",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PRICE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"AMT",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"BIGO",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"BIGO",Row))
	End With
End Sub

Function DataValidation ()
	dim i,j
	DataValidation = false
	with frmThis
		'�����Ͱ� ����Ǿ����� �˻�
		if not mobjSCGLSpr.IsDataChanged(.sprSht) then
			gErrorMsgBox "����� " & meNO_DATA,"�������"
			exit function
		end if

   		'=================== ����üũ����
   		Dim intCnt
   		Dim strArray
   		Dim Rowcnt
   		Dim Colcnt
   		Dim strMedAndReal
   		Dim strERR
   		Dim strMEDCODENAME
   		Dim strMEDCODE
   		Dim strREALMEDCODE
   		Dim strCLIENTNAME
   		Dim intVal
   		Dim intRtn
   		Dim vntData
   		Dim vntData2
   		Dim vntData3
   		Dim strCLIENTCODE
   		Dim strDEPTCODE
   		Dim strSEQCODE
   		Dim lngAMT
   		Dim lngREAL_AMT
   		Dim lngBONUS
   		Dim strCLIENTSUBNAME
   		Dim strCLIENTSUBCODE
   		Dim strDEPT_CD
   		Dim strSUBSEQ
   		Dim strMPPNAME, strMPPCODE
   		
   		
   		intVal = 0
   		'����üũ�κ��� �ϴ� �������� �����.
   		 For intCnt = 1 To .sprSht.MaxRows
   			mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,""
   		 Next
   		 
   		 '�������ڵ�üũ
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt),1,1) = "A" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt)) = 6 Then
				vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_CLIENTCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
				if not gDoErrorRtn ("SelectRtn_CODE") then
					IF mlngRowCnt <> 1 Then
						strERR = "�������ڵ����"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strCLIENTNAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
   				vntData = mobjccPDDCCHARGEEXCOM.SelectRtn_CLIENTNAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTNAME)
	   			
				if not gDoErrorRtn ("SelectRtn_CLIENTNAME") then
					If mlngRowCnt = 1 Then
						strCLIENTCODE = vntData(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",intCnt,strCLIENTCODE
					Else
						strERR = "�������ڵ����"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		
	   	 If intVal Then Exit Function
	   	 	
   		'=================================
	end with
	
	DataValidation = true
	
End Function

Sub DeleteRtn()
	Dim vntData
	Dim intSelCnt, intRtn, i
	On error resume next
	with frmThis
		'�� �Ǿ� ������ ���
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)
		if gDoErrorRtn ("DeleteRtn") then exit sub
		if intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		if intRtn <> vbYes then exit sub
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			if not gDoErrorRtn ("DeleteRtn_SC_USER_COL") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
			end if
		next
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
	end with
End Sub
'======================================
'��Ÿ�Լ�
'======================================
Sub makePageData
     Dim vntData
     
     With frmThis        
        
        gSetSheetDefaultColor() 
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, 7, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, "ITEMNAME|STD|QTY|PRICE|AMT|BIGO|ATTR01"
        mobjSCGLSpr.SetHeader           .sprSht, "�����׸�|�԰�|����|�ܰ�|�ݾ�|���|��������"
        mobjSCGLSpr.SetColWidth .sprSht, "-1","         20|  14|  14|  14|  16|  28|15"
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, "ITEMNAME|STD|BIGO"     , , ,200
        mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "QTY|PRICE|AMT", -1, -1, 0
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"        
       
    End With
End Sub

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" height="100%" width="100%" >
				<TR>
					<TD>
						<TABLE id="tblTitle" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE" id="tblTitleName"><FONT face="����">&nbsp;���� �������ε�</FONT></td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right"  height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="0" border="0">
										<TR>
											<TD width="3"><IMG id="ImgFind" onmouseover="JavaScript:this.src='../../../images/imginitOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imginit.gif'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imginit.gif" border="0" name="imgFind"></TD>
											<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �����մϴ�."
													src="../../../images/imgSave.gif" width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gif'"
													height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgDelete.gif" width="54" border="0"
													name="imgDelete"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'"
													height="20" alt="�ڷḦ �ݽ��ϴ�." src="../../../images/imgClose.gIF" width="54" border="0"
													name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" height="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="HEIGHT: 17px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="1040" border="0" align="LEFT">
										<TR>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtJOBNO, txtJOBNAME)"><FONT face="����">Job&nbsp;No</FONT></TD>
											<TD class="DATA" width="420"><INPUT class="INPUT_L" id="txtJOBNAME" title="�ڵ��" style="WIDTH: 256px; HEIGHT: 22px" type="text"
													maxLength="100" align="left" size="37" name="txtJOBNAME"><IMG id="ImgJOBNO" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23"
													align="absMiddle" border="0" name="ImgJOBNO"><INPUT class="INPUT" id="txtJOBNO" title="jobno" style="WIDTH: 88px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="6" size="9" name="txtJOBNO"></TD>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtOUTSCODE, txtOUTSNAME)"><FONT face="����">����ó</FONT></TD>
											<TD class="DATA"><INPUT class="INPUT_L" id="txtOUTSNAME" title="�ڵ��" style="WIDTH: 256px; HEIGHT: 22px"
													type="text" maxLength="100" align="left" size="37" name="txtOUTSNAME"><IMG id="imgOUTSCODE" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
													border="0" name="imgOUTSCODE"><INPUT class="INPUT" id="txtOUTSCODE" title="jobno" style="WIDTH: 88px; HEIGHT: 22px" accessKey=",M"
													type="text" maxLength="6" size="9" name="txtOUTSCODE"></TD>
										</TR>
										<TR>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtBUDGETDATE, '')"><FONT face="����">������</FONT></TD>
											<TD class="DATA" width="420"><INPUT class="INPUT" id="txtBUDGETDATE" title="������" style="WIDTH: 128px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="100" align="left" size="16" name="txtBUDGETDATE"><IMG id="imgCalEndar" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle" border="0"
													name="imgCalEndar"></TD>
											<TD class="LABEL" style="WIDTH: 100px" onclick="vbscript:Call gCleanField(txtMEMO, '')"><FONT face="����">��&nbsp;&nbsp; 
													�� </FONT>
											</TD>
											<TD class="DATA" colSpan="2"><FONT face="����"><INPUT class="INPUT_L" id="txtMEMO" title="���" style="WIDTH: 336px; HEIGHT: 22px" type="text"
														maxLength="100" align="left" size="50" name="txtMEMO"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<tr>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 100%" vAlign="top" align="center">
									<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%;height:95%; POSITION: relative" 
									ms_positioning="GridLayout">
										<OBJECT id=sprSht style="WIDTH: 100%; HEIGHT: 95%" classid=clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5>
	<PARAM NAME="_Version" VALUE="393216">
	<PARAM NAME="_ExtentX" VALUE="25321">
	<PARAM NAME="_ExtentY" VALUE="18680">
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
								</td>
							</tr>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
