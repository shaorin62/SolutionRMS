<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMEXCUTIONPOP.aspx.vb" Inherits="MD.MDCMEXCUTIONPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>������ ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : ������ �����˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : MDCMEXCUTIONPOP.aspx
'��      �� : JOBNO ��ȸ�� ���� �˾�
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/18 By Kim Tae Ho
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI ���� ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- �������� ���� Ŭ���̾�Ʈ ��ũ��Ʈ�� Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT CLASSID="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" 
			id="Microsoft_Licensed_Class_Manager_1_0">
			<PARAM NAME="LPKPath" VALUE="fpSpread60.lpk">
		</OBJECT>  
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjMDCMGET
Dim mobjMDCMPRINTEXCUTION
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode

'-----------------------------
' �̺�Ʈ ���ν��� 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub
Sub imgClose_onclick()
	Window_OnUnload
End Sub

Sub imgSave_onclick ()
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

Sub txtJOBNO_onkeydown
	if window.event.keyCode = meEnter then
		Call imgQuery_onclick()	
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
End Sub



Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------
' Spread Sheet Event
'-----------------------------	
'onblour �̺�Ʈ
Sub txtDEMANDAMT_onblur
	with frmThis
		call gFormatNumber(.txtDEMANDAMT,0,true)
	end with
End Sub
Sub sprSht_change(ByVal Col,ByVal Row)
	
	Dim vntData
   	Dim i, strCols
   	Dim strCode, strCodeName
   	Dim strQTY,strPRICE,strAMT 
   	Dim intCnt,intCnt0,intCnt1
   	Dim lngSUSU
   	Dim lngSUSUAMT
   	Dim lngRATE
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		strCode = ""
		strCodeName = ""
		
		IF  Col = 4 Then
			strCode		= ""'mobjSCGLSpr.GetTextBinding( .sprSht,"CUSTCODE",frmThis.sprSht.ActiveRow)
			strCodeName = mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row)
			
			vntData = mobjMDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,strCode,strCodeName)

			if not gDoErrorRtn ("GetCUSTNO") then
				If mlngRowCnt = 1 Then
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntData(0,0)
					mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntData(1,0)			
					'mobjSCGLSpr.CellChanged .sprSht, frmThis.sprSht.ActiveCol-1,frmThis.sprSht.ActiveRow
					.txtYEARMON.focus
					.sprSht.focus 
					'mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
				Else
					mobjSCGLSpr_ClickProc .sprSht, 4, Row
					.txtYEARMON.focus
					.sprSht.focus 
				End If
   			end if
   		
		end if
		If Col = 5 Then
			if 100 =< mobjSCGLSpr.GetTextBinding( .sprSht,"SUSURATE",Row) Then
				msgbox "�й����� 100 ���� Ŭ�� �����ϴ�."
				Exit Sub
			End if
			
			if (mobjSCGLSpr.GetTextBinding( .sprSht,"SUSURATE",Row) <> "" Or mobjSCGLSpr.GetTextBinding( .sprSht,"SUSURATE",Row) <> 0) Then 
			lngSUSUAMT = (.txtAMT.value * mobjSCGLSpr.GetTextBinding( .sprSht,"SUSURATE",Row) ) * 0.01
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSU",Row,lngSUSUAMT
			end if
		End if
		
		If Col = 6 Then
			if .txtAMT.value < mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) Then
				msgbox "�й谪 �� ������ ���� Ŭ�� �����ϴ�."
				Exit Sub
			End if
		
			if (mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) <> "" Or mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) <> 0) Then 
			lngRATE = (mobjSCGLSpr.GetTextBinding( .sprSht,"SUSU",Row) / .txtAMT.value) * 100
			mobjSCGLSpr.SetTextBinding .sprSht,"SUSURATE",Row,lngRATE
			end if
		
		End if
   	end with
   	mobjSCGLSpr.CellChanged frmThis.sprSht, Col,Row
	SUM_AMT
End Sub	
sub sprSht_DblClick (Col,Row)
	'���õ� �ο� ��ȯ
	'window.returnvalue = mobjSCGLSpr.GetClip (frmThis.sprSht,1,frmThis.sprSht.ActiveRow,frmThis.sprSht.MaxCols,1,1)
	'call Window_OnUnload()
end sub
sub imgAddRow_onclick ()
	With frmThis
		call sprSht_Keydown(meINS_ROW, 0)
	End With 
end sub
sub imgDelRow_onclick ()
	'�������� ���λ���
	With frmThis
		call sprSht_Keydown(meDEL_ROW, 0)
	End With 
	'call sprSht_Keydown(meDEL_ROW, 0)
	'DeleteRtn_Dtl
end sub
Sub sprSht_Keydown(KeyCode, Shift) 
    Dim intRtn
    if KeyCode <> meINS_ROW and KeyCode <> meDEL_ROW and KeyCode <> meCR then exit sub  
    if KeyCode = meCR Then
		if frmThis.sprSht.ActiveRow = frmThis.sprSht.MaxRows and frmThis.sprSht.ActiveCol = 7 Then
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		End if
	Else 
		intRtn = mobjSCGLSpr.InsDelRow(frmThis.sprSht, cint(KeyCode), cint(Shift), -1, 1)
		Select Case intRtn
			Case meINS_ROW':
					'SetDefaultNewRow
			Case meDEL_ROW: DeleteRtn_DTL
		End Select
    End if
End Sub



Sub sprSht_ButtonClicked (Col,Row,ButtonDown)
	dim vntRet, vntInParams
	Dim strGUBUN
	with frmThis
		strGUBUN = ""
		IF Col = 3 Then
			IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN") then exit Sub
		
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		
		end if
		.txtYEARMON.focus
		.sprSht.focus 

	End with
	
End Sub

Sub mobjSCGLSpr_ClickProc(sprSht, Col, Row)
dim vntRet, vntInParams
	with frmThis
		IF Col = 9 Then			
			'IF Col <> mobjSCGLSpr.CnvtDataField(.sprSht,"BTN1") then exit Sub
			Dim strGUBUN
			strGUBUN = ""
			vntInParams = array("", mobjSCGLSpr.GetTextBinding( .sprSht,"CLIENTNAME",Row))
			
			vntRet = gShowModalWindow("../MDCO/MDCMCUSTPOP.aspx",vntInParams , 413,425)
			
			IF isArray(vntRet) then
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row, vntRet(0,0)
				mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTNAME",Row, vntRet(1,0)			
				mobjSCGLSpr.CellChanged .sprSht, Col,Row
				.txtYEARMON.focus
				.sprSht.focus 
				mobjSCGLSpr.ActiveCell .sprSht, Col+2,Row
			End IF
		end if
		.txtYEARMON.focus	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
		.sprSht.Focus
	end with
End Sub
'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	Dim intNo,i,vntInParam
	
	set mobjMDCMPRINTEXCUTION = gCreateRemoteObject("cMDPT.ccMDPTPRINTEXCUTION")
	set mobjMDCMGET = gCreateRemoteObject("cMDCO.ccMDCOGET")
	
	with frmThis
		.txtJOBYEARMON.style.visibility = "hidden"
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		'strEXYEARMON, strEXSEQ,strEXAMT,strEXSUSU
		for i = 0 to intNo
			select case i
				case 0 : .txtYEARMON.value = vntInParam(i)	
				case 1 : .txtSEQ.value = vntInParam(i)
				case 2 : .txtAMT.value = vntInParam(i)
				case 3 : .txtSUSU.value = vntInParam(i)
			end select
		next
		'msgbox .txtJOBYEARMON.value
		'SpreadSheet ������
		gSetSheetDefaultColor()
	End with
        With frmThis
			'���ν�Ʈ
            gSetSheetColor mobjSCGLSpr, .sprSht 
			mobjSCGLSpr.SpreadLayout .sprSht, 7, 0
			mobjSCGLSpr.AddCellSpan  .sprSht, 2, SPREAD_HEADER, 2, 1
			mobjSCGLSpr.SpreadDataField .sprSht, "SUBSEQNTCODE | BTN | CLIENTNAME | SUSURATE | SUSU | NOTE"
			'mobjSCGLSpr.SetHeader .sprSht, "����|�ڵ�| ������ | CLIENTCODE | BTN | CLIENTNAME | SUSURATE | SUSU | NOTE"
			mobjSCGLSpr.SetHeader .sprSht, "����|�ڵ�| ������ | ��������|������| ���"
			mobjSCGLSpr.SetColWidth .sprSht, "-1", " 0| 6 |2| 18 | 10 |11 |17"
			mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
			mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
			mobjSCGLSpr.SetCellTYpeButton2 .sprSht,"..", "BTN"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTCODE"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "CLIENTNAME"
			mobjSCGLSpr.SetCellTypeEdit2 .sprSht, "NOTE"
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUSU", -1, -1, 0
			mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "SUSURATE", -1, -1, 1
			'SetCellTypeFloat2
			'SetCellTypeStatic2
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "SEQ",-1,-1,1,2,false
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTCODE",-1,-1,2,2,false
			'mobjSCGLSpr.SetCellAlign2 .sprSht, "CUSTNAME",-1,-1,0,2,false
			mobjSCGLSpr.ColHidden .sprSht, "SUBSEQ", true
			mobjSCGLSpr.SetScrollBar .sprSht,2,False,0,-1
			'Sum ��Ʈ
			gSetSheetColor mobjSCGLSpr, .sprShtSum
			mobjSCGLSpr.SpreadLayout .sprShtSum, 7, 1, 0,0,1,1,1,false,true,true,1
			mobjSCGLSpr.SpreadDataField .sprShtSum, "SUBSEQ | CLIENTCODE | BTN | CLIENTNAME | SUSURATE | SUSU | NOTE"
			mobjSCGLSpr.AddCellSpan  .sprShtSum, 2, 1, 2, 1
			mobjSCGLSpr.SetText .sprShtSum, 2, 1, "�� ��"
			mobjSCGLSpr.SetScrollBar .sprShtSum, 0
			mobjSCGLSpr.SetBackColor .sprShtSum,"1|2",rgb(205,219,215),false
			mobjSCGLSpr.SetCellTypeFloat2 .sprShtSum, "SUSU", -1, -1, 0
			mobjSCGLSpr.ColHidden .sprShtSum, "SUBSEQ", true
			mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
			mobjSCGLSpr.SetRowHeight .sprShtSum, "-1", "15"
			.sprSht.focus
        End With
        
        SelectRtn
end sub

Sub EndPage()
	set mobjMDCMPRINTEXCUTION = Nothing
	gEndPage
End Sub

sub SelectRtn ()
   	Dim vntData
   	Dim i, strCols

	'On error resume next
	with frmThis
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)

		vntData = mobjMDCMPRINTEXCUTION.GetGETDIVAMT(gstrConfigXml,mlngRowCnt,mlngColCnt,.txtYEARMON.value,.txtSEQ.value)

		if not gDoErrorRtn ("SelectRtn") then
			mobjSCGLSpr.SetClip .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
			'If mlngRowCnt < 1 Then
			'frmThis.sprSht.MaxRows = 20 '���ʷο찳�� �����Һκ�
			'mobjSCGLSpr.SetFlag  .sprSht,meINS_TRANS
			'End If
   			gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			mobjSCGLSpr.DeselectBlock .sprSht
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			Call SUM_AMT ()
   		end if
   	end with
end sub
Sub DeleteRtn_DTL
	Dim vntData
	Dim intSelCnt, intRtn, i,intCnt,intCnt2
	dim strYEARMON,strSEQ,strSUBSEQ
	Dim lngSUMAMT,lngSUMAMT2
	On error resume next
	
	with frmThis
		'�� �Ǿ� ������ ���
		intSelCnt = 0
		vntData = mobjSCGLSpr.GetSelectedItemNo(.sprSht,intSelCnt)

		if gDoErrorRtn ("DeleteRtn_Dtl") then exit sub

		if intSelCnt < 1 then
			gErrorMsgBox "������ �ڷ�" & meMAKE_CHOICE, ""
			Exit sub
		end if
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		if intRtn <> vbYes then exit sub
		strYEARMON = ""
	
		strSEQ = 0
		strSUBSEQ = 0
		'�հ谡 �´��� ���ΰ˻�
		'��������Ǿ� �ִ� �ݾ�
		''���� �հ�ݾ� ������ ���� ���� �߰�
		'for intCnt = 1 To .sprSht.MaxRows
		'lngSUMAMT = lngSUMAMT + mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",intCnt)	
		'Next
		''������ �ݾ�
		'for intCnt2 = intSelCnt-1 to 0 step -1
		'lngSUMAMT2 = lngSUMAMT2 + mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",vntData(intCnt2))
		'Next
		''(���� ����Ǿ��ִ� �ݾ� - �����ұݾ�) �� û���ݾװ� �ٸ��ٸ�,
		''û���ݾ��� ũ�ٸ� ���� ȭ�� ���� ���Է� ���·� ����
		
		'If lngSUMAMT - lngSUMAMT2 <> CDBL(.txtSUSU.value) Then 
		'gErrorMsgbox "������ �ݾ��� û���հ� �ݾװ� ��ġ���� �ʽ��ϴ�." & vbCrlf & "�������� �ʴ� �ݾ��� û���հ�ݾװ� ��ġ��Ű����","�Է¿���"
		'exit Sub
		'End If 
		''�հ� ���� ��
		
		'���õ� �ڷḦ ������ ���� ����
		for i = intSelCnt-1 to 0 step -1
			strYEARMON = Trim(.txtYEARMON.value) 
			strSEQ = Trim(.txtSEQ.value) 
			strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",vntData(i))	
			
			'Insert Transaction�� �ƴ� ��� ���� ������ü ȣ��
			if mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",vntData(i)) <> "" then
				intRtn = mobjMDCMPRINTEXCUTION.DeleteRtn(gstrConfigXml,strYEARMON,strSEQ,strSUBSEQ)
			end if
			if not gDoErrorRtn ("DeleteRtn") then
				mobjSCGLSpr.DeleteRow .sprSht,vntData(i)
				'�հ�����
				gWriteText "", "�ڷᰡ ����" & mePROC_DONE
   			end if
		next
		'ProcessRtn
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
	end with
End Sub

'�⺻�׸����� ���WIDTH�� ���ҽÿ� �հ� �׸��嵵 �Բ����Ѵ�.
sub sprSht_ColWidthChange(ByVal Col1, ByVal Col2)
	With frmThis
		mobjSCGLSpr.SameColWidth .sprSht, .sprShtSum
	End with
end sub
'��ũ���̵��� �հ� �׸����� �Բ� �����δ�.
Sub sprSht_TopLeftChange(ByVal OldLeft, ByVal OldTop, ByVal NewLeft, ByVal NewTop)
    mobjSCGLSpr.TopLeftChange frmThis.sprShtSum, NewTop, NewLeft
End Sub
Sub SUM_AMT()
	Dim lngCnt
	Dim strSUMDEMANDAMT
	Dim strDIVAMT
	strSUMDEMANDAMT = 0
	With frmThis
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		
		mobjSCGLSpr.SetTextBinding .sprShtSum,"SUSU",1, strSUMDEMANDAMT
	End With
End Sub
Sub ProcessRtn ()
    Dim intRtn
   	dim vntData
	Dim strMasterData
	Dim strYEARMON,strSEQ,strSUSU,strAMT
	Dim strSUMDEMANDAMT
   	Dim strDIVAMT
	Dim lngCnt,intCnt
	
	with frmThis
   		'������ Validation
		if DataValidation =false then exit sub
		'On error resume next
		
		For lngCnt = 1 To .sprSht.MaxRows
				strDIVAMT = 0
				strDIVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",lngCnt)
				strSUMDEMANDAMT = strSUMDEMANDAMT + strDIVAMT
		Next
		'ȸ�ǰ�� �޶� ����ɼ� ����.. �д�ݾ��� û���ݾ׺��� ũ�ٸ� ����,,
		'���� �۴ٸ� �ٷ����� û���ݾ��� ���꿡�� ���� �Ǵ� �谨 �Ǹ� ���� �д� PD_GROUP_DIVAMT �� ���� ���� 
		'�հ�ݾ� ���߾�� �Ұ�� ������ ������ �߰��Ѵ�.
		If CDBL(.txtSUSU.value) < strSUMDEMANDAMT Then
   			msgbox "����ó�����ݾ� �� ���� �������հ�ݾ׺��� ����� �մϴ�.."
   			Exit Sub
   		End IF

		'����� ��ο� ������ ����
   		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,2,intCnt) = "" then
			mobjSCGLSpr.DeleteRow .sprSht,intCnt
			End If
		Next
		
		'��Ʈ�� ����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"SUBSEQ | CLIENTCODE | BTN | CLIENTNAME | SUSURATE | SUSU | NOTE")
		
		if .sprSht.MaxRows = 0 Then
			MsgBox "������ �����͸� �Է� �Ͻʽÿ�"
			Exit Sub
		end if
		'strYEARMON,strSEQ,strSUSU,strAMT
		strYEARMON	 = .txtYEARMON.value
		strSEQ = .txtSEQ.value
		strSUSU = .txtSUSU.value
		strAMT = .txtAMT.value
		'if strDEMANDAMT = "" Then
		'strDEMANDAMT = 0
		'End If
		'strITEMCODESEQ  = ""
		'������ �����͸� ���� �´�.
		
		'strMasterData = gXMLGetBindingData (xmlBind)
		'ó�� ������ü ȣ��
		'intRtn = mobjEXE_HDR.ProcessRtn()
		
		intRtn = mobjMDCMPRINTEXCUTION.ProcessRtn(gstrConfigXml,vntData,strYEARMON,strSEQ,strSUSU,strAMT)
	
		if not gDoErrorRtn ("ProcessRtn") then
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			SelectRtn
   		end if
   		
   	end with
End Sub
'------------------------------------------
' ������ ó���� ���� ����Ÿ ����
'------------------------------------------
Function DataValidation ()
	DataValidation = false
	
	Dim vntData
   	Dim i, strCols
    Dim intCnt,strValidationFlag
	'On error resume next
	with frmThis
  			
		'Master �Է� ������ Validation : �ʼ� �Է��׸� �˻�
   		IF not gDataValidation(frmThis) then exit Function
   		strValidationFlag = ""
  		for intCnt = 1 to .sprSht.MaxRows
			 if mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt) <> "" _
			 AND (mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",intCnt) = "" _
			 Or mobjSCGLSpr.GetTextBinding(.sprSht,"SUSU",intCnt) = 0) Then 
					gErrorMsgBox intCnt & " ��° ���� �Է³��� �� Ȯ���Ͻʽÿ�","�Է¿���"
					Exit Function
			 End if
		next
   		
   		
   	End with
	DataValidation = true
End Function
-->
		</script>
	</HEAD>
	<body class="base"  bottomMargin="0"
		leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="573" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" 
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle" valign=bottom>
												������&nbsp;����
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 225px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 108px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD style="WIDTH: 126px"><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgSave.gIF'" height="20" alt="�ڷḦ �����մϴ�." src="../../../images/imgSave.gIF"
													width="54" border="0" name="imgSave"></TD>
											<TD><IMG id="imgClose" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'" style="CURSOR: hand"
													onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20" alt="�ڷḦ �ݽ��ϴ�."
													src="../../../images/imgClose.gIF" width="54" border="0" name="imgClose"></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblTitle2" height="1" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/PopupBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="1"></td>
							</tr>
						</table>
						<TABLE id="tblBody" style=" HEIGHT: 340px" cellSpacing="0" cellPadding="0" width="100%"
							border="0">
							<TR>
								<TD class="TOPSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" style="HEIGHT: 20px" vAlign="middle" height="20"><FONT face="����">
										<TABLE class="SEARCHDATA" id="tblKey" cellSpacing="0" cellPadding="0" width="100%" align="right"
											border="0">
											<TBODY>
												<TR>
													<TD class="SEARCHLABEL" >
														������ȣ&nbsp;
													</TD>
													<td class="SEARCHDATA" style="WIDTH: 138px"><INPUT class="NOINPUT" id="txtYEARMON" style="WIDTH: 80px; HEIGHT: 22px" type="text" size="8"
															name="txtYEARMON"><INPUT class="NOINPUT" id="txtSEQ" style="WIDTH: 56px; HEIGHT: 22px" type="text" size="4"
															name="txtSEQ">
													</td>
													<TD class="SEARCHLABEL" >
													����ݾ�&nbsp;
													<td class="SEARCHDATA" style="WIDTH: 90px"><INPUT class="NOINPUT" id="txtAMT" style="WIDTH: 120px; HEIGHT: 22px" tabIndex="1" type="text"
															size="14" name="txtAMT">
													</td>
													<TD class="SEARCHLABEL" >
													&nbsp;&nbsp;&nbsp;&nbsp; ������&nbsp;
													<td class="SEARCHDATA"><INPUT class="NOINPUT" id="txtSUSU" style="WIDTH: 126px; HEIGHT: 22px" tabIndex="1" type="text"
															size="15" name="txtSUSU">
													</td>
												</TR>
											</TBODY>
										</TABLE>
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD style="HEIGHT: 26px" vAlign="bottom" align="right" width="100%"><INPUT class="NOINPUT" id="txtJOBYEARMON" style="WIDTH: 122px; HEIGHT: 22px" tabIndex="1"
										type="text" size="15" name="txtJOBYEARMON"><IMG id="ImgAddRow" onmouseover="JavaScript:this.src='../../../images/imgAddRowOn.gif'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgAddRow.gif'" alt="�� �� �߰�" src="../../../images/imgAddRow.gif"
										width="54" border="0" name="imgAddRow"><IMG id="ImgDelRow" onmouseover="JavaScript:this.src='../../../images/imgDelRowOn.gif'"
										style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelRow.gif'" alt="�� �� ����" src="../../../images/imgDelRow.gif"
										width="54" border="0" name="imgDelRow">
								</TD>
							</TR>
							<TR>
								<TD align="center"><FONT face="����">
										<OBJECT id="sprSht" style="WIDTH: 574px; HEIGHT: 251px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
											>
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15187">
											<PARAM NAME="_ExtentY" VALUE="6641">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
										<OBJECT id="sprShtSum" style="WIDTH: 574px; HEIGHT: 23px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5">
											<PARAM NAME="_Version" VALUE="393216">
											<PARAM NAME="_ExtentX" VALUE="15187">
											<PARAM NAME="_ExtentY" VALUE="609">
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
											<PARAM NAME="MaxCols" VALUE="5">
											<PARAM NAME="MaxRows" VALUE="500">
											<PARAM NAME="MoveActiveOnFocus" VALUE="-1">
											<PARAM NAME="NoBeep" VALUE="0">
											<PARAM NAME="NoBorder" VALUE="0">
											<PARAM NAME="OperationMode" VALUE="0">
											<PARAM NAME="Position" VALUE="0">
											<PARAM NAME="ProcessTab" VALUE="-1">
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
									</FONT>
								</TD>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
							</TR>
						</TABLE>
						<FONT face="����"></FONT>
				</TD>
				</FORM></TR>
		</TABLE>
	</body>
</HTML>
