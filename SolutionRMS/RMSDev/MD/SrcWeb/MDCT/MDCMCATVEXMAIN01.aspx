<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCATVEXMAIN01.aspx.vb" Inherits="MD.MDCMCATVEXMAIN01" %>
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
    Dim sprSht_DataFields
    Dim vntData_DataFields	
    Dim sprSht_DisplayFields
    Dim sprSht_NotNull
    Dim vntData_Nullable
    Dim sprSht_DefualtValueFields
    Dim vntData_DefaultValue
    Dim vntData_DataType
    Dim vntData_DataLength
    Dim mdblTAB_ID, mstrTAB_NAME, mstrTAB_USER_NAME, mstrTAB_TYPE, mstrTAB_DESC 
    Dim mobjccMDCATVEXCOM  , mobjccMDELECEXBrowse
    Dim mobjPDCMJOBNOREG
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
    '����������ü ����	
    Set mobjccMDCATVEXCOM = gCreateRemoteObject("cMDCT.ccMDCTCATVEXCOM")
    Set mobjccMDELECEXBrowse = gCreateRemoteObject("cMDCT.ccMDCTCATVEXBrowse")

   '���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	With frmThis
	.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
	End With
   'InsOKFlag �� false ������ �����Ѵ�.
	mInsOKFlag   =  false
	
	gSetSheetDefaultColor
	with frmThis
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout    .sprSht, 1, 0
		mobjSCGLSpr.SetHeader       .sprSht, "* �ʱ�ȭ�� �Դϴ�. "& vbcrlf & vbcrlf &"* ����: ����� �����Ͽ� �ֽð�, �ݵ�� ó����ư�� �����ʽÿ�."
		mobjSCGLSpr.SetRowHeight    .sprSht, "0", "40" 
		mobjSCGLSpr.SetColWidth     .sprSht, "-1", "70"
		
	end with
	
	Call imgFind_onclick
end Sub

Sub EndPage()
	set mobjccMDCATVEXCOM = Nothing
	set mobjccMDELECEXBrowse = Nothing
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
		with frmThis
			If .txtYEARMON.value <> ""  Then
			Else
				gErrorMsgBox "����� �ʼ� �Դϴ�.",""
				exit sub
			End If
			
			If LEN(.txtYEARMON.value) <> 6 Then
				gErrorMsgBox "����� 6�ڸ� �Դϴ�.",""
				exit sub
			End If
		End with
		
		'DataProcedure
		
		mdblTAB_ID        = 33
		mstrTAB_NAME      = "MD_CATV_MEDIUM"
		mstrTAB_USER_NAME = "CATV�����Ź���"
		mstrTAB_TYPE      = "TABLE"
		mstrTAB_DESC      = "CATV�����Ź���ε�"
		
		gFlowWait meWAIT_ON
		makePageData
		gFlowWait meWAIT_OFF
		
		'�߰��κ�
		Dim i, RowNum, intRows
		RowNum = 501
	    
		mobjSCGLSpr.SetMaxRows frmThis.sprSht, RowNum 
		intRows = Ubound(vntData_DefaultValue,1) +1
	    
		For i=1 To intRows
			mobjSCGLSpr.SetText frmThis.sprSht, i , -1, vntData_DefaultValue(i-1) 
		Next 
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
		If mobjSCGLSpr.GetTextBinding(frmThis.sprSht,8,500) <> "" then
			gErrorMsgbox "�ϰ�û��� �ѹ��� ���԰����� �����ʹ� 500���Դϴ�. �ٽ� �÷��ֽʽÿ�.",""
			mobjSCGLSpr.ClearText frmThis.sprSht , -1, -1, -1, -1 
			exit sub
		End If
	end if
end Sub


'==================================================
'�⺻�Է��� ����� �����Ϳ� �������� ������Ʈ
'==================================================
Sub DataProcedure()
	Dim intRtn
	Dim strYEARMON 
	Dim strCLIENTCODE
	Dim strSUBSCRIPTION_AMT
	Dim strSEQ
	with frmThis
		On error resume next		
		strYEARMON= .txtYEARMON.value
		strSEQ = "33"
		
		
		intRtn = mobjccMDCATVEXCOM.ElecExcelUpload(gstrConfigXML,strYEARMON)
		if not gDoErrorRtn ("DataProcedure") then
  		end if
 	end with
   	
End Sub


Function SelectRtn_Dup ()
	SelectRtn_Dup = False
	
	Dim intRtn
	Dim intRtn2
	Dim intRtnDup
	Dim intDelete
	Dim intCnt
	Dim strMEDNAME
	Dim strCLIENTCODE
	Dim strYEARMON

	With frmThis	
		strYEARMON = .txtYEARMON.value
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		intRtn = mobjccMDCATVEXCOM.SelectRtn_Dup(gstrConfigXML,mlngRowCnt,mlngColCnt,strYEARMON)
		if not gDoErrorRtn ("SelectRtn_Dup") then
			If mlngRowCnt <> 0 Then
				intRtn2 = gYesNoMsgbox("�ش��� �ڷ� �� �����մϴ�." & vbcrlf &" �ƴϿ� �� �����Ͻø� �����ڷ� �� �Բ� ���� �˴ϴ�." & vbcrlf &"���� �ڷḦ �����Ͻð� ���� �Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
				if intRtn2 <> vbYes then
				SelectRtn_Dup = True
				Else
					intRtnDup = mobjccMDCATVEXCOM.SelectRtn_TransFlag(gstrConfigXML,strYEARMON)
					If intRtnDup = "Y" Then
					'�ŷ������� ����
					gErrorMsgBox "�ŷ����� �� �����մϴ�." & vbcrlf & "�ش��ڷ��� �ŷ������� ���� �Ͻñ� �ٶ��ϴ�.","����ȳ�!"
					Exit Function
					Else
					'������ True �¿�
					intDelete = mobjccMDCATVEXCOM.DeleteRtn_Medium(gstrConfigXML,strYEARMON)
					if not gDoErrorRtn ("DeleteRtn_Medium") then
					Else
						gErrorMsgBox "�����ͻ��� ERROR!","�����ȳ�!"
						Exit Function
					End If
					SelectRtn_Dup = True
					End If
				End If
			Else
			'�ߺ��� �����Ƿ� ���
			SelectRtn_Dup = True
			End If
		End if
		
	End With
End Function

Sub ProcessRtn ()
	Dim intRtn   'Return ��
   	Dim vntData  'Insert �� ������
   	Dim vntData2
   	Dim intCnt
   	Dim lngAMT
   	Dim lngCOMMI_RATE
   	Dim strCOMMISSION
   	Dim strYEARMON
   	Dim lngREAL_AMT
   	dIM lngBONUS
   	'������ Validation
   	with frmThis
		'���� Rows ����ó��
		For intCnt = 1 to .sprSht.MaxRows
			If mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt) = "" AND mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt) = "" then 
				mobjSCGLSpr.DeleteRow .sprSht,intCnt
			else
				CALL SetTrim (intCnt) ' ���鹮�ڿ� ����
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"AMT",intCnt,0
				End If
				If mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_AMT",intCnt) = ""  Then
					mobjSCGLSpr.SetTextBinding .sprSht,"REAL_AMT",intCnt,0
				End If
				
				If mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt) <> "" AND mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_AMT",intCnt) <> "" _
					AND mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_AMT",intCnt) <> 0 Then
	                
					lngAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)
					lngREAL_AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"REAL_AMT",intCnt)
					'(����ݾ� - û��ݾ�) / û��ݾ�
					lngBONUS = gRound(((lngREAL_AMT - lngAMT) / lngREAL_AMT),2)
	                
					mobjSCGLSpr.SetTextBinding .sprSht,"BONUS",intCnt,lngBONUS
				else
					mobjSCGLSpr.SetTextBinding .sprSht,"BONUS",intCnt,0
				End If
			End If
		Next
		
	
		
		'==================��������
		if DataValidation =false then exit sub
		'Exit SUb
		'==================��������
		 For intCnt = 1 To .sprSht.MaxRows
            lngAMT =  mobjSCGLSpr.GetTextBinding(.sprSht,"AMT",intCnt)  
            strCOMMISSION = gRound((lngAMT * 15 / 100),0)     
            mobjSCGLSpr.SetTextBinding .sprSht,"COMMISSION",intCnt,strCOMMISSION
            mobjSCGLSpr.SetTextBinding .sprSht,"COMMI_RATE",intCnt,15
         Next
   		'==================�ߺ�ó�� By KTH
		IF not SelectRtn_Dup () Then 
			Exit Sub
		End If
		
		strYEARMON = .txtYEARMON.value
		'On error resume next
		'����� �����͸� �����´�.
		vntData = mobjSCGLSpr.GetDataRows(.sprSht, sprSht_DataFields)
 	    if  not IsArray(vntData) then 
		    gErrorMsgBox "����� " & meNO_DATA,"�������"
		    exit sub
        end if
  	    Dim STime, ETime
  	    STime = Time
			intRtn = mobjccMDCATVEXCOM.ProcessRtn(gstrConfigXML, vntData, strYEARMON, mstrTAB_NAME, sprSht_DataFields, vntData_DataType, vntData_DataLength,  false)
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
		mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTCODE",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"MPP",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"PROGNAME",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"CNT",Row,trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CNT",Row))
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDSTDATE",Row,trim(REPLACE(REPLACE(mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",Row),"-",""),".",""))
		mobjSCGLSpr.SetTextBinding .sprSht,"TBRDEDDATE",Row,trim(REPLACE(REPLACE(mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",Row),"-",""),".",""))
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
				vntData = mobjccMDCATVEXCOM.SelectRtn_CLIENTCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
				if not gDoErrorRtn ("SelectRtn_CODE") then
					IF mlngRowCnt <> 1 Then
						strERR = "�������ڵ����"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					END IF
				END IF 
   			Else 
   				strCLIENTNAME = trim(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt))
   				vntData = mobjccMDCATVEXCOM.SelectRtn_CLIENTNAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTNAME)
	   			
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
   		 
   		
   		 '�귣�� �����۾�  SEQNO, CUSTCODE, DEPTCD, CLIENTSUBCODE 
   		 For intCnt = 1 To .sprSht.MaxRows
   			If  mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",intCnt) <> "" Then
   				If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt),1,1) = "A" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt)) = 6 Then
   					strCLIENTCODE = mobjSCGLSpr.GetTextBinding(.sprSht,"CLIENTCODE",intCnt)
   					strSUBSEQ = mobjSCGLSpr.GetTextBinding(.sprSht,"SUBSEQ",intCnt)
   					
   					vntData = mobjccMDCATVEXCOM.SelectRtn_SUBSEQ(gstrConfigXML,mlngRowCnt,mlngColCnt,strCLIENTCODE, strSUBSEQ)
					if not gDoErrorRtn ("SelectRtn_SUBSEQ") then
						IF mlngRowCnt <> 1 Then
							strERR = "�귣���ڵ����"
							mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
							intVal = 1
						else
							IF strCLIENTCODE <> vntData(1,0) then
								strERR = "�ش籤������ �귣���ڵ�Ȯ��"
								mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
								intVal = 1
							ELSE
								strSUBSEQ = vntData(0,0)
								strDEPT_CD = vntData(2,0)
								strCLIENTSUBCODE = vntData(3,0)
								mobjSCGLSpr.SetTextBinding .sprSht,"SUBSEQ",intCnt,strSUBSEQ
								mobjSCGLSpr.SetTextBinding .sprSht,"DEPT_CD",intCnt,strDEPT_CD
								mobjSCGLSpr.SetTextBinding .sprSht,"CLIENTSUBCODE",intCnt,strCLIENTSUBCODE
							END IF
						END IF
					END IF 
				end if
   			Else 
				strERR = "�귣���ڵ����"
				mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
				intVal = 1
			End If
   		 Next
   		 
   		 'mpp�ڵ�üũ
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",intCnt),1,1) = "P" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",intCnt)) = 6 Then
   			Else 
   				if mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",intCnt) <> "" then
   				
   					strMPPNAME = mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",intCnt)
   					vntData = mobjccMDCATVEXCOM.SelectRtn_MPPNAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strMPPNAME)
					if not gDoErrorRtn ("SelectRtn_MPPNAME") then
						If mlngRowCnt = 1 Then
							strMPPCODE = vntData(0,0)
							mobjSCGLSpr.SetTextBinding .sprSht,"MPP",intCnt,strMPPCODE
						Else
							mobjSCGLSpr.SetTextBinding .sprSht,"MPP",intCnt,""
							'strERR = "MPP(��)�ڵ����"
							'mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
							'intVal = 1
						End If
					End If
				end if
			End If
   		 Next 		
   		   
   		 'ä���ڵ�üũ
   		 For intCnt = 1 To .sprSht.MaxRows
   			If MID(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt),1,1) = "B" AND LEN(mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt)) = 6 Then
   			Else 
   				strMEDCODENAME = mobjSCGLSpr.GetTextBinding(.sprSht,"MEDCODE",intCnt)
   				vntData = mobjccMDCATVEXCOM.SelectRtn_MEDCODENAME(gstrConfigXML,mlngRowCnt,mlngColCnt,strMEDCODENAME)
				if not gDoErrorRtn ("SelectRtn_MEDCODENAME") then
					If mlngRowCnt = 1 Then
						strMEDCODE = vntData(0,0)
						
						mobjSCGLSpr.SetTextBinding .sprSht,"MEDCODE",intCnt,strMEDCODE
						
						if mobjSCGLSpr.GetTextBinding(.sprSht,"MPP",intCnt) = "" then
							mobjSCGLSpr.SetTextBinding .sprSht,"MPP",intCnt,vntData(1,0)
						end if
						vntData2 = mobjccMDCATVEXCOM.SelectRtn_REALMEDCODE(gstrConfigXML,mlngRowCnt,mlngColCnt,strMEDCODE)
						strREALMEDCODE = vntData2(0,0)
						mobjSCGLSpr.SetTextBinding .sprSht,"REAL_MED_CODE",intCnt,strREALMEDCODE
						
					Else
						strERR = "ä���ڵ����"
						mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
						intVal = 1
					End If
				End If
			End If
   		 Next
   		 '���α׷�üũ(�ڸ���,�̱������̼�)
   		 For intCnt = 1 To .sprSht.MaxRows
                If mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",intCnt) <> "" Then
                    If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",intCnt)) < 255 Then
                    mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",intCnt),"'","") 
                    mobjSCGLSpr.SetTextBinding .sprSht,"PROGRAM",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGRAM",intCnt),",","") 
                    Else
                        strERR = "���α׷����ڱ���"
                        mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                        intVal = 1
                    End If
                End If
         Next
   		 '�����üũ(�ڸ���,�̱������̼�)
   		 For intCnt = 1 To .sprSht.MaxRows
                If mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",intCnt) <> "" Then
                    If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",intCnt)) < 255 Then
                    mobjSCGLSpr.SetTextBinding .sprSht,"PROGNAME",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",intCnt),"'","") 
                    mobjSCGLSpr.SetTextBinding .sprSht,"PROGNAME",intCnt,Replace(mobjSCGLSpr.GetTextBinding(.sprSht,"PROGNAME",intCnt),",","") 
                    Else
                        strERR = "�������ڱ���"
                        mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                        intVal = 1
                    End If
                End If
         Next
         
         '��۽����ϱ���üũ
   		 For intCnt = 1 To .sprSht.MaxRows
                If mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",intCnt) <> "" Then
                    If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDSTDATE",intCnt)) <> 8 Then
                        strERR = "��۽����ϱ���"
                        mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                        intVal = 1
                    End If
                End If
         Next
         
         '��������ϱ���üũ
         For intCnt = 1 To .sprSht.MaxRows
                If mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",intCnt) <> "" Then
                    If Len(mobjSCGLSpr.GetTextBinding(.sprSht,"TBRDEDDATE",intCnt)) <> 8 Then
                        strERR = "��������ϱ���"
                        mobjSCGLSpr.SetTextBinding .sprSht,"ATTR01",intCnt,strERR
                        intVal = 1
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
        mlngRowCnt=Clng(0): mlngColCnt=Clng(0)
        vntData = mobjccMDCATVEXCOM.getTABCOLINFO(gstrConfigXml,mlngRowCnt,mlngColCnt,mdblTAB_ID)

        sprSht_DataFields    = mChangeData (vntData,2,"|")
        vntData_DataFields   = gArray2Single(vntData,1)	  
        
        sprSht_DisplayFields = mChangeData (vntData,1,"|")
         
        sprSht_DefualtValueFields = mDefaultValueField(vntData,2,4,"|")

        vntData_DefaultValue = gArray2Single (vntdata,4)
        vntData_DataType     = gArray2Single (vntData,5)
        vntData_DataLength   = gArray2Single (vntData,6)
        vntData_Nullable	 = gArray2Single (vntData,8)

        sprSht_NotNull = fNotNull(vntData_Nullable, vntData_DataFields, mlngRowCnt)
        
        gSetSheetDefaultColor() 
        gSetSheetColor mobjSCGLSpr,     .sprSht
        mobjSCGLSpr.SpreadLayout        .sprSht, mlngRowCnt, 0
        mobjSCGLSpr.SpreadDataField     .sprSht, sprSht_DataFields
        mobjSCGLSpr.SetHeader           .sprSht, sprSht_DisplayFields
        mobjSCGLSpr.SetCellTypeEdit2    .sprSht, sprSht_DataFields, , ,200
        mobjSCGLSpr.SetCellTypeFloat2   .sprSht, "AMT|REAL_AMT", -1, -1, 0
        mobjSCGLSpr.SetCellsLock2		.sprSht, true, sprSht_DefualtValueFields
        mobjSCGLSpr.SetRowHeight        .sprSht, "-1", "13"
        mobjSCGLSpr.SetRowHeight        .sprSht, "0", "13"        
        mobjSCGLSpr.SetColWidth         .sprSht, "-1", 11
        mobjSCGLSpr.ColHidden .sprSht,"REAL_MED_CODE|CLIENTSUBCODE|DEPT_CD|COMMI_RATE|COMMISSION|TRU_TAX_FLAG|COMMI_TAX_FLAG",true
    End With
End Sub

Function mChangeData(vntData, colidx, splitMark)
    Dim strRtn, lngRowCnt, i
     
    lngRowCnt=-1: lngRowCnt=ubound(vntData,2)
    for i = 0 to lngRowCnt
      if i=lngRowCnt then splitMark = ""
        strRtn = strRtn & vntData(colidx,i) & splitMark
    Next
    mChangeData = strRtn
End Function

Function mDefaultValueField(vntData, colidx, CheckColidx, splitMark)
	Dim strRtn, lngRowCnt, i
	
	lngRowCnt=-1: lngRowCnt=Ubound(vntData,2)
	
	for i=0 to lngRowCnt
		if i=lngRowCnt then splitMark = ""
		if vntData(CheckColidx, i) <> "" then
			strRtn = strRtn & vntData(colidx,i) & splitMark
		end if
	next
	mDefaultValueField = strRtn
End Function

Function mHaveID()
	mHaveID = false
	Dim i, intRowCnt, temp
	temp = "ID"
	intRowCnt = Ubound(vntData_DefaultValue,1)
	For i=0 To intRowCnt
	  if temp = vntData_DefaultValue(i) then 
	     mHaveID = true
	     exit Function		
	  end if	
	Next
End Function

Function fNotNull(vntData_Nullable, vntData_DataFields, intRows)
	Dim i, sprSht_NotNull
	sprSht_NotNull = ""
	For i=0 to intRows-1
		if vntData_Nullable(i) = "N" then
		  sprSht_NotNull = sprSht_NotNull & vntData_DataFields(i) & "|"
		end if  
	Next
	fNotNull = sprSht_NotNull
End Function
-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%">
				<TR>
					<TD>
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gif"
							border="0">
							<TR>
								<td style="WIDTH: 400px" align="left" width="400" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
											<td align="left" height="4"><FONT face="����"></FONT></td>
										</tr>
										<tr>
											<td class="TITLE" id="tblTitleName"><FONT face="����">&nbsp;���̺� û����� (�ϰ� û��)</FONT></td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
										<TR>
											<TD class="" id="tblWait" style="Z-INDEX: 200"><IMG id="imgWaiting" style="CURSOR: wait" height="23" alt="ó�����Դϴ�." src="../../../images/Waiting.GIF"
													border="0" name="imgWaiting">
											</TD>
										</TR>
									</TABLE>
									<TABLE id="tblButton" style="WIDTH: 114px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										width="114" border="0">
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
										</TR>
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD class="TOPSPLIT" style="HEIGHT: 17px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="LABEL" style="WIDTH: 90px" onclick="vbscript:Call gCleanField(txtYEARMON, '')">���</TD>
											<TD class="DATA"><INPUT class="INPUT" id="txtYEARMON" title="�ش�⵵" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM,M"
													type="text" maxLength="6" size="9" name="txtYEARMON"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</TD>
							</TR>
							<TR>
								<TD class="BODYSPLIT"><FONT face="����"></FONT></TD>
							</TR>
							<tr>
								<td>
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 750px" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5"
										VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="27464">
										<PARAM NAME="_ExtentY" VALUE="16616">
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
								</td>
							</tr>
						</TABLE>
					</TD>
				</TR>
			</TABLE>
		</form>
	</body>
</HTML>
