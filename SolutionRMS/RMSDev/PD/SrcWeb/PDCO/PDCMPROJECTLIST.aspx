<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMPROJECTLIST.aspx.vb" Inherits="PD.PDCMPROJECTLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB ���� ��Ȳ ��ȸ</title> 
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
Dim mobjPDCMGET
Dim mInsOKFlag 'Insert Flag 
Dim mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode '�˾�����
Dim mobjPONOLIST
    
'=============================
' �̺�Ʈ���ν��� 
'=============================
Sub imgQuery_onclick
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

Sub imgExcel1_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht1
	end with
	gFlowWait meWAIT_OFF
End Sub

Sub imgDetail_onclick()
	Dim strJOBNO, strPRONO
	Dim vntInParams
	Dim vntRet
	Dim strRow, strCol
	with frmThis
		IF .sprSht1.MaxRows >0 then
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow)
			strPRONO = mobjSCGLSpr.GetTextBinding( .sprSht1,"PROJECTNO",.sprSht1.ActiveRow)
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",.sprSht1.ActiveRow),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow))
			vntRet = gShowModalWindow("PDCMESTDTLSRC.aspx",vntInParams , 1060,780)
			strRow = .sprSht1.ActiveRow
			strCol = .sprSht1.ActiveCol
			'���⼭ ���� ���� ���� ȭ�� ȣ��
			.txtCLIENTCODE1.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht1.Focus
			
			SelectRtn_DBLHDR(strPRONO)
			mobjSCGLSpr.ActiveCell .sprSht1, strCol, strRow		
		end if	
	end with
End Sub

Sub window_onload
    Initpage()
End Sub

Sub Window_OnUnload()
    EndPage()
End Sub

Sub InitPage()
    '����������ü ����	
    set mobjPDCMGET = gCreateRemoteObject("cPDCO.ccPDCOGET")
    set mobjPONOLIST	= gCreateRemoteObject("cPDCO.ccPDCOPONOLIST")

   '���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	mInsOKFlag   =  false
	
	gSetSheetDefaultColor() 
	with frmThis
		'������Ʈ ����Ʈ ��Ʈ����
		gSetSheetColor mobjSCGLSpr, .sprSht
		mobjSCGLSpr.SpreadLayout .sprSht, 15, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPDEPTNAME|CPEMPNO|CPEMPNAME|MEMO"
		mobjSCGLSpr.SetHeader .sprSht,		"������Ʈ�ڵ�|������Ʈ��|�������ڵ�|������|������ڵ�|�����|�귣���ڵ�|�귣��|�׷챸��|�����|�μ��ڵ�|���μ�|���|�����|���"
		mobjSCGLSpr.SetColWidth .sprSht, "-1","         0|        16|         0|    20|         0|    20|         0|    18|      10|     8|       0|      10|   0|    10| 10"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		'mobjSCGLSpr.SetCellTypeCheckBox2 .sprSht, "CHK"
		'mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "MCCOMMI_RATE", -1, -1, 2
		mobjSCGLSpr.SetCellTypeDate2 .sprSht, "CREDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht, true, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CPDEPTCD|CPEMPNO|MEMO"
		mobjSCGLSpr.ColHidden .sprSht, "PROJECTNO|CLIENTCODE|CLIENTSUBCODE|SUBSEQ", true
		mobjSCGLSpr.SetCellAlign2 .sprSht, "PROJECTNO|PROJECTNM|CLIENTCODE|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|GROUPGBN|CREDAY|CPDEPTCD|CPDEPTNAME|CPEMPNO|CPEMPNAME|MEMO",-1,-1,0,2,false
        
        
       
        'job����Ʈ ��Ʈ����
        gSetSheetColor mobjSCGLSpr, .sprSht1
		mobjSCGLSpr.SpreadLayout .sprSht1, 16, 0, 0, 0,2
		mobjSCGLSpr.SpreadDataField .sprSht1, "PROJECTNO|JOBNO|JOBNAME|CLIENTNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|ENDFLAG|JOBGUBN|CREPART|CREGUBN|REQDAY|COMMITION|CLIENTCODE|PREESTNO"
		mobjSCGLSpr.SetHeader .sprSht1,		   "������Ʈ��ȣ|Job No|���۰Ǹ�|������|�����|����θ�|�귣��|�귣���|����|��ü�κ�|��ü�з�|�ű�|�ۼ���|��������|�������ڵ�|Ȯ�������ڵ�"
		mobjSCGLSpr.SetColWidth .sprSht1, "-1", "          0|     7|      19|13    |6     |12      |6     |13      |   6|12      |12      |5   |10    |0       |0         |0"
		mobjSCGLSpr.SetRowHeight .sprSht1, "-1", "13"
		mobjSCGLSpr.SetRowHeight .sprSht1, "0", "15"
		mobjSCGLSpr.SetCellTypeDate2 .sprSht1, "REQDAY", -1, -1, 10
		mobjSCGLSpr.SetCellsLock2 .sprSht1, true, "PROJECTNO|JOBNO|JOBNAME|CLIENTSUBCODE|CLIENTSUBNAME|SUBSEQ|SUBSEQNAME|JOBGUBN|CREPART|CREGUBN|REQDAY|ENDFLAG|CLIENTNAME|PREESTNO"
		mobjSCGLSpr.ColHidden .sprSht1, "PROJECTNO|COMMITION|CLIENTCODE|PREESTNO", true
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "JOBNAME|CLIENTSUBNAME|SUBSEQNAME|CLIENTNAME",-1,-1,0,2,false
		mobjSCGLSpr.SetCellAlign2 .sprSht1, "CLIENTSUBCODE|SUBSEQ|JOBGUBN|CREPART|CREGUBN|JOBNO|ENDFLAG",-1,-1,2,2,false
        
        		
	end with
	InitPageData
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
		DateClean
	End with
	'���ο� XML ���ε��� ����
	gXMLNewBinding frmThis,xmlBind,"#xmlBind"	
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

Sub EndPage()
	set mobjPDCMGET = Nothing
	set mobjPONOLIST = Nothing
	gEndPage
End Sub

'=============================
' ��ɹ�ưŬ���̺�Ʈ
'=============================
Sub imgClose_onclick()
    Window_OnUnload()
End Sub

'-----------------------------------------------------------------------------------------
' JOB �˾� ��ư[��ȸ��]
'-----------------------------------------------------------------------------------------
'�̹�����ư Ŭ����
Sub ImgCLIENTCODE1_onclick
	Call CLIENTCODE1_POP()
End Sub

'���� ������List ��������
Sub CLIENTCODE1_POP
	dim vntRet
	Dim vntInParams
	with frmThis
		vntInParams = array(trim(.txtCLIENTCODE1.value), trim(.txtCLIENTNAME1.value))
		vntRet = gShowModalWindow("PDCMCUSTPOP.aspx",vntInParams , 413,435)
		
		if isArray(vntRet) then
			if .txtCLIENTCODE1.value = vntRet(0,0) and .txtCLIENTNAME1.value = vntRet(1,0) then exit Sub ' ����� �����Ͱ� ���ٸ� exit
			.txtCLIENTCODE1.value = trim(vntRet(0,0))       ' Code�� ����
			.txtCLIENTNAME1.value = trim(vntRet(1,0))       ' �ڵ�� ǥ��
			gSetChangeFlag .txtCLIENTCODE1                  ' gSetChangeFlag objectID	 Flag ���� �˸�
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
			
			vntData = mobjPDCMGET.GetCUSTNO(gstrConfigXml,mlngRowCnt,mlngColCnt,trim(.txtCLIENTCODE1.value),trim(.txtCLIENTNAME1.value))
			
			if not gDoErrorRtn ("txtCLIENTNAME1_onkeydown") then
				If mlngRowCnt = 1 Then
					.txtCLIENTCODE1.value = trim(vntData(0,0))
					.txtCLIENTNAME1.value = trim(vntData(1,0))
				Else
					Call CLIENTCODE1_POP()
				End If
   			end if
   		end with
		window.event.returnValue = false
		window.event.cancelBubble = true
	end if
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

'****************************************************************************************
' ������� �޷�
'****************************************************************************************
'��ȸ��
Sub imgCalEndarFROM1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtFROM,frmThis.imgCalEndarFROM1,"txtFROM_onchange()"
		gSetChange
	end with
End Sub

Sub imgCalEndarTO1_onclick
	WITH frmThis
		'CalEndar�� ȭ�鿡 ǥ��
		gShowPopupCalEndar frmThis.txtTo,frmThis.imgCalEndarTO1,"txtTo_onchange()"
		gSetChange
	end with
End Sub


Sub txtFROM_onchange
	gSetChange
End Sub


Sub txtTo_onchange
	gSetChange
End Sub

'=============================
'SheetEvent
'=============================
sub sprSht1_DblClick (ByVal Col, ByVal Row)
	Dim strJOBNO, strPRONO
	Dim vntInParams
	Dim vntRet
	Dim strRow
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht1, ""
		Else
			strJOBNO = mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",.sprSht1.ActiveRow)
			strPRONO = mobjSCGLSpr.GetTextBinding( .sprSht1,"PROJECTNO",.sprSht1.ActiveRow)
			
			vntInParams = array(mobjSCGLSpr.GetTextBinding( .sprSht1,"PREESTNO",Row),mobjSCGLSpr.GetTextBinding( .sprSht1,"JOBNO",Row))
			vntRet = gShowModalWindow("PDCMESTDTLSRC.aspx",vntInParams , 1060,780)
			strRow = Row
			'���⼭ ���� ���� ���� ȭ�� ȣ��
			.txtCLIENTCODE1.focus()	'�˾�â�� ���� ���鼭 �Ҿ���� ��Ŀ���� �ٽ� ��Ʈ�� �Ű��ش�
			.sprSht1.Focus
			
			SelectRtn_DBLHDR(strPRONO)
			mobjSCGLSpr.ActiveCell .sprSht1, Col, strRow			
		end if
	end with
end sub

Sub sprSht_Click(ByVal Col, ByVal Row)
	Dim strPROJECTNO	
	Dim vntInParams
	Dim vntRet
	Dim strRow
	with frmThis
		strPROJECTNO = mobjSCGLSpr.GetTextBinding( .sprSht,"PROJECTNO",.sprSht.ActiveRow)
		SelectRtn_DBLHDR(strPROJECTNO)
	End with
End Sub
'------------------------------------------
' ������ ��ȸ
'------------------------------------------
Sub SelectRtn ()
	Dim vntData
	Dim strYEARMON, strREAL_MED_CODE
	Dim strFROM,strTO
	Dim strTAXNO
   	Dim i, strCols
   	
	'On error resume next
	with frmThis
		'Sheet�ʱ�ȭ
		.sprSht.MaxRows = 0
		
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strFROM = MID(.txtFROM.value,1,4) &  MID(.txtFROM.value,6,2) &  MID(.txtFROM.value,9,2)
		strTO =  MID(.txtTO.value,1,4) &  MID(.txtTO.value,6,2) &  MID(.txtTO.value,9,2)
		
		'���ݰ�꼭 �Ϸ���ȸ
		vntData = mobjPONOLIST.SelectRtn(gstrConfigXml,mlngRowCnt,mlngColCnt,strFROM,strTO,Trim(.txtPROJECTNM1.value),Trim(.txtPROJECTNO1.value),Trim(.txtCLIENTNAME1.value),Trim(.txtCLIENTCODE1.value))
		If not gDoErrorRtn ("SelectRtn") then
			'��ȸ�� �����͸� ���ε�
			call mobjSCGLSpr.SetClipBinding (frmThis.sprSht,vntData,1,1,mlngColCnt,mlngRowCnt,True)
			'�ʱ� ���·� ����
			mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
			If mlngRowCnt < 1 Then
				.sprSht.MaxRows = 0	
			else
				Call sprSht_Click(1,1)
			End If
			gWriteText lblstatus, "������ �ڷῡ ���ؼ� " & mlngRowCnt & " ���� �ڷᰡ �˻�" & mePROC_DONE			
		End If		
	END WITH
	'��ȸ�Ϸ�޼���
	gWriteText "", "�ڷᰡ �˻�" & mePROC_DONE
End Sub


Sub SelectRtn_DBLHDR (ByVal strPONO)
	Dim vntData
	Dim strFROM,strTO
	Dim intCnt
	'on error resume next
	with frmThis
		mlngRowCnt=clng(0): mlngColCnt=clng(0)
		
		vntData = mobjPONOLIST.SelectRtn_JOB(gstrConfigXml,mlngRowCnt,mlngColCnt, strPONO)
		
		If not gDoErrorRtn ("SelectRtn_JOB") then
			If mlngRowCnt > 0 Then
				call mobjSCGLSpr.SetClipBinding (frmThis.sprSht1,vntData,1,1,mlngColCnt,mlngRowCnt,True)
				mobjSCGLSpr.SetFlag  frmThis.sprSht1,meCLS_FLAG
			ELSE
				.sprSht1.MaxRows = 0
			END IF
		END IF
	End with
End SUB

-->
		</SCRIPT>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<form id="frmThis">
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="100%" height="100%">
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
											<td class="TITLE" id="tblTitleName"><FONT face="����">&nbsp;���� ����</FONT></td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" width="640" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 350px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
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
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0" height="100%">
							<TR>
								<TD class="TOPSPLIT" style="HEIGHT: 10px"></TD>
							</TR>
							<TR>
								<TD class="KEYFRAME" vAlign="middle" align="center">
									<TABLE class="DATA" id="tblKey0" cellSpacing="1" cellPadding="0" width="100%" border="0">
										<TR>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call DateClean()"
												width="80">
												�����</TD>
											<TD class="SEARCHDATA" width="230"><INPUT class="INPUT" id="txtFROM" title="�Ⱓ�˻�(FROM)" style="WIDTH: 80px; HEIGHT: 22px"
													accessKey="DATE" type="text" maxLength="10" size="6" name="txtFROM"><IMG id="imgCalEndarFROM1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF" width="23" align="absMiddle"
													border="0" name="imgCalEndarFROM1">&nbsp;~ <INPUT class="INPUT" id="txtTO" title="�Ⱓ�˻�(TO)" style="WIDTH: 80px; HEIGHT: 22px" accessKey="DATE"
													type="text" maxLength="10" size="7" name="txtTO"><IMG id="imgCalEndarTO1" onmouseover="JavaScript:this.src='../../../images/imgCalEndarOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgCalEndar.gIF'" height="20" src="../../../images/imgCalEndar.gIF"
													width="23" align="absMiddle" border="0" name="imgCalEndarTO1"></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtCLIENTCODE1, txtCLIENTNAME1)"
												width="80"><FONT face="����">������</FONT></TD>
											<TD class="SEARCHDATA" width="260"><FONT face="����"><FONT face="����"><INPUT class="INPUT_L" id="txtCLIENTNAME1" title="�ڵ��" style="WIDTH: 179px; HEIGHT: 22px"
															type="text" maxLength="100" size="24" name="txtCLIENTNAME1"></FONT><IMG id="ImgCLIENTCODE1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgCLIENTCODE1"><INPUT class="INPUT" id="txtCLIENTCODE1" title="�ڵ��Է�" style="WIDTH: 56px; HEIGHT: 22px"
														type="text" maxLength="6" size="4" name="txtCLIENTCODE1"></FONT></TD>
											<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtPROJECTNO1, txtPROJECTNM1)"
												width="80">������Ʈ��</TD>
											<TD class="SEARCHDATA"><FONT face="����"><INPUT class="INPUT_L" id="txtPROJECTNM1" title="�ڵ��" style="WIDTH: 144px; HEIGHT: 22px"
														type="text" maxLength="100" size="18" name="txtPROJECTNM1"><IMG id="ImgPROJECTNO1" onmouseover="JavaScript:this.src='../../../images/imgPopupOn.gIF'"
														style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgPopup.gIF'" height="20" src="../../../images/imgPopup.gIF" width="23" align="absMiddle"
														border="0" name="ImgPROJECTNO1"><INPUT class="INPUT" id="txtPROJECTNO1" title="�ڵ�" style="WIDTH: 56px; HEIGHT: 22px" type="text"
														maxLength="6" size="4" name="txtPROJECTNO1"></FONT></TD>
											<td class="SEARCHDATA"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20" alt="�ڷḦ �˻��մϴ�."
													src="../../../images/imgQuery.gIF" align="right" border="0" name="imgQuery"></td>
										</TR>
									</TABLE>
									<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"><FONT face="����"></FONT></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;������Ʈ ����Ʈ</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></TD>
													</TR>
												</TABLE>
												<!--Common Button End--></TD>
										</TR>
									</TABLE>
								</TD>
							<!--BodySplit Start-->
							<TR>		
								<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
							</TR>
							<TR>
								<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
									<DIV id="pnlTab1" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
									<OBJECT id="sprSht" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="27464">
										<PARAM NAME="_ExtentY" VALUE="8467">
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
									<!--/DIV--></TD>
							</TR>
							<TR>
								<TD>
									<TABLE height="13" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0">
										<TR>
											<TD class="TOPSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
									<TABLE height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
										border="0"> <!--background="../../../images/TitleBG.gIF"-->
										<TR>
											<TD align="left" height="20">
												<table cellSpacing="0" cellPadding="0" width="100%" border="0">
													<tr>
														<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIF" width="14"></td>
														<td align="left" height="4"><FONT face="����"></FONT></td>
													</tr>
													<tr>
														<td class="TITLE">&nbsp;JOB ����Ʈ</td>
													</tr>
												</table>
											</TD>
											<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
												<!--Common Button Start-->
												<TABLE id="tblButton1" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
													<TR>
														<TD><IMG id="imgDetail" onmouseover="JavaScript:this.src='../../../images/imgDetailOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDetail.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgDetail.gIF" border="0" name="imgDetail"></TD>
														<TD><IMG id="imgExcel1" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel1"></TD>
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
							<!--BodySplit End-->
							<!--List Start-->
							<TR>
									<TD class="LISTFRAME" style="WIDTH: 100%; HEIGHT: 50%" vAlign="top" align="left">
									<DIV id="pnlTab2" style="VISIBILITY: visible; WIDTH: 100%; HEIGHT: 95%; POSITION: relative" ms_positioning="GridLayout">
									<!--DIV id="pnlTab2" style="VISIBILITY: hidden; WIDTH: 1038px; POSITION: relative" ms_positioning="GridLayout"-->
									<OBJECT id="sprSht1" style="WIDTH: 100%; HEIGHT: 95%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" VIEWASTEXT>
										<PARAM NAME="_Version" VALUE="393216">
										<PARAM NAME="_ExtentX" VALUE="27464">
										<PARAM NAME="_ExtentY" VALUE="8467">
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
									<!--/DIV--></TD>
							</TR>
							<!--BodySplit Start-->
							<TR>
								<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 13px"><FONT face="����"></FONT></TD>
							</TR>
							<!--BodySplit End-->
							<!--Bottom Split Start-->
							<TR>
								<TD class="BOTTOMSPLIT" id="lblStatus" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
							</TR>
							<!--Bottom Split End--></TABLE>
						<!--Input Define Table End--></TD>
				</TR>
				<!--Top TR End--></TABLE>
			<!--Main End--></form>
		</TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></TR></TBODY></TABLE></FORM>
	</body>
</HTML>
