<%@ Page Language="vb" AutoEventWireup="false" Codebehind="MDCMCOMMITLIST.aspx.vb" Inherits="MD.MDCMCOMMITLIST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>����Ʈ �����ڷ� ���� �� ��ȸ</title>
		<META http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<!--
'****************************************************************************************
'�ý��۱��� : �μ��ü
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMTRANSCONF.aspx
'��      �� : �ۼ��� �ŷ����� �� Confirm �� �Ѵ�.
'�Ķ�  ���� : 
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2008/08/29 By Kim Tae Ho
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
Dim mobjMDSRREPORTLIST
'=========================================================================================
' �̺�Ʈ ���ν��� 
'=========================================================================================
Sub window_onload
	Initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgClose_onclick()
	EndPage
End Sub

Sub imgQuery_Onclick()
	gFlowWait meWAIT_ON
	SelectRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgDelete_onclick
	gFlowWait meWAIT_ON
	DeleteRtn
	gFlowWait meWAIT_OFF
End Sub


Sub imgExcel_onclick ()
	gFlowWait meWAIT_ON
	with frmThis
		mobjSCGLSpr.ExportExcelFile .sprSht
	end with
	gFlowWait meWAIT_OFF
End Sub


Sub EndPage()
	set mobjMDSRREPORTLIST = Nothing
	gEndPage
End Sub

'=========================================================================================
' UI���� ���ν��� 
'=========================================================================================
'-----------------------------------------------------------------------------------------
' ������ ȭ�� ������ �� �ʱ�ȭ 
'-----------------------------------------------------------------------------------------
Sub InitPage()

	'����������ü ����	
	set mobjMDSRREPORTLIST	= gCreateRemoteObject("cMDSC.ccMDSCREPORTLIST") '��ȸ

	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"
	
	'�� ��ġ ���� �� �ʱ�ȭ
	pnlTab1.style.position = "absolute"
	pnlTab1.style.top = "225px"
	'pnlTab1.style.height ="300px"
	pnlTab1.style.left= "7px"
	
	mobjSCGLCtl.DoEventQueue
	
    'Sheet �⺻Color ����
    gSetSheetDefaultColor
    with frmThis
		'Sheet Į�� ����
	    gSetSheetColor mobjSCGLSpr, .sprSht
		
		'Sheet Layout ������
		mobjSCGLSpr.SpreadLayout .sprSht, 21, 0,6
		'YEARMON|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|AMT
	    mobjSCGLSpr.SpreadDataField .sprSht, "YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP"
		mobjSCGLSpr.SetHeader .sprSht,        "���|��ǥ����|������|��ü��|��ü���|�����|�귣��|��ü����|������ڵ�|��޾�|�ΰ���|�����|�������ڵ�|��ü�ڵ�|��ü���ڵ�|������ڵ�|�귣���ڵ�|��ü�����ڵ�|��ǥ�����ڵ�|�ΰ�������|MPP",0,1,true
		mobjSCGLSpr.SetColWidth .sprSht, "-1","   7|       9|    15|    15|      15|    15|    15|       7|         0|    11|    11|    10|         0|       0|         0|         0|         0|           0|           0|         0|0"
		mobjSCGLSpr.SetCellTypeFloat2 .sprSht, "AMT|VAT", -1, -1, 0
		mobjSCGLSpr.SetCellsLock2 .sprSht,true,"YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP"
		mobjSCGLSpr.SetRowHeight .sprSht, "0", "15"
		mobjSCGLSpr.SetRowHeight .sprSht, "-1", "13"
		mobjSCGLSpr.SetCellAlign2 .sprSht, "YEARMON|VOCH_GBNNAME|MEDFLAGNAME",-1,-1,2,2,false
		mobjSCGLSpr.ColHidden .sprSht, "CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|EXCLIENTCODE|MPP", true
	End with

	pnlTab1.style.visibility = "visible" 
	
	'ȭ�� �ʱⰪ ����
	InitPageData	
End Sub

Sub SelectRtn()
	Dim vntData
	Dim i, strCols
	Dim strYEARMON
	Dim strVOCH_GBN
	Dim intCnt
	with frmThis
	'ON ERROR RESUME NEXT
		.sprSht.MaxRows = 0
		
		'���� ���ʵ����� ������� state�� �����ش�.
		SelectRtn_STATECHK
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = .txtYEARMON.value
		
		IF .rdT.checked = TRUE THEN
			strVOCH_GBN = "VOCH"
		ELSE
			strVOCH_GBN = "NOVOCH"
		END IF
		
		IF .cmbCOMMITGBN.value = 0 THEN
			vntData = mobjMDSRREPORTLIST.SelectRtn_LOW(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strVOCH_GBN)

			if not gDoErrorRtn ("SelectRtn_LOW") then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE			
   			end if
		ELSE
			vntData = mobjMDSRREPORTLIST.SelectRtn_REPORT(gstrConfigXml,mlngRowCnt,mlngColCnt,strYEARMON,strVOCH_GBN)

			if not gDoErrorRtn ("SelectRtn_REPORT") then
				mobjSCGLSpr.SetClipBinding .sprSht, vntData, 1, 1, mlngColCnt, mlngRowCnt, True
				mobjSCGLSpr.SetFlag  frmThis.sprSht,meCLS_FLAG
					
   				gWriteText lblStatus, mlngRowCnt & "���� �ڷᰡ �˻�" & mePROC_DONE
   			end if
		END IF
		
   		AMT_SUM
	End With
End Sub


Sub SelectRtn_STATECHK()
	Dim vntData
	Dim i, strCols
	Dim strYEAR
	Dim strVOCH_GBN
	Dim intCnt
	with frmThis
	'ON ERROR RESUME NEXT
		'Long Type�� ByRef ������ �ʱ�ȭ
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEAR = MID(.txtYEARMON.value,1,4)
		
		vntData1 = mobjMDSRREPORTLIST.SelectRtn_STATECHK(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEAR)
		
		If mlngRowCnt > 0 then
			for i = 1 to mlngRowCnt
				if vntData1(1,i) <> "" then
					if vntData1(1,i) = "VOCH" THEN
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = true
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
						document.getElementById("txtVOCH" & vntData1(0,i)).value = vntData1(2,i)
					elseif vntData1(1,i) = "NOVOCH" then
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = true
						document.getElementById("txtVOCH" & vntData1(0,i)).value = vntData1(2,i)
					else
						document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
						document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
						document.getElementById("txtVOCH" & vntData1(0,i)).value = ""
					end if
				else
					document.getElementById("rdVOCH" & vntData1(0,i)).checked = false
					document.getElementById("rdNOVOCH" & vntData1(0,i)).checked = false
					document.getElementById("txtVOCH" & vntData1(0,i)).value = ""
				end if
			next
		else
			for i = 1 to 12
				if i < 10 then
					document.getElementById("rdVOCH0" & i).checked = false
					document.getElementById("rdNOVOCH0" & i).checked = false
					document.getElementById("txtVOCH0" & i).value = ""
				else
					document.getElementById("rdVOCH" & i).checked = false
					document.getElementById("rdNOVOCH" & i).checked = false
					document.getElementById("txtVOCH" & i).value = ""
				end if
			next
		END IF
	
	End With
End Sub


'-----------------------------------------------------------------------------------------
' ȭ�� ó�� SCRIPT
'-----------------------------------------------------------------------------------------
sub sprSht_DblClick (ByVal Col, ByVal Row)
	Dim vntRet
	Dim vntInParams
	Dim strTRANSYEARMON
	Dim strTRANSNO
	
	with frmThis
		if Row = 0 and Col >1 then
			mobjSCGLSpr.SetSheetSortUser  .sprSht, ""
		end if
	end with
end sub

Sub sprSht_Change(ByVal Col, ByVal Row)
	'���� �÷��� ����
	mobjSCGLSpr.CellChanged frmThis.sprSht, Col, Row
End Sub

'-----------------------------------------------------------------------------------------
' Field üũ
'-----------------------------------------------------------------------------------------
'------------------------------------------
' ��������
'------------------------------------------
Sub DeleteRtn ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strYEARMON
	with frmThis
		if .sprSht.MaxRows = 0 then
			gErrorMsgBox "������ �����Ͱ� �����ϴ�.","ó���ȳ�!"
			Exit Sub
		end if
		
		if .cmbCOMMITGBN.value <> 1 then
			gErrorMsgBox "�̻����� �ڷ�� ������ �� �����ϴ�.","ó���ȳ�!"
			Exit Sub
		end if
		
		
		
		intRtn = gYesNoMsgbox("�ڷḦ �����Ͻðڽ��ϱ�?","�ڷ���� Ȯ��")
		IF intRtn <> vbYes then exit Sub
		intCnt = 0
		
		'���õ� �ڷḦ ������ ���� ����
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
	
		intRtn = mobjMDSRREPORTLIST.DeleteRtn(gstrConfigXml,strYEARMON)
		IF not gDoErrorRtn ("DeleteRtn") then
			gErrorMsgBox "�ڷᰡ �����Ǿ����ϴ�.","�����ȳ�!"
			.sprSht.MaxRows = 0
   		End IF
   		
		'���� ���� ����
		mobjSCGLSpr.DeselectBlock .sprSht
		SelectRtn_STATECHK
		'SelectRtn
	End with
	err.clear	
End Sub

Sub DeleteRtn_process ()
	Dim vntData
	Dim intCnt, intRtn, i
	Dim intCnt2
	Dim strYEARMON
	with frmThis
		
		'���õ� �ڷḦ ������ ���� ����
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
	
		intRtn = mobjMDSRREPORTLIST.DeleteRtn(gstrConfigXml,strYEARMON)
		IF not gDoErrorRtn ("DeleteRtn") then
			If strDESCRIPTION <> "" Then
				gErrorMsgBox strDESCRIPTION,"�����ȳ�!"
				Exit Sub
			End If
   		End IF
		
		IF not gDoErrorRtn ("DeleteRtn") then
			gWriteText "", intCnt & "���� ����" & mePROC_DONE
   		End IF
   		
	End with
	err.clear	
End Sub
'-----------------------------------------------------------------------------------------
' ȭ���� �ʱ���� ������ ����
'-----------------------------------------------------------------------------------------
Sub InitPageData
	with frmThis
		.txtYEARMON.value = Mid(gNowDate,1,4) & Mid(gNowDate,6,2)
		
		.sprSht.MaxRows = 0			
	end With
End Sub

Sub imgSave_onclick
	IF frmThis.cmbCOMMITGBN.value = "1" then
		gErrorMsgBox "�̻��� ���¿����� ���尡���մϴ�.","����ȳ�!"
		Exit Sub
	end if
	gFlowWait meWAIT_ON
	ProcessRtn
	gFlowWait meWAIT_OFF
End Sub

'------------------------------------------
' ���� �������
'------------------------------------------
Sub ProcessRtn
	Dim intRtn
   	dim vntData
   	Dim vntData1
   	dIM strYEARMON
	
	with frmThis
   		if .sprSht.MaxRows = 0 Then
			gErrorMsgBox "��ȸ�� ���� �����Ƿ� ������ �Ұ��� �մϴ�.","����ȳ�!"
			Exit Sub
		end if
		
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		
		strYEARMON = mobjSCGLSpr.GetTextBinding(.sprSht,"YEARMON",1)
		vntData1 = mobjMDSRREPORTLIST.SelectRtn_EXISTREPORT(gstrConfigXml,mlngRowCnt,mlngColCnt, strYEARMON)
		
		If mlngRowCnt > 0 then
			IF vntData1(1,1) = "VOCH" THEN
				intRtn = gYesNoMsgbox("�̹� ��ǥ�Ϸ���·� ����� �ڷᰡ �����մϴ�. �ٽ� �����Ͻðڽ��ϱ�?","�ڷ�����")
				IF intRtn <> vbYes then exit Sub
				DeleteRtn_process
			elseif vntData1(1,1) = "NOVOCH" THEN
				intRtn = gYesNoMsgbox("�̹� ��ǥ�̿Ϸ���·� ����� �ڷᰡ �����մϴ�. �ٽ� �����Ͻðڽ��ϱ�?","�ڷ�����")
				IF intRtn <> vbYes then exit Sub
				DeleteRtn_process
			END IF
		END IF
		
		
		'On error resume next
		'��Ʈ�� ����� �����͸� �����´�.
		for i=1 to .sprSht.MaxRows
			mobjSCGLSpr.CellChanged frmThis.sprSht, 1, i
		Next
		
		vntData = mobjSCGLSpr.GetDataRows(.sprSht,"YEARMON|VOCH_GBNNAME|CLIENTNAME|MEDNAME|REAL_MED_NAME|CLIENTSUBNAME|SUBSEQNAME|MEDFLAGNAME|EXCLIENTCODE|AMT|VAT|PROGNAME|CLIENTCODE|MEDCODE|REAL_MED_CODE|CLIENTSUBCODE|SUBSEQ|MEDFLAG|VOCH_GBN|TRU_TAX_FLAG|MPP")
		
		intRtn = mobjMDSRREPORTLIST.ProcessRtn(gstrConfigXml,vntData)
		
		if not gDoErrorRtn ("ProcessRtn") then 'EXCUTION_ProcessRtn ProcessRtn_Confirm_OK
			'��� �÷��� Ŭ����
			mobjSCGLSpr.SetFlag  .sprSht,meCLS_FLAG
			gOkMsgBox "����Ǿ����ϴ�.","���强��!"
			'gWriteText "", intRtn & "���� �ڷᰡ ����" & mePROC_DONE
			.cmbCOMMITGBN.value ="1"
			SelectRtn
   		end if
   	end with
End Sub

Sub AMT_SUM
	Dim lngCnt
	Dim lntTVAMT,		lntTVAMTSUM
	Dim lntRDAMT,		lntRDAMTSUM
	Dim lntDMBAMT,		lntDMBAMTSUM
	Dim lntCATVAMT,		lntCATVAMTSUM
	Dim lntINTERNETAMT, lntINTERNETAMTSUM
	Dim lntOUTDOORAMT,	lntOUTDOORAMTSUM
	Dim lntMP01AMT,		lntMP01AMTSUM
	Dim lntMP02AMT,		lntMP02AMTSUM

	With frmThis
		lntTVAMTSUM = 0
		lntRDAMTSUM = 0
		
		lntDMBAMTSUM = 0
		lntCATVAMTSUM = 0
		
		lntINTERNETAMTSUM = 0
		lntOUTDOORAMTSUM = 0
		
		lntMP01AMTSUM = 0
		lntMP02AMTSUM = 0
		
		'������ �׸��� �հ�׸��� ���ֱ�
		For lngCnt = 1 To .sprSht.MaxRows
			lntTVAMT = 0
			lntRDAMT = 0
			lntDMBAMT = 0
			lntCATVAMT = 0
			lntINTERNETAMT = 0
			lntOUTDOORAMT = 0
			lntMP01AMT = 0
			lntMP02AMT = 0
                
			IF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "TV" THEN
				lntTVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntTVAMTSUM = lntTVAMTSUM  + lntTVAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "RD" THEN
				lntRDAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntRDAMTSUM = lntRDAMTSUM  + lntRDAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "DMB" THEN
				lntDMBAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntDMBAMTSUM = lntDMBAMTSUM  + lntDMBAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "CATV" THEN
				lntCATVAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntCATVAMTSUM = lntCATVAMTSUM  + lntCATVAMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "�Ź�" THEN
				lntMP01AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntMP01AMTSUM = lntMP01AMTSUM  + lntMP01AMT
			
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "����" THEN
				lntMP02AMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntMP02AMTSUM = lntMP02AMTSUM  + lntMP02AMT
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "���ͳ�" THEN
				lntINTERNETAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntINTERNETAMTSUM = lntINTERNETAMTSUM  + lntINTERNETAMT
			ELSEIF mobjSCGLSpr.GetTextBinding(.sprSht,"MEDFLAGNAME", lngCnt) = "����" THEN
				lntOUTDOORAMT = mobjSCGLSpr.GetTextBinding(.sprSht,"AMT", lngCnt)
				lntOUTDOORAMTSUM = lntOUTDOORAMTSUM  + lntOUTDOORAMT
			END IF
		Next
		
		if .sprSht.MaxRows >0 Then
			.txtTV.value = lntTVAMTSUM
			.txtRD.value = lntRDAMTSUM
			
			.txtDMB.value = lntDMBAMTSUM
			.txtCATV.value = lntCATVAMTSUM
			
			.txtINTERNET.value = lntINTERNETAMTSUM
			.txtOUTDOOR.value = lntOUTDOORAMTSUM
			
			.txtMP01.value = lntMP01AMTSUM
			.txtMP02.value = lntMP02AMTSUM
			
			call gFormatNumber(.txtTV,0,true)
			call gFormatNumber(.txtRD,0,true)
			call gFormatNumber(.txtDMB,0,true)
			call gFormatNumber(.txtCATV,0,true)
			call gFormatNumber(.txtINTERNET,0,true)
			call gFormatNumber(.txtOUTDOOR,0,true)
			call gFormatNumber(.txtMP01,0,true)
			call gFormatNumber(.txtMP02,0,true)
		ELSE
			.txtTV.value = 0
			.txtRD.value = 0
			
			.txtDMB.value = 0
			.txtCATV.value = 0
			
			.txtINTERNET.value = 0
			.txtOUTDOOR.value = 0
			
			.txtMP01.value = 0
			.txtMP02.value = 0
		end if
	End With
End Sub

		</script>
	</HEAD>
	<body class="base">
		<XML id="xmlBind"></XML>
		<FORM id="frmThis" method="post" runat="server">
			<!--Main Start-->
			<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="1040" border="0">
				<!--Top TR Start-->
				<TBODY>
					<TR>
						<TD style="HEIGHT: 54px">
							<!--Top Define Table Start-->
							<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
								border="0">
								<TR>
									<TD align="left" width="400" height="28">
										<table cellSpacing="0" cellPadding="0" width="100%" border="0">
											<tr>
												<td align="left" width="14" rowSpan="2"><IMG height="28" src="../../../images/TitleIcon.gIf" width="14"></td>
												<td align="left" height="4"><FONT face="����"></FONT></td>
											</tr>
											<tr>
												<td class="TITLE">&nbsp;����Ʈ ����</td>
											</tr>
										</table>
									</TD>
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
									</TD>
								</TR>
							</TABLE>
							<!--Top Define Table End-->
							<!--Input Define Table End-->
							<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0"> <!--TopSplit Start->
								<!--TopSplit Start-->
								<TR>
									<TD class="TOPSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--TopSplit End-->
								<!--Input Start-->
								<TR>
									<TD class="KEYFRAME" style="WIDTH: 1040px; HEIGHT: 20px" vAlign="top" align="center">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="SEARCHLABEL" style="CURSOR: hand" onclick="vbscript:Call gCleanField(txtYEARMON,'')"
													width="90">���&nbsp;
												</TD>
												<TD class="SEARCHDATA" width="180"><INPUT class="INPUT" id="txtYEARMON" title="���" style="WIDTH: 88px; HEIGHT: 22px" accessKey="NUM"
														type="text" maxLength="6" size="9" name="txtYEARMON">
												</TD>
												<TD class="SEARCHLABEL" width="90">��������</TD>
												<TD class="SEARCHDATA" width="120"><SELECT id="cmbCOMMITGBN" title="��������" style="WIDTH: 112px" name="cmbCOMMITGBN">
														<OPTION value="0" selected>�̻���</OPTION>
														<OPTION value="1">����</OPTION>
													</SELECT>
												</TD>
												<TD class="SEARCHLABEL" width="90">��ǥ����</TD>
												<TD class="SEARCHDATA">&nbsp;&nbsp;<INPUT id="rdT" title="Ȯ��������ȸ" type="radio" CHECKED value="rdT" name="rdGBN">&nbsp;��ǥȮ��&nbsp;&nbsp;&nbsp;&nbsp;
													<INPUT id="rdF" title="��Ȯ��������ȸ" type="radio" value="rdF" name="rdGBN">&nbsp;��ǥ��Ȯ��
												</TD>
												<td class="SEARCHDATA" width="50"><IMG id="imgQuery" onmouseover="JavaScript:this.src='../../../images/imgQueryOn.gIF'"
														style="CURSOR: hand; HEIGHT: 20px" onmouseout="JavaScript:this.src='../../../images/imgQuery.gIF'" height="20"
														alt="�ڷḦ �˻��մϴ�." src="../../../images/imgQuery.gIF" align="absMiddle" border="0" name="imgQuery">
												</td>
											</TR>
										</TABLE>
										<TABLE height="10" cellSpacing="0" cellPadding="0" width="100%" background="../../../images/TitleBG.gIF"
											border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 25px"><FONT face="����"></FONT></TD>
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
															<td class="TITLE">&nbsp;����Ʈ �����ڷ� ���� �� ��ȸ</td>
														</tr>
													</table>
												</TD>
												<TD style="WIDTH: 640px" vAlign="middle" align="right" height="20">
													<!--Common Button Start-->
													<TABLE id="tblButton" style="HEIGHT: 20px" cellSpacing="0" cellPadding="2" border="0">
														<TR>
															<TD><IMG id="imgSave" onmouseover="JavaScript:this.src='../../../images/imgSaveOn.gif'" style="CURSOR: hand"
																	onmouseout="JavaScript:this.src='../../../images/imgSave.gif'" height="20" alt="�ڷḦ �����մϴ�."
																	src="../../../images/imgSave.gIF" border="0" name="imgSave"></TD>
															<td><IMG id="imgDelete" onmouseover="JavaScript:this.src='../../../images/imgDeleteOn.gIF'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgDelete.gIF'"
																	height="20" alt="�ڷḦ �����մϴ�.." src="../../../images/imgDelete.gif" border="0" name="imgDelete"></td>
															<td><IMG id="imgExcel" onmouseover="JavaScript:this.src='../../../images/imgExcelOn.gif'"
																	style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgExcel.gif'"
																	height="20" alt="�ڷḦ ������ �޽��ϴ�." src="../../../images/imgExcel.gIF" border="0" name="imgExcel"></td>
														</TR>
													</TABLE>
													<!--Common Button End--></TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1250px; HEIGHT: 10px">
										<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="BODYSPLIT" style="WIDTH: 1040px; HEIGHT: 10px">
													<TABLE class="DATA" id="tblKey" cellSpacing="1" cellPadding="0" width="100%" border="0">
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">����</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">1��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">2��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">3��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">4��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">5��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">6��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">7��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">8��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">9��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">10��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">11��</TD>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">12��</TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">��ǥȮ��</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH01" title="1��" disabled type="checkbox" name="rdVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH02" title="2��" disabled type="checkbox" name="rdVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH03" title="3��" disabled type="checkbox" name="rdVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH04" title="4��" disabled type="checkbox" name="rdVOCH04"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH05" title="5��" disabled type="checkbox" name="rdVOCH05"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH06" title="6��" disabled type="checkbox" name="rdVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH07" title="7��" disabled type="checkbox" name="rdVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH08" title="8��" disabled type="checkbox" name="rdVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH09" title="9��" disabled type="checkbox" name="rdVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH10" title="10��" disabled type="checkbox" name="rdVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH11" title="11��" disabled type="checkbox" name="rdVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdVOCH12" title="12��" disabled type="checkbox" name="rdVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">��Ȯ��</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH01" title="1��" disabled type="checkbox" name="rdNOVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH02" title="2��" disabled type="checkbox" name="rdNOVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH03" title="3��" disabled type="checkbox" name="rdNOVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH04" title="4��" disabled type="checkbox" name="rdNOVOCH04"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH05" title="5��" disabled type="checkbox" name="rdNOVOCH05"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH06" title="6��" disabled type="checkbox" name="rdNOVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH07" title="7��" disabled type="checkbox" name="rdNOVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH08" title="8��" disabled type="checkbox" name="rdNOVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH09" title="9��" disabled type="checkbox" name="rdNOVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH10" title="10��" disabled type="checkbox" name="rdNOVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH11" title="11��" disabled type="checkbox" name="rdNOVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT id="rdNOVOCH12" title="12��" disabled type="checkbox" name="rdNOVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD--></TR>
														<TR>
															<TD class="LABEL" style="TEXT-ALIGN: center" width="80">������</TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH01" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="7" name="txtVOCH01"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH02" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH02"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH03" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH03"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH04" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH04"></TD>
															<TD class="DATA" width="80" style="TEXT-ALIGN: center"><INPUT class="NOINPUT_L" id="txtVOCH05" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH05"></TD>
															<TD class="DATA" width="80" style="TEXT-ALIGN: center"><INPUT class="NOINPUT_L" id="txtVOCH06" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH06"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH07" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH07"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH08" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH08"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH09" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH09"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH10" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH10"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH11" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH11"></TD>
															<TD class="DATA" style="TEXT-ALIGN: center" width="80"><INPUT class="NOINPUT_L" id="txtVOCH12" title="������" style="WIDTH: 70px; HEIGHT: 22px" readOnly
																	type="text" maxLength="10" size="2" name="txtVOCH12"></TD>
															<!--TD class="DATA"><INPUT class="NOINPUTB" id="txtSTATE" style="WIDTH: 159px; HEIGHT: 21px" readOnly type="text"
														size="21" name="txtSTATE"></TD-->
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<!--Input End-->
								<!--BodySplit Start-->
								<TR>
									<TD class="BODYSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<!--BodySplit End-->
								<!--List Start-->
								<TR>
									<TD class="LISTFRAME" style="WIDTH: 1038px; HEIGHT: 555px" vAlign="top" align="center">
										<DIV id="pnlTab1" style="VISIBILITY: hidden; WIDTH: 100%; POSITION: relative; HEIGHT: 556px"
											ms_positioning="GridLayout">
											<OBJECT id="sprSht" style="Z-INDEX: 101; LEFT: 0px; WIDTH: 100%; POSITION: absolute; TOP: 0px; HEIGHT: 556px"
												width="100%" classid="clsid:41F841C1-AE16-11D5-8817-0050DA6EF5E5" name="sprSht" VIEWASTEXT>
												<PARAM NAME="_Version" VALUE="393216">
												<PARAM NAME="_ExtentX" VALUE="27437">
												<PARAM NAME="_ExtentY" VALUE="15505">
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
								<!--List End-->
								<!--BodySplit Start-->
								<!--Brench End-->
								<!--Bottom Split Start-->
								<TR>
									<TD class="BOTTOMSPLIT" style="WIDTH: 1040px"><FONT face="����"></FONT></TD>
								</TR>
								<TR>
									<TD>
										<TABLE cellSpacing="0" cellPadding="0" width="100%" border="0">
											<TR>
												<TD class="KEYFRAME" style="WIDTH: 1040px" vAlign="middle" align="center">
													<TABLE class="DATA" cellSpacing="1" cellPadding="0" width="100%" border="0">
														<TR>
															<TD class="LABEL" width="90">TV</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtTV" title="TV�����" style="WIDTH: 152px; HEIGHT: 22px" type="text"
																	size="20" name="txtTV" readOnly></TD>
															<TD class="LABEL" width="90">RD</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtRD" title="���������" style="WIDTH: 152px; HEIGHT: 22px" type="text"
																	size="20" name="txtRD"></TD>
															<TD class="LABEL" width="90">CATV</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtCATV" title="���̺����" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtCATV" readOnly></TD>
															<TD class="LABEL" width="90">������DMB</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtDMB" title="������DMB�����" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtDMB" readOnly></TD>
														</TR>
														<TR>
															<TD class="LABEL" width="90">�Ź�</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtMP01" title="�Ź������" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="20" name="txtMP01" readOnly></TD>
															<TD class="LABEL" width="90">����</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtMP02" title="���������" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="21" name="txtMP02" readOnly></TD>
															<TD class="LABEL" width="90">���ͳ�</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtINTERNET" title="���ͳݱ����" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtINTERNET" readOnly></TD>
															<TD class="LABEL" width="90">����</TD>
															<TD class="DATA" width="170"><INPUT class="NOINPUT_R" id="txtOUTDOOR" title="���ܱ����" style="WIDTH: 152px; HEIGHT: 22px"
																	type="text" size="12" name="txtOUTDOOR" readOnly></TD>
														</TR>
													</TABLE>
												</TD>
											</TR>
										</TABLE>
									</TD>
								</TR>
								<TR>
									<TD class="BOTTOMSPLIT" id="lblStatus"><FONT face="����"></FONT></TD>
								</TR>
								<!--Bottom Split End--></TABLE>
							<!--Input Define Table End--></TD>
					</TR>
					<!--Top TR End--></TBODY></TABLE>
			<!--Main End--></FORM>
		</TR></TBODY></TABLE>
	</body>
</HTML>
