<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCCOINITPOP.aspx.vb" Inherits="SC.SCCOINITPOP" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��������</title> 
		<!--
'****************************************************************************************
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCOMPPPOP.aspx
'��      �� : MPP �˾�
'Ư��  ���� : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2009/07/07 By KTY
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
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331">
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjSCCOGET 
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

Sub imgCancel_onclick
	call Window_OnUnload()
End Sub

'-----------------------------
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()
	dim vntInParam
	dim intNo,i
	
	'����������ü ����	
	set mobjSCCOGET = gCreateRemoteObject("cSCCO.ccSCCOGET")
	'���Ѽ���/�����Ķ����/ȭ������ ���� �⺻ �۾��� ����
	gInitComParams mobjSCGLCtl,"MC"

	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'�⺻�� ����
		mstrFields = "": mblnUseOnly = true: mstrUseDate="" : mblnLikeCode = true
		
		for i = 0 to intNo
			select case i
				case 0 : mstrFields = vntInParam(i)			'��ȸ�߰��ʵ�
				case 1 : mblnUseOnly = vntInParam(i)		'���� ������� �͸�
				case 2 : mstrUseDate = vntInParam(i)		'�ڵ� ��� ����
				case 3 : mblnLikeCode = vntInParam(i)		'��ȸ�� �ڵ带 Like���� ����
			end select
		next
	end with	
end sub

Sub EndPage()
	set mobjSCCOGET = Nothing
	gEndPage
End Sub


-->
		</script>
	</HEAD>
	<body class="base" bottomMargin="0" leftMargin="0" topMargin="0" rightMargin="0">
		<TABLE id="tblForm" cellSpacing="0" cellPadding="0" width="373" border="0">
			<TR>
				<TD>
					<FORM id="frmThis">
						<TABLE id="tblTitle" height="28" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<td style="WIDTH: 148px" align="left" width="148" height="28">
									<table cellSpacing="0" cellPadding="0" width="100%" border="0">
										<tr>
											<td align="left" width="49" rowSpan="2"><IMG height="28" src="../../../images/title_icon1.gif" width="49"></td>
											<td align="left" height="4"></td>
										</tr>
										<tr>
											<td class="TITLE" id="objTitle">
												��������
											</td>
										</tr>
									</table>
								</td>
								<TD vAlign="middle" align="right" height="28">
									<TABLE class="" id="tblWaitP" style="Z-INDEX: 200; LEFT: 150px; VISIBILITY: hidden; WIDTH: 65px; POSITION: absolute; TOP: 0px; HEIGHT: 23px"
										cellSpacing="1" cellPadding="1" width="75%" border="0">
									</TABLE>
								</TD>
							</TR>
						</TABLE>
						<TABLE cellSpacing="0" cellPadding="0" width="100%" background="../../../images/back_p.gIF"
							border="0">
							<TR>
								<TD align="left" width="100%" height="2"></TD>
							</TR>
						</TABLE>
						<TABLE id="tblBody" cellSpacing="0" cellPadding="0" width="100%" border="0">
							<TR>
								<TD align="center" class="TITLE">
									<FONT face="����" size="2">���� </FONT>
								</TD>
								<td class="TITLE">
									&nbsp; !�߿� [RMS] ���Ӱ��� ���� �Դϴ�.
								</td>
							</TR>
							<TR>
								<TD class="BOTTOMSPLIT"></TD>
							</TR>
							<tr>
								<td colspan="2" class="TITLE" align="left">
									<P>
										<br>
										�ȳ��ϼ��� RMS ����� �Դϴ�.
										<BR>
										���� SK_P �� ��� �ý��� ������ �����Ͽ� RMS ����
										<BR>
										���̵��� ��� ���� M&amp;C ���� ����Ͻô� ID �� PW ��
										<BR>
										�ӽ� ���� �Ǿ��ֽ��ϴ�.
										<BR>
										���� ������ �������� �α��� �ϼž� �ϴ°�쿡��
										<BR>
										RMS ����ڿ��� �����Ͽ� �ֽʽÿ�..
										<BR>
										<br>
										�ű� SK_P �� �������� �α��� �Ͻǰ�쿡 �н����尡
										<BR>
										�ű� ����� �����ϰ� �Է��Ͻø� ������ �����մϴ�.
										<BR>
										�����մϴ�. <font color="red">[�� -&gt; �űԻ��: 123456 �н�����:123456]</font>
										<br>
										<br>
										<font color="red">RMS - ���ǻ���, 02-6390-3981</font>
									</P>
								</td>
							</tr>
							<tr>
								<td align="right" colspan="2">
									<TABLE id="tblButton" style="WIDTH: 52px; HEIGHT: 20px" cellSpacing="0" cellPadding="0"
										border="0">
										<TR>
											<TD><IMG id="imgCancel" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gif'"
													style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gif'"
													height="20" alt="ȭ���� �ݽ��ϴ�." src="../../../images/imgClose.gif" width="54" border="0"
													name="imgCancel"></TD>
											<TD width="15"><FONT face="����"></FONT></TD>
										</TR>
									</TABLE>
								</td>
							</tr>
						</TABLE>
				</TD>
				</FORM>
			</TR>
		</TABLE>
	</body>
</HTML>
