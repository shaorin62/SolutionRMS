<%Response.Buffer = True%>
<%Response.Expires=0%>
<%
'1. �������� ����
	Dim mstrMessage			'�޼���

'2. �Ķ���� �Ҵ�
	mstrMessage = Request.QueryString("MSG")
	If Len(Trim(mstrMessage)) = 0 Then
		mstrMessage = "�������� �˼� ���� ������ �߻��Ͽ����ϴ�."
	End If
'==================================================================================
' ��� ����...
'==================================================================================
Set session("oPageEngine") = Nothing
Set session("oApp") = Nothing
Set session("oRpt") = Nothing
%>
<html>
	<head>
		<title>Message</title>
		<meta name="VI60_defaultClientScript" content="JavaScript">
		<meta NAME="GENERATOR" Content="Microsoft Visual Studio 6.0">
		<meta HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=ks_c_5601">
		<meta HTTP-EQUIV="Pragma" CONTENT="no-cache">
		<meta HTTP-EQUIV="Cache-Control" CONTENT="no-cache">
		<meta HTTP-EQUIV="Erpires" CONTENT="-1">
	</head>
	<body bottomMargin="0" leftMargin="0" rightMargin="0" topMargin="0">
		<!--���� �׵θ� ����-->
		<table align="center" border="1" cellPadding="0" cellSpacing="0" width="100%" height="100%" bgcolor="#E2E0E0" style="font-family:����; font-size:9pt;">
			<tr>
				<td>
					<!--���� ȭ�� ����-->
					<table align="center" border="1" cellPadding="0" cellSpacing="0" width="100%" height="100%" style="font-family:����; font-size:9pt;">
						<tr>
							<td>
								<!--���� ���� ����-->
								<table align="center" cellspacing="0" cellSpacing="0" width="500" height="300" style="font-family:����; font-size:9pt;" background="./images/confirm.jpg">
									<tr>
										<td valign="top" align="left">
											<!--�α��� ���� �Է� ����-->
											<table cellspacing="3" cellSpacing="3" style="font-family:����; font-size:9pt;" border="0">
												<tr style="HEIGHT: 60px">
													<td width="160">&nbsp;</td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>
												</tr>
												<tr style="HEIGHT: 20px">
													<td>&nbsp;</td>
													<th colspan="3" align="center">
														<font size="2" color="blue">�ý��� ������ �߻��߽��ϴ�.<br>
															<br>
															����Ƿ� ������ �ּ���. </font>
													</th>
												</tr>
												<tr style="HEIGHT: 42px">
													<td>&nbsp;</td>
													<td colspan="3" align="center">
														<font size="2" color="Red">������:
															<%=mstrMessage%>
														</font>
													</td>
												</tr>
												<tr>
													<td>&nbsp;</td>
													<th colspan="3">
														<img id="imgConfirm" name="imgConfirm" src="./images/CON_BT.GIF" height="19" width="60" style="CURSOR: hand" alt="ȭ���� �ݽ��ϴ�." LANGUAGE="javascript" onclick="return imgConfirm_onclick()">
													</th>
												</tr>
											</table>
											<!--�α��� ���� �Է� ����-->
										</td>
									</tr>
								</table>
								<!--���� ���� ����-->
							</td>
						</tr>
					</table>
					<!--���� ȭ�� ����-->
				</td>
			</tr>
		</table>
		<!--���� �׵θ� ����-->
	</body>
</html>
<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
//========================================
// Ȯ�θ޼���ó��
//========================================
function imgConfirm_onclick() {
	window.blur();
}

//-->
</SCRIPT>
