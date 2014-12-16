<%Response.Buffer = True%>
<%Response.Expires=0%>
<%
'1. 전역변수 선언
	Dim mstrMessage			'메세지

'2. 파라메터 할당
	mstrMessage = Request.QueryString("MSG")
	If Len(Trim(mstrMessage)) = 0 Then
		mstrMessage = "서버에서 알수 없는 에러가 발생하였습니다."
	End If
'==================================================================================
' 쎄션 해제...
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
		<!--전제 테두리 시작-->
		<table align="center" border="1" cellPadding="0" cellSpacing="0" width="100%" height="100%" bgcolor="#E2E0E0" style="font-family:굴림; font-size:9pt;">
			<tr>
				<td>
					<!--전제 화면 시작-->
					<table align="center" border="1" cellPadding="0" cellSpacing="0" width="100%" height="100%" style="font-family:굴림; font-size:9pt;">
						<tr>
							<td>
								<!--업무 내용 시작-->
								<table align="center" cellspacing="0" cellSpacing="0" width="500" height="300" style="font-family:굴림; font-size:9pt;" background="./images/confirm.jpg">
									<tr>
										<td valign="top" align="left">
											<!--로그인 정보 입력 시작-->
											<table cellspacing="3" cellSpacing="3" style="font-family:굴림; font-size:9pt;" border="0">
												<tr style="HEIGHT: 60px">
													<td width="160">&nbsp;</td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>
													<td>&nbsp;</td>
												</tr>
												<tr style="HEIGHT: 20px">
													<td>&nbsp;</td>
													<th colspan="3" align="center">
														<font size="2" color="blue">시스템 오류가 발생했습니다.<br>
															<br>
															전산실로 연락을 주세요. </font>
													</th>
												</tr>
												<tr style="HEIGHT: 42px">
													<td>&nbsp;</td>
													<td colspan="3" align="center">
														<font size="2" color="Red">오류명:
															<%=mstrMessage%>
														</font>
													</td>
												</tr>
												<tr>
													<td>&nbsp;</td>
													<th colspan="3">
														<img id="imgConfirm" name="imgConfirm" src="./images/CON_BT.GIF" height="19" width="60" style="CURSOR: hand" alt="화면을 닫습니다." LANGUAGE="javascript" onclick="return imgConfirm_onclick()">
													</th>
												</tr>
											</table>
											<!--로그인 정보 입력 종료-->
										</td>
									</tr>
								</table>
								<!--업무 내용 종료-->
							</td>
						</tr>
					</table>
					<!--전제 화면 종료-->
				</td>
			</tr>
		</table>
		<!--전제 테두리 종료-->
	</body>
</html>
<SCRIPT ID="clientEventHandlersJS" LANGUAGE="javascript">
<!--
//========================================
// 확인메세지처리
//========================================
function imgConfirm_onclick() {
	window.blur();
}

//-->
</SCRIPT>
