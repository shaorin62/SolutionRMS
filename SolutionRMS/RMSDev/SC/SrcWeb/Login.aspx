<%@ Page Language="vb" AutoEventWireup="false" Codebehind="Login.aspx.vb" Inherits="SC.Login" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>열정과 패기로 ! Beyond SK ! RMS</TITLE>
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="/css/style.css" type="text/css" rel="STYLESHEET">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../Etc/SCClient.inc" -->
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../Etc/SCUIClass.inc" -->
		<SCRIPT ID="clientEventHandlersVBS" LANGUAGE="vbscript">
<!--

option explicit
Const meTab = 9
Dim mlngRowCnt
Dim mlngColCnt
Dim mobjSCCOLOGIN
Dim mlngPreRowCnt
Dim mlngPreColCnt
Dim mlngClRowCnt
Dim mlngClColCnt
Dim mstrLOGINCHK

mstrLOGINCHK = "True"

Sub window_onload()
	InitPage
	'call initpopup()
End Sub

sub initpopup
	Dim vntRet
	Dim vntInParams
	On error resume next
	
	With frmThis
		vntInParams = array("", "")
	    vntRet = gShowModalWindow("SCCO/SCCOINITPOP.aspx",vntInParams , 380,300)
	End With
	gSetChange
end sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub MainFrame_Open()
	Dim strWith
	Dim strHeight
	
	strWith =  Screen.width
	strHeight =  Screen.height - 30
	ShowWindow "main.asp", "work",strWith,strHeight,""
End Sub

'로그인 텍스트 박스
Sub txtLOGIN_onfocus
	With frmThis
		.txtLOGIN.value = ""
		.txtLOGIN.focus()
	End With
End Sub

'패스워드 텍스트 박스
Sub txtPASSWORD_onfocus
	With frmThis
		.txtPASSWORD.value = ""
		.txtPASSWORD.focus()
	End With
End Sub

Sub txtPASSWORD_onkeydown
	If window.event.keyCode = meEnter Then
		If frmThis.txtLOGIN.value <> "" and frmThis.txtPASSWORD.value <> "" Then
			imgLOGIN_onclick
		End If	
	End If
End Sub

'----------------------------------------------------------------------------------
'로그인 모듈 1) 사용자 ID,PWD 를 리턴 하여 MAIN.aspx 호출
'----------------------------------------------------------------------------------
Sub imgLOGIN_onclick
	Dim intRtn

	If CheckWebClient Then
	Else
		intRtn = gYesNoMsgbox("현재 클라이언트 모듈이 설치 되어있지 않습니다."& vbcrlf &"설치하시겠습니까?"& vbcrlf &"설치후 새로고침을 눌러 다시 접속 하십시오.","설치확인")
		If intRtn <> vbYes Then exit Sub
		location.href = "http://10.110.10.89/DownLoad/SCGLCom.exe"
		Exit Sub
	End If
	
	Dim strLoginIdx
	Dim strPWD
	Dim vntData
	Dim vntPreData
	Dim vntDataClipping
	Dim strSTARTLOGINDATECHK
	Dim strSTARTLOGINDATE
	Dim strTERMDATE
	Dim strDate
	Dim vntInParams
	Dim vntRet
	Dim strSAVEPWD
	Dim strNOWID
	Dim strNOWPWD
	Dim strID
	gstrUsrID = ""
	gstrEmpNo = ""
	gstrUsrName = ""

	'권한설정/공통파라메터/화면조정 등의 기본 작업을 수행
	gInitPageSetting mobjSCGLCtl,"MC"
 
	With frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		mlngPreRowCnt =clng(0)
		mlngPreColCnt =clng(0)
		mlngClRowCnt =clng(0)
		mlngClColCnt =clng(0)
		
		strLoginIdx = Trim(.txtLOGIN.value)
		strPWD = Trim(.txtPASSWORD.value)
		
		If Len(strLoginIdx) = 5 Then strLoginIdx = "000" & strLoginIdx
		
		vntPreData = mobjSCCOLOGIN.SelectRtn_IDX(gstrConfigXml,mlngPreRowCnt,mlngPreColCnt,strLoginIdx)
		If not gDoErrorRtn ("SelectRtn_IDX") Then
			If mlngPreRowCnt > 0  Then
				If vntPreData(0,1) = "N" Then 
					gErrorMsgbox "입력하신 ID 는 사용이 중지된 아이디입니다." & vbcrlf & "관리자에게 문의 하십시오.","로그인안내!" 
					.txtLOGIN.value = ""
					.txtPASSWORD.value = ""
					.txtLOGIN.focus()
					exit Sub
				End If
 				
				vntData = mobjSCCOLOGIN.SelectRtn_LOGINIDX(gstrConfigXml,mlngRowCnt,mlngColCnt,strLoginIdx,strPWD)
				If not gDoErrorRtn ("SelectRtn_LOGINIDX") Then
					If mlngRowCnt = 1 Then
						gstrUsrID=vntData(0,1)
						gstrEmpNo=vntData(0,1)
						strSAVEPWD = vntData(1,1)
						gstrUsrName=vntData(2,1)
						strSTARTLOGINDATE = vntData(4,1)
						
						strDate = dateAdd("M",-6,Date)
						strDate = Replace(strDate,"-","") 
						
						gSetSession gstrUsrID,gstrEmpNo,gstrUsrName
						
						gInitPageSetting mobjSCGLCtl,"MC"
						'정상로그인 이기때문에, ClippingLevel 을 0으로 교체 한다.
						intRtn = mobjSCCOLOGIN.ClippingCleanRtn(gstrConfigXml,strLoginIdx)
						
						'실제로그인
						If strSTARTLOGINDATE = "" Or (strSTARTLOGINDATE < strDate) Then
							'패스워드 강제 변경 mstrLOGINCHK 을 false 로 만들고 pwd2 인풋박스 보여주고 Login_Change 로직을 태운다
							'여기서 바로 팝업을 띄운다
							gErrorMsgBox "최초접속 및 6개월 이내 비밀번호 변경 내역이 없는 사용자 께서는" & vbcrlf & "비밀번호를 변경 하셔야 합니다.","로그인안내!"
							If .txtLOGIN.value = "ID" Or .txtLOGIN.value = "" Then
								strID = ""
							Else
								strID = .txtLOGIN.value 
							End If
							vntInParams = array(strID)
							vntRet = gShowModalWindow("passwordChange.aspx",vntInParams , 380,300)
							If vntRet = "T" Then
								MainFrame_Open 
							End If
						Else
							MainFrame_Open 
						End If 
					Else
						vntDataClipping = mobjSCCOLOGIN.SelectRtn_Clipping(gstrConfigXml,mlngClRowCnt,mlngClColCnt,strLoginIdx)

						If not gDoErrorRtn ("SelectRtn_Clipping") Then
							'로그인 연속 5회 이상 실패경우
							If vntDataClipping(0,1)+1 = 5 Then 
								'USE_YN 을 "N" 으로 변경
								intRtn = mobjSCCOLOGIN.ClippingEndRtn(gstrConfigXml,strLoginIdx)
								gErrorMsgBox "비밀번호 연속5회 오류이므로 계정이 사용정지 되었습니다." & vbcrlf & "사용을 계속 하시려면, 관리자 에게 문의 하십시오.", "로그인안내!"
								Exit Sub
							Else
								intRtn = mobjSCCOLOGIN.ClippingRtn(gstrConfigXml,strLoginIdx)
								gErrorMsgbox "비밀번호 입력오류 입니다." & vbcrlf & "입력 " & vntDataClipping(0,1)+1 & " 회 오류!","로그인안내!"
								Exit Sub
							End If
						End If
					End If
   				End If
   			Else
   				gErrorMsgbox "입력하신 ID 는 존재하지 않는 ID 입니다." & vbcrlf & "관리자에게 문의 하십시오.","로그인안내!"
   			End If
   		End If
	End With
End Sub

'각 웹 페이지를 오픈한다.
Sub ShowWindow (byval strPageURL, byval strWindowName, byval lngWidth, byval lngHeight, byval strOptions)
	Dim lngTop, lngLeft
	
	'화면의 중앙에 위치시킨다.
	lngTop = 0
	lngLeft = (window.screen.width - lngWidth) / 2

	strOptions = "toolbar=no, location=no, menubar=no, scrollbars=Yes, status=yes, resizable=yes, top=" & lngTop & ", left=" & lngLeft & ", width=" & lngWidth-10 & ", height=" & lngHeight-50 
	window.open  strPageURL,strWindowName,strOptions
	
	Call pageunload()
End Sub

Sub PWDCHANGE
	'여기서도 바로 로그인 변경 팝업을 띄운다
	Dim vntInParams
	Dim vntRet
	Dim strID
	With frmThis
		If .txtLOGIN.value = "ID" Or .txtLOGIN.value = "" Then
		strID = ""
		Else
		strID = .txtLOGIN.value 
		End If
		vntInParams = array(strID)
		vntRet = gShowModalWindow("passwordChange.aspx",vntInParams , 380,300)
		If vntRet = "T" Then
			MainFrame_Open 
		End If
	End With
End Sub

Function CheckWebClient()
	On error resume next   
	Dim strVer : strVer = mobjSCGLCtl.Version

	If Err.number <> 0 Then
		CheckWebClient = false
	Else
		CheckWebClient = true
	End If
	err.clear	
End function

Sub InitPage()
	'서버업무객체 생성	
	set mobjSCCOLOGIN = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process
End Sub

Sub EndPage()
	set mobjSCCOLOGIN = Nothing
	gEndPage
End Sub

-->
		</SCRIPT>
		<script language="javascript">

function Login_Enter(){
	imgLOGIN_onclick()
}
//function pageunload (){
//	top.window.opener = top;
//	top.window.close();
//}


function pageunload (){
	//top.window.opener = top;
	//http://10.110.10.86:4350/SC/SrcWeb/Login.htm
//	if(navigator.appVersion.indexOf("MSIE 7.0") >= 0) {
//		window.open("http://10.110.10.86:4350/SC/SrcWeb/" + "Login.htm","_self").close();
//	}else if(navigator.appVersion.indexOf("MSIE 8.0") >=0){
//		window.open("http://10.110.10.86:4350/SC/SrcWeb/" + "Login.htm","_self").close();
//	}else{
//		self.close();
//	}
	if(navigator.appVersion.indexOf("MSIE 7.0") >= 0) {
		window.open("blank.html","_top").close();
	}else if(navigator.appVersion.indexOf("MSIE 8.0") >=0){
		window.open("blank.html","_top").close();
	}else{
		top.window.opener = top;
		top.window.close();
	}	
	//top.window.close();
}


		</script>
	</HEAD>
	<body>
		<XML id="xmlBind"></XML>
		<form name="frmThis">
			<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				<tr>
					<td>
						<table width="900" height="466" border="0" align="center" cellpadding="0" cellspacing="0">
							<tr>
								<td background="../../../images/newLogin/login_bg.jpg"><table width="900" height="460" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td height="221" colspan="3"><FONT face="굴림"></FONT></td>
										</tr>
										<tr>
											<td width="424" height="22"></td>
											<td width="174"><span class="SEARCHDATA"> <input class="INPUT_R3" id="txtLOGIN" title="ID" style="WIDTH: 163px; HEIGHT: 20px" type="text"
														maxlength="100" value="ID" size="20" name="txtLOGIN"> </span>
											</td>
											<td width="302" rowspan="3" align="left"><img src="../../../images/newLogin/bt_login.gif" id="imgLOGIN" name="imgLOGIN" style="CURSOR:hand"
													width="76" height="47"></td>
										</tr>
										<tr>
											<td height="3"></td>
											<td></td>
										</tr>
										<tr>
											<td height="22"></td>
											<td><span class="SEARCHDATA"> <input type="password" class="INPUT_R3" id="txtPASSWORD" title="Password" style="WIDTH: 163px; HEIGHT: 20px"
														maxlength="100" value="Password" size="20" name="txtPASSWORD"> </span>
											</td>
										</tr>
										<tr>
											<td></td>
											<td></td>
											<td height="10"></td>
										</tr>
										<tr>
											<td></td>
											<td class="text" style="CURSOR:hand" onclick="vbscript:Call PWDCHANGE()">* 패스워드 변경</td>
											<td></td>
										</tr>
										<tr>
											<td height="170" colspan="3"><FONT face="굴림"></FONT></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
