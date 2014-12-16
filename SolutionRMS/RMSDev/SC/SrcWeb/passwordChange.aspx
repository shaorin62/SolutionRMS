<%@ Page Language="vb" AutoEventWireup="false" Codebehind="passwordChange.aspx.vb" Inherits="SC.passwordChange" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>비밀번호 변경</title> 
		<!--
'****************************************************************************************
'시스템구분 : SFAR/공통/공통코드 팝업
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMPOP1.aspx
'기      능 : JOBNO 조회를 위한 팝업
'파라  메터 : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , 조회추가필드, 현재 사용중인 것만 조회할지 여부,
'			  코드 사용시점, 코드Like할지 여부
'특이  사항 : 
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2003/05/21 By ParkJS
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<style type="text/css"> .login { width:141; height: 18px; padding: 2px 1px 0px 2px; border:1 solid #9bb7d9; background-color: #6994c7; font-size: 12px;color:#edebeb; }
		.text1 { font-size: 8pt; color: #717171; font-family: 돋움; height: 10px; background-color: none;	text-align: left;text-valign: middle; }
		.INPUT_R3{ border: 1px solid #999999; color: #303030; font-family:돋움; font-size:9pt; background-color: #FFFFFF; }
		</style>
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../Etc/SCUIClass.inc" -->
		<!-- #INCLUDE VIRTUAL="../../Etc/SCClient.inc" -->
		<!-- Farpoint SpreadSheet License :spr32x60.ocx -->
		<OBJECT id="Microsoft_Licensed_Class_Manager_1_0" classid="clsid:5220cb21-c88d-11cf-b347-00aa00a28331" >
		</OBJECT>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--
option explicit

Dim mobjSCCOLOGIN 
Dim mlngRowCnt, mlngColCnt
DIm mblnUseOnly,mstrUseDate,mstrFields,mblnLikeCode
Const meTab = 9
Dim mlngPreRowCnt
Dim mlngPreColCnt
Dim mlngClRowCnt
Dim mlngClColCnt
Dim mstrLOGINCHK

'-----------------------------
' 이벤트 프로시져 
'-----------------------------	
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub imgSave_onclick
	ProcessRtn
End Sub

Sub ImgCancel_onclick
	Window_OnUnload
End Sub

Sub imgClose_onclikc
	EndPage
End Sub

'-----------------------------
' UI업무 프로시져 
'-----------------------------	
sub InitPage()

	dim vntInParam
	dim intNo,i
	set mobjSCCOLOGIN = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process

	with frmThis
		
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		
		for i = 0 to intNo
			select case i
				case 0 : .txtLOGIN.value = vntInParam(i)	
			end select
		next
		
		If .txtLOGIN.value = "" Then
			.txtLOGIN.focus()
		Else
			.txtPWD.focus() 
		End If
      
	end with	
end sub

Sub EndPage()
	set mobjSCCOLOGIN = Nothing
	gEndPage
End Sub

Sub ProcessRtn
	call window.execScript("checkForSubmit()","JavaScript")
End Sub

Sub WorkEndchk
	Dim intRtn
	Dim strDate
	Dim strLoginIdx , strPwdIdx
	Dim strClipping
	Dim vntData , vntPreData
	
	with frmThis
		mlngRowCnt=clng(0)
		mlngColCnt=clng(0)
		mlngPreRowCnt =clng(0)
		mlngPreColCnt =clng(0)
		
		strDate = gNowDate
		strDate = replace(strDate,"-","")
		
		strLoginIdx = Trim(.txtLOGIN.value)
		strPwdIdx = Trim(.txtPWD.value)
		If Len(strLoginIdx) = 5 Then strLoginIdx = "000" & strLoginIdx
		
		vntData = mobjSCCOLOGIN.SelectRtn_PASSWORDCHANGEIDX(gstrConfigXml,mlngRowCnt,mlngColCnt,strLoginIdx,strPwdIdx)
		
		if not gDoErrorRtn ("SelectRtn_PASSWORDCHANGEIDX") then
			If mlngRowCnt > 0  Then
				'아이디와 패스워드는 맞지만 중지되었을수도있음.
				If vntData(0,1) = "N" Then 
					gErrorMsgbox "입력하신 ID 는 사용이 중지된 아이디입니다." & vbcrlf & "관리자에게 문의 하십시오.","로그인안내!"
					.txtLOGIN.value = ""
					.txtPWD.value = ""
					.txtLOGIN.focus()
					exit Sub
				End If
 				
				gstrUsrID = vntData(1,1)
				gstrEmpNo = vntData(1,1)
				gstrUsrName = vntData(2,1)
				strClipping = vntData(4,1)
				
				gSetSession gstrUsrID,gstrEmpNo,gstrUsrName
				
				gInitPageSetting mobjSCGLCtl,"MC"
			ELSE 
				'SelectRtn_PASSWORDCHANGEIDX 에서 mlngRowCnt= 0 일때 비밀번호 오류일수 있으나 아예 id자체가 없을수도 있으므로 id로만 검색하는 SelectRtn_IDX를 한다.
				vntPreData = mobjSCCOLOGIN.SelectRtn_IDX(gstrConfigXml,mlngPreRowCnt,mlngPreColCnt,strLoginIdx)
				
				'SelectRtn_IDX이 0보다 클경우는 아이디는 존재하는 것이므로 변경전password 오류메세지
				if mlngPreRowCnt > 0 then
					gErrorMsgbox "입력하신 변경전PW 는 일치하지 않습니다." & vbcrlf & "관리자에게 문의 하십시오.","로그인안내!"
					Exit Sub
					
				'0일경우는 아이디가 존재하지 않으므로 id존재 오류메세지
				else
					gErrorMsgbox "입력하신 ID 는 존재하지 않는 ID 입니다." & vbcrlf & "관리자에게 문의 하십시오.","로그인안내!"
					exit Sub
				end if
			End If
		End If
	
		intRtn = mobjSCCOLOGIN.ProcessRtn_PwdUpdate(gstrConfigXml,strLoginIdx,trim(.txtCHGPWD.value),strDate)
		
		if not gDoErrorRtn ("ProcessRtn_PwdUpdate") then
			gOkMsgBox "비밀번호가 변경 되었습니다.","변경안내"
			EndPage 
			window.returnvalue = "T"
		End If
		
	End with
End Sub


-->
	</script>
	<SCRIPT language="JavaScript">
<!--

	function checkForSubmit() {
		
		var frm = document.forms[0];
		var bln = true;
		var regexp = /^[a-z\d]{8,12}$/i;
		var regexp_str = /[a-z]/i;
		var regexp_num = /[\d]/i;

		if (frm.txtCHGPWD.value.length < 8 ) {
			alert("변경할 비밀번호는 8~12자 사이로 입력하세요");
			return false ;
		}
		if (frm.txtPWD.value == frm.txtCHGPWD.value) {
			alert("기존과 동일한 비밀번호로 설정할 수 없습니다.");
			return false;
		}
		if (frm.txtCHGPWD.value == frm.txtLOGIN.value){
			alert("아이디와 동일한 비밀번호는 설정할 수 없습니다..");
			return false ;
		}
		if (frm.txtCHGPWD.value != frm.txtCONFIRMPWD.value){
			alert("변경비밀번호 와 확인비밀번호가 동일하지 않습니다.");
			return false ;
		}
		if (!(regexp.test(frm.txtCHGPWD.value) && regexp_str.test(frm.txtCHGPWD.value) && regexp_num.test(frm.txtCHGPWD.value))) {
			alert("비밀번호는 영문자,숫자의 조합만으로 작성하세요.");
			return false ;
		}
		
		WorkEndchk();
	}
//-->
		</SCRIPT>
	</HEAD>
		<body class="base" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM id="frmThis">
			<table width="372" height="244" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="85" align="left" valign="top" background="/images/passwordchange/pass_bg.gif"><table width="372" border="0" cellspacing="3" cellpadding="0">
							<tr>
								<td width="117">&nbsp;</td>
								<td width="165">&nbsp;</td>
								<td width="78" height="101">&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">변경전 ID
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" id="txtLOGIN"  title="변경 전 PW" style="WIDTH: 163px; HEIGHT: 18px"
											type="text" maxlength="100" size="20" name="txtLOGIN" value=""> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">변경전 PW
								</td>
								<td width="165"><span class="SEARCHDATA"> <input class="INPUT_R3" id=txtPWD type="password"  title="변경 전 PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtPWD"> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">변경후 PW
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" name="txtCHGPWD" type="password"  title="변경 후 PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtCHGPWD"> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">확인 PW
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" id=txtCONFIRMPWD type="password" title="확인 PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtCONFIRMPWD"> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td height="3" colspan="3"></td>
							</tr>
							<tr>
								<td>&nbsp;</td>
								<td align="center"><img src="/images/passwordchange/btn_save.gif" width="57" height="23" id="ImgSave">&nbsp;<img src="/images/passwordchange/btn_cancel.gif" width="57" height="23" id="ImgCancel"></td>
								<td>&nbsp;</td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</FORM>
	</body>
</HTML>
