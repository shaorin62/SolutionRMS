<%@ Page Language="vb" AutoEventWireup="false" Codebehind="passwordChange_old.aspx.vb" Inherits="SC.passwordChange_old" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>��й�ȣ ����</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : SFAR/����/�����ڵ� �˾�
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : PDCMPOP1.aspx
'��      �� : JOBNO ��ȸ�� ���� �˾�
'�Ķ�  ���� : CATEGORY ID OR NAME, SC_CATEGORY_GROUP , ��ȸ�߰��ʵ�, ���� ������� �͸� ��ȸ���� ����,
'			  �ڵ� ������, �ڵ�Like���� ����
'Ư��  ���� : 
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
	.pass { width:141; height: 18px; padding: 2px 1px 0px 2px; border:1 solid #9bb7d9; background-color: #6994c7; font-size: 12px;color:#edebeb; }
		</style>
		<!-- UI ���� ActiveX COM -->
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
' �̺�Ʈ ���ν��� 
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
' UI���� ���ν��� 
'-----------------------------	
sub InitPage()

	dim vntInParam
	dim intNo,i
	set mobjSCCOLOGIN		 = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '�α��� ��� Process

	with frmThis
		
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
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
	Dim strLoginIdx
	strDate = gNowDate
	strDate = replace(strDate,"-","")
	
	with frmThis
	strLoginIdx = Trim(.txtLOGIN.value)
	If Len(strLoginIdx) = 5 Then strLoginIdx = "000" & strLoginIdx
	
		intRtn = mobjSCCOLOGIN.ProcessRtn_PwdUpdate(gstrConfigXml,strLoginIdx,trim(.txtCHGPWD.value),strDate)
		if not gDoErrorRtn ("ProcessRtn_PwdUpdate") then
			gOkMsgBox "��й�ȣ�� ���� �Ǿ����ϴ�.","����ȳ�"
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
			alert("������ ��й�ȣ�� 8~12�� ���̷� �Է��ϼ���");
			return false ;
		}
		if (frm.txtPWD.value == frm.txtCHGPWD.value) {
			alert("������ ������ ��й�ȣ�� ������ �� �����ϴ�.");
			return false;
		}
		if (frm.txtCHGPWD.value == frm.txtLOGIN.value){
			alert("���̵�� ������ ��й�ȣ�� ������ �� �����ϴ�..");
			return false ;
		}
		if (frm.txtCHGPWD.value != frm.txtCONFIRMPWD.value){
			alert("�����й�ȣ �� Ȯ�κ�й�ȣ�� �������� �ʽ��ϴ�.");
			return false ;
		}
		if (!(regexp.test(frm.txtCHGPWD.value) && regexp_str.test(frm.txtCHGPWD.value) && regexp_num.test(frm.txtCHGPWD.value))) {
			alert("��й�ȣ�� ������,������ ���ո����� �ۼ��ϼ���.");
			return false ;
		}
		
		WorkEndchk();
	}
//-->
		</SCRIPT>
	</HEAD>
	<body class="base" leftMargin="0" topMargin="0" rightMargin="0">
		<FORM id="frmThis">
			<table width="100%" height="100%" border="0" cellspacing="0" cellpadding="0">
				
				<tr>
					<td align="center" valign="top">
						<!-- �α��ν��� ----------------------------------->
						<table width="590" border="0" cellspacing="0" cellpadding="0">
							<tr>
								<td height="126" background="/images/login/log_bg01.gif">&nbsp;</td>
							</tr>
							<tr>
								<td height="56" background="/images/login/log_confirmbg02.gif"><table width="367" height="56" border="0" cellpadding="0" cellspacing="0">
										<tr>
											<td width="95">&nbsp;</td>
											<td width="150"><table border="0" cellspacing="0" cellpadding="0">
													<tr>
														<td width="141"><input id=txtLOGIN name="txtLOGIN" type="text" class="login" value=""></td>
													</tr>
													<tr>
														<td height="8"></td>
													</tr>
													<tr>
														<td><input id=txtPWD name="txtPWD" type="password" class="pass" value=""></td>
													</tr>
												</table>
											</td>
											<td width="56"><img src="/images/login/btn_confirmSave.gif" id="ImgSave"></td>
											<td width="61"><img src="/images/login/btn_confirmcancel.gif" id="ImgCancel"></td>
										</tr>
									</table>
								</td>
							</tr>
							<tr>
								<td height="166" valign="top" background="/images/login/log_confirmbg03.gif"><table width="304" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="4" height="3"></td>
										</tr>
										<tr>
											<td></td>
											<td width="90" height="11"></td>
											<td width="88"><input id=txtCHGPWD name="txtCHGPWD" type="password" class="pass" value=""></td>
											<td width="71"></td>
										</tr>
									</table>
									<table>
										<tr>
											<td heigh="3"></td>
										</tr>
									</table>
									<table width="304" border="0" cellspacing="0" cellpadding="0">
										<tr>
											<td colspan="4" height="3"></td>
										</tr>
										<tr>
											<td></td>
											<td width="90" height="11"></td>
											<td width="88"><input id=txtCONFIRMPWD name="txtCONFIRMPWD" type="password" class="pass" value=""></td>
											<td width="71"></td>
										</tr>
									</table>
								</td>
							</tr>
						</table>
						<!-- �α��γ�----------------------------------->
					</td>
				</tr>

			</table>
		</FORM>
	</body>
</HTML>
