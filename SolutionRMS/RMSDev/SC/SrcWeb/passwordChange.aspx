<%@ Page Language="vb" AutoEventWireup="false" Codebehind="passwordChange.aspx.vb" Inherits="SC.passwordChange" %>
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
		.text1 { font-size: 8pt; color: #717171; font-family: ����; height: 10px; background-color: none;	text-align: left;text-valign: middle; }
		.INPUT_R3{ border: 1px solid #999999; color: #303030; font-family:����; font-size:9pt; background-color: #FFFFFF; }
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
	set mobjSCCOLOGIN = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '�α��� ��� Process

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
				'���̵�� �н������ ������ �����Ǿ�����������.
				If vntData(0,1) = "N" Then 
					gErrorMsgbox "�Է��Ͻ� ID �� ����� ������ ���̵��Դϴ�." & vbcrlf & "�����ڿ��� ���� �Ͻʽÿ�.","�α��ξȳ�!"
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
				'SelectRtn_PASSWORDCHANGEIDX ���� mlngRowCnt= 0 �϶� ��й�ȣ �����ϼ� ������ �ƿ� id��ü�� �������� �����Ƿ� id�θ� �˻��ϴ� SelectRtn_IDX�� �Ѵ�.
				vntPreData = mobjSCCOLOGIN.SelectRtn_IDX(gstrConfigXml,mlngPreRowCnt,mlngPreColCnt,strLoginIdx)
				
				'SelectRtn_IDX�� 0���� Ŭ���� ���̵�� �����ϴ� ���̹Ƿ� ������password �����޼���
				if mlngPreRowCnt > 0 then
					gErrorMsgbox "�Է��Ͻ� ������PW �� ��ġ���� �ʽ��ϴ�." & vbcrlf & "�����ڿ��� ���� �Ͻʽÿ�.","�α��ξȳ�!"
					Exit Sub
					
				'0�ϰ��� ���̵� �������� �����Ƿ� id���� �����޼���
				else
					gErrorMsgbox "�Է��Ͻ� ID �� �������� �ʴ� ID �Դϴ�." & vbcrlf & "�����ڿ��� ���� �Ͻʽÿ�.","�α��ξȳ�!"
					exit Sub
				end if
			End If
		End If
	
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
			<table width="372" height="244" border="0" cellpadding="0" cellspacing="0">
				<tr>
					<td height="85" align="left" valign="top" background="/images/passwordchange/pass_bg.gif"><table width="372" border="0" cellspacing="3" cellpadding="0">
							<tr>
								<td width="117">&nbsp;</td>
								<td width="165">&nbsp;</td>
								<td width="78" height="101">&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">������ ID
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" id="txtLOGIN"  title="���� �� PW" style="WIDTH: 163px; HEIGHT: 18px"
											type="text" maxlength="100" size="20" name="txtLOGIN" value=""> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">������ PW
								</td>
								<td width="165"><span class="SEARCHDATA"> <input class="INPUT_R3" id=txtPWD type="password"  title="���� �� PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtPWD"> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">������ PW
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" name="txtCHGPWD" type="password"  title="���� �� PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtCHGPWD"> </span>
								</td>
								<td>&nbsp;</td>
							</tr>
							<tr>
								<td class="text1" style="PADDING-RIGHT:0px; PADDING-LEFT:50px; PADDING-BOTTOM:0px; PADDING-TOP:0px">Ȯ�� PW
								</td>
								<td><span class="SEARCHDATA"> <input class="INPUT_R3" id=txtCONFIRMPWD type="password" title="Ȯ�� PW" style="WIDTH: 163px; HEIGHT: 18px" maxlength="100" size="20" name="txtCONFIRMPWD"> </span>
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
