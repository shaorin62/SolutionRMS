<%@ Page CodeBehind="SCCalendar.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="SC.SCCalendar" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>�޷�[����]</title> 
		<!--
'****************************************************************************************
'�ý��۱��� : ����
'����  ȯ�� : ASP.NET, VB.NET, COM+ 
'���α׷��� : SCCalendar.aspx
'��      �� : �޷¼����� ���� �˾�
'�Ķ�  ���� : ��������
'Ư��  ���� : Codebehind�� ���� - dll�� �����ϹǷ� ��𼭵��� �̿밡��
'----------------------------------------------------------------------------------------
'HISTORY    :1)  
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- Calendar�� ����ϱ� ���� css �� js �߰� -->
		<LINK href="../../../Etc/SCCalendar.css" type="text/css" rel="stylesheet">
		<script language="JavaScript" src="../../../Etc/SCCalendarPop.js"></script>
		<script language="vbscript" id="clientEventHandlersVBS">
<!--

Sub window_onload
	InitPage
End Sub

sub window_onunload
	EndPage
end sub

Sub InitPage
	Dim intNo,i
	DIm vntInParam

	with frmThis
		'IN �Ķ���� �� ��ȸ�� ���� �߰� �Ķ���� 
		vntInParam = window.dialogArguments
		if isarray(vntInParam) then
			intNo = ubound(vntInParam)					
			for i = 0 to intNo
				select case i
					case 0:	.txtDATE.value = vntInParam(0)
				end select
			next
		else
			.txtDATE.value = ""
		end if
	end with	
	'Calendar�� ȭ�鿡 ǥ��
	gshowCalendar "frmThis","txtDATE","","window_onunload()"	
End Sub

Sub EndPage
	'���õ� ���� ��ȯ
	window.returnvalue = frmThis.txtDATE.value

	window.close
End Sub
-->
		</script>
	</HEAD>
	<body>
		<!-- Calendar�� ����ϱ� ���� DIV �߰� -->
		<div class="CALTEXT" id="PopupCalendar" style="Z-INDEX: 101; WIDTH: 16px; HEIGHT: 24px"></div>
		<form id="frmThis">
			&nbsp;&nbsp; <INPUT id="txtDATE" style="Z-INDEX: 100; LEFT: 17px; WIDTH: 107px; POSITION: absolute; TOP: 2px; HEIGHT: 22px" type="hidden" size="12" readOnly>&nbsp;
		</form>
	</body>
</HTML>
