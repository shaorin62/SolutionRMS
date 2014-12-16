<%@ Page CodeBehind="SCCalendar.aspx.vb" Language="vb" AutoEventWireup="false" Inherits="SC.SCCalendar" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>달력[선택]</title> 
		<!--
'****************************************************************************************
'시스템구분 : 공통
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : SCCalendar.aspx
'기      능 : 달력선택을 위한 팝업
'파라  메터 : 현재일자
'특이  사항 : Codebehind를 없앰 - dll과 무관하므로 어디서든지 이용가능
'----------------------------------------------------------------------------------------
'HISTORY    :1)  
'****************************************************************************************
-->
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<!-- Calendar를 사용하기 위해 css 및 js 추가 -->
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
		'IN 파라메터 및 조회를 위한 추가 파라메터 
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
	'Calendar를 화면에 표시
	gshowCalendar "frmThis","txtDATE","","window_onunload()"	
End Sub

Sub EndPage
	'선택된 날자 반환
	window.returnvalue = frmThis.txtDATE.value

	window.close
End Sub
-->
		</script>
	</HEAD>
	<body>
		<!-- Calendar를 사용하기 위해 DIV 추가 -->
		<div class="CALTEXT" id="PopupCalendar" style="Z-INDEX: 101; WIDTH: 16px; HEIGHT: 24px"></div>
		<form id="frmThis">
			&nbsp;&nbsp; <INPUT id="txtDATE" style="Z-INDEX: 100; LEFT: 17px; WIDTH: 107px; POSITION: absolute; TOP: 2px; HEIGHT: 22px" type="hidden" size="12" readOnly>&nbsp;
		</form>
	</body>
</HTML>
