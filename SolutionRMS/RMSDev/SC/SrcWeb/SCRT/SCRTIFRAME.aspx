<%@ Page Language="vb" AutoEventWireup="false" Codebehind="SCRTIFRAME.aspx.vb" Inherits="SC.SCRTIFRAME" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>iframeTEST</title>
		<meta name="GENERATOR" content="Microsoft Visual Studio .NET 7.1">
		<meta name="CODE_LANGUAGE" content="Visual Basic .NET 7.1">
		<meta name="vs_defaultClientScript" content="VBScript">
		<meta name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
        <!--#INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script id=clientEventHandlersVBS language=vbscript>
<!--

Sub showReport_onclick
		'gShowIFrameReport(아이프레임네임, 모듈,레포트명,파라미터,옵션)
    	'gShowiFrameReport(iFrameName, Module,ReportName,Params, Opt)
     	 gShowiFrameReport ifrTest, "SC", "SCMENU.rpt", "HPC:영업", "A"
End Sub
-->
</script>
</HEAD>
	
	<body MS_POSITIONING="FlowLayout">
		<form id="Form1" method="post" runat="server">    		
		<input type="button" id="showReport" value="Iframe Test Button">
			<iframe id="ifrtest" src="" style="WIDTH: 880px; HEIGHT: 608px"></iframe>
		</form>
	</body>
</HTML>
