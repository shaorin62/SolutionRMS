
<%
		
		ModuleDir  =  Request.QueryString("ModuleDir")
		'ModuleDir  =  "SC"
		'Response.Write ModuleDir
		ReportName =  Request.QueryString("ReportName")
		'ReportName =  "SCMENU.rpt"
		'Response.Write ReportName
		Params     =  Request.QueryString("Params")
		'Params     =  "SJCC:생산"
		'Response.Write Params
%>
<%
'====================================================================================
' AlwaysRequiredSteps.asp
' Create the Crystal Reports Objects
'====================================================================================
' CREATE THE APPLICATION OBJECT                                                                     
If Not IsObject (session("oApp")) Then                              
  Set session("oApp") = Server.CreateObject("CrystalRuntime.Application.9")
End If                                                               

'=========================================================
'Report Path 지정
'=========================================================
Path = Request.ServerVariables("PATH_TRANSLATED")                     
basePath = Right(Path, 27)
Path = Replace(Path,basePath,"") & ModuleDir &"\Rpt\"

'======================================================================
' Crystal Report 파일이 존재하는 체크한다.
'======================================================================
on error resume next 
Set objFileSys = Server.CreateObject("Scripting.FileSystemObject")
   
Dim strtemp  
    strtemp =   reportname & " 이 없습니다." 
If Not objFileSys.FileExists(path & reportname ) Then
	   
	Set strFilePath = Nothing   		
	Response.Redirect "SCRTERR.asp?MSG=" & Replace("리포트의 파일명[" & ReportName & "]이 잘못되었습니다.", " ", "%20")
Else
	Set strFilePath = Nothing
End If
	

If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if

On error resume next

Set session("oRpt") = session("oApp").OpenReport(Path & ReportName, 1)

If Err.Number <> 0 Then
  Response.Write "Error Occurred creating Report Object: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Abandon
  Response.End
End If

session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False
session("oRpt").DiscardSavedData

%>
<% 
'==================================================================
'==================================================================
' 파라미터 세팅
'==================================================================
'==================================================================

Dim vntParam, iRow
vntParam = Split(Params, ":")    

If IsArray(vntParam) Then
	For iRow = 0 To Ubound(vntParam,1)
		''Response.Write "vntParam"& iRow &":" & vntParam(iRow)
		session("oRpt").ParameterFields.Item(iRow + 1).SetCurrentValue Cstr(vntParam(iRow))
	Next
End IF

'===========================수정필요
Const DBServer = "SFAR"
Const DBName   = ""
Const UserID   = "sysdba"
Const UserPWD  = "sysdba"
'===========================수정필요

set Database=session("oRpt").Database
 '=========================================================================
 '>>>>>>>>>>>>>>>>>>>>>>>>>>>> ServerName , DataBase Name,  User ID, Passwd 
 '=========================================================================
Call Database.Tables.Item(1).SetLogOnInfo( DBServer , DBName , UserID , UserPWD )
 
Set Database = nothing
'==================================================================
'==================================================================
%>
<%
'====================================================================================
' MoreRequiredSteps.asp
' Retrieve the Records and Create the "Page on Demand" Engine Object
'====================================================================================

On Error Resume Next

session("oRpt").ReadRecords

If Err.Number <> 0 Then                                               
  Response.Write "Error Occurred Reading Records: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Abandon
  Response.End
Else
  If IsObject(session("oPageEngine")) Then                              
  	set session("oPageEngine") = nothing
  End If
  set session("oPageEngine") = session("oRpt").PageEngine
End If
%>
<%
'====================================================================================
' SmartViewerActiveX.aspx
'====================================================================================
%>
<HTML>
	<HEAD>
		<TITLE>Crystal Reports ActiveX Viewer</TITLE>
	</HEAD>
	<BODY BGCOLOR="C6C6C6" ONUNLOAD="CallDestroy();" leftmargin="0" topmargin="0" rightmargin="0" bottommargin="0">
		<OBJECT ID="CRViewer" CLASSID="CLSID:2DEF4530-8CE6-41c9-84B6-A54536C90213" WIDTH="100%" HEIGHT="100%" CODEBASE="/SFARDev/SC/Setup/activexviewer.cab#Version=9,2,0,442" >
			<PARAM NAME="lastProp" VALUE="600">
			<PARAM NAME="_cx" VALUE="265">
			<PARAM NAME="_cy" VALUE="185">
			<PARAM NAME="DisplayGroupTree" VALUE="-1">
			<PARAM NAME="DisplayToolbar" VALUE="-1">
			<PARAM NAME="EnableGroupTree" VALUE="0">
			<PARAM NAME="EnableNavigationControls" VALUE="-1">
			<PARAM NAME="EnableStopButton" VALUE="-1">
			<PARAM NAME="EnablePrintButton" VALUE="-1">
			<PARAM NAME="EnableZoomControl" VALUE="-1">
			<PARAM NAME="EnableCloseButton" VALUE="-1">
			<PARAM NAME="EnableProgressControl" VALUE="-1">
			<PARAM NAME="EnableSearchControl" VALUE="-1">
			<PARAM NAME="EnableRefreshButton" VALUE="0">
			<PARAM NAME="EnableDrillDown" VALUE="-1">
			<PARAM NAME="EnableAnimationControl" VALUE="-1">
			<PARAM NAME="EnableSelectExpertButton" VALUE="-1">
			<PARAM NAME="EnableToolbar" VALUE="-1">
			<PARAM NAME="DisplayBorder" VALUE="-1">
			<PARAM NAME="DisplayTabs" VALUE="-1">
			<PARAM NAME="DisplayBackgroundEdge" VALUE="-1">
			<PARAM NAME="SelectionFormula" VALUE="">
			<PARAM NAME="EnablePopupMenu" VALUE="-1">
			<PARAM NAME="EnableExportButton" VALUE="-1">
			<PARAM NAME="EnableSearchExpertButton" VALUE="0">
			<PARAM NAME="EnableHelpButton" VALUE="0">
			<PARAM NAME="LaunchHTTPHyperlinksInNewBrowser" VALUE="-1">
			<PARAM NAME="EnableLogonPrompts" VALUE="-1">
			<!--   
<PARAM NAME="EnableRefreshButton" VALUE="-1">
<PARAM NAME="EnableGroupTree" VALUE=1>
<PARAM NAME="DisplayGroupTree" VALUE=1>
<PARAM NAME="EnablePrintButton" VALUE=1>
<PARAM NAME="EnableExportButton" VALUE=1>
<PARAM NAME="EnableDrillDown" VALUE=1>
<PARAM NAME="EnableSearchControl" VALUE=-1>
<PARAM NAME="EnableAnimationControl" VALUE=1>
<PARAM NAME="EnableZoomControl" VALUE=1>
	 -->
		</OBJECT>
		<SCRIPT LANGUAGE="VBScript">
<!--
Sub Window_Onload	
	On Error Resume Next
	Dim webBroker
	Set webBroker = CreateObject("WebReportBroker9.WebReportBroker")
	if ScriptEngineMajorVersion < 2 then
		window.alert "IE 3.02 users on NT4 need to get the latest version of VBScript or install IE 4.01 SP1. IE 3.02 users on Win95 need DCOM95 and latest version of VBScript, or install IE 4.01 SP1. These files are available at Microsoft's web site."
	else
		Dim webSource
		Set webSource = CreateObject("WebReportSource9.WebReportSource")
		webSource.ReportSource = webBroker
		webSource.URL = "SCRTSERVER.asp"
		webSource.PromptOnRefresh = True
		CRViewer.ReportSource = webSource
	end if
	CRViewer.ViewReport
End Sub
-->
		</SCRIPT>
		<script language="javascript">
function CallDestroy()
{
	window.open("SCRTCLEAN.asp");
}
		</script>
	</BODY>
</HTML>