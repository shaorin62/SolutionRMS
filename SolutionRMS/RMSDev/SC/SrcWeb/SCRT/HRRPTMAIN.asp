<%@ LANGUAGE="VBSCRIPT" %>
<%
'----------------------------------------------------------------------
' ĳ�� Clear
'----------------------------------------------------------------------
Response.Expires = 0
Response.Buffer = true
'----------------------------------------------------------------------
' Crystal Report ȭ�ϸ� ����
'----------------------------------------------------------------------
DSN	    =  Request.QueryString("DSN")
Module	    =  Request.QueryString("ModuleDir")
RptName     =  Request.QueryString("ReportName")
Params      =  Request.QueryString("Params")
Params      =  Replace(Params,"*","%")
Opt	    =  Request.QueryString("Opt")

%>
<%  
'----------------------------------------------------------------------
' Crystal Report ������ �����ϴ� üũ�Ѵ�.
'----------------------------------------------------------------------
Path = Request.ServerVariables("PATH_TRANSLATED")                     
basePath = Right(Path, 28)
Path = Replace(Path, basePath, "") & Module &"\Rpt\"


'Response.Write ("DSN:" & DSN & " Module:" & Module & " RptName:" & RptName & " Params:" & Params & " Opt:" & Opt & " Path:" & Path)

On error resume next     
Dim strtemp              
   Set objFileSys = Server.CreateObject("Scripting.FileSystemObject")
   strtemp =   RptName & " �� �����ϴ�." 
    If Not objFileSys.FileExists(path & RptName ) Then
	 	Set objFileSys = Nothing   		
		Response.Redirect "SCRTERRMSG.asp?MSG=" & Replace("����Ʈ�� ���ϸ�[" & RptName & "]�� �߸��Ǿ����ϴ�.", " ", "%20")
	Else
		Set objFileSys = Nothing
	End If
	

	
'===================================================================================
'Create the Crystal Reports Objects
'=================================================================================== '
' CREATE THE APPLICATION OBJECT        
If Not IsObject (session("oApp")) Then                              
  Set session("oApp") = Server.CreateObject("CrystalRuntime.Application.10")
End If                                                               
' CREATE THE REPORT OBJECT                                            
' The Report object is created by calling the Application object's OpenReport method.
If IsObject(session("oRpt")) then
	Set session("oRpt") = nothing
End if    
Set session("oRpt") = session("oApp").OpenReport(Path  & RptName, 1)

If Err.Number <> 0 Then
  Response.Write "Error Occurred creating Report Object: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Contents.Remove("oRpt")
  Session.Contents.Remove("oApp")
  Response.End
End If

session("oRpt").MorePrintEngineErrorMessages = False
session("oRpt").EnableParameterPrompting = False
session("oRpt").DiscardSavedData

'--------------------
' �Ķ���� ����
'--------------------
Dim vntParam, iRow
    vntParam = Split(Params, ":") '------>>������ üũ�Ͻÿ�(Default=":")
If IsArray(vntParam) Then
	For iRow = 0 To Ubound(vntParam,1)
		session("oRpt").ParameterFields.Item(iRow + 1).SetCurrentValue Cstr(vntParam(iRow))
	Next
End IF
'--------------------
' Database ����
'--------------------
'Dim vntDBParam
'	vntDBParam = Split(DSN, ";")  '------>>DB Connection ����(Default=";")  
'	strDB = vntDBParam(2) : strUID = vntDBParam(0) : strPWD = vntDBParam(1)

'=========================================================================
'>>>>>>>>>>>>>>>>>>>>>>>>>>>> ServerName , DataBase Name,  User ID, Passwd 
'=========================================================================
Set Database=session("oRpt").Database    
Call Database.Tables.Item(1).SetLogOnInfo("PHOTO","","SYSDBA","SYSDBA")

If Err.Number <> 0 Then
  Response.Write "Error Occurred creating Report Object: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Contents.Remove("oRpt")
  Session.Contents.Remove("oApp")
  Set Database = nothing
  Response.End
End If

Set Database = nothing

'=========================================================================
' Retrieve the Records and Create the "Page on Demand" Engine Object
'========================================================================= 
session("oRpt").ReadRecords

If Err.Number <> 0 Then                                               
  Response.Write "Error Occurred Reading Records: " & Err.Description
  Set Session("oRpt") = nothing
  Set Session("oApp") = nothing
  Session.Contents.Remove("oRpt")
  Session.Contents.Remove("oApp")
  Response.End
Else
  If IsObject(session("oPageEngine")) Then                              
  	set session("oPageEngine") = nothing
  End If
  
  set session("oPageEngine") = session("oRpt").PageEngine
End If

%>		 
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<META name="GENERATOR" content="Microsoft Visual Studio .NET 7.0">
		<META name="CODE_LANGUAGE" content="Visual Basic 7.0">
		<META name="vs_defaultClientScript" content="VBScript">
		<META name="vs_targetSchema" content="http://schemas.microsoft.com/intellisense/ie5">
		<TITLE>Crystal Reports ActiveX Viewer</TITLE>
		 <SCRIPT language="vbscript" id="clientEventHandlersVBS">      
		<!--
			Option Explicit

			Dim rTimer
			Dim PrinterTimer
			Dim PageOne
			Dim webBroker
			Dim webSource

			PageOne = True

			Sub Window_Onload()     			    
   				On Error Resume Next          									
				If "<%= Request.QueryString("OPT")%>" = "B" THEN            '����(A-Crystal View ��� �μ�,B-�����μ�)
					CRViewer.style.width = "1"
					CRViewer.style.height = "1"
				else
					CRViewer.style.width = "100%"   '�̸����� ���
					CRViewer.style.height = "100%"
				End If   
				
				Dim webBroker
				Set webBroker = CreateObject("CrystalReports10.WebReportBroker.1")
				if ScriptEngineMajorVersion < 2 then
					window.alert "IE 3.02 users on NT4 need to get the latest version of VBScript or install IE 4.01 SP1. IE 3.02 users on Win95 need DCOM95 and latest version of VBScript, or install IE 4.01 SP1. These files are available at Microsoft's web site."
				else
					Dim webSource
					Set webSource = CreateObject("CrystalReports10.WebReportSource.1")
					webSource.ReportSource = webBroker
					webSource.URL = "RDCrptserver10.asp"
					webSource.PromptOnRefresh = True
					CRViewer.ReportSource = webSource
				end if
				CRViewer.ViewReport
			End Sub

			'============================================
			' CRViewer�� Download�� ��������....
			'============================================
			Sub CRViewer_DownloadFinished(ByVal downloadType)
			' On Error Resume next
				if "<%= Request.QueryString("OPT")%>" = "B" THEN  
					If downloadType = 1 and PageOne Then
						PageOne = False
						rTimer = window.settimeout ("OnMyTimeOut()",1000)
					End If
				Else 
					    rTimer = window.setTimeout ("OnMyTimeOut()",600000)
				End if
			End Sub

			Sub OnMyTimeOut()
				'On Error Resume next
				If Not CRViewer.IsBusy Then
					Window.ClearTimeout(rTimer)
					CRViewer.PrintReport
					PrinterTimer = window.SetTimeOut( "OnPrinterTimeOut", 1000)
					parent.opener = parent 
					parent.close
				End If
			end sub

			Sub OnPrinterTimeOut()
				If Not CRViewer.IsBusy then
					window.ClearTimeOut(PrinterTimer)
				End If
			end sub

			Sub window_onUnload()
				call EndPage()
			End Sub

			Sub EndPage()
				Set webBroker = Nothing
				Set webSource = Nothing
			
				window.open "SCRTCLEAN.asp","","left=10000"  'SESSION CLEAN
			End Sub

			Sub CRBug_onclick()
				self.close()
				window.open("../../Setup/SetupCrystal.htm")	
			End Sub
			-->
		</SCRIPT>
	</HEAD>
	<BODY leftmargin="0" topmargin="0">
		<TABLE border="0" width="100%" height="100%"  cellpadding="0" cellspacing="0" ID="Table1">
			<!--<TR bgcolor="#EEEEEE"><TD align="center">  			
			<A id="CRBug" href="#"><FONT size="2" face="����" color="#0000FF">�� ũ����Ż ����Ʈ ���α׷� �ȳ�</FONT></A></TD></TR>
			-->
			<TR valign="top" bgcolor="#EEEEEE">
				<TD width="100%" height="100%"><!--cabȭ���� ��ο� �������� �� Codebase�������-->
					<OBJECT ID="CRViewer" CLASSID="CLSID:A1B8A30B-8AAA-4a3e-8869-1DA509E8A011"	WIDTH="100%" HEIGHT="100%"
					 CODEBASE="/crystalreportviewers10/ActiveXControls/ActiveXViewer.cab#Version=10,0,5,650" VIEWASTEXT>
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
					</OBJECT>
				</TD>
			</TR>
		</TABLE>
	</BODY>
</HTML>