<%@ Page Language="vb" AutoEventWireup="false" Codebehind="leftmenu_productreport.aspx.vb" Inherits="SC.leftmenu_productreport" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>열정과 패기로 ! Beyond SK ! RMS</TITLE>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<LINK href="/css/style.css" type="text/css" rel="stylesheet">
			<!-- #INCLUDE VIRTUAL="../../Etc/SCClient.inc" -->
			<!-- #INCLUDE VIRTUAL="../../Etc/SCUIClass.inc" -->
			<script language="vbscript" id="clientEventHandlersVBS">
<!--
Dim mlngRowCnt
Dim mlngColCnt
Dim mobjSCCOLOGIN

Sub window_onload
	'if (parent.leftFrame_Check.document.all.chkLeftHidden.checked) then
	'	menuhide()
	'else
	'	menuVisible()
	'end if
	menuVisible()
	initpage
	Call left_auth()
End Sub

Sub initpage
	Dim vntData
	Dim vntPreData 
	Dim strDEFAULTPGM
	
	set mobjSCCOLOGIN	 = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process
	
	gInitComParams mobjSCGLCtl,"MC"
	
	'on error resume next
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	vntData = mobjSCCOLOGIN.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
		if mlngRowCnt > 0 Then
		'msgbox vntData(0,1)
		'msgbox vntData(2,1)
		'msgbox vntData(3,1)
		document.getElementById("loginname").innerHTML = vntData(2,1) & " 님"
		end if
   	end if
   	strDEFAULTPGM = "PDPR1"
	PGM_DefaultMenuAuth(strDEFAULTPGM)
   	
End Sub

Sub EndPage()
	set mobjSCCOLOGIN = Nothing
	gEndPage
End Sub

Sub imgLogOut_onclick
	gInitPageSetting mobjSCGLCtl,"MD"
	Call Win_MainClose()
End Sub

sub ShowWindow (byval strPageURL, byval strWindowName, byval lngWidth, byval lngHeight, byval strOptions)
	dim lngTop, lngLeft
	if strOptions="" then
		strOptions = "toolbar=no, location=no, menubar=no, scrollbars=no, status=yes, resizable=yes"
	end if
	'화면의 중앙에 위치시킨다.
	'lngTop = (window.screen.height - lngHeight) / 2
	lngTop = 105
	'lngLeft = (window.screen.width - lngWidth) / 2
	lngLeft = 0
	strOptions = strOptions & ", top=" & lngTop & ", left=" & lngLeft & ", width=" & lngWidth-10 & ", height=" & lngHeight
	window.open  strPageURL,strWindowName,strOptions
end sub

Sub PGM_Auth(byval strMENU) 
	Dim vntData
	Dim vntPreData 
	Dim strVAL
	'on error resume next
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	vntData = mobjSCCOLOGIN.SelectRtn_AUTH(gstrConfigXml,mlngRowCnt,mlngColCnt,strMENU)
	if not gDoErrorRtn ("SelectRtn_AUTH") then	
		if mlngRowCnt > 0 Then
			strVAL = "T"
		Else
			strVAL = "F"
		'document.getElementById("loginname").innerHTML = vntData(2,1) & " 님"
		end if
		Call auth(strVAL,strMENU) 
   	end if
End Sub

Sub PGM_DefaultMenuAuth(byval strMENU) 
	Dim vntData
	Dim vntPreData 
	Dim strVAL
	'on error resume next
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	vntData = mobjSCCOLOGIN.SelectRtn_AUTH(gstrConfigXml,mlngRowCnt,mlngColCnt,strMENU)
	if not gDoErrorRtn ("SelectRtn_AUTH") then	
		if mlngRowCnt > 0 Then
			strVAL = "T"
		Else
			strVAL = "F"
		
		end if
		'Call defaultauth(strVAL,strMENU) 
   	end if
End Sub
-->
</script>
<script language="javascript">
<!--
	var gStrLeftmenu;
	var gStrHidei;
	gStrLeftmenu = "";
	gStrHidei = 184;
	var strImg = 0;

	function Win_MainClose(){
		top.close();
	}	

	function menuclick(A){
		if(A.style.display=="none"){
			A.style.display="";
			strImg = 1;
		}else{
			A.style.display="none";
			strImg = 0;
		}
	}

	function plusminus(){
		if(smenu01.style.display=="none"){
			smenu01.style.display="";
		}else{
			smenu01.style.display="none";
		}
		/*
		if(smenu02.style.display=="none"){
			smenu02.style.display="";
		}else{
			smenu02.style.display="none";
		}
		if(smenu03.style.display=="none"){
			smenu03.style.display="";
		}else{
			smenu03.style.display="none";
		}
		if(smenu04.style.display=="none"){
			smenu04.style.display="";
		}else{
			smenu04.style.display="none";
		}
		if(smenu05.style.display=="none"){
			smenu05.style.display="";
		}else{
			smenu05.style.display="none";
		}
		*/
	}

	function detailclick (strURL, strLOC) {
		var i;
		var lngWidth;
		var lngHeight;
		var lngTop;
		var lngLeft;
		var strOptions;
		
		if (parent.leftFrame_Check.document.all.chkWindowOpen.checked) {
		
			lngWidth = "1100";
			lngHeight = "768";
			
			lngTop = (window.screen.height - lngHeight) / 2;
			lngLeft = (window.screen.width - lngWidth) / 2;
			
			strOptions = " toolbar=no, location=no, menubar=no, scrollbars=no, status=no, resizable=yes, width=" + lngWidth + " ,height=" + lngHeight + " ,top=" + lngTop + " ,left=" + lngLeft ;
						
			window.open (strURL,"",strOptions);
			
		}else{
			parent.mainFrame.location.href = strURL;
		}
		
		
		for(i=1; i<7; i++) {
			if (strLOC.substr(1) == i ){
				document.getElementById("b"+i).style.backgroundColor = "F7DEDE";
				document.getElementById("b"+i).style.fontWeight = "bold";
			}else{
				document.getElementById("b"+i).style.backgroundColor = "";
				document.getElementById("b"+i).style.fontWeight = "";
			}
		}

	}
	function left_auth() {
	var strAuth1;
	var i;
	var strMENU
		for(i=1;i<7; i++){
			strMENU = "PDPR" + i;
			PGM_Auth(strMENU);
		}
		
		
	}

	function auth(strTT,strMENU) {
	var i;	
		if(strTT == "F") {
			document.getElementById(strMENU+"_V").style.display = "none";
		}
		//alert(strTT);
		//alert(strMENU);
	}
	function defaultauth(strTT,strMENU) {
		if(strTT == "F") {
			parent.mainFrame.location.href="http://10.110.10.89:8080/SC/SrcWeb/SCNT/GList.asp"
		} else
		{
			parent.mainFrame.location.href="http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMCUSTLISTSEARCH.aspx"
   			document.getElementById("b1").style.backgroundColor = "F7DEDE"
		}
		
	}
	
	//메뉴숨기기
	function menuhide(){
	var strColEnd,strCols;
	var i
	var strColv;
		strColv="2%,98%";
		
			//gStrHidei = gStrHidei-50;
			//strColEnd = gStrHidei + strColv;	
			parent.strSetTime.cols = strColv;
			//window.setTimeout("menuhide()", 1)
			document.getElementById("Table_close1").style.display = "none"
			document.getElementById("Table_close2").style.display = "none"
			document.getElementById("Table_open").style.display = "inline"
			//parent.leftFrame_Check.document.all.chkLeftHidden.checked = true
	}

	//메뉴보이기
	function menuVisible(){
	var strColEnd,strColv;
	strColv = "184,*"
			gStrHidei = gStrHidei +50;
			strColEnd = gStrHidei + strColv;
			parent.strSetTime.cols = strColv;
			//window.setTimeout("menuVisible()", 1)
			document.getElementById("Table_open").style.display = "none"
			document.getElementById("Table_close1").style.display = "inline"
			document.getElementById("Table_close2").style.display = "inline"
			//parent.leftFrame_Check.document.all.chkLeftHidden.checked = false
	}


//-->	
			</script>
		</SCRIPT>
	</HEAD>
	<body> <!--onLoad="bluring"-->
		<form name="frmThis">
			<table cellSpacing="0" cellPadding="0" width="21" border="0" ID="Table_open"  style="display: none;">
				<tr>
					<td align="right" height="164"><a href="javascript:menuVisible();"><IMG  width="21" height="164" src="../../images/newleftmenu/bt_m_open.gif" border="0"></a></td>
				</tr>
			</table>
			<table cellSpacing="0" cellPadding="0" width="184" border="0" ID="Table_close1">
				<tr>
					<td align="right" height="20"><a href="javascript:menuhide();"><IMG  width="103" height="20" src="../../images/newleftmenu/bt_m_close.gif" border="0"></a></td>
				</tr>
			</table>
			<table width="184" height="100%" cellSpacing="0" cellPadding="0" border="0" ID="Table_close2">
				<tr>
					<td background="../../images/newleftmenu/left_m_bg.gif" height="34" valign="top">
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td height="5"></td>
							</tr>
							<tr>
								<td class="left_menu" style="PADDING-RIGHT: 0px; PADDING-LEFT: 15px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">제작 보고서</td>
							</tr>
							<tr>
								<td class="user_txt02" style="PADDING-RIGHT: 0px; PADDING-LEFT: 129px; PADDING-BOTTOM: 0px; PADDING-TOP: 14px" id="loginname"></td>
							</tr>
							<tr>
								<td height="45"></td>
							</tr>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu07);" class="left_menu2" onclick="blur()">
									보고서</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu07">		
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px" 
									id="PDPR1_V">
									<a id="b1" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMCUSTLISTSEARCH.aspx',this.id);"
									class="left_menu3">광고처별 결산서</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR2_V">
									<a id="b2" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMOUTSLISTSEARCH.aspx',this.id);"
									class="left_menu3">외주비 사용내역</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR3_V">
									<a id="b3" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMACCSEARCH.aspx',this.id);"
									class="left_menu3">진행비 사용내역</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR4_V">
									<a id="b4" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMPREDEMANDSEARCH.aspx',this.id);"
									class="left_menu3">선청구 누계현황</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR5_V">
									<a id="b5" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMPREPAYMENTSEARCH.aspx',this.id);"
									class="left_menu3">선급금 현황</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR6_V">
									<a id="b6" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMACCRUEDSEARCH.aspx',this.id);"
									class="left_menu3">미지급금 현황</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="PDPR7_V">
									<a id="b7" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMTOTALSEARCH.aspx',this.id);"
									class="left_menu3">제작원가 투입현황<br>
							</tr>
							
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
