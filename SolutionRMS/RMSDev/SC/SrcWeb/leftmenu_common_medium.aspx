<%@ Page Language="vb" AutoEventWireup="false" Codebehind="leftmenu_common_medium.aspx.vb" Inherits="SC.leftmenu_common_medium" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>열정과 패기로 ! Beyond SK ! RMS</TITLE>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
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
	
	set mobjSCCOLOGIN		 = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process
	
	'gInitPageSetting mobjSCGLCtl,"MD"
	gInitComParams mobjSCGLCtl,"MC"

	'on error resume next
	mlngRowCnt=clng(0)
	mlngColCnt=clng(0)
	vntData = mobjSCCOLOGIN.SelectRtn_USER(gstrConfigXml,mlngRowCnt,mlngColCnt)
	if not gDoErrorRtn ("SelectRtn_USER") then	
		if mlngRowCnt > 0 Then
		document.getElementById("loginname").innerHTML = vntData(2,1) & " 님"
		end if
   	end if
   
End Sub

Sub EndPage()
	set mobjSCCOLOGIN = Nothing
	gEndPage
End Sub

Sub imgLogOut_onclick
	gInitPageSetting mobjSCGLCtl,"SC"
	Call Win_MainClose()
End Sub

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
		end if
		Call auth(strVAL,strMENU) 
   	end if
End Sub

-->
		</script>
		<script language="JavaScript" type="text/JavaScript">
<!--
	var gStrLeftmenu;
	var gStrHidei;
	var strImg = 0;
	gStrLeftmenu = "";
	gStrHidei = 184;

	
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
		
		
		for(i=1; i<=6; i++) {
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
		for(i=1;i<=6; i++){
			strMENU = "SCMD" + i;
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
	</HEAD>
	<body> <!--onLoad="bluring"-->
		<form name="frmThis">
			<table cellSpacing="0" cellPadding="0" width="21" border="0" ID="Table_open" style="DISPLAY: none">
				<tr>
					<td align="right" height="164"><a href="javascript:menuVisible();"><IMG width="21" height="164" src="../../images/newleftmenu/bt_m_open.gif" border="0"></a></td>
				</tr>
			</table>
			<table cellSpacing="0" cellPadding="0" width="184" border="0" ID="Table_close1">
				<tr>
					<td align="right" height="20"><a href="javascript:menuhide();"><IMG width="103" height="20" src="../../images/newleftmenu/bt_m_close.gif" border="0"></a></td>
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
								<td class="left_menu" style="PADDING-RIGHT: 0px; PADDING-LEFT: 15px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">공통매체관리</td>
							</tr>
							<tr>
								<td class="user_txt02" style="PADDING-RIGHT: 0px; PADDING-LEFT: 125px; PADDING-BOTTOM: 0px; PADDING-TOP: 14px" id="loginname"></td>
							</tr>
							<tr>
								<td height="45"></td>
							</tr>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu01);" class="left_menu2" onclick="blur()">AOR 청약관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu01">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD1_V">
									<a id="b1" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDSC/MDCMAOR_MEDIUM.aspx',this.id);"
										class="left_menu3">AOR 개별청약</a></td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD2_V">
									<a id="b2" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDSC/MDCMAORCOMMIAL.aspx',this.id);"
										class="left_menu3">거래명세서</a></td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD3_V">
									<a id="b3" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDSC/MDCMAORCONFLIST.aspx',this.id);"
										class="left_menu3">거래명세서 승인</a></td>
							</tr>							
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu02);" class="left_menu2" onclick="blur()">AOR 정산 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu02">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD4_V">
									<a id="b4" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDSC/MDCMAORCOMMITAX.aspx',this.id);"
										class="left_menu3">수수료 청구</a></td>
							</tr>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu03);" class="left_menu2" onclick="blur()">통합 청구 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu03">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD5_V">
									<a id="b5" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMMERGETAX.aspx',this.id);"
										class="left_menu3">통합청구 계산서</a></td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
								 id="SCMD6_V">
									<a id="b6" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/PD/SrcWeb/PDCO/PDCMMERGEVOCH.aspx',this.id);"
										class="left_menu3">통합청구 전표</a></td>
							</tr>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
