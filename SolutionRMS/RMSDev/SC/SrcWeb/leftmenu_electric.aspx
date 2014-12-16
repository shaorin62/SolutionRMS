<%@ Page Language="vb" AutoEventWireup="false" Codebehind="leftmenu_electric.aspx.vb" Inherits="SC.leftmenu_electric" %>
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
	
	set mobjSCCOLOGIN		 = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process

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
   	strDEFAULTPGM = "MDEL5"
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
		if(smenu00.style.display=="none"){
			smenu00.style.display="";
		}else{
			smenu00.style.display="none";
		}
		if(smenu01.style.display=="none"){
			smenu01.style.display="";
		}else{
			smenu01.style.display="none";
		}
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
		
		
		for(i=1; i<=23; i++) {
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
		for(i=1;i<=23; i++){
			strMENU = "MDEL" + i;
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
			parent.mainFrame.location.href="http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRIC.aspx"
   			document.getElementById("b6").style.backgroundColor = "F7DEDE"
   			document.getElementById("b6").style.fontWeight = "bold"
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
								<td class="left_menu" style="PADDING-RIGHT: 0px; PADDING-LEFT: 15px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">공중파</td>
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
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px">
									<a href="javascript:menuclick(smenu00);" class="left_menu2" onclick="blur()">
									매체 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu00">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px" 
									id="MDEL1_V">
									<a id="b1" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMKOBACOCLIENTCODE.aspx',this.id);"
									class="left_menu3">KOBACO/SBS 광고주 등록</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL2_V">
									<a id="b2" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMREALMEDCODEMST.aspx',this.id);"
									class="left_menu3">KOBACO/SBS 매체 등록</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL3_V">
									<a id="b3" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECMATTERMST.aspx',this.id);"
									class="left_menu3">소재관리</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL20_V">
									<a id="b20" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECMATTERMST_SRC.aspx',this.id);"
									class="left_menu3">소재관리-등록</a><br>
							</tr>
							<!--<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL4_V">
									<a id="b4" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMEXRATE.aspx',this.id);"
									class="left_menu3">제작대행사 수수료율 등록</a><br>
							</tr>-->
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu01);" class="left_menu2" onclick="blur()">
									청약관리</a></td>
									<!--<span onclick="javascript:menuclick(smenu01);" class="left_menu2" onfocus="blur()" style="CURSOR: hand; LINE-HEIGHT: 13px; LETTER-SPACING: -0.01em"> 
										청약관리</span>-->
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu01">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px" 
									id="MDEL4_V">
									<a id="b4" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECEXMAIN01.aspx',this.id);"
									class="left_menu3">일괄청약</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL5_V">
									<a id="b5" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRIC.aspx',this.id);"
									class="left_menu3">개별청약</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL6_V">
									<a id="b6" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRICSEARCHLIST.aspx',this.id);"
									class="left_menu3">청약내용검증</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL7_V">
									<a id="b7" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRICLIST.aspx',this.id);"
									class="left_menu3">KOBACO/SBS신탁/수수료</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL8_V">
									<a id="b8" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRANSAL.aspx',this.id);"
									class="left_menu3">거래명세서 생성</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL9_V">
									<a id="b9" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRANSLIST.aspx',this.id);"
									class="left_menu3">거래명세서 검증</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL22_V">
									<a id="b22" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRICPPLLIST.aspx',this.id);"
									class="left_menu3">가상/간접 등록</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL23_V">
									<a id="b23" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRICPPLCONFIRM.aspx',this.id);"
									class="left_menu3">가상/간접 승인</a><br>
							</tr>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu02);" class="left_menu2" onclick="blur()">
									정산 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu02">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px" 
									id="MDEL10_V">
									<a id="b10" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECTRUTAX.aspx',this.id);"
									class="left_menu3">광고비청구</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL19_V">
									<a id="b19" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMSENDELECTRUTAX.aspx',this.id);"
									class="left_menu3">위수탁전자세금계산서</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL11_V">
									<a id="b11" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECCOMMIALL.aspx',this.id);"
									class="left_menu3">KOBACO수수료거래/세금계산서생성</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL12_V">
									<a id="b12" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECCOMMILIST.aspx',this.id);"
									class="left_menu3">수수료 거래명세서 조회</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL13_V">
									<a id="b13" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECCOMMITAXLIST.aspx',this.id);"
									class="left_menu3">수수료 세금계산서 조회</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL14_V">
									<a id="b14" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECSPONCOMMIAL.aspx',this.id);"
									class="left_menu3">협찬 수수료거래명세서 생성</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL15_V">
									<a id="b15" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECSPONCOMMITAX.aspx',this.id);"
									class="left_menu3">협찬 수수료세금계산서 생성</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL16_V">
									<a id="b16" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECVOCH.aspx',this.id);"
									class="left_menu3">전표 처리</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL21_V">
									<a id="b21" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDCO/MDCMALLOUTVOCH.aspx',this.id);"
									class="left_menu3">매체매입전표처리</a><br>
							</tr>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu03);" class="left_menu2" onclick="blur()">
									리포트</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
						<table width="184" border="0" cellpadding="0" cellspacing="0" id="smenu03">
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px" 
									id="MDEL17_V">
									<a id="b17" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECMULTILIST.aspx',this.id);"
									class="left_menu3">광고주/매체사별</a><br>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px" 
									id="MDEL18_V">
									<a id="b18" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.89:8080/MD/SrcWeb/MDET/MDCMELECREALMEDLIST.aspx',this.id);"
									class="left_menu3">년간 매체사별 실적</a><br>
							</tr>
							
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
