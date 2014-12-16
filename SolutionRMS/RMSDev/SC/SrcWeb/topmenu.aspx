	<%@ Page Language="vb" AutoEventWireup="false" Codebehind="topmenu.aspx.vb" Inherits="SC.topmenu" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
  <HEAD>
		<TITLE>열정과 패기로 ! Beyond SK ! RMS</TITLE>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<link href="/css/style.css" rel="stylesheet" type="text/css">
			<!-- #INCLUDE VIRTUAL="../../Etc/SCClient.inc" -->
			<!-- #INCLUDE VIRTUAL="../../Etc/SCUIClass.inc" -->
			<script language="vbscript" id="clientEventHandlersVBS">
<!--
<!--
Dim mlngRowCnt
Dim mlngColCnt
Dim mobjSCCOLOGIN

Sub window_onload
	initpage
End Sub

Sub Window_OnUnload()
	EndPage
End Sub

Sub initpage
	Dim vntData
	Dim vntPreData 
	set mobjSCCOLOGIN = gCreateRemoteObject("cSCCO.ccSCCOLOGIN") '로그인 모듈 Process
	gInitComParams mobjSCGLCtl,"MC"
	
	parent.mainFrame.location.href="http://10.110.10.89:8080/SC/SrcWeb/SCNT/GList.asp"
End Sub

Sub EndPage()
	set mobjSCCOLOGIN = Nothing
	gEndPage
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
gStrLeftmenu = "";
gStrHidei = 200;

function MM_preloadImages() { //v3.0
	TD1.style.display="inline";
	TD2.style.display="none";
	TD3.style.display="none";
	TD4.style.display="none";
	TD5.style.display="none";
    var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}



function auth(strTT,strMENU) {
	var i;	
	if(strTT == "F") {
		alert("메뉴에 대한 권한이 없습니다.");
	} else {
		if(strMENU == "SCCM0"){
			parent.leftFrame.location.href='leftmenu_common.aspx'; 
			gStrLeftmenu = 'leftmenu_common.aspx';
		}
		if(strMENU == "SCMD0"){
			parent.leftFrame.location.href='leftmenu_common_medium.aspx'; 
			gStrLeftmenu = 'leftmenu_common_medium.aspx';
		}
		if(strMENU == "MDEL0"){
			parent.leftFrame.location.href="leftmenu_electric.aspx";
			gStrLeftmenu = 'leftmenu_electric.aspx';
		}
		if(strMENU == "MDCA0"){
			parent.leftFrame.location.href="leftmenu_catv.aspx";
			gStrLeftmenu = 'leftmenu_catv.aspx';
		}	
		if(strMENU == "MDPR0"){
			parent.leftFrame.location.href="leftmenu_print.aspx";
			gStrLeftmenu = 'leftmenu_print.aspx';
		}
		if(strMENU == "MDIN0"){
			parent.leftFrame.location.href="leftmenu_internet.aspx";
			gStrLeftmenu = 'leftmenu_internet.aspx';
		}
		if(strMENU == "MDOU0"){
			parent.leftFrame.location.href="leftmenu_outdoor.aspx";
			gStrLeftmenu = 'leftmenu_outdoor.aspx';
		}
		if(strMENU == "PDCM0"){
			parent.leftFrame.location.href="leftmenu_productdemand.aspx";
			gStrLeftmenu = 'leftmenu_productdemand.aspx';
		}
		if(strMENU == "PDMA0"){
			parent.leftFrame.location.href="leftmenu_productmanage.aspx";
			gStrLeftmenu = 'leftmenu_productmanage.aspx';
		}
		if(strMENU == "READ0"){
			parent.leftFrame.location.href="leftmenu_reporttotal.aspx";
			gStrLeftmenu = 'leftmenu_reporttotal.aspx';
		}
		if(strMENU == "REME0"){
			parent.leftFrame.location.href="leftmenu_reportmeddtl.aspx";
			gStrLeftmenu = 'leftmenu_reportmeddtl.aspx';
		}
		if(strMENU == "PDPR0"){
			parent.leftFrame.location.href="leftmenu_productreport.aspx";
			gStrLeftmenu = 'leftmenu_productreport.aspx';
		}
		if(strMENU == "SCRP0"){
			parent.leftFrame.location.href="leftmenu_totalreport.aspx";
			gStrLeftmenu = 'leftmenu_totalreport.aspx';
		}
		if(strMENU == "SCCT0"){
			parent.leftFrame.location.href="leftmenu_contract.aspx";
			gStrLeftmenu = 'leftmenu_contract.aspx';
		}
		if(strMENU == "MDCG0"){
			parent.leftFrame.location.href="leftmenu_cloud.aspx";
			gStrLeftmenu = 'leftmenu_cloud.aspx';
		}
		if(strMENU == "MDTO0"){
			parent.leftFrame.location.href="leftmenu_gentotal.aspx";
			gStrLeftmenu = 'leftmenu_gentotal.aspx';
		}
		if(strMENU == "MDAD0"){
			parent.leftFrame.location.href="leftmenu_kakao.aspx";
			gStrLeftmenu = 'leftmenu_kakao.aspx';
		}
		if(strMENU == "MDIM0"){
			parent.leftFrame.location.href="leftmenu_ifcmall.aspx";
			gStrLeftmenu = 'leftmenu_ifcmall.aspx';
		}
		if(strMENU == "MDMP0"){
			parent.leftFrame.location.href="leftmenu_MMP.aspx";
			gStrLeftmenu = 'leftmenu_MMP.aspx';
		}
	}
}

function Tubmenu1(){
	if(TD1.style.display=="none"){
		TD1.style.display="inline";
		TD2.style.display="none";
		TD3.style.display="none";
		TD4.style.display="none";
		TD5.style.display="none";
		//PGM_Auth("SCCM0");	
	}
}

function Tubmenu2(){
	if(TD2.style.display=="none"){
		TD2.style.display="inline";
		TD1.style.display="none";
		TD3.style.display="none";
		TD4.style.display="none";
		TD5.style.display="none";
	}
}

function Tubmenu3(){
	if(TD3.style.display=="none"){
		TD3.style.display="inline";
		TD2.style.display="none";
		TD1.style.display="none";
		TD4.style.display="none";
		TD5.style.display="none";
	}
}

function Tubmenu4(){
	if(TD4.style.display=="none"){
		TD4.style.display="inline";
		TD3.style.display="none";
		TD2.style.display="none";
		TD1.style.display="none";
		TD5.style.display="none";
	}
}

// 옥외광고클릭시 하단 내역 
function Tubmenu5(){
	if(TD5.style.display=="none"){
		TD5.style.display="inline";
		TD2.style.display="inline";
		TD4.style.display="none";
		TD3.style.display="none";
		TD1.style.display="none";
	}
}

//공중파
function LeftSub1() {
	PGM_Auth("MDEL0");	
	TD5.style.display="none";
}
//케이블
function LeftSub2(){
	PGM_Auth("MDCA0");	
	TD5.style.display="none";
}
//인쇄
function LeftSub3(){
	PGM_Auth("MDPR0");	
	TD5.style.display="none";
}
//인터넷
function LeftSub4(){
	PGM_Auth("MDIN0");	
	TD5.style.display="none";
}
//옥외
function LeftSub5(){
	PGM_Auth("MDOU0");	
	TD5.style.display="none";
}
function LeftSub6(){
	PGM_Auth("PDCM0");	
}
function LeftSub7(){
	PGM_Auth("PDMA0");	
}
function LeftSub8(){
	PGM_Auth("READ0");	
}
function LeftSub9(){
	PGM_Auth("REME0");	
}
function LeftSub10(){
	PGM_Auth("PDPR0");	
}
function LeftSub11(){
	PGM_Auth("SCRP0");	
}
function LeftSub12(){
	PGM_Auth("SCCT0");	
}
function LeftSub13(){
	PGM_Auth("MDCG0");	
}
//종합편성방송
function LeftSub14(){
	PGM_Auth("MDTO0");	
	TD5.style.display="none";
}
function LeftSub15(){
	PGM_Auth("SCCM0");
}
function LeftSub16(){
	PGM_Auth("SCMD0");
}

function LeftSub17(){
	PGM_Auth("MDAD0");
}

function LeftSub18(){
	PGM_Auth("MDIM0");
}
//라이나생명 MMP
function LeftSub19(){
	PGM_Auth("MDMP0");
}

function Location_Home(){
	parent.topFrame.location.href="topmenu.aspx"
	parent.leftFrame.location.href="leftmenu_common.aspx"
}

//메뉴숨기기
function menuhide(){
	var strColEnd,strCols;
	var i
	var strColv;
	strColv=",100%";
	
	if (gStrHidei > 0){
		gStrHidei = gStrHidei-50;
		strColEnd = gStrHidei + strColv;	
		parent.strSetTime.cols = strColEnd;
		window.setTimeout("menuhide()", 1)
	}
}

//메뉴보이기
function menuVisible(){
	var strColEnd,strColv;
   strColv = ",*"
   if (gStrHidei < 181){
		gStrHidei = gStrHidei +50;
		strColEnd = gStrHidei + strColv;
		parent.strSetTime.cols = strColEnd;
		window.setTimeout("menuVisible()", 1)
   }
}

function allblur(){
for(i = 0;i < document.links.length;i++)
document.links[i].onfocus = document.links[i].blur;
}

//<--모든 링크이벤트시 onfocus=this.blur();를 붙인 효과가 난다.
function bluring(){ 
if(event.srcElement.tagName=="A"||event.srcElement.tagName=="IMG") document.body.focus(); 
} 
document.onfocusin=bluring; 
//-->
		</script>
</HEAD><!--아래그림 없는것 : top_logo_new2,top_back_2 라인없애기top_logo_new3,top_back1-->
	<body onload="javascript:MM_preloadImages('../../../images/topmenu/menu01_on.gif','../../../images/topmenu/menu03_on.gif','../../../images/topmenu/menu04_on.gif');">
		<table height="81" cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
			<tr>
				<td width="254" height="81" valign="top" ><IMG src="../../../images/topmenu/top_logo_new2.gif" useMap="#ImageMap1" border="0">
				</td>
				
				<td valign="top" background="../../../images/topmenu/top_back_3.gif" style="PADDING-RIGHT:0px; PADDING-EFT:0px; PADDING-BOTTOM:0px; PADDING-TOP:55px;background-repeat:repeat-x">
					<table height="33" cellSpacing="0" cellPadding="0" width="100%" align="left" border="0">
						<tr>
							<td>
								<table height="20" cellSpacing="0" cellPadding="0" width="582" align="left" border="0">
									<tr>
										<td width="120"><A  onmouseover="javascript:MM_swapImage('Image1111','','../../../images/topmenu/menu1_on.gif',1);"
											onmouseout="javascript:MM_swapImgRestore();" href="javascript:Tubmenu1();"><IMG id="Image11" src="../../../images/topmenu/menu1_off.gif" border="0" name="Image1111"></A></td>
										<td width="120"><A  onmouseover="javascript:MM_swapImage('Image2111','','../../../images/topmenu/menu2_on.gif',1);"
												onmouseout="javascript:MM_swapImgRestore();" href="javascript:Tubmenu2();"><IMG id="Image21" src="../../../images/topmenu/menu2_off.gif" border="0" name="Image2111"></A></td>
										<td width="120"><A  onmouseover="javascript:MM_swapImage('Image3111','','../../../images/topmenu/menu3_on.gif',1);"
												onmouseout="javascript:MM_swapImgRestore();" href="javascript:Tubmenu3();"><IMG id="Image31" src="../../../images/topmenu/menu3_off.gif" border="0" name="Image3111"></A></td>
										<td width="120"><A  onmouseover="javascript:MM_swapImage('Image4111','','../../../images/topmenu/menu4_on.gif',1);"
												onmouseout="javascript:MM_swapImgRestore();" href="javascript:Tubmenu4()"><IMG id="Image41" src="../../../images/topmenu/menu4_off.gif" border="0" name="Image4111"></A></td>
										<td width="120"></td>
									</tr>	
								</table>
							</td>
						</tr>
						<tr>
							<td valign=top background="#000000">
								<table width="100%" border="0" cellpadding="0" cellspacing="0">
									<tr id="Submenu">
										<td id="TD1" style="PADDING-LEFT: 0px; PADDING-TOP: 5px">
											<table id="Table7" cellSpacing="0" cellPadding="0" border="0">
												<tr>
													<td >
														<td><A href="javascript:LeftSub15();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">공통 관리</A><IMG src="../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub16();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">공통매체관리</A><IMG src="../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub19();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">라이나생명 MMP</A></td>
													</td>
												</tr>
											</table>
										</td>
										<td id="TD2" style="PADDING-LEFT: 126px; PADDING-TOP: 5px">
											<table id="Table6" cellSpacing="0" cellPadding="0" border="0">
												<tr>
													<td><A href="javascript:LeftSub1();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">공중파</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub2();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">케이블</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub14();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">종합편성방송</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub3();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">인쇄</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub4();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">인터넷</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub5();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">옥외</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:Tubmenu5();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">기타 광고</A></td>
												</tr>
											</table>
										</td>
										<td id="TD3" style="PADDING-LEFT: 253px; PADDING-TOP: 5px">
											<table id="Table8" cellSpacing="0" cellPadding="0" border="0">
												<tr>
													<td><A href="javascript:LeftSub6();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">제작의뢰</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub7();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">제작관리</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub12();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">계약서관리</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub10();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">제작보고서</A></td>
												</tr>
											</table>
										</td>
										<td id="TD4" style="PADDING-LEFT: 370px; PADDING-TOP: 5px">
											<table id="Table9" cellSpacing="0" cellPadding="0" border="0">
												<tr>
													<td><A href="javascript:LeftSub8();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">광고비 집행 종합</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub9();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">매체별집행내역</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub11();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">결산보고서</A></td>
												</tr>
											</table>
										</td>
									</tr>
									<tr>
										<td id="TD5" style="PADDING-LEFT: 380px; PADDING-TOP: 1px" colspan = "4">
											<table id="Table10" cellSpacing="0" cellPadding="0" border="0">
												<tr>
													<td><A href="javascript:LeftSub13();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">CGV클라우드</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub17();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">포인트친구 AD</A><IMG src="../../../images/topmenu/2dep_bg_sh.gif"><A href="javascript:LeftSub18();" style="FONT-SIZE: 12px; COLOR: #717171; FONT-FAMILY: 돋움; font-weight:bold;">IFC MALL</A></td>
												</tr>
											</table>
										</td>
									</tr>
								</table>
							</td>
						</tr>

					</table>
				</td>
			</tr>
			
		</table>
		<map name="ImageMap1">
			<area style="CURSOR: hand" shape="RECT" coords="0,10,280,70" href="javascript:Location_Home();">
		</map>
	</body>
</HTML>