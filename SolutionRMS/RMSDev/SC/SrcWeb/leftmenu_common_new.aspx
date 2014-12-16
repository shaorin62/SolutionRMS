<%@ Page Language="vb" AutoEventWireup="false" Codebehind="leftmenu_common_new.aspx.vb" Inherits="SC.leftmenu_common_new" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<TITLE>열정과 패기로 ! Beyond SK ! RMS</TITLE>
		<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
		<LINK href="/css/style.css" type="text/css" rel="stylesheet">
			<!-- #INCLUDE VIRTUAL="../../Etc/SCClient.inc" -->
			<!-- #INCLUDE VIRTUAL="../../Etc/SCUIClass.inc" -->
			<script language="vbscript" id="clientEventHandlersVBS">

			</script>
	</HEAD>
	<body> <!--onLoad="bluring"-->
		<form name="frmThis">
			<table width="184" height="100%" cellSpacing="0" cellPadding="0" border="" ID="Table1">
				<tr>
					<td height="20">
						<table cellSpacing="0" cellPadding="0" width="184" border="0">
							<tr>
								<td align="right"><IMG height="20" src="../../images/newleftmenu/bt_m_close.gif" width="103"></td>
							</tr>
						</table>
					</td>
				</tr>
				<tr>
					<td background="../../images/newleftmenu/left_m_bg.gif" height="34" valign="top">
						<table width="184" border="0" cellpadding="0" cellspacing="0">
							<tr>
								<td height="44"></td>
							</tr>
							<tr>
								<td class="left_menu" style="PADDING-RIGHT: 0px; PADDING-LEFT: 50px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px">공통</td>
							</tr>
							<tr>
								<td height="15"></td>
							</tr>
							<tr>
								<td class="user_txt02" style="PADDING-RIGHT: 0px; PADDING-LEFT: 129px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px">관리자 
									님</td>
							</tr>
							<tr>
								<td height="45"></td>
							</tr>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px"
									id="SCCM1_V">
									<a id="b1" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCNT/List.asp',this.id);"
										class="left_menu2">공지사항</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu01);" class="left_menu2" onclick="blur()">거래처 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM2_V">
									<a id="b2" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOMEDLIST.aspx',this.id);"
										class="left_menu3">매체사</a></td>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM3_V">
									<a id="b3" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTLIST.aspx',this.id);"
										class="left_menu3">광고주</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM4_V">
									<a id="b4" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTEXELIST.aspx',this.id);"
										class="left_menu3">대대행사</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM5_V">
									<a id="b5" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTOUTLIST.aspx',this.id);"
										class="left_menu3">외주처</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 0px"
									id="SCCM6_V">
									<a id="b6" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTMPPLIST.aspx',this.id);"
										class="left_menu3">MPP</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM7_V">
									<a id="b7" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTCRELIST.aspx',this.id);"
										class="left_menu3">크리조직</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM8_V">
									<a id="b8" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOCUSTGREATLIST.aspx',this.id);"
										class="left_menu3">광고처</a>
								</td>
							</tr>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px">
									<a href="javascript:menuclick(smenu02);" class="left_menu2" onclick="blur()">브랜드관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM9_V">
									<a id="b9" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOBRANDHDRLIST.aspx',this.id);"
										class="left_menu3">대표브랜드관리</a>
								</td>
							</tr>
							<tr>
								<td style="PADDING-RIGHT: 0px; PADDING-LEFT: 26px; PADDING-BOTTOM: 0px; PADDING-TOP: 2px"
									id="SCCM10_V">
									<a id="b10" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOBRANDDTLLIST.aspx',this.id);"
										class="left_menu3">브랜드관리</a>
								</td>
							</tr>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM11_V">
									<a id="b11" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCO/SCCOFEE.aspx',this.id);"
										class="left_menu2">Fee 거래광고주 관리</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM12_V">
									<a id="b12" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCD/SCCDCODE.aspx',this.id);"
										class="left_menu2">공통코드</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM13_V">
									<a id="b13" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCD/SCCDDEPTMST.aspx',this.id);"
										class="left_menu2">조직정보</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM14_V">
									<a id="b14" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCD/SCCDEMPMST.aspx',this.id);"
										class="left_menu2">사용자</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM15_V">
									<a id="b15" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCD/SCCDROLELIST.aspx',this.id);"
										class="left_menu2">권한설정</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
							<tr>
								<td class="left_menu2" style="PADDING-RIGHT: 0px; PADDING-LEFT: 20px; PADDING-BOTTOM: 0px; PADDING-TOP: 10px"
									id="SCCM16_V">
									<a id="b16" style="CURSOR:hand" onclick="javascript:detailclick('http://10.110.10.86:4350/SC/SrcWeb/SCCD/SCCDBATCHLOGLIST.aspx',this.id);"
										class="left_menu2">일괄작업</a></td>
							</tr>
							<TR>
								<td bgcolor="#e3e3e3" height="1"></td>
							</TR>
						</table>
					</td>
				</tr>
			</table>
		</form>
	</body>
</HTML>
