<%@ Page Language="vb" AutoEventWireup="false" Codebehind="PDCMJOBMST.aspx.vb" Inherits="PD.PDCMJOBMST" %>
<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.0 Transitional//EN">
<HTML>
	<HEAD>
		<title>JOB관리</title> 
		<!--
'****************************************************************************************
'시스템구분 : RMS/PD/PDCO
'실행  환경 : ASP.NET, VB.NET, COM+ 
'프로그램명 : PDCMJOBMST.aspx
'기      능 : JOBLIST 에서 선택된 JOB관련 프로그램을 호출하는 Main Frame 이다.
'파라  메터 : 
'특이  사항 : SIZE 100% PopUp
'----------------------------------------------------------------------------------------
'HISTORY    :1) 2011/05/04 By kty
'****************************************************************************************
-->
		<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
		<meta content="Microsoft Visual Studio .NET 7.0" name="GENERATOR">
		<meta content="Visual Basic 7.0" name="CODE_LANGUAGE">
		<meta content="VBScript" name="vs_defaultClientScript">
		<meta content="http://schemas.microsoft.com/intellisense/ie5" name="vs_targetSchema">
		<LINK href="../../Etc/STYLES.CSS" type="text/css" rel="STYLESHEET">
		<!-- UI 공통 ActiveX COM -->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCUIClass.inc" -->
		<!-- 공통으로 사용될 클라이언트 스크립트를 Include-->
		<!-- #INCLUDE VIRTUAL="../../../Etc/SCClient.inc" -->
		<script language="vbscript" id="clientEventHandlersVBS">

option explicit
Dim mobjPDCOPREESTDTL
Dim mlngRowInputCnt
Dim mlngColInputCnt
Dim mstrSEQ

'=========================================================================================
' 이벤트 프로시져 
'=========================================================================================
Sub window_onload()
	InitPage
end sub

Sub Window_OnUnload()
	Dim vntData
	
	with frmThis
		
		mlngRowInputCnt=clng(0) : mlngColInputCnt=clng(0)
		
		'set mobjPDCOPREESTDTL = gCreateRemoteObject("cPDCO.ccPDCOPREESTDTL")
		
		vntData = mobjPDCOPREESTDTL.Delete_CloseInput(gstrConfigXml,mlngRowInputCnt,mlngColInputCnt)
		
		'Set mobjPDCOPREESTDTL = Nothing
	
	end with
	set mobjPDCOPREESTDTL = Nothing
End Sub

Sub EndPage()
	gEndPage
End Sub

'닫기버튼
Sub imgClose_MST_onclick ()
	EndPage
End Sub

Sub initpage
	Dim vntInParam
	Dim intNo,i
	
	gInitComParams mobjSCGLCtl,"MC"
	set mobjPDCOPREESTDTL = gCreateRemoteObject("cPDCO.ccPDCOPREESTDTL")

	with frmThis
		
		'IN 파라메터 및 조회를 위한 추가 파라메터 
		vntInParam = window.dialogArguments
		intNo = ubound(vntInParam)
		'기본값 설정
		
		mstrSEQ = ""
	
		For i = 0 to intNo
			Select case i
				case 0 : .txtJOBNO.value = vntInParam(i)	
				case 1 : mstrSEQ = vntInParam(i)
				case 2 : .txtJOBNAME.value = vntInParam(i)
				case 3 : .txtPREESTNO.value = vntInParam(i)	
				case 4 : .txtPRIJOBNAME.value = vntInParam(i)	
				case 5 : .txtPROJECTNM.value = vntInParam(i)	
				case 6 : .txtCLIENTNAME.value = vntInParam(i)	
				case 7 : .txtJOBGUBNNAME.value = vntInParam(i)
				case 8 : .txtCLIENTCODE.value = vntInParam(i)
				case 9 : .txtTIMCODE.value = vntInParam(i)
				case 10 :.txtSUBSEQ.value = vntInParam(i)
				case 11 :.txtJOBGUBN.value = vntInParam(i)
				case 12 :.txtJOBPARTNAME.value = vntInParam(i)
			End select
		Next
		
		
		'.txtJOBNO.value = "C110066"
		'mstrSEQ = "1"
		'.txtJOBNAME.value = "네이처셋 2011 TVCF 제작비"
		'.txtPREESTNO.value = "1106010003"
		'.txtPRIJOBNAME.value = "네이처셋 2011 TVCF 제작비"
		'.txtPROJECTNM.value = "네이처셋 2011 TVCM 제작비"
		'.txtCLIENTNAME.value = "주식회사한독약품"
		'.txtJOBGUBNNAME.value = "CF"
		'.txtCLIENTCODE.value = "A00180"
		'.txtTIMCODE.value = "A00322"
		'.txtSUBSEQ.value = "S1100737"
		'.txtJOBGUBN.value = "PA02"
		'.txtJOBPARTNAME.value = "TV-CF"
		

		'페이지 표기
		.txtPRIJOBVIEW.value	= .txtPRIJOBNAME.value
		.txtJOBVIEW.value		= .txtJOBNAME.value
		.txtJOBNOVIEW.value		= .txtJOBNO.value
		
		
		
		'프레임 전체 가동		
		Set_AllFrameOpen

		'초기 모든 프레임 숨김
		document.getElementById("frmMain_1").style.display = "none"
		document.getElementById("frmMain_2").style.display = "none"
		document.getElementById("frmMain_3").style.display = "none"
		document.getElementById("frmMain_4").style.display = "none"
		document.getElementById("frmMain_5").style.display = "none"
		document.getElementById("frmMain_6").style.display = "none"
		
		jobMst1_onclick
	End With
End Sub

Sub jobMst1_onclick
	document.getElementById("frmMain_1").style.display = "inline"
	document.getElementById("frmMain_2").style.display = "none"
	document.getElementById("frmMain_3").style.display = "none"
	document.getElementById("frmMain_4").style.display = "none"
	document.getElementById("frmMain_5").style.display = "none"
	document.getElementById("frmMain_6").style.display = "none"
End Sub

Sub jobMst2_onclick
	document.getElementById("frmMain_1").style.display = "none"
	document.getElementById("frmMain_2").style.display = "inline"
	document.getElementById("frmMain_3").style.display = "none"
	document.getElementById("frmMain_4").style.display = "none"
	document.getElementById("frmMain_5").style.display = "none"
	document.getElementById("frmMain_6").style.display = "none"
End Sub

Sub jobMst3_onclick
	document.getElementById("frmMain_1").style.display = "none"
	document.getElementById("frmMain_2").style.display = "none"
	document.getElementById("frmMain_3").style.display = "inline"
	document.getElementById("frmMain_4").style.display = "none"
	document.getElementById("frmMain_5").style.display = "none"
	document.getElementById("frmMain_6").style.display = "none"
End Sub

Sub jobMst4_onclick
	document.getElementById("frmMain_1").style.display = "none"
	document.getElementById("frmMain_2").style.display = "none"
	document.getElementById("frmMain_3").style.display = "none"
	document.getElementById("frmMain_4").style.display = "inline"
	document.getElementById("frmMain_5").style.display = "none"
	document.getElementById("frmMain_6").style.display = "none"
End Sub

Sub jobMst5_onclick
	document.getElementById("frmMain_1").style.display = "none"
	document.getElementById("frmMain_2").style.display = "none"
	document.getElementById("frmMain_3").style.display = "none"
	document.getElementById("frmMain_4").style.display = "none"
	document.getElementById("frmMain_5").style.display = "inline"
	document.getElementById("frmMain_6").style.display = "none"
End Sub

Sub jobMst6_onclick
	document.getElementById("frmMain_1").style.display = "none"
	document.getElementById("frmMain_2").style.display = "none"
	document.getElementById("frmMain_3").style.display = "none"
	document.getElementById("frmMain_4").style.display = "none"
	document.getElementById("frmMain_5").style.display = "none"
	document.getElementById("frmMain_6").style.display = "inline"
End Sub


Sub jobMst_Call
	jobMst2_onclick
	call frmMain_2.SelectRtn()
	Call TAB2_Click()
End Sub

Sub jobMst_Tab1Search
	Call frmMain_1.SelectRtn()
End Sub

'다른좝을 선택하여 좝을 생성 하였을 경우,,, 올바른 좝을 선택 하고 조회!
'각각의 프레임에서 재 조회를 위한 공통 조회함수들 - 각 프레임에서 호출할 수 있다.
Sub jobMst_Tab1Search_EstCopy
	Call frmMain_1.PreSelectData()
End Sub

Sub jobMst_Tab2Search
	Call frmMain_2.SelectRtn()
End Sub

Sub jobMst_Tab3Search
	Call frmMain_3.SelectRtn()
End Sub

Sub jobMst_Tab4Search
	Call frmMain_4.SelectRtn()
End Sub

Sub jobMst_Tab5Search
	Call frmMain_5.SelectRtn()
End Sub

Sub jobMst_Tab6Search
	Call frmMain_6.SelectRtn()
End Sub
'--여기까지 각각의 프레임 조회함수 호출

Sub Set_AllFrameOpen
	Dim strIframe
	
	strIframe = "<iframe id='frmMain_1'  frameborder='0' width='100%' height='100%' src='PDCMJOBMST_ESTLIST.aspx' style='position:absolute;top:88px;left:0px;width:100%'></iframe>"
	strIframe = strIframe & "<iframe id='frmMain_2'  frameborder='0' width='100%' height='100%' src='PDCMJOBMST_ESTDTL.aspx' style='position:absolute;top:88px;left:0px;width:100%'></iframe>"
	strIframe = strIframe & "<iframe id='frmMain_3'  frameborder='0' width='100%' height='100%' src='PDCMCHARGEBASICLIST.aspx' style='position:absolute;top:88px;left:0px;'></iframe>"
	strIframe = strIframe & "<iframe id='frmMain_4'  frameborder='0' width='100%' height='100%' src='PDCMCHARGEDIVLIST.aspx' style='position:absolute;top:88px;left:0px;'></iframe>"
	strIframe = strIframe & "<iframe id='frmMain_5'  frameborder='0' width='100%' height='100%' src='PDCMEXELIST.aspx' style='position:absolute;top:88px;left:0px;'></iframe>"
	strIframe = strIframe & "<iframe id='frmMain_6'  frameborder='0' width='100%' height='100%' src='PDCMSUMMARYLIST.aspx' style='position:absolute;top:88px;left:0px;'></iframe>"
	
	document.getElementById("frmMain").innerHTML = strIframe
end Sub

		</script>
		<script type="text/JavaScript">
<!--
function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.01
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && d.getElementById) x=d.getElementById(n); return x;
}

function MM_nbGroup(event, grpName) { //v6.0
  var i,img,nbArr,args=MM_nbGroup.arguments;
  if (event == "init" && args.length > 2) {
    if ((img = MM_findObj(args[2])) != null && !img.MM_init) {
      img.MM_init = true; img.MM_up = args[3]; img.MM_dn = img.src;
      if ((nbArr = document[grpName]) == null) nbArr = document[grpName] = new Array();
      nbArr[nbArr.length] = img;
      for (i=4; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
        if (!img.MM_up) img.MM_up = img.src;
        img.src = img.MM_dn = args[i+1];
        nbArr[nbArr.length] = img;
	  }
	}
  } else if (event == "over") {
    document.MM_nbOver = nbArr = new Array();
    for (i=1; i < args.length-1; i+=3) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = (img.MM_dn && args[i+2]) ? args[i+2] : ((args[i+1])? args[i+1] : img.MM_up);
      nbArr[nbArr.length] = img;
    }
  } else if (event == "out" ) {
    for (i=0; i < document.MM_nbOver.length; i++) {
      img = document.MM_nbOver[i]; img.src = (img.MM_dn) ? img.MM_dn : img.MM_up; }
  } else if (event == "down") {
    nbArr = document[grpName];
    if (nbArr)
      for (i=0; i < nbArr.length; i++) { img=nbArr[i]; img.src = img.MM_up; img.MM_dn = 0; }
    document[grpName] = nbArr = new Array();
    for (i=2; i < args.length-1; i+=2) if ((img = MM_findObj(args[i])) != null) {
      if (!img.MM_up) img.MM_up = img.src;
      img.src = img.MM_dn = (args[i+1])? args[i+1] : img.MM_up;
      nbArr[nbArr.length] = img;
    } 
  }
}

//MenuFocus Move....
function TAB2_Click() {	
	MM_nbGroup('down','group1','jobMst2','../../../images/jobMst2_On.gif',1);
}

//-->
		</script>
	</HEAD>
	<body class="Base" style="MARGIN-TOP: 0px; MARGIN-LEFT: 0px" onload="javascript:MM_preloadImages('../../../images/jobMst1_On.gif','../../../images/jobMst2_On.gif','../../../images/jobMst3_On.gif','../../../images/jobMst4_On.gif','../../../images/jobMst5_On.gif','../../../images/jobMst6_On.gif');MM_nbGroup('down','group1','jobMst1','../../../images/jobMst1_On.gif',1);">
		<BR>
		<form id="frmThis">
			<table class="SEARCHDATA" cellSpacing="0" cellPadding="0" width="99%" border="0">
				<tr>
					<td style="CURSOR: hand" align="left" width="120"><A onmouseover="javascript:MM_nbGroup('over','jobMst1','../../../images/jobMst1_On.gif','../../../images/jobMst1_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst1','../../../images/jobMst1_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst1" height="20" alt="" src="../../../images/jobMst1_Off.gif" width="120"
								onload="" border="0" name="jobMst1"></A></td>
					<td style="CURSOR: hand" align="left" width="120"><A onmouseover="javascript:MM_nbGroup('over','jobMst2','../../../images/jobMst2_On.gif','../../../images/jobMst2_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst2','../../../images/jobMst2_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst2" height="20" alt="" src="../../../images/jobMst2_Off.gif" width="120"
								onload="" border="0" name="jobMst2"></A></td>
					<td style="CURSOR: hand" align="left" width="120"><A onmouseover="javascript:MM_nbGroup('over','jobMst4','../../../images/jobMst4_On.gif','../../../images/jobMst4_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst4','../../../images/jobMst4_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst4" height="20" alt="" src="../../../images/jobMst4_Off.gif" width="120"
								onload="" border="0" name="jobMst4"></A></td>
					<td style="CURSOR: hand" align="left" width="120"><A onmouseover="javascript:MM_nbGroup('over','jobMst5','../../../images/jobMst5_On.gif','../../../images/jobMst5_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst5','../../../images/jobMst5_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst5" height="20" alt="" src="../../../images/jobMst5_Off.gif" width="120"
								onload="" border="0" name="jobMst5"></A></td>
					<td style="CURSOR: hand" align="left" width="120"><A onmouseover="javascript:MM_nbGroup('over','jobMst3','../../../images/jobMst3_On.gif','../../../images/jobMst3_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst3','../../../images/jobMst3_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst3" height="20" alt="" src="../../../images/jobMst3_Off.gif" width="120"
								onload="" border="0" name="jobMst3"></A></td>
					<td style="WIDTH: 5px; CURSOR: hand" align="left" width="5"><A onmouseover="javascript:MM_nbGroup('over','jobMst6','../../../images/jobMst6_On.gif','../../../images/jobMst6_On.gif',1);"
							onclick="javascript:MM_nbGroup('down','group1','jobMst6','../../../images/jobMst6_On.gif',1);" onmouseout="javascript:MM_nbGroup('out');"
							target="_top"><IMG id="jobMst6" height="20" alt="" src="../../../images/jobMst6_Off.gif" width="120"
								onload="" border="0" name="jobMst6"></A></td>
					<td id="lblJOBNAME" style="FONT-SIZE: 9pt; WIDTH: 185px; FONT-FAMILY: 굴림체" width="185"></td>
					<td align="right"><INPUT id="txtJOBPARTNAME" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtJOBPARTNAME"><INPUT id="txtJOBGUBN" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtJOBGUBN"><INPUT id="txtSUBSEQ" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtSUBSEQ"><INPUT id="txtTIMCODE" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtTIMCODE"><INPUT id="txtCLIENTCODE" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtCLIENTCODE"><INPUT id="txtCOMMITIONVALUE" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtCOMMITIONVALUE"><INPUT id="txtSELECT" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtSELECT"><INPUT id="txtPRIJOBNAME" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtPRIJOBNAME"><INPUT id="txtPREESTNO" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtPREESTNO"
							size="1"><INPUT id="txtJOBNAME" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" name="txtJOBNAME"><INPUT id="txtJOBNO" style="WIDTH: 10px; HEIGHT: 10px" type="hidden" size="1" name="txtJOBNO">&nbsp;</td>
				</tr>
				<tr>
					<td class="SEARCHDATA" style="WIDTH: 911px" width="911" colSpan="7">&nbsp;프로젝트 <INPUT class="NOINPUTB_L" id="txtPROJECTNM" title="프로젝트 명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="32" name="txtPROJECTNM">&nbsp;&nbsp;&nbsp;대표제작명
						<INPUT class="NOINPUTB_L" id="txtPRIJOBVIEW" title="프로젝트 명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="29" name="txtPRIJOBVIEW"> &nbsp;&nbsp;&nbsp;ActivityJOB
						<INPUT class="NOINPUTB_L" id="txtJOBVIEW" title="프로젝트 명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="30" name="txtJOBVIEW">&nbsp;</td>
					<td align="right"><IMG id="imgClose_MST" onmouseover="JavaScript:this.src='../../../images/imgCloseOn.gIF'"
							style="CURSOR: hand" onmouseout="JavaScript:this.src='../../../images/imgClose.gIF'" height="20"
							alt="화면을 닫습니다." src="../../../images/imgClose.gIF" width="54" align="absMiddle" border="0" name="imgClose_MST">&nbsp;</td>
				</tr>
				<tr>
					<td class="SEARCHDATA" style="WIDTH: 911px" width="911" colSpan="7">&nbsp;광고주명 <INPUT class="NOINPUTB_L" id="txtCLIENTNAME" title="프로젝트 명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="32" name="txtCLIENTNAME">&nbsp;&nbsp; 
						매체부문명 <INPUT class="NOINPUTB_L" id="txtJOBGUBNNAME" title="프로젝트 명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="29" name="txtJOBGUBNNAME">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
						JOBNo&nbsp;<INPUT class="NOINPUTB_L" id="txtJOBNOVIEW" title="JOB명" style="WIDTH: 224px; HEIGHT: 20px"
							readOnly type="text" maxLength="10" size="30" name="txtJOBNOVIEW"></td>
					<td class="SEARCHDATA"></td>
				</tr>
			</table>
		</form>
		<span id="frmMain" style="WIDTH: 100%; HEIGHT: 88%"></span>
	</body>
</HTML>
