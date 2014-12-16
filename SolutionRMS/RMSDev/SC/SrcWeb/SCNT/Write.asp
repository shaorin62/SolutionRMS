<% 
   Option Explicit
   
   Dim part,tail_part 
   part = request("part")
   tail_part = request("tail_part") 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>게시판</title>
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<link href="CSS/new.css" rel="stylesheet" type="text/css">
<script>
<!--
function Send(){

pjs=document.Board_Write.name.value;
   if(pjs=="") {
     alert("이름을 입력하세요!");
     document.Board_Write.name.focus();
	 return false;
	 }
	
pjs=document.Board_Write.title.value;
   if(pjs=="") {
     alert("글제목을 입력하세요!");
     document.Board_Write.title.focus();
	 return false;
	 }
 pjs=document.Board_Write.content.value;
  if(pjs=="") {
     alert("내용을 입력하세요!");
     document.Board_Write.content.focus();
	 return false;
	 }

pjs=document.Board_Write.pwd.value;
   if(pjs=="") {
     alert("패스워드를 입력하세요! \n\n수정과 삭제시 반드시 필요합니다!");
     document.Board_Write.pwd.focus();
	 return false;
	 }
	 
}

function MM_popupMsg(msg) { //v1.0
  //alert(msg);
}
//-->
</script>
<!-- ####################  빠른검색   ###########################-->

<script language="JavaScript" type="text/JavaScript" src="board.js"></script>
<!-- ####################  빠른검색   ###########################-->
</head>
<body leftmargin="10" topmargin="0" marginwidth="0" marginheight="0" onload="document.Board_Write.name.focus()">
 <form action="write_ok.asp" method="post" name="Board_Write" >
 <img src="Image/title_br.jpg" border=0>		
  <table width="940" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="white" bordercolorlight="#EDECEC" bordercolordark="white">
    <tr> 
      <td align="center"><table width="940" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
          <tr> 
            <td height="20" colspan="2" align="center">&nbsp; </td>
          </tr>
          <tr> 
            <td width="104" height="25" align="right" bgcolor="FFFFFF">글쓴이</td>
            <td width="494" height="25"> &nbsp; <input name="name" type="text" id="name2" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" size="20" >
            </td>
          </tr>
          <!--
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">이메일</font></td>
            <td height="25">&nbsp; <input name="email" type="text" id="email" style="border:#67949E 1 solid ; background-color:#FFFFFF; color:#666666; height:20" size="40" onFocus='myF(this);' onblur='myB(this);'  title="이메일 주소를 입력하세요"></td>
          </tr>
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">홈페이지</font></td>
            <td height="25">&nbsp; <input name="homepage" type="text" id="homepage" style="border:#67949E 1 solid ; background-color:#FFFFFF; color:#666666; height:20"  title="홈페이지 주소를 입력하세요" onFocus='myF(this);' onblur='myB(this);' value="http://" size="40"></td>
          </tr>
          -->
          <tr> 
            <td height="25" align="right" bgcolor="FFFFFF">제목</td>
            <td height="25">&nbsp; <input name="title" type="text" id="title" size="113" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" maxlength=255></td>
          </tr>
          <!--
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">태그사용</font></td>
            <td height="25">&nbsp; <input name="tag" type="radio"  title="HTML 모드로  입력하세요"   onFocus='myF(this);' onblur='myB(this);' onClick="MM_popupMsg('WEBEZ 쥔장이  알려드립니다!\n\nHTML 태그 사용시 정확한 태그소스를 입력하여 주시기 바랍니다.\nHTML 소스 열고 닫는 부분을 체크하신후 글쓰기를 하시길 바랍니다.\n\n만약, 태그소스가 부적절하게올려지게되면 웹페이지 테이블이 트러질수 있으므로 보다 나은 커뮤니티\n\n활성화를 위해 예고없이 삭제조치될수있음을 알려드립니다.\n\n여기 오신분들은 태그소스정도는 잘할수 있으리라 믿어 의심치 않습니다. (^_^) \n\n항상 최선을 다하고 노력하는 WEBEZ 홈피지기--오일박-- 이 한마디 올렸습니다.\n                                        --감사합니당--')"  value="html_tag"> 
              <font color="67949E">HTML 태그&nbsp;</font> <input name="tag" type="radio"  title="일반 텍스트를 입력하세요" onFocus='myF(this);' onblur='myB(this);' value="text_tag" checked> 
              <font color="67949E">TEXT 모드&nbsp;&nbsp;&nbsp;<font color="#FF6600">(HTML 
              태그입력시 width:600 반드시 지켜주시기 바랍니다.)</font> </font></td>
          </tr>
          -->
          <tr> 
            <td align="right" bgcolor="FFFFFF">글내용</td>
            <td style="padding:1"><br> &nbsp; <textarea name="content" wrap="hard" style="font-family:돋움; COLOR:#666666; width:800; height:201; border:#B9B9B9 1 solid ; background-image: url(image/textline.gif); "></textarea> 
              </td>
          </tr>
          <tr> 
            <td height="25" align="right" bgcolor="FFFFFF">비밀번호</td>
            <td height="25">&nbsp; <input name="pwd" type="password" id="pwd" size="15" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" > 
              &nbsp;<font color="B9B9B9">(수정과 삭제시 필요합니다)</font></td>
          </tr>
          <tr align="left" valign="top"> 
            <td colspan="2">&nbsp; </td>
          </tr>
          <tr> 
            <td height="25" colspan=2 align="right">
            <table  height="25" border="0" cellpadding="0" cellspacing="0">
                <td  align="center"><input name="image" type="image" onclick="return Send()" src="Image/write_icon.gif"  border="0"></td>
                <td align="left" width=70><a href="List.asp"><img src="Image/list_icon.gif" border="0"></a></td>
                </tr>
             </table>
           </td>
             
          </tr>
        </table></td>
    </tr>
  </table>
</form>     
</body>
</html>