<!--#include file="dbcon.asp"-->

 <%
              Dim part,tail_part,num,page,sql,rs,content
			  part = request("part")
              tail_part = request("tail_part")

              num=request("num")
              page=request("page")
   
              sql="select * from web_board where board_num='"&num&"'"
              set rs=server.CreateObject("ADODB.recordset")
              rs.open sql,dbcon
			  
            
              content=rs("content")
			
          %>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>게시판</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="CSS/new.css" rel="stylesheet" type="text/css">
<script>
function Send(){
pjs=document.board_modify.name.value;
   if(pjs=="") {
     alert("이름을 입력하세요!");
     document.board_modify.name.focus();
	 return ;
	 }
pjs=document.board_modify.title.value;
   if(pjs=="") {
     alert("글제목을 입력하세요!");
     document.board_modify.title.focus();
	 return ;
	 }
	 
pjs=document.board_modify.content.value;
   if(pjs=="") {
     alert("글내용을 입력하세요!");
     document.board_modify.content.focus();
	 return ;
	 }	 

pjs=document.board_modify.pwd.value;
   if(pjs=="") {
     alert("패스워드를 입력하세요! \n\n수정과 삭제시 반드시 필요합니다!");
     document.board_modify.pwd.focus();
	 return ;
	 }
	 document.board_modify.submit();
}	 
</script>

</head>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
				

<form action="modify_ok.asp" method="post" name="board_modify">
 <img src="Image/title_br.jpg" border=0>		
                    <table width="940" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="white" bordercolorlight="#EDECEC" bordercolordark="white">
                      <tr> 
                        <td><table width="940" border="1" align="center" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
                            <tr> 
                              <td height="15" colspan="2" align="center"><input type="hidden" name="num" value="<%=request("num")%>">
                        <input name="writeday" type="hidden" id="writeday" value="<%=rs("writeday")%>">
                        <input name="part" type="hidden" id="part" value="<%=part%>"> 
                        <input name="tail_part" type="hidden" id="tail_part" value="<%=tail_part%>"> 
                      </td>
                            </tr>
                            <tr> 
                              <td width="92" height="25" align="right" bgcolor="#FFFFFF">글쓴이</td>
                              <td width="488" height="25"> &nbsp; <input name="name" type="text" id="name2" value="<%=rs("name")%>" size="20" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20"></td>
                            </tr>
                            
                            <tr> 
                              <td height="25" align="right" bgcolor="#FFFFFF">제목</td>
                              <td height="25">&nbsp; <input name="title" type="text" id="title" value="<%=rs("title")%>" size="113" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20"></td>
                            </tr>
							
                            <tr> 
                              <td align="right" bgcolor="#FFFFFF">글내용</td>
                              <td> &nbsp;<br>
                        &nbsp; <textarea name="content" wrap="hard" style="font-family:돋움; COLOR:#666666; width:800; height:211; border:#B9B9B9 1 solid ; background-image: url(image/textline.gif); "><%=content%></textarea> 
                        </td>
                            </tr>
                            <tr> 
                              <td height="25" align="right" bgcolor="#FFFFFF">비밀번호</td>
                              <td height="25">&nbsp; <input name="pwd" type="password" id="pwd" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" value="<% if session("admin") <> "" or session("admin") <> "" then%><%=rs("pwd")%><% end if %>" size="15"> 
                                &nbsp;<font color="#FFFFFF">(수정과 삭제시 필요합니다)</font></td>
                            </tr>
                            <tr> 
                              <td height="25" colspan="2" align="right"><table height="30" border="0" cellpadding="0" cellspacing="0">
                                  <tr> 
                                    <td  align="center"><a href="javascript:Send()"><img src="Image/modify_icon.gif"  border="0"></a></td>
                                    <td width="80" align="left"><a href="List.asp?part=<%=part%>&tail_part=<%=tail_part%>"><img src="Image/list_icon.gif"  border="0"></a></td>
                                  </tr>
                                </table></td>
                            </tr>
                          </table></td>
                      </tr>
                    </table>
                  </form>
</body>
</html>
<% rs.close
   dbcon.close
   set rs = nothing
   set dbcon = nothing
 %>