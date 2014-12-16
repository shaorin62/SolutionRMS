<% 
Option Explicit

Dim part,tail_part,num,key,t_num,types
part = request("part")
tail_part = request("tail_part") 
num=request("num") 
key=request("key") 
t_num=request("t_num") 

if request("key")="comment" then  ' 코멘트 지우기라면
   types="t_pwd"
else
   types="pwd"   ' 부모글 지우기라면
end if


%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>게시판</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="CSS/new.css" rel="stylesheet" type="text/css">
 <script>
 function win_focus(){
  document.board_del.<%=types%>.focus();
  }
 
  function del() {
  var pjs;
  pjs=document.board_del.<%=types%>.value;
  if (pjs=="") {
  alert ("패스워드를 입력하세요.");
  document.board_del.<%=types%>.focus();
  return;
     }
  document.board_del.submit();
          }

</script>
</head>
<body leftmargin="0" topmargin="80" marginwidth="0" marginheight="0" onload="document.board_del.<%=types%>.focus();">
<form name="board_del" method="post" action="delete_ok.asp?part=<%=part%>&tail_part=<%=tail_part%>">
            <table width="300" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#CCCCCC">
              <tr> 
                <td height="40" align="center" bgcolor="#FFFFFF"><font color="#B9B9B9" size="5">Delete</font></td>
              </tr>
              <input type="hidden" name="num" value="<%=num%>">
              <input type="hidden" name="t_num" value="<%=t_num%>">
              <tr> 
                <td height="32" align="center" valign="middle" bgcolor="#FFFFFF">삭제된 
                  글은 복구하실수 없습니다.<br>
                  비밀번호를 입력하세요...</td>
              </tr>
              <tr> 
                <td align="center" valign="middle" bgcolor="#FFFFFF">비밀번호: 
                  <input name="<%=types%>" type="password" size="12" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" > </td>
              </tr>
               <tr> 
                <td height="50" align="center" valign="bottom" bgcolor="#FFFFFF"> 
                  <a href="List.asp?part=<%=part%>&tail_part=<%=tail_part%>"><img src="Image/list_icon.gif"  border="0"></a>&nbsp;&nbsp;<a href="javascript:del()"><img src="Image/delete_icon.gif"  border="0"></a> 
                </td>
              </tr>
            </table>
          </form>
</body>
</html>
