<!--#include file="dbcon.asp"-->
<%

   Dim part,tail_part,num 
   part = request("part")
   tail_part = request("tail_part") 
   num=request("num")    '�۹�ȣ ����
%>






<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>�Խ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="CSS/new.css" rel="stylesheet" type="text/css">
<script>
function check_n(vv){
 if (vv.value == ''){
     vv.value =vv.defaultValue;
 }
}
 function clear_n(vv){
 if (vv.value == vv.defaultValue){
 vv.value = '';}
 }
 

</script>

<script>

function sendit(){
 if (document.tail.name.value == document.tail.name.defaultValue){
    alert("�̸��� �Է��ϼ���!");
	document.tail.name.focus();
	return;
	}

if (document.tail.pwd.value == document.tail.pwd.defaultValue){
    alert("��й�ȣ�� �Է��ϼ���!");
	document.tail.pwd.focus();
	return;
	}

if (document.tail.content.value == document.tail.content.defaultValue){
    alert("�ϰ��¸��� �����ּ���!");
	document.tail.content.focus();
	return;
	}
	document.tail.submit();
 }
</script>



<script>
//function board_delete(num,part,tail_part)
 // {
//  window.open("Delete.asp?num="+num+"&part="+part+"&tail_part="+tail_part,"","width=300,height=200,menubar=no,toolbar=no");
//  }  
 
//function comment_delete(num,t_num,key,part,tail_part)
//  {
//  window.open("Delete.asp?num="+num+"&t_num="+t_num+"&key="+key+"&part="+part+"&tail_part="+tail_part,"","width=300,height=200,menubar=no,toolbar=no");
//  } 
function reply(){
  document.content_view.action="write.asp";
  document.content_view.submit();
  }
  
function chocheon(sessions,num,part,tail_part){

  if(confirm("������ �̱��� ��õ�Ͻðڽ��ϱ�?")){
    location.href="write_ok.asp?sessions="+sessions+"&num="+num+"&part="+part+"&tail_part="+tail_part;
   }
  }
</script>

</head>
<body leftmargin="10" topmargin="0" marginwidth="0" marginheight="0">

			<%
			      Dim updatesql,sql,rs,page,content,r_num,sessions
			      '��ȸ�� ������Ʈ
                  updatesql="update web_board set readnum=readnum+1 where board_num='"&num&"'"
                  dbcon.execute(updatesql)
  
  
                  sql="select * from web_board where board_num= '" & num & "'"
                  set rs=server.CreateObject("ADODB.recordset")
                  rs.open sql,dbcon,1,3
                  page=request("page")

				  if rs("tag")=1 then
                     content=rs("content")
			      else
				    content=replace(rs("content"),chr(13)& chr(10),"<br>")
				  end if

				  r_num=rs("r_num")   ''�ڸ�Ʈ ����
			  
			%>	
				
	<img src="Image/title_br.jpg" border=0>		
	<table width="940" border="0" align="left" cellpadding="0" cellspacing="0">
  <tr>
  <td bgcolor="#B9B9B9" height=2 colspan =5></td>
  </tr>	
  <tr> 
    <td height="25" colspan="3" align="right" ><table width="100%" height="25" border="0" cellpadding="0" cellspacing="0">
        <tr bgcolor="#F8F8F8"> 
          <td  width="60" align="right">�ۼ���&nbsp;:&nbsp;</td>
          <td  align="left" ><%=rs("name")%></td>
          <td  align="left"></td>
          <td  align="center"></td>
          <td  width="120" align="center">�ۼ���&nbsp;:&nbsp;<%=rs("writeday")%>
          </td>
        </tr>
      </table></td>
  </tr>
   <tr>
  <td bgcolor="#E5E5E5" height=1 colspan =5></td>
  </tr>	
  <tr> 
    <td height="25" colspan="3" align="center" bgcolor="#F8F8F8"><!--2��?TD��--><table width="100%" height="25"  border="0" cellpadding="0" cellspacing="0">
       
        <tr> 
          <td width="60" height="25" align="right">&nbsp;&nbsp;����&nbsp;:&nbsp;</td>
          <td  height="25"><%=rs("title")%></td>
          <td  height="25" align="center"> <font color="#67949E"> 
            <% if session("id")="sharini" then %>
            IP : <%=rs("u_ip")%>&nbsp;</font> <font color="#67949E"> 
            <% end if %>
            </font> </td>
        </tr>
      </table></td>
  </tr>
   <tr>
  <td bgcolor="#E5E5E5" height=1 colspan =5></td>
  </tr>	
  <tr align="left"> 
    <td height="30" colspan="3" valign="top"> <table width="100%" height="30"  border="0" cellpadding="0" cellspacing="0"  bordercolor="white" bordercolordark="white" bordercolorlight="#EDECEC">
        <tr> 
          <td valign="top"><br> <%=content%> <br> &nbsp; </td>
        </tr>
      </table></td>
  </tr>
  
  <tr align="right"> 
    <td colspan="3">
      <% if session("admin") <> "" then %>
      �θ�� ��й�ȣ : <font color="#FF6600"><%=rs("pwd")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</font> 
      <% end if %>
    </td>
  </tr>
 
  <tr> 
    <td height="2" colspan="3" align="center" bgColor="#B9B9B9"></td>
  </tr>
  <!--############## ������ �����ֱ� ���� #######################-->
  <%
	
	                 sql="select * from web_board where b_num='"& request("b_num")-1&"'"
	                 set rs=server.CreateObject("ADODB.recordset")
	                 rs.open sql,dbcon
					 
				%>
  <tr> 
    <td width="100" height="25" align="center" ><table width="96%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
        <tr> 
          <td height="25" align="center" ><font color="#FFFFFF"><img src="Image/lastList.gif" align=absmiddle border=0></font></td>
        </tr>
      </table></td>
    <td width="700" height="25" align="left">&nbsp;
      <% if (rs.eof or rs.bof) then %>
      �������� �����ϴ�. 
      <% else %>
      <a href="view.asp?num=<%=rs("board_num")%>&b_num=<%=rs("b_num")%>&part=<%=part%>&tail_part=<%=tail_part%>"><%=rs("title")%></a> 
      <%  '������ڸ� 24�ð��� ���ؼ� ������ ���̹����� ��´�(�ֽű� ����Ʈ)  
		   if datediff("h",rs("writeday"),now()) < 24 then %>
      <img src="Image/i_new.gif" width="12" height="12" align="absmiddle" > 
      <% end if %>
      &nbsp; 
      <%  '�ڸ�Ʈ ���� 
	                       if rs("comment_count") > 0 then %>
      &nbsp; ...<img src="Image/comment_icon.gif" width="10" height="10" align="absmiddle">&nbsp;<font color="#67949E">(<%=rs("comment_count")%>)</font> 
      <% end if %>
      <!--############## ������ �����ֱ� �� #######################-->
    </td>
    <td width="163" height="25" align="right"><font color="#FF9900"><%=rs("name")%>&nbsp;&nbsp;<%=rs("writeday")%></font></td>
  </tr>
  <% end if %>
  
  <tr> 
    <td height="1" colspan="3" align="center" bgcolor="#B9B9B9"></td>
  </tr>
  
  <!--############## ������ �����ֱ� ���� #######################-->
  <%
		
	                  sql="select * from web_board where b_num='"& request("b_num")+1 &"'"
	                  set rs=server.CreateObject("ADODB.recordset")
	                  rs.open sql,dbcon
	                 %>
  <tr> 
    <td width="100" height="25" align="center" ><table width="96%" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
        <tr> 
          <td height="25" align="center" ><font color="#FFFFFF"><img src="Image/nextList.gif" align=absmiddle border=0></font></td>
        </tr>
      </table></td>
    <td width="700" height="25" align="left">&nbsp;
      <% if rs.eof then %>
      �������� �����ϴ�. 
      <% else %>
      <a href="view.asp?num=<%=rs("board_num")%>&b_num=<%=rs("b_num")%>&part=<%=part%>&tail_part=<%=tail_part%>"><%=rs("title")%></a> 
      <%  '������ڸ� 24�ð��� ���ؼ� ������ ���̹����� ��´�(�ֽű� ����Ʈ)  
		                   if datediff("h",rs("writeday"),now()) < 24 then %>
      <img src="Image/i_new.gif" width="12" height="12" align="absmiddle" > 
      <% end if %>
      &nbsp; 
      <%  '�ڸ�Ʈ ���� 
	                       if rs("comment_count") > 0 then %>
      &nbsp; ...<img src="Image/comment_icon.gif" width="10" height="10" align="absmiddle">&nbsp;<font color="#67949E">(<%=rs("comment_count")%>)</font> 
      <% end if %>
      <!--############## ������ �����ֱ� �� #######################-->
    </td>
    <td height="25" align="right"><font color="#FF9900"><%=rs("name")%>&nbsp;&nbsp;<%=rs("writeday")%></font></td>
  </tr>
  <% end if %>
  
  <tr> 
   <td height="2" colspan="3" align="center" bgcolor="#B9B9B9"></td>
  </tr>
  <tr> 
    <td colspan="3" align="center">&nbsp;</td>
  </tr>
  <tr align="right"> 
    <td height="30" colspan="3"><table width="170" height="30" border="0" cellpadding="0" cellspacing="0">
        <tr align="center"> 
          <td><a href="List.asp?part=<%=part%>&tail_part=<%=tail_part%>"><img src="Image/list_icon.gif"  border="0"></a></td>
          <td ><a href="modify.asp?page=<%=page%>&num=<%=num%>&part=<%=part%>&tail_part=<%=tail_part%>"><img src="Image/modify_icon.gif"  border="0"></a></td>
          <td ><a href="delete.asp?num=<%=num%>&part=<%=part%>&tail_part=<%=tail_part%>"><img src="Image/delete_icon.gif" border="0"></a></td>
        </tr>
      </table></td>
  </tr>
  <tr align="right"> 
    <td colspan="3">&nbsp;</td>
  </tr>
</table>	

</body>
</html>
<% rs.close
   dbcon.close
   set rs = nothing
   set dbcon = nothing
 %>

