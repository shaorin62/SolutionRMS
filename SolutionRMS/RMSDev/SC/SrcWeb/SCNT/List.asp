<!--#include file="dbcon.asp"-->
<% 
   Dim part,tail_part 
   '���߰Խ����� ��� ���̺�� ����
   part = request("part")
   ' �ڸ�Ʈ ���̺�� ����
   tail_part = request("tail_part")

   Dim page
   if request("page")="" then '������ ����
      page=1
   else	
	  page=request("page")
   end if
  
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>�Խ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="css/new.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="10" topmargin="0" marginwidth="0" marginheight="0">
<img src="Image/title_br1.jpg" border=0>
<table width="940" border="0" align="left" cellpadding="0" cellspacing="0" >
 
  <tr align="left"> 
    <td colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!-- ##################  ���ڵ� ��� ���� #####################-->
        <% 
		                  Dim str,search,sql,rs,count,totalpage,rowcount,mrnum,num,i,trgbcolor,content
					      str = request("str")
						  search = request("search")
						  
					      if request("search")= "" then	                           
    					  sql="select * from web_board order by board_num desc"
						  else
						  sql="select * from web_board where "& str &" like '%"& search &"%'  order by board_num desc"
						  end if
						   
						  set rs=server.CreateObject("ADODB.recordset")
						  rs.pagesize=15   '''����������� �����ش�.�ݵ�� ���ڵ�� �������� �������־���Ѵ�.
                          rs.open sql,dbcon,1 ''���ڵ��� Ŀ��Ÿ���� �������־�� �Ѵ�.������ �ȵǸ� ����¡�� �ȵȴ�.
   
                          if not(rs.eof or rs.bof) then
						  
						  count=rs.recordcount         ''''''''���ڵ� ī���� ����
                          totalpage = rs.pagecount     ''''''''�������� ����
                          rs.absolutepage = page       ''''''''���������� ����
						  
						  end if
					    %>
        <tr> 
          <td width="26%" height="20"><font color="#B9B9B9">Total Count :&nbsp;<font color="#FF6600"><%=count%></font></font></td>
          <td width="74%" height="20" align="right"><font color="#B9B9B9">Total Page :<font color="#FF6600"> 
            <%=totalpage%></font></font>&nbsp;&nbsp;</td>
        </tr>
      </table></td>
  </tr>
   <tr>
	<td colspan=5 height="2" bgcolor="#B9B9B9"></td>
  </tr>
  <tr height=30> 
    <td align="center" width="40">��ȣ</td>
    <td align="center" width = "600">����</td>
    <td align="center" width="100">�ۼ���</td>
    <td align="center"  width="120">�ۼ�����</td>
    <td align="center" width="80">��ȸ��</td>
  </tr>
  <tr>
	<td colspan=5 height="1" bgcolor="#B9B9B9"></td>
  </tr>
  <%  
					
					 
			             if not(rs.eof or rs.bof) then							         
                         i=1   '''pagesize ����
                         
						 rowcount =rs.pagesize
						 
						 do until rs.eof or i > rs.pagesize    ''''''''�����������ŭ ������ ����.
					    
						 mrnum = (page - 1) * rs.PageSize + i - 1
						 num = rs.recordcount - mrnum
						 
						 if i mod 2 = "0" then
                         trgbcolor = "#F6F6F6"
                         else
                         trgbcolor = "#FFFFFF"
                         end if
						 content=rs("content")
                    %>
  <!--####################   ���� ������ ���     #####################-->
  <tr onmouseover="this.style.backgroundColor='<%=trgbcolor%>'" onmouseout="this.style.backgroundColor='#FFFFFF'"> 
    <td  height="25" align="center" bgcolor="<%=trgbcolor%>"><font color="#46676F"><%=num%></font></td>
    <td   height="25" align="left" bgcolor="<%=trgbcolor%>">&nbsp;<a  href="view.asp?page=<%=page%>&num=<%=rs("board_num")%>&b_num=<%=rs("b_num")%>&part=<%=part%>&tail_part=<%=tail_part%>" title="<%=rs("title")%>">&nbsp;
    <%
    Dim strTitle
    Dim intCnt
    intCnt = len(rs("title"))
	  if intCnt > 50 Then
	  strTitle = Mid(rs("title"),1,50) + "..."	  
	  Else 
	  
	  strTitle = rs("title")
	  End If	  
	  ResPonse.Write strTitle
    %></a> 
      <%  '������ڸ� 24�ð��� ���ؼ� ������ ���̹����� ��´�(�ֽű� ����Ʈ)  
		                   if datediff("h",rs("writeday"),now()) < 24 then %>
      <img src="Image/i_new.gif" width="12" height="12" align="absmiddle" > 
      <% end if %>
      &nbsp; 
      <%  '�ڸ�Ʈ ���� 
	                       if rs("comment_count") > 0 then %>
      &nbsp; ...<img src="Image/comment_icon.gif" width="10" height="10" align="absmiddle">&nbsp;<font color="#67949E">(<%=rs("comment_count")%>)</font> 
      <% end if %>
      <!--####################   ���� ������ ���     #####################-->
    </td>
    <td  height="25" align="center" bgcolor="<%=trgbcolor%>"><%=rs("name")%></td>
    <td height="25" align="center" bgcolor="<%=trgbcolor%>"><%=rs("writeday")%></td>
    <td  height="25" align="center" bgcolor="<%=trgbcolor%>"><%=rs("readnum")%></td>
  </tr>
 <!-- <tr> 
    <td height="1" colspan="5" background="Image/line2.gif"></td>
  </tr>-->
  <% 
					   rowcount=rowcount-1   
					   rs.movenext ''���� ���ڵ�� �̵�
					   i=i+1       ''pagesize�� ���� ���� "i"�� ������Ų�� 
                       loop        ''������ �ݺ��Ѵ�
                    %>
  <% else %>
  <tr> 
    <td height="30" colspan="5" align="center"><font color="#67949E">�Խõ� ���� �����ϴ�.</font></td>
  </tr>
  <% end if %>
 
  <tr> 
    <td height="2" colspan="5" bgcolor="#B9B9B9"></td>
  </tr>
  <tr align="right" height=30> 
  <form action="list.asp" name="board_list" method="post" ID="Form1">
	<td height="30"  align="left" colspan=2> 
	<table width="580" border="0" cellpadding="0" cellspacing="0" ID="Table1">
          <tr > 
            <td   align="right" width=80> 
              <select name="str" id="str">
                <option value="title">����</option>
                <option value="content">����</option>
                <option value="name">�۾���</option>
              </select> &nbsp; </td>
            <td  align="left" width=160> <input name="search" type="text" id="search" style="border:#B9B9B9 1 solid ; background-color:#ffffff; color:000000; height:20"> 
            </td>
            <td  ><a href="#"><input type="image" src="Image/search_icon.gif"  border="0" align=absmiddle ID="Image1" NAME="Image1"></a><a href="Write.asp"><img src="Image/write_icon.gif"  border="0"></a><!--<a href="List.asp"><img src="Image/list_icon.gif"  border="0"></a>--></td>
          </tr>
        </table>
	</td></form>
    <td height="30" colspan="3"> 
      <%  %>
      <!-- ����¡ ���� -->
      <!--#include file="pageing.asp"-->
      <!-- ����¡ ����  -->
    </td>
  </tr>
  <tr> 
    <td height="1" colspan="5" align="center" bgcolor="#B9B9B9"></td>
  </tr>
</table>
</body>
</html>
<% rs.close
   dbcon.close
   set rs = nothing
   set dbcon = nothing
 %>
