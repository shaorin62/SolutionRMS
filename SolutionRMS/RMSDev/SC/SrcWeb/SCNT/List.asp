<!--#include file="dbcon.asp"-->
<% 
   Dim part,tail_part 
   '다중게시판일 경우 테이블명 변수
   part = request("part")
   ' 코멘트 테이블명 변수
   tail_part = request("tail_part")

   Dim page
   if request("page")="" then '페이지 설정
      page=1
   else	
	  page=request("page")
   end if
  
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>게시판</title>
<meta http-equiv="Content-Type" content="text/html; charset=euc-kr">
<link href="css/new.css" rel="stylesheet" type="text/css">
</head>

<body leftmargin="10" topmargin="0" marginwidth="0" marginheight="0">
<img src="Image/title_br1.jpg" border=0>
<table width="940" border="0" align="left" cellpadding="0" cellspacing="0" >
 
  <tr align="left"> 
    <td colspan="5"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <!-- ##################  레코드 출력 시작 #####################-->
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
						  rs.pagesize=15   '''페이지사이즈를 정해준다.반드시 레코드셋 오픈전에 지정해주어야한다.
                          rs.open sql,dbcon,1 ''레코드의 커서타입을 지정해주어야 한다.지정이 안되면 페이징이 안된다.
   
                          if not(rs.eof or rs.bof) then
						  
						  count=rs.recordcount         ''''''''레코드 카운터 셋팅
                          totalpage = rs.pagecount     ''''''''총페이지 셋팅
                          rs.absolutepage = page       ''''''''현재페이지 설정
						  
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
    <td align="center" width="40">번호</td>
    <td align="center" width = "600">제목</td>
    <td align="center" width="100">작성자</td>
    <td align="center"  width="120">작성일자</td>
    <td align="center" width="80">조회수</td>
  </tr>
  <tr>
	<td colspan=5 height="1" bgcolor="#B9B9B9"></td>
  </tr>
  <%  
					
					 
			             if not(rs.eof or rs.bof) then							         
                         i=1   '''pagesize 변수
                         
						 rowcount =rs.pagesize
						 
						 do until rs.eof or i > rs.pagesize    ''''''''페이지사이즈만큼 루프를 돈다.
					    
						 mrnum = (page - 1) * rs.PageSize + i - 1
						 num = rs.recordcount - mrnum
						 
						 if i mod 2 = "0" then
                         trgbcolor = "#F6F6F6"
                         else
                         trgbcolor = "#FFFFFF"
                         end if
						 content=rs("content")
                    %>
  <!--####################   실제 데이터 출력     #####################-->
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
      <%  '등록일자를 24시간을 비교해서 작으면 뉴이미지를 찍는다(최신글 리스트)  
		                   if datediff("h",rs("writeday"),now()) < 24 then %>
      <img src="Image/i_new.gif" width="12" height="12" align="absmiddle" > 
      <% end if %>
      &nbsp; 
      <%  '코멘트 갯수 
	                       if rs("comment_count") > 0 then %>
      &nbsp; ...<img src="Image/comment_icon.gif" width="10" height="10" align="absmiddle">&nbsp;<font color="#67949E">(<%=rs("comment_count")%>)</font> 
      <% end if %>
      <!--####################   실제 데이터 출력     #####################-->
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
					   rs.movenext ''다음 레코드로 이동
					   i=i+1       ''pagesize와 비교할 변수 "i"를 증가시킨다 
                       loop        ''루프를 반복한다
                    %>
  <% else %>
  <tr> 
    <td height="30" colspan="5" align="center"><font color="#67949E">게시된 글이 없습니다.</font></td>
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
                <option value="title">제목</option>
                <option value="content">내용</option>
                <option value="name">글쓴이</option>
              </select> &nbsp; </td>
            <td  align="left" width=160> <input name="search" type="text" id="search" style="border:#B9B9B9 1 solid ; background-color:#ffffff; color:000000; height:20"> 
            </td>
            <td  ><a href="#"><input type="image" src="Image/search_icon.gif"  border="0" align=absmiddle ID="Image1" NAME="Image1"></a><a href="Write.asp"><img src="Image/write_icon.gif"  border="0"></a><!--<a href="List.asp"><img src="Image/list_icon.gif"  border="0"></a>--></td>
          </tr>
        </table>
	</td></form>
    <td height="30" colspan="3"> 
      <%  %>
      <!-- 페이징 파일 -->
      <!--#include file="pageing.asp"-->
      <!-- 페이징 파일  -->
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
