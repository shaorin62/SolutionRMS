<!--#include file="dbcon.asp"-->

<%

  Dim part,tail_part,num,pwd,t_pwd,t_num,sql,rs,str
  part=request("part")
  tail_part=request("tail_part")  
  num=request("num")       ''''''글번호값  
  pwd=request("pwd") ''''''패스워드값
  t_pwd=request("t_pwd")  
  t_num=request("t_num")    ''''' 코멘트 번호 변수



 ' response.Write t_num&"<br>"
 ' response.Write t_pwd&"<br>"
'   response.Write num&"<br>"

  
'''''''''' 부모의 글 지우기라면   
  
  if (num <> "") and (pwd <> "") then

'''''''''' 넘어온 변수에 해당하는 레코드셋을 만든다  
  sql="select * from web_board where board_num='"& num &"'" 
  set rs=server.CreateObject("ADODB.recordset")
  rs.open sql,dbcon

'''''''''' 해당하는 레코드에 패스워드 컬럼에 값과 넘어온 패스워드 값과 비교한다.  
  if Lcase(pwd) = Lcase(rs("pwd")) then
'''''''''' 패스워드가 맞으면 삭제한다 
  if rs("comment_count") <> 0 then
   str="<script language='javascript'>"
   str=str& "alert('이 글엔 답변글이 있으므로 삭제하실수 없습니다.\n\n다시 확인하신후 시도하세요!');"
   str=str& "history.back(-1);"
   str=str& "</script>"
   response.Write str
  
  else
  part=request("part")
  tail_part=request("tail_part")
  sql="delete  from web_board where board_num='"& num &"'"
  dbcon.execute (sql)

  
  
'''''' 패스워드가 일치하면 윈도우창을 열었던 부모창으로 돌아간다 opener 함수 이용  
   response.redirect "list.asp?num="&request("num")&"&part="&part&"&tail_part="&tail_part  


   end if
 
   else 
   
 '''''''''''' 패스워드가 일치하지 않으면 
 str="<script language='javascript'>"
 str=str& "alert('패스워드가 틀립니다 \n\n다시확인하십시오.');"
 str=str& "history.back(-1);"
 str=str& "</script>"
 response.Write str
  
  
  
 end if 
 end if 
 '''''''''''''''''''''여기까지 부모글 삭제여부 처리

 '''''''''''  코멘트 지우기라면   
 Dim updatesql  
   
 if t_pwd <> "" and t_num <> "" then
    
   sql="select * from web_tail where t_num='"& t_num &"'"	
   set rs=dbcon.execute(sql)
   
   if Lcase(rs("pwd")) = Lcase(t_pwd) then
  
    ''''''''''''  해당 번호 레코드를 지운다
   sql="delete  from web_tail where t_num='"& t_num &"'"
   dbcon.execute (sql)
   
  '''''''''''' 참조되는 부모 테이블에 코멘트 값도 삭제한만큼 수정한다.
   updatesql="update web_board set comment_count=comment_count-1 where board_num='"& num &"'"
   dbcon.execute (updatesql)
  
  
   else
%>

<script>
alert("비밀번호가 틀립니다");
history.back();
</script>

<%  
   '''''''''''' 패스워드가 일치하지 않으면 
  
 end if
%>

<script>
self.location="view.asp?num=<%=request("num")%>&part=<%=part%>&tail_part=<%=tail_part%>";
</script>

<%

'예전소스
'response.redirect "view.asp?num="&request("num")&"&part="&part&"&tail_part="&tail_part    


 end if
  
    rs.close
   dbcon.close
   set rs=nothing
   set dbcon=nothing    
 
%>