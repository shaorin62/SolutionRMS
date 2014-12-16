<!--#include file="dbcon.asp"-->
<%

   Dim part,tail_part,num,name,pwd,email,content,u_ip,sql
   part = request("part")
   tail_part = request("tail_part") 
   num=request("num")
   name=request("name")
   pwd=request("pwd")
   email=request("email")
   content=request("content")
   u_ip=request.ServerVariables("REMOTE_HOST")
   

   
 ''''''''''  코멘트 부모테이블에 코멘트값을 +1 업데이트한다 조건은 글번호..
   sql="update web_board set comment_count=comment_count+1 where board_num='"&num&"'"  
   dbcon.execute sql
   
''''''''''' 코멘트  테이블에 저장한다.  
   sql="insert  web_tail (tail_num,name,pwd,email,content,writeday,u_ip)  values"
   sql=sql& "('"&num&"','"
   sql=sql&name&"','"
   sql=sql&pwd&"','"
   sql=sql&email&"','"
   sql=sql&content&"','"
   sql=sql&date()&"','"&u_ip&"')"

   dbcon.execute sql
   
   
''''''''''' 넘어온 글번호 변수값을 요청페이지로 넘겨준다.  
   response.Redirect "view.asp?num="&num&"&part="&part&"&tail_part="&tail_part


  
   dbcon.close
   set dbcon=nothing    
 


%>