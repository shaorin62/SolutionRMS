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
   

   
 ''''''''''  �ڸ�Ʈ �θ����̺� �ڸ�Ʈ���� +1 ������Ʈ�Ѵ� ������ �۹�ȣ..
   sql="update web_board set comment_count=comment_count+1 where board_num='"&num&"'"  
   dbcon.execute sql
   
''''''''''' �ڸ�Ʈ  ���̺� �����Ѵ�.  
   sql="insert  web_tail (tail_num,name,pwd,email,content,writeday,u_ip)  values"
   sql=sql& "('"&num&"','"
   sql=sql&name&"','"
   sql=sql&pwd&"','"
   sql=sql&email&"','"
   sql=sql&content&"','"
   sql=sql&date()&"','"&u_ip&"')"

   dbcon.execute sql
   
   
''''''''''' �Ѿ�� �۹�ȣ �������� ��û�������� �Ѱ��ش�.  
   response.Redirect "view.asp?num="&num&"&part="&part&"&tail_part="&tail_part


  
   dbcon.close
   set dbcon=nothing    
 


%>