<!--#include file="dbcon.asp"-->

<% 
  
  Dim part,tail_part,num,name,email,homepage,tag,title,content,CheckValue,pwd,u_ip,sessions
  part = request("part")
  tail_part = request("tail_part") 
  num=request("num")   '##############��õ����
  name=request("name")
  email=request("email")
  homepage=request("homepage") 
  
  if request("tag")= "html_tag"  then 
  tag="1"
  title=replace(request("title"),"'","''") 
  content=replace(request("content"),"'","''")

  
  else
   tag="2"
  %>	
 <script LANGUAGE="VBScript" RUNAT="Server">
Function CheckWord(CheckValue)
	CheckValue = replace(CheckValue, "&" , "&amp;")
	CheckValue = replace(CheckValue, "<", "&lt;")
	CheckValue = replace(CheckValue, ">", "&gt;")
	CheckValue = replace(CheckValue, "'", "''")
	CheckWord=CheckValue
End Function
</script>

 <%
  title=CheckWord(request("title"))
   ' title=replace(request("title"),"'","''") 
 ' title=Replace(Replace(title,"<","&lt;"),">","&gt;")
  content=CheckWord(request("content"))
 ' content=replace(request("content"),"'","''")
 ' content= Replace(Replace(content,"<","&lt;"),">","&gt;")
  end if 
 %>
 
 <%
  pwd=request("pwd")
  u_ip=Request.Servervariables("remote_host")
  sessions=request("sessions")  '############## ��õ ���Ǿ��̵� �ߺ���������
 %>


<%
  
  
'############################################### �ߺ� ��õ�̸� ����
   
   Dim sql,rs,updatesql
   
   if (sessions <> "") and (num <> "")  then 
   
   sql="select * from web_board where board_num='"&num&"'"
   set rs=server.CreateObject("adodb.recordset")
   rs.open sql,dbcon
   
''''' �ѹ� ��õ�� �ϸ� ���̺� ��õ���Ǿ��̵� ����..�װ��� ���Ѵ�  

   if sessions = rs("sessions") then
   
%>


   <script>
   alert("��õ ����Ʈ�� �Ϸ翡 �� �ۿ� ���Ͽ�\n\n�ѹ����ۿ� �����ʽ��ϴ�");
   history.back();   
   </script>

   
<% '############################################### �ߺ���õ�̸� ������� %>


<% 
   else   '############################################### �ߺ��� �ƴϸ� ��õ���� ���ǰ���  ������Ʈ �Ѵ�
   
   updatesql="update web_board set r_num=r_num+1,sessions='"&sessions&"' where board_num='"&num&"'"
   dbcon.execute(updatesql)
   response.redirect "view.asp?num="&num&"&part="&part&"&tail_part="&tail_part
   
   end if
 end if
 
 
  '############################################### �ߺ� ��õ���� ��
%> 


	
<% 
    '############################################### ��õ Ÿ���� �ƴϸ� �Ϲ� �۾��� ���� 
    
   Dim b_num	
	
   if (sessions = "") and (num = "")  then 
   sql="select * from web_board order by b_num desc"
   set rs=server.CreateObject("adodb.recordset")
   rs.open sql,dbcon
   
   if not(rs.eof or rs.bof) then
   b_num=rs("b_num")+1
   else
   b_num=1
   end if
  

   sql="insert  web_board (b_num,name,email,homepage,title,content,pwd,writeday,sessions,r_num,readnum,comment_count,u_ip,tag)values"
   sql=sql& "("&b_num&",'"&name&"','"
   sql=sql&email&"','"
   sql=sql&homepage&"','"
   sql=sql&title&"','"
   sql=sql&content&"','"
   sql=sql&pwd&"','"
   sql=sql&date()&"','"&sessions&"',0,0,0,'"&u_ip&"','"& tag &"')"
 
  dbcon.execute(sql)


response.redirect "list.asp?part="&part&"&tail_part="&tail_part
 
 end if
 
 '############################################### ��õ Ÿ���� �ƴϸ� �Ϲ� �۾��� �� 


   rs.close
   dbcon.close
   set rs = nothing
   set dbcon=nothing    
 
 %>