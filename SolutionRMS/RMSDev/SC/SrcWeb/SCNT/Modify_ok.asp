<!--#include file="dbcon.asp"-->

<%
   
   Dim part,tail_part,name,email,homepage,writeday,pwd,u_ip,sql,rs,tag,title,content
   part = request("part")
   tail_part = request("tail_part") 

    name=request("name")
  ' name=CheckWord(request("name"))
   email=request("email")
   'title=CheckWord(request("title")) 
  ' title=Replace(Replace(title,"<","&lt;"),">","&gt;")
   'content=CheckWord(request("content"))
  ' content= Replace(Replace(content,"<","&lt;"),">","&gt;")]
'   title=replace(request("title"),"'","''") 
 '  content=replace(request("content"),"'","''")
   homepage=request("homepage")
   writeday=request("writeday")
   pwd=request("pwd")
   u_ip=request.ServerVariables("REMOTE_HOST")
  
  ''''''''''' �۹�ȣ ���� �ش��ϴ� ���ڵ带 ���� 
   sql="select * from web_board  where board_num='"&request("num")&"'"
   set rs=server.CreateObject("ADODB.recordset")
   rs.open sql,dbcon
   
    if request("tag")="html_tag" then 
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
   
   Dim rs_pwd
   rs_pwd=rs("pwd")
 
   



''''''''''' ����� �н������ Ŭ���̾�Ʈ���Է� ���� �Է¹��� �н������ ��  
   if request("pwd") = rs_pwd then

'''''''''   ���� ������ �����Ѵ�
   sql="update web_board  set name='"&name&"'"
   sql=sql &",email='"&email&"'"
   sql=sql &",title='"&title&"'"
   sql=sql &",content='"&content&"'"
   sql=sql &",homepage='"&homepage&"'"
   sql=sql &",pwd='"&pwd&"'"
   sql=sql &",writeday='"& writeday &"'"
   sql=sql &",u_ip='"&u_ip&"'"
   sql=sql &",tag='"&tag&"'"
   sql=sql &"where board_num='"&request("num")&"'"
   
   dbcon.execute (sql)
 
 response.redirect "list.asp?part="&part&"&tail_part="&tail_part
   
   
 ''''''' �׷��� ������
 else
 Dim str
 '''''' �н����尡 ��ġ���� ������ ���â�� ����.  
  str="<script language='javascript'>"
  str=str& "alert('�н����尡 Ʋ���ϴ�.\n\n�ٽ� Ȯ���Ͻʽÿ�.');"
  str=str& "history.back(-1);"
  str=str& "</script>"

 response.Write str
 
 %>
 
<!-- �ڹٽ�ũ��Ʈ�� �������...  
 <script>
   alert("�н����尡 Ʋ���ϴ�.\n\n�ٽ� Ȯ���Ͻʽÿ�.");
   history.back(-1);
 </script>	
 -->   
 
<% end if 


   rs.close
   dbcon.close
   set rs=nothing
   set dbcon=nothing    
 


%>    
  