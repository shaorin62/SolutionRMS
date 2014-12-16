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
  
  ''''''''''' 글번호 값에 해당하는 레코드를 선택 
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
 
   



''''''''''' 디비의 패스워드와 클라이언트에게로 부터 입력받은 패스워드와 비교  
   if request("pwd") = rs_pwd then

'''''''''   만약 같으면 수정한다
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
   
   
 ''''''' 그렇지 않으면
 else
 Dim str
 '''''' 패스워드가 일치하지 않으면 경고창을 띄운다.  
  str="<script language='javascript'>"
  str=str& "alert('패스워드가 틀립니다.\n\n다시 확인하십시오.');"
  str=str& "history.back(-1);"
  str=str& "</script>"

 response.Write str
 
 %>
 
<!-- 자바스크립트의 구현방법...  
 <script>
   alert("패스워드가 틀립니다.\n\n다시 확인하십시오.");
   history.back(-1);
 </script>	
 -->   
 
<% end if 


   rs.close
   dbcon.close
   set rs=nothing
   set dbcon=nothing    
 


%>    
  