<!--#include file="dbcon.asp"-->

<%

  Dim part,tail_part,num,pwd,t_pwd,t_num,sql,rs,str
  part=request("part")
  tail_part=request("tail_part")  
  num=request("num")       ''''''�۹�ȣ��  
  pwd=request("pwd") ''''''�н����尪
  t_pwd=request("t_pwd")  
  t_num=request("t_num")    ''''' �ڸ�Ʈ ��ȣ ����



 ' response.Write t_num&"<br>"
 ' response.Write t_pwd&"<br>"
'   response.Write num&"<br>"

  
'''''''''' �θ��� �� �������   
  
  if (num <> "") and (pwd <> "") then

'''''''''' �Ѿ�� ������ �ش��ϴ� ���ڵ���� �����  
  sql="select * from web_board where board_num='"& num &"'" 
  set rs=server.CreateObject("ADODB.recordset")
  rs.open sql,dbcon

'''''''''' �ش��ϴ� ���ڵ忡 �н����� �÷��� ���� �Ѿ�� �н����� ���� ���Ѵ�.  
  if Lcase(pwd) = Lcase(rs("pwd")) then
'''''''''' �н����尡 ������ �����Ѵ� 
  if rs("comment_count") <> 0 then
   str="<script language='javascript'>"
   str=str& "alert('�� �ۿ� �亯���� �����Ƿ� �����ϽǼ� �����ϴ�.\n\n�ٽ� Ȯ���Ͻ��� �õ��ϼ���!');"
   str=str& "history.back(-1);"
   str=str& "</script>"
   response.Write str
  
  else
  part=request("part")
  tail_part=request("tail_part")
  sql="delete  from web_board where board_num='"& num &"'"
  dbcon.execute (sql)

  
  
'''''' �н����尡 ��ġ�ϸ� ������â�� ������ �θ�â���� ���ư��� opener �Լ� �̿�  
   response.redirect "list.asp?num="&request("num")&"&part="&part&"&tail_part="&tail_part  


   end if
 
   else 
   
 '''''''''''' �н����尡 ��ġ���� ������ 
 str="<script language='javascript'>"
 str=str& "alert('�н����尡 Ʋ���ϴ� \n\n�ٽ�Ȯ���Ͻʽÿ�.');"
 str=str& "history.back(-1);"
 str=str& "</script>"
 response.Write str
  
  
  
 end if 
 end if 
 '''''''''''''''''''''������� �θ�� �������� ó��

 '''''''''''  �ڸ�Ʈ �������   
 Dim updatesql  
   
 if t_pwd <> "" and t_num <> "" then
    
   sql="select * from web_tail where t_num='"& t_num &"'"	
   set rs=dbcon.execute(sql)
   
   if Lcase(rs("pwd")) = Lcase(t_pwd) then
  
    ''''''''''''  �ش� ��ȣ ���ڵ带 �����
   sql="delete  from web_tail where t_num='"& t_num &"'"
   dbcon.execute (sql)
   
  '''''''''''' �����Ǵ� �θ� ���̺� �ڸ�Ʈ ���� �����Ѹ�ŭ �����Ѵ�.
   updatesql="update web_board set comment_count=comment_count-1 where board_num='"& num &"'"
   dbcon.execute (updatesql)
  
  
   else
%>

<script>
alert("��й�ȣ�� Ʋ���ϴ�");
history.back();
</script>

<%  
   '''''''''''' �н����尡 ��ġ���� ������ 
  
 end if
%>

<script>
self.location="view.asp?num=<%=request("num")%>&part=<%=part%>&tail_part=<%=tail_part%>";
</script>

<%

'�����ҽ�
'response.redirect "view.asp?num="&request("num")&"&part="&part&"&tail_part="&tail_part    


 end if
  
    rs.close
   dbcon.close
   set rs=nothing
   set dbcon=nothing    
 
%>