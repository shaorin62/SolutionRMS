<%  

%>


<% 
  '무분별한 변수사용을 막기위해
  Option explicit
  
  Dim Dbcon
  Set DbCon = Server.CreateObject("ADODB.Connection")  
  DbCon.open "provider=sqloledb; data source=10.110.10.88\mcrmsdb; initial catalog=mcdev_new; user id=advsa; password = advsa1234"    

  
%>