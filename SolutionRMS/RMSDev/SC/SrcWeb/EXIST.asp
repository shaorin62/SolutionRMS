<%
Set fso = CreateObject("Scripting.FileSystemObject")
dim kkk
if fso.fileexists("C:\Program Files\SCGLCom\SCGLCom.dll") then 
Response.Redirect "Login.aspx"
else  
%>
<script language="javascript">
if(answer = confirm("RMS 시스템을 이용하시기 위해서는 \n Client Module 을 설치해야합니다. 설치하시겠습니까?")){
 location.href = "/DownLoad/SCGLCom.exe"
}
else {  
window.close();
} 
</script>
<%
end if
%>
