<%
Set fso = CreateObject("Scripting.FileSystemObject")
dim kkk
if fso.fileexists("C:\Program Files\SCGLCom\SCGLCom.dll") then 
Response.Redirect "Login.aspx"
else  
%>
<script language="javascript">
if(answer = confirm("RMS �ý����� �̿��Ͻñ� ���ؼ��� \n Client Module �� ��ġ�ؾ��մϴ�. ��ġ�Ͻðڽ��ϱ�?")){
 location.href = "/DownLoad/SCGLCom.exe"
}
else {  
window.close();
} 
</script>
<%
end if
%>
