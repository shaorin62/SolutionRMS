<% 
   Option Explicit
   
   Dim part,tail_part 
   part = request("part")
   tail_part = request("tail_part") 
%>

<!DOCTYPE HTML PUBLIC "-//W3C//DTD HTML 4.01 Transitional//EN">
<html>
<head>
<title>�Խ���</title>
<meta http-equiv="Content-Type" content="text/html; charset=ks_c_5601-1987">
<link href="CSS/new.css" rel="stylesheet" type="text/css">
<script>
<!--
function Send(){

pjs=document.Board_Write.name.value;
   if(pjs=="") {
     alert("�̸��� �Է��ϼ���!");
     document.Board_Write.name.focus();
	 return false;
	 }
	
pjs=document.Board_Write.title.value;
   if(pjs=="") {
     alert("�������� �Է��ϼ���!");
     document.Board_Write.title.focus();
	 return false;
	 }
 pjs=document.Board_Write.content.value;
  if(pjs=="") {
     alert("������ �Է��ϼ���!");
     document.Board_Write.content.focus();
	 return false;
	 }

pjs=document.Board_Write.pwd.value;
   if(pjs=="") {
     alert("�н����带 �Է��ϼ���! \n\n������ ������ �ݵ�� �ʿ��մϴ�!");
     document.Board_Write.pwd.focus();
	 return false;
	 }
	 
}

function MM_popupMsg(msg) { //v1.0
  //alert(msg);
}
//-->
</script>
<!-- ####################  �����˻�   ###########################-->

<script language="JavaScript" type="text/JavaScript" src="board.js"></script>
<!-- ####################  �����˻�   ###########################-->
</head>
<body leftmargin="10" topmargin="0" marginwidth="0" marginheight="0" onload="document.Board_Write.name.focus()">
 <form action="write_ok.asp" method="post" name="Board_Write" >
 <img src="Image/title_br.jpg" border=0>		
  <table width="940" border="0" align="left" cellpadding="0" cellspacing="0" bordercolor="white" bordercolorlight="#EDECEC" bordercolordark="white">
    <tr> 
      <td align="center"><table width="940" border="1" cellpadding="0" cellspacing="0" bordercolor="#FFFFFF">
          <tr> 
            <td height="20" colspan="2" align="center">&nbsp; </td>
          </tr>
          <tr> 
            <td width="104" height="25" align="right" bgcolor="FFFFFF">�۾���</td>
            <td width="494" height="25"> &nbsp; <input name="name" type="text" id="name2" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" size="20" >
            </td>
          </tr>
          <!--
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">�̸���</font></td>
            <td height="25">&nbsp; <input name="email" type="text" id="email" style="border:#67949E 1 solid ; background-color:#FFFFFF; color:#666666; height:20" size="40" onFocus='myF(this);' onblur='myB(this);'  title="�̸��� �ּҸ� �Է��ϼ���"></td>
          </tr>
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">Ȩ������</font></td>
            <td height="25">&nbsp; <input name="homepage" type="text" id="homepage" style="border:#67949E 1 solid ; background-color:#FFFFFF; color:#666666; height:20"  title="Ȩ������ �ּҸ� �Է��ϼ���" onFocus='myF(this);' onblur='myB(this);' value="http://" size="40"></td>
          </tr>
          -->
          <tr> 
            <td height="25" align="right" bgcolor="FFFFFF">����</td>
            <td height="25">&nbsp; <input name="title" type="text" id="title" size="113" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" maxlength=255></td>
          </tr>
          <!--
          <tr> 
            <td height="25" align="center" bgcolor="67949E"><font color="#FFFFFF">�±׻��</font></td>
            <td height="25">&nbsp; <input name="tag" type="radio"  title="HTML ����  �Է��ϼ���"   onFocus='myF(this);' onblur='myB(this);' onClick="MM_popupMsg('WEBEZ ������  �˷��帳�ϴ�!\n\nHTML �±� ���� ��Ȯ�� �±׼ҽ��� �Է��Ͽ� �ֽñ� �ٶ��ϴ�.\nHTML �ҽ� ���� �ݴ� �κ��� üũ�Ͻ��� �۾��⸦ �Ͻñ� �ٶ��ϴ�.\n\n����, �±׼ҽ��� �������ϰԿ÷����ԵǸ� �������� ���̺��� Ʈ������ �����Ƿ� ���� ���� Ŀ�´�Ƽ\n\nȰ��ȭ�� ���� ������� ������ġ�ɼ������� �˷��帳�ϴ�.\n\n���� ���źе��� �±׼ҽ������� ���Ҽ� �������� �Ͼ� �ǽ�ġ �ʽ��ϴ�. (^_^) \n\n�׻� �ּ��� ���ϰ� ����ϴ� WEBEZ Ȩ������--���Ϲ�-- �� �Ѹ��� �÷Ƚ��ϴ�.\n                                        --�����մϴ�--')"  value="html_tag"> 
              <font color="67949E">HTML �±�&nbsp;</font> <input name="tag" type="radio"  title="�Ϲ� �ؽ�Ʈ�� �Է��ϼ���" onFocus='myF(this);' onblur='myB(this);' value="text_tag" checked> 
              <font color="67949E">TEXT ���&nbsp;&nbsp;&nbsp;<font color="#FF6600">(HTML 
              �±��Է½� width:600 �ݵ�� �����ֽñ� �ٶ��ϴ�.)</font> </font></td>
          </tr>
          -->
          <tr> 
            <td align="right" bgcolor="FFFFFF">�۳���</td>
            <td style="padding:1"><br> &nbsp; <textarea name="content" wrap="hard" style="font-family:����; COLOR:#666666; width:800; height:201; border:#B9B9B9 1 solid ; background-image: url(image/textline.gif); "></textarea> 
              </td>
          </tr>
          <tr> 
            <td height="25" align="right" bgcolor="FFFFFF">��й�ȣ</td>
            <td height="25">&nbsp; <input name="pwd" type="password" id="pwd" size="15" style="border:#B9B9B9 1 solid ; background-color:#FFFFFF; color:#666666; height:20" > 
              &nbsp;<font color="B9B9B9">(������ ������ �ʿ��մϴ�)</font></td>
          </tr>
          <tr align="left" valign="top"> 
            <td colspan="2">&nbsp; </td>
          </tr>
          <tr> 
            <td height="25" colspan=2 align="right">
            <table  height="25" border="0" cellpadding="0" cellspacing="0">
                <td  align="center"><input name="image" type="image" onclick="return Send()" src="Image/write_icon.gif"  border="0"></td>
                <td align="left" width=70><a href="List.asp"><img src="Image/list_icon.gif" border="0"></a></td>
                </tr>
             </table>
           </td>
             
          </tr>
        </table></td>
    </tr>
  </table>
</form>     
</body>
</html>