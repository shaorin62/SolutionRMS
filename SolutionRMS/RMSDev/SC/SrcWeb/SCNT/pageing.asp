				  
		 <!--####################  ����¡ �κ�  ����   #####################-->	
		 

	           <%
				   Dim cdivide,blockpage,x	
                   cdivide = 10 '������ ������ ����
                   'response.write "cdivide='"&cdivide&"'<br>"

                  blockPage=Int((page-1)/cdivide)*cdivide+1
        '************************ ���� 10 ������ ���� ***************************
                if blockPage = 1 Then
                   Response.Write ""
                Else
                %>
                <a href="list.asp?page=<%=blockPage-cdivide%>"> <img src="image/i_pp.gif" align=absmiddle  border="0"></a> 
                <%
                End If
        '************************ ���� 10 �� ���� ��***************************

        '---�������� ����-------------------------------------------------------
               if page=1 and int(page)<>int(totalpage) then
              %>
                <img src="image/i_pre.gif"  border="0" align=absmiddle > 
                <% elseif page=1 and int(page)=int(totalpage) then %>
                <img src="image/i_pre.gif"  border="0" align=absmiddle > <!--width="16" height="12"-->
                <% elseif int(page)=int(totalpage) then %>
                <a href="list.asp?page=<%=page-1%>"> <img src="image/i_pre.gif" align=absmiddle  border="0"></a> 
                <% else %>
                <a href="list.asp?page=<%=page-1%>"> <img src="image/i_pre.gif" align=absmiddle  border="0"></a> 
                <% end if
       '---�������� ���� ��---------------------------------------------------


             x=1
       
	         Do Until x > cdivide or blockPage > totalpage
             If blockPage=int(page) Then
             %>
                <font color="#FF9900"><%=blockPage%></font> 
                <%Else%>
                <a href="list.asp?page=<%=blockPage%>"><%=blockPage%></a> 
                <%
    End If
         
    blockPage=blockPage+1
    x = x + 1
    Loop


'----�������� ����---------------------------------------------------
if page=1 and int(page)<>int(totalpage) then
%>
                <a href="list.asp?page=<%=page+1%>"><img src="image/i_next.gif" align=absmiddle  border="0"></a> 
                <%elseif page=1 and int(page)=int(totalpage) then%>
                <img src="image/i_next.gif"  border="0" align=absmiddle > 
                <%elseif int(page)=int(totalpage) then%>
                <img src="image/i_next.gif" border="0" align=absmiddle > 
                <%else%>
                <a href="list.asp?page=<%=page+1%>"> <img src="image/i_next.gif" align=absmiddle  border="0"></a> 
                <%end if
'-----�������� ���� ��-------------------------------------------------

'************************ ���� 10 �� ���� ����*************************** 
if blockPage > totalpage Then
   Response.Write ""
Else
%>
                <a href="list.asp?page=<%=blockPage%>"> <img src="image/i_ff.gif" align=absmiddle  border="0"></a> 
                <%
End If
'************************ ���� 10 �� ���� ��***************************         
%>


       <!--####################  ����¡ �κ� ��    #####################-->
						