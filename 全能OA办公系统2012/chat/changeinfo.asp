<!--#INCLUDE FILE="inc_config.asp"-->
<!-- #include file="inc_dbconn.asp" -->
<html>
<head>
<title>�޸�����</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body      {  font-size: 10.4pt }
td        {  font-size: 10.4pt }
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>

<body bgcolor="#cccccc">
<%
  password=request("password")
  if password="" then %>
  <br><br><center><form method="POST" action="changeinfo.asp">������������
   <input type="password" name="password" style="font-size:9pt"><br><br>
   <input type="submit" value="����" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
   <input type="button" value="�ر�" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center></form>
</body>
</html>
<% response.end %>
<% end if  %>

<% if password<>"" then %>
<% if request("email")<>"" then %>

<%
my_conn.execute "update "&dbtable_user&" set "&dbfield_user_password&"='" & request("password") & "',"&dbfield_user_email&"='" & request("email") &"',"&dbfield_user_oicq&"='" & request("oicq")& "',"&dbfield_user_homepage&"='" & request("homepage") &"',"&dbfield_user_comefrom&"='" & request("comefrom") & "',"&dbfield_user_sex&"='" &  request("sex") & "' where "&dbfield_user_username&"='" & session("user") & "'"
%>
<br><br><center>[<% =session("user") %>]�������޸ĳɹ�!<br><br>
<input type="button" value="�رմ���" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
</body>
</html>
<% response.end %>
<% end if %>

<% set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&" ='" & session("user") & "'")
if rs(dbfield_user_password)<>password then %>
<form method="POST" action="changeinfo.asp">
 <br><br><center>�����,������<br>
   <input type="password" name="password" style="font-size:9pt"><br>
   <input type="submit" value="����" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
   <input type="button" value="�ر�" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
</form>
</body>
</html>
<% else %>
<html>
<head>
<title>�޸�����</title>
</head>

<body bgcolor="#cccccc"><br>
<table width="280" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999" align="center">
<form method="POST" action="changeinfo.asp">
  <tr bgcolor="<%=m_bg1%>"> 
   <td align="center" style="color:<%=m_text1%>">�� �� �� ��</td>
  </tr>
  <tr> 
   <td>�� �� ��: <% =session("user") %></td>
  </tr>
  <tr> 
   <td>�û�����:<input type="text" name="password" value=<% =rs(dbfield_user_password) %>></td>
  </tr>
  <tr>
  <%
   if rs(dbfield_user_sex)="boy" then %>
   <td>�ԡ�����:<input type="radio" name="sex"  value="boy" checked>boy <input type="radio" name="sex" value="girl">girl</td>
   <% else %>
   <td>�ԡ�����:<input type="radio" name="sex"  value="boy">boy <input type="radio" name="sex" value="girl"  checked>girl</td>
   <% end if %>
  </tr>
  <tr> 
   <td>��������:<input type="text" name="email" value="<% =rs(dbfield_user_email) %>"></td>
  </tr>
  <tr> 
   <td>��ҳ��ַ:<input type="text" name="homepage" value=<% =rs(dbfield_user_homepage) %>></td>
  </tr>
  <tr> 
   <td >��������:<input type="text" name="comefrom" value=<% =rs(dbfield_user_comefrom) %>></td>
  </tr>
  <tr> 
   <td>O I C Q :<input type="text" name="oicq" value=<% =rs(dbfield_user_oicq) %>></td>
  </tr>
  <tr> 
   <td align="center"><input type="submit" value="����" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
       <input type="reset" value="��ԭ" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
       <input type="button" value="�ر�" onclick="window.close()" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'"></td>
  </tr>
</form>
</table>
</body>
</html>
<% end if %>
<% my_conn.close 
   set my_conn=nothing %>
<% end if %>