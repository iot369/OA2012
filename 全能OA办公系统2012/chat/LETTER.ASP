<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp" -->
<% response.expires=0 %>
<html>
<head>
<title>��������Ү!</title>
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

<body bgcolor="#cccccc" text="<%=m_text2%>">
   <%
if request("id")<>"" then
my_conn.execute ("delete from "&dbtable_gbook&" where "&dbfield_gbook_id&"="&trim(request("id")))
%>
<br><br>
<center><strong><font color="<%=m_text2%>">ɾ�����Գɹ�!</font></strong><br><br>
<input type="button" value="����" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="top.location.href='letter.asp'"> 
<input type="button" value="�ر�" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()">
</center>
<% else %> 
<table width="380" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999">
  <tr align="center" bgcolor="<%=m_bg1%>">
    <td colspan="2" style="color:<%=m_text1%>">�� �� �� ��</td>
  </tr>
<%
set rs=my_conn.execute("select * from "&dbtable_gbook&" where "&dbfield_gbook_name&"='"&session("user")&"' order by "&dbfield_gbook_id&" desc")
if not rs.eof then
  do while not rs.eof %>
  <tr> 
    <td align="center" valign="center"><img src="img/<%=rs(dbfield_gbook_picture)%>.gif"></td>
    <td>
	<% if rs(dbfield_gbook_comefrom)<>"" then %>
	���� <font color="blue"><%=rs(dbfield_gbook_comefrom)%></font> ��
	<% end if %> <font color="red"><%=rs(dbfield_gbook_lyname)%></font> 
		<% if rs(dbfield_gbook_homepage)<>"" then %>
		<a href="<%=rs(dbfield_gbook_homepage)%>"><img src="img/home.gif" width="16" height="16" border="0"></a>
		<% end if %>
		<% if rs(dbfield_gbook_email)<>"" then %>
	    <a href="mailto:<%=rs(dbfield_gbook_email)%>"><img src="img/mail.gif" width="16" height="16" border="0"></a>
		<% end if %>
	    �� <font color="yellow"><%=rs(dbfield_gbook_addtime)%></font> ����˵:<br>
	    <%=rs(dbfield_gbook_message)%>
  <tr>
    <td colspan="2" align="right">
	<a href="letter.asp?id=<%=rs(dbfield_gbook_id)%>" onclick="return confirm('��ȷ��Ҫɾ������������')">ɾ����</a>
	<a href="reletter.asp?n=<%=rs(dbfield_gbook_lyname)%>" onclick="javascript:window.parent.resizeTo(620,200)");">������</a></td>
  </tr>
  <%
  rs.movenext
  loop
else %>
  <tr>
     <td colspan="2" align="center">�㻹û������Ү!</td>
  </tr>
<% end if
rs.close
set rs=nothing
my_conn.close
set my_conn=nothing
end if
%>
</table>
</body>
</html>
<% session("havebook")=1 %>
