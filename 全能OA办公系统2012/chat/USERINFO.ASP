<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title>聊友信息</title>
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
<% name=request("name")
if name <>""  and session("userlevel")<>"" then%>
<% set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&"='"& name &"'")
if not rs.eof then %>
  <table width="280" border="1" cellspacing="0" cellpadding="3" bgcolor="<%=c_canbg%>" bordercolorlight="#FFFFFF" bordercolordark="#999999" align="center">
    <tr bgcolor="<%=m_bg1%>"> 
      <td colspan="2" align="center" style="color:<%=m_text1%>">聊友信息</td>
    </tr>
	<%if session("userlevel")>7 then%>
<tr>
 <td> 用户IP: </td>
 <td> <% =rs(dbfield_user_ip) %></td>
</tr>
    <tr>
      <td> ID: </td>
      <td> <% =rs(dbfield_user_id) %></td>
    </tr>
	<%end if%>
    <tr> 
      <td> 呢称: </td>
      <td> <% =rs(dbfield_user_username) %></td>
    </tr>
    <tr> 
      <td> 信箱: </td>
      <td> <% =rs(dbfield_user_email) %></td>
    </tr>
    <tr> 
      <td> OICQ: </td>
      <td> <% =rs(dbfield_user_oicq) %></td>
    </tr>
    <tr> 
      <td> 主页: </td>
      <td> <% =rs(dbfield_user_homepage) %></td>
    </tr>
    <tr> 
      <td> 来自: </td>
      <td> <% =rs(dbfield_user_comefrom) %></td>
    </tr>
    <tr> 
      <td> 积分: </td>
      <td> <% =rs(dbfield_user_rate) %></td>
    </tr>
    <tr> 
      <td> 性别: </td>
      <td> <% =rs(dbfield_user_sex) %></td>
    </tr>
	<% if session("userlevel")>7 then %>
<tr>
 <td> 最后进入时间: </td>
 <td> <% =rs(dbfield_user_lasttime) %></td>
</tr>

<% end if
 else %>
<script language="javascript">
<!--
alert("没有这个用户的资料！");
-->
</script>
<%
end if
rs.close
set rs=nothing
end if
my_conn.close
set my_conn=nothing%>
</table>
<br>
<form method="POST" action="userinfo.asp">
      <table border=0 align="center">
        <tr> 
          <td align="center">用户昵称：<input type="text" name="name" size="20" maxlength=20>
            <input type="submit" value="查看" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
            <input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()">
          </td>
        </tr>
      </table>
</form>
</body>
</html>