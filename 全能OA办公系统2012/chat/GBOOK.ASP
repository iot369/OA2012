<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp" -->
<html>
<head>
<title><%=r_title%> 留 言 簿</title>
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
towho   = trim(request("p_name"))
who     = trim(request("p_lyname"))
message = trim(request("p_message"))
if towho<>"" and who<>"" and message<>"" then 
   message=replace(message,chr(13),"<br>")
   if towho<>"all" then towho=trim(request("p_tname"))
   set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&"='"& who &"'")
   homepage=""
   email=""
   comefrom=""
   if not rs.eof then
      homepage=rs(dbfield_user_homepage)
      email=rs(dbfield_user_email)
      comefrom=rs(dbfield_user_comefrom)
   end if
   insertsql = "'" & now & "','" & towho & "', '" & who & "', '" & email & "','" & homepage & "','"  &  message & "','" & comefrom & "','" & request("R1") & "'"
   my_conn.execute("insert into "&dbtable_gbook&" ("&dbfield_gbook_addtime&","&dbfield_gbook_name&","&dbfield_gbook_lyname&","&dbfield_gbook_email&","&dbfield_gbook_homepage&","&dbfield_gbook_message&","&dbfield_gbook_comefrom&","&dbfield_gbook_picture&") values (" & insertsql & ")")
   rs.close
   set rs=nothing
   if towho <>"all" then
%><br><br><br>
<center><strong><font color="<%=m_text2%>">您的留言已发送</font></strong><br><br>
<input type="button" value="返回" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:history.go(-1)"> 
<input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()">
</center>
<%
   response.end
   end if 
   end if %>
<table width="560" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999" align="center">
  <form action="gbook.asp" method="post">
  <tr bgcolor="<%=m_bg1%>"> 
    <td colspan="2" align="center" style="color:<%=m_text1%>"><%=r_title%> 留 言 簿</td>
  </tr>
  <tr> 
    <td width="300">留言者：<%=session("user")%><br>
	  <input type="hidden" name="p_lyname" value="<%=session("user")%>">
      留言给：<input type="radio" name="p_name" value="all" checked>
      大家<input type="radio" name="p_name" value="notall">
      留言给
      <input type="text" name="p_tname" size="14" style="font-size:9pt">
      <br>
      留　言：
<textarea name="p_message" cols="33" wrap="VIRTUAL" style="font-size:9pt"></textarea>
      <br>
      <div align="right"> 
        <input type="submit" value=" 留 言 " class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
        <input type="reset" value=" 重 填 " class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
        <input type="button" value=" 关 闭 " onclick="window.close();" class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
        &nbsp;&nbsp;</div>
    </td>
    <td valign="top">表情： <br>
      <input type="radio" name="R1" value="1" checked>
      <img src="img/1.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="3">
      <img src="img/3.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="5">
      <img src="img/5.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="7">
      <img src="img/7.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="9">
      <img src="img/9.gif" width="15" height="15">
      <input type="radio" name="R1" value="11">
      <img src="img/11.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="13">
      <img src="img/13.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="15">
      <img src="img/15.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="17">
      <img src="img/17.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="19">
      <img src="img/19.gif" width="15" height="15">
      <input type="radio" name="R1" value="2">
      <img src="img/2.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="4">
      <img src="img/4.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="6">
      <img src="img/6.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="8">
      <img src="img/8.gif" width="15" height="15">
      <input type="radio" name="R1" value="10">
      <img src="img/10.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="12">
      <img src="img/12.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="14">
      <img src="img/14.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="16">
      <img src="img/16.gif" width="15" height="15">
      <input type="radio" name="R1" value="18">
      <img src="img/18.gif" width="15" height="15"> 
      <input type="radio" name="R1" value="20">
      <img src="img/20.gif" width="15" height="15">
    </td>
  </tr>
 </form>
</table>
<br>
<table width="560" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999" align="center">
  <tr align="center" bgcolor="<%=m_bg1%>"> 
    <td colspan="2" style="color:<%=m_text1%>">大 家 的 留 言</td>
  </tr>
<%
set rs=my_conn.execute("select * from "&dbtable_gbook&" where "&dbfield_gbook_name&"='all' order by "&dbfield_gbook_id&" desc")
if not rs.eof then
  page=cint(request("page"))
  do while npage<page
     npage=npage+1
     rs.movenext
  loop
  do while not rs.eof
     if i=10 then exit do
     i=i+1%>
  <tr> 
    <td align="center" valign="center"><img src="img/<%=rs(dbfield_gbook_picture)%>.gif"></td>
    <td>
	<% if rs(dbfield_gbook_comefrom)<>"" then %>
	来自 <font color="blue"><%=rs(dbfield_gbook_comefrom)%></font> 的
	<% end if %> <font color="red"><%=rs(dbfield_gbook_lyname)%></font> 
		<% if rs(dbfield_gbook_homepage)<>"" then %>
		<a href="<%=rs(dbfield_gbook_homepage)%>"><img src="img/home.gif" width="16" height="16" border="0"></a>
		<% end if %>
		<% if rs(dbfield_gbook_email)<>"" then %>
	    <a href="mailto:<%=rs(dbfield_gbook_email)%>"><img src="img/mail.gif" width="16" height="16" border="0"></a>
		<% end if %>
	于 <font color="yellow"><%=rs(dbfield_gbook_addtime)%></font> 对大家说:<br>
	<%=rs(dbfield_gbook_message)%>
  </tr>
  <tr>
    <td colspan="2" height="5"></td>
  </tr>
<%
rs.movenext
loop
%>
  <tr>
    <td colspan="2" align="right">
    <% if page=0 then %>
       [上一页]
    <% else %>
       <a href="gbook.asp?page=<% =paeg-10 %>">[上一页]</a>
    <% end if %>
    <% if rs.eof then %>
       [下一页]
    <% else %>
       <a href="gbook.asp?page=<%=page+10%>">[下一页]</a>
	<% end if %>
	</td>
  </tr>
<%
rs.close
set rs=nothing
my_conn.close
set my_conn=nothing
end if
%>
</table>
</body>
</html>
