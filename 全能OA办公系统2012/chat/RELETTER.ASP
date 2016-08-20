<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp" -->
<html>
<head>
<title>回复留言</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body      {  font-size: 10.4pt }
td        {  font-size: 10.4pt }
.normal   {  <%=m_button%> ; font-size: 9pt}
.over     {  <%=m_buttonover%> ; font-size: 9pt}
.down     {  <%=m_buttondown%> ; font-size: 9pt}
-->
</style>
</head>

<body text="<%=m_text2%>" bgcolor="#cccccc">
<%
towho   = trim(request("p_towho"))
message = trim(request("p_message"))
message=replace(message,chr(13),"<br>")
if message<>"" and  towho<>"" then
   set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&"='"& session("user") &"'")
   homepage=""
   email=""
   comefrom=""
   if not rs.eof then
      homepage=rs(dbfield_user_homepage)
      email=rs(dbfield_user_email)
      comefrom=rs(dbfield_user_comefrom)
   end if
   insertsql = "'" & now & "','" & towho & "', '" & session("user") & "', '" & email & "','" & homepage & "','"  &  message & "','" & comefrom & "','" & request("R1") & "'"
   my_conn.execute("insert into "&dbtable_gbook&" ("&dbfield_gbook_addtime&","&dbfield_gbook_name&","&dbfield_gbook_lyname&","&dbfield_gbook_email&","&dbfield_gbook_homepage&","&dbfield_gbook_message&","&dbfield_gbook_comefrom&","&dbfield_gbook_picture&") values (" & insertsql & ")")
   rs.close
   set rs=nothing
   my_conn.close
   set my_conn=nothing
%><br><br>
<center><strong><font color="<%=m_text2%>">您的留言已发送</font></strong><br><br>
<input type="button" value="返回" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:history.go(-1)"> 
<input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()">
</center>
<%
   response.end
   end if 
%>
<table width="560" border="1" cellspacing="0" cellpadding="3" bordercolorlight="#FFFFFF" bordercolordark="#999999">
 <form action="reletter.asp" method="post">
  <tr bgcolor="<%=m_bg1%>"> 
    <td colspan="2" align="center" style="color:<%=m_text1%>">回 复 留 言</td>
  </tr>
  <tr> 
    <td width="300">留言者：<%=session("user")%><br>
      留言给：<%=request("n")%>
	  <input type="hidden" name="p_towho" value="<%=trim(request("n"))%>"><br>
      留　言：<textarea name="p_message" cols="33" wrap="VIRTUAL" style="font-size:9pt"></textarea>
      <br>
      <div align="right"> 
        <input type="submit" value=" 留 言 " class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
        <input type="reset" value=" 重 填 " class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
		<input type="button" value=" 返 回 " onclick="location.href='letter.asp'" class="normal" onMouseOver="this.className='over';" onMouseDown="this.className='down';" onMouseOut="this.className='normal';">
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
</body>
</html>