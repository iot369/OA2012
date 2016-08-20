<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title><%=r_title%></title>
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
<script language="javascript">
<!--  
  function checkvalue(theform)
  
{ if (theform.p_username.value == "" || theform.p_username.value.length < 1 || theform.p_username.value.length > 10)
{ alert("用户名不能为空或用户名过长（10位以内）！");
  theform.p_username.focus();
  return false; }
  
  for (var i=0; i < theform.p_username.value.length; i++)
{ var pu=theform.p_username.value.substring(i,i+1)
      if (pu == '"' || pu == "'" || pu == "<" || pu == ">")
	  { alert("用户名中有非法字符，请重新输入！");
	    theform.p_username.focus();
	    return false; }  }
  
  if (theform.p_password.value == "" || theform.p_password.value.length < 1 || theform.p_password.value.length > 10)
{ alert("密码不能为空或密码过长（10位以内）！");
  theform.p_password.focus();
  return false; }

  for (var i=0; i < theform.p_password.value.length; i++)
{ var pp=theform.p_password.value.substring(i,i+1)
      if (pp == '"' || pp == "'" || pp == "<" || pp == ">")
	  { alert("密码中有非法字符，请重新输入！");
	    theform.p_password.focus();
	    return false; }  }
  
  if (theform.p_email.value == "" || theform.p_email.value.length < 1 || theform.p_email.value.length >30 )
{ alert("信箱不能为空或信箱过长（30位以内）！");
  theform.p_email.focus();
  return false; }

  if (theform.p_email.value.indexOf('@') == -1 || theform.p_email.value.indexOf('@') == 0 || theform.p_email.value.indexOf('@') == theform.p_email.value.length - 1)
{ alert("非法电子邮箱，请重新输入！");
  theform.p_email.focus();
  return false; }
  
  for (var i=0; i < theform.p_email.value.length; i++)
{ var pe=theform.p_email.value.substring(i,i+1)
      if (pe == '"' || pe == "'" || pe == "<" || pe == ">")
	  { alert("信箱中有非法字符，请重新输入！");
	    theform.p_email.focus();
	    return false; }  }
		
  return true;
}
//-->
</script>
</head>

<body bgcolor="<%=m_bg%>" text="<%=m_text%>">
<br><br><br>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <tr bgcolor="<%=m_bg1%>"> 
    <td height="3"></td>
  </tr>
  <tr bgcolor="#cccccc" align="center">
	  <%
  If Request.ServerVariables("Request_Method")="POST" and trim(request("p_username"))<>"" and trim(request("p_password"))<>"" then
     g_username = trim(request("p_username"))
     g_password = trim(request("p_password"))
     g_sex      = trim(request("p_sex"))
     g_email    = trim(request("p_email"))
     g_homepage = trim(request("p_homepage"))
     g_comefrom = trim(request("p_comefrom"))
     g_oicq     = trim(request("p_oicq"))
	 if g_username="大家" then
	    response.write "<td><br><br>此名称已被系统占用，请换个名字再注册！<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'><br><br></td>"
	 else
     set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&"='"&g_username&"'")
     if not rs.eof then
        response.write "<td><br><br>用户名已存在，请选择其它用户名！<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'><br><br></td>"
     else
        badword = Split(r_badword, "|")
	    if UBound(badword) <> 0 then
	       for each b_word in badword
		       if instr(g_username,b_word)<>0 then
		          response.write "<td><br><br>用户名含有非法字符！<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'><br><br></td>"
		          badwordcheck=1
			   end if
			next
        end if
		if badwordcheck <> 1 then
		   insertsql="'"&g_username&"','"&g_password&"','"&g_sex&"','"&g_email&"','"&g_homepage&"','"&g_comefrom&"','"&g_oicq&"'"
           my_conn.execute("insert into "&dbtable_user&" ("&dbfield_user_username&","&dbfield_user_password&","&dbfield_user_sex&","&dbfield_user_email&","&dbfield_user_homepage&","&dbfield_user_comefrom&","&dbfield_user_oicq&") values ("&insertsql&")")
		   response.write "<td><br><br>恭喜, 您已注册成功！<br><form method='post' action='chat.asp'><input type='hidden' name='l_username' value="&g_username&"><input type='hidden' name='l_password' value="&g_password&"><input type='submit' value='进入聊天室' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"'></form><br></td>"
        end if
    end if
	rs.close
	set rs=nothing
	my_conn.close
	set my_conn=nothing
	end if
  else %>
  <form action="regsiter.asp" method="POST" onsubmit="return checkvalue(this);"> 
	<td>
        <br><%=r_title%> 新 用 户 注 册<br>
        <br>姓 名：<input type="text" name="p_username" maxlength="10" value="<%=trim(request("f_username"))%>">
        <br>密 码：<input type="text" name="p_password" maxlength="10">
        <br>性 别：<input type="radio" name="p_sex" value="boy" checked>boy
                   <input type="radio" name="p_sex" value="girl">girl
		<br>信 箱：<input type="text" name="p_email" maxlength="30">
        <br>主 页：<input type="text" name="p_homepage" value="http://">
        <br>来 自：<input type="text" name="p_comefrom">
        <br>OICQ ：<input type="text" name="p_oicq">
        <br>
        <br>
        <input type="submit" value="确定" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        <input type="reset"  value="重填" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        <input type="button" value="返回" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="location='index.asp'">
        <br>
        <br>
      </td>
    </form>
	<% end if %>
  </tr>
  <tr bgcolor="<%=m_bg1%>"> 
    <td height="3"></td>
  </tr>
</table>
</body>
</html>
