<%response.expires=0%>
<!--#include file="sqlstr.asp"-->
<!--#include file="opendb.asp"-->
<!--#include file="../inc/public.asp"-->
<%
function strlength(inputstr)
	dim length,i
	length=0
	for i=1 to len(inputstr)
		if asc(mid(inputstr,i,1))<0 then
			length=length+2
		else
			length=length+1
		end if
	next
	strlength=length
end function
sub userinf(href)
	oabusyusername=request.cookies("oabusyusername")
	oabusyuserdept=request.cookies("oabusyuserdept")
	oabusyuserlevel=request.cookies("oabusyuserlevel")
	if request("submit")="确认修改" then
		errorinfo=""
		password=request("password")
		if strlength(password)>20 then
			errorinfo=errorinfo&"密码太长，不能超过20个字符！<br>"
		end if
		name=request("name")
		if strlength(name)>10 then
			errorinfo=errorinfo&"姓名太长，不能超过5个汉字或10个英文字符！<br>"
		end if
		if errorinfo="" then
			on error resume next
			set conn=opendb("oabusy","conn","accessdsn")
			conn.begintrans
			sql = "update userinf set "
			sql = sql & "password=" & SqlStr(password) & ", "
			sql = sql & "name=" & SqlStr(name) & " where username=" & sqlstr(oabusyusername)
sqlstr(oabusyusername)
			conn.Execute(sql)
			if err.number<>0 then
				conn.rollbacktrans
				call DispErrorInfo1("修改用户信息出错！出错原因："&err.description)
				conn.close
				set conn=nothing
				response.end
			else
				conn.committrans
%>
<br><br>
<font color="red" size="+1">用户资料修改成功！</font>
<br><br>
<%
			end if
		else
%>
<div align="center">
<table width="80%" border="0">
<tr><td>
<center><b><font color="red" size="+1">出错了</font></b></center><br><br>
<font color="#ee0000" size="+1"><%=errorinfo%></font>
<center><input type="button" value="返回" onclick="history.go( -1 );return true;"></center>
</td></tr></table>
</div>
<%
		end if
	else
		
%>
<script Language="JavaScript">
function maxlength(str,minl,maxl)
{
	if(str.length <= maxl && str.length >= minl)
	{
		return true;
	}
	else
	{
		return false;
	}
}

function form_check()
{
	var l2=maxlength(document.form2.password.value,1,20);
	if(!l2)
	{
		window.alert("密码的长度大于1位小于20位");
		document.form2.password.focus();
		return (false);
	}

	var a1=document.form2.password.value;
	var a2=document.form2.repassword.value;
	if(a1!=a2)
	{
		window.alert("两次输入的密码应相同");
		document.form2.repassword.focus();
		return (false);
	}
	
	var l3=maxlength(document.form2.name.value,1,10);
	if(!l3)
	{
		window.alert("姓名的长度不能超过5个汉字或10个字符！");
		document.form2.name.focus();
		return (false);
	}
}
</script>
<%
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	on error resume next
rs.open "select count(*) as countss from userinf",conn,1,1
usercount=rs("countss")
if usercount >500 then
   rs.close
   set rs=nothing
   %>
   <script language=javascript>
       window.alert("对不起，超过了最大使用用户数，请删除部分用户！");
       location.href="/usercontrol.asp";
   </script>
<%
end if
rs.close
	sql="select * from userinf where username=" & sqlstr(oabusyusername)
	rs.open sql,conn,1
%>
<br><br>
<form action="<%=href%>" method=post name="form2" onsubmit="return form_check();">
<table border="1" cellspacing="0" cellpadding="5" bordercolorlight="#808080" bordercolordark="#D4D0C8" width="400">
<tr height="25">
<td width="166" align="right" height="23" bgcolor="#D4D0C8">
用&nbsp;户&nbsp;名：
</td>
<td width="228" height="23" bgcolor="#D4D0C8">
<%=oabusyusername%>　
</td>
</tr>
<tr height="25">
<td width="166" align="right" height="23">
密&nbsp;&nbsp;&nbsp;&nbsp;码：
</td>
<td width="228" height="23">
<input type="password" name="password" size=20 value="<%=rs("password")%>" maxlength="20">
</td>
</tr>
<tr height="25">
<td width="166" align="right" height="23">
密码确认：
</td>
<td width="228" height="23">
<input type="password" name="repassword" size=20 value="<%=rs("password")%>" maxlength="20">
</td>
</tr height="25">
<tr>
<td width="166" align="right" height="23">
姓&nbsp;&nbsp;&nbsp;&nbsp;名：
</td>
<td width="228" height="23">
<input type="text" name="name" size=20 value="<%=rs("name")%>" maxlength="10">
</td>
</tr>
<tr height="25">
<td width="166" align="right" height="23">
部&nbsp;&nbsp;&nbsp;&nbsp;门：
</td>
<td width="228" height="23">
<%=oabusyuserdept%>　
</td>
</tr>
<tr height="25">
<td width="166" align="right" height="23">
职&nbsp;&nbsp;&nbsp;&nbsp;位：
</td>
<td width="228" height="23">
<%=oabusyuserlevel%>　
</td>
</tr>
<tr height="25">
<td align=center colspan="2" height="25" width="396">
<input type="submit" name="submit" value="确认修改">
</td>
</tr>
</table>
  </center>
</div>
</form>
<%
	end if
end sub
%>