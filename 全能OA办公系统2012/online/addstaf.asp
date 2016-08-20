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
sub addstaf(href)
	oabusyuserdept=request.cookies("oabusyuserdept")
	if request("submit")="增加" then
		errorinfo=""
		username=request("username")
		if strlength(username)>10 then
			errorinfo=errorinfo&"用户名太长，不能超过5个汉字或10个英文字符！<br>"
		end if
		password=request("password")
		if strlength(password)>20 then
			errorinfo=errorinfo&"密码太长，不能超过20个字符！<br>"
		end if
		name=request("name")
		if strlength(name)>10 then
			errorinfo=errorinfo&"姓名太长，不能超过5个汉字或10个英文字符！<br>"
		end if
		userdept=oabusyuserdept
		userlevel="员工"
		if errorinfo="" then
		'判断是否有与申请的用户名相同的
			on error resume next
			set conn=opendb("oabusy","conn","accessdsn")
			conn.begintrans
			set rs=server.createobject("adodb.recordset")
			sql="select * from userinf where username=" & sqlstr(username) & " or password=" & sqlstr(password)
			rs.open sql,conn,1
			if not rs.eof and not rs.bof then
%>
<center><font color="red" size="+1">
<p align="center">用户名为<%=keepformat(username)%>的用户已经存在，请选择其他用户名</font><br><br>
<input type="button" onclick="history.go( -1 );return true;" value="返回"></center>
<%   
			else
				sql = "Insert Into userinf (username,password,name,userdept,userlevel) Values( "
				sql = sql & SqlStr(username) & ", "
				sql = sql & SqlStr(password) & ", "
				sql = sql & SqlStr(name) & ", "
				sql = sql & SqlStr(userdept) & ", "
				sql = sql & SqlStr(userlevel) & ")"
				conn.Execute(sql)
				set rs1=server.createobject("adodb.recordset")
				sql="SELECT @@IDENTITY AS IdSum from userinf"
				rs1.open sql,conn,1
				IdSum=rs1("IdSum")
				set rs1=nothing
				if err.number<>0 then
					conn.rollbacktrans
					call DispErrorInfo1("对不起，测试版本不能添加用户！")
					conn.close
					set conn=nothing
					 
					response.end
				else
					conn.committrans
				end if
%>
<br><br><font color="red" size="+1">用户名为<%=keepformat(username)%>的用户注册成功！</font><br><br>
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
 	var l1=maxlength(document.form2.username.value,1,10);
	if(!l1)
	{
		window.alert("用户名的长度不能超过5个汉字或10个英文字符！");
		document.form2.username.focus();
		return (false);
	}

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
		window.alert("姓名的长度不能超过5个汉字或10个英文字符！");
		document.form2.name.focus();
		return (false);
	}
}
</script>
<form action="<%=href%>" method=post name="form2" onsubmit="return form_check();">
<div align="center">
  <center>
<table border="1" cellspacing="0" cellpadding="2" bordercolorlight="#808080" bordercolordark="#D4D0C8" width="300">
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
用&nbsp;户&nbsp;名：
</td>
<td width="249">
<input type=text name="username" size=20 maxlength="10"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
密&nbsp;&nbsp;&nbsp;&nbsp;码：
</td>
<td width="249">
<input type="password" name="password" size=20 maxlength="20"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
密码确认：
</td>
<td width="249">
<input type="password" name="repassword" size=20 maxlength="20"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
姓&nbsp;&nbsp;&nbsp;&nbsp;名：
</td>
<td width="249">
<input type="text" name="name" size=20 maxlength="10"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
部&nbsp;&nbsp;&nbsp;&nbsp;门：
</td>
<td width="249">
<%=oabusyuserdept%>　
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
职&nbsp;&nbsp;&nbsp;&nbsp;位：
</td>
<td width="249">
员工
</td>
</tr>
<tr height="25">
<td align=center colspan="2">
<input type="submit" name="submit" value="增  加">&nbsp; <input type="reset" value="取  消">
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
