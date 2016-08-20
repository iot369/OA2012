<!--#include file="sqlstr.asp"-->
<!--#include file="opendb.asp"-->
<%
sub userinf(href)
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")

if request("submit")="更改" then
password=request("password")
name=request("name")
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql = "update userinf set "
sql = sql & "password=" & SqlStr(password) & ", "
sql = sql & "name=" & SqlStr(name) & " where username=" & sqlstr(oabusyusername)
conn.Execute sql
%>
<br><br>
<font color=red>用户资料维护成功！</font>
<%
else
%>

<script Language="JavaScript">
 function maxlength(str,minl,maxl) {
    if(str.length <= maxl && str.length >= minl){return true;}else{return false;}
                                    }

 function form_check(){

   var l2=maxlength(document.form2.password.value,1,20);
   if(!l2){window.alert("密码的长度大于1位小于20位");document.form2.password.focus();return (false);}

   var a1=document.form2.password.value;
   var a2=document.form2.repassword.value;
   if(a1!=a2){window.alert("两次输入的密码应相同");document.form2.repassword.focus();return (false);}

   var l3=maxlength(document.form2.name.value,1,20);
   if(!l3){window.alert("姓名的长度大于1位小于20位");document.form2.name.focus();return (false);}

                    }

</script>


<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
%>
<br><br>
<form action="<%=href%>" method=post name="form2" onsubmit="return form_check();">
  <table border=0 cellpadding="0" cellspacing="1" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="B0C8EA">
    <tr>
<td height="25" bgcolor="#F3F9FA">
用&nbsp;户&nbsp;名：<%=oabusyusername%></td>
</tr>
<tr>
<td height="25" bgcolor="#F3F9FA">
密&nbsp;&nbsp;&nbsp;&nbsp;码：
<input type="password" name="password" size=20 value="<%=rs("password")%>"></td>
</tr>
<tr>
<td height="25" bgcolor="#F3F9FA">
密码确认：
  <input type="password" name="repassword" size=20 value="<%=rs("password")%>"></td>
</tr>
<tr>
<td height="25" bgcolor="#F3F9FA">
姓&nbsp;&nbsp;&nbsp;&nbsp;名：
<input type="text" name="name" size=20 value="<%=rs("name")%>"></td>
</tr>
<tr>
<td height="25" bgcolor="#F3F9FA">
部&nbsp;&nbsp;&nbsp;&nbsp;门：<%=oabusyuserdept%></td>
</tr>
<tr>
<td height="25" bgcolor="#F3F9FA">
职&nbsp;&nbsp;&nbsp;&nbsp;位：<%=oabusyuserlevel%></td>
</tr>
<tr>
<td height="25" align=center bgcolor="#F3F9FA">
<input type="submit" name="submit" value="更改"></td>
</tr>
</table>
</form>
<%
end if
end sub
%>