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
	if request("submit")="����" then
		errorinfo=""
		username=request("username")
		if strlength(username)>10 then
			errorinfo=errorinfo&"�û���̫�������ܳ���5�����ֻ�10��Ӣ���ַ���<br>"
		end if
		password=request("password")
		if strlength(password)>20 then
			errorinfo=errorinfo&"����̫�������ܳ���20���ַ���<br>"
		end if
		name=request("name")
		if strlength(name)>10 then
			errorinfo=errorinfo&"����̫�������ܳ���5�����ֻ�10��Ӣ���ַ���<br>"
		end if
		userdept=oabusyuserdept
		userlevel="Ա��"
		if errorinfo="" then
		'�ж��Ƿ�����������û�����ͬ��
			on error resume next
			set conn=opendb("oabusy","conn","accessdsn")
			conn.begintrans
			set rs=server.createobject("adodb.recordset")
			sql="select * from userinf where username=" & sqlstr(username) & " or password=" & sqlstr(password)
			rs.open sql,conn,1
			if not rs.eof and not rs.bof then
%>
<center><font color="red" size="+1">
<p align="center">�û���Ϊ<%=keepformat(username)%>���û��Ѿ����ڣ���ѡ�������û���</font><br><br>
<input type="button" onclick="history.go( -1 );return true;" value="����"></center>
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
					call DispErrorInfo1("�Բ��𣬲��԰汾��������û���")
					conn.close
					set conn=nothing
					 
					response.end
				else
					conn.committrans
				end if
%>
<br><br><font color="red" size="+1">�û���Ϊ<%=keepformat(username)%>���û�ע��ɹ���</font><br><br>
<%
			end if
		else
%>
<div align="center">
<table width="80%" border="0">
<tr><td>
<center><b><font color="red" size="+1">������</font></b></center><br><br>
<font color="#ee0000" size="+1"><%=errorinfo%></font>
<center><input type="button" value="����" onclick="history.go( -1 );return true;"></center>
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
		window.alert("�û����ĳ��Ȳ��ܳ���5�����ֻ�10��Ӣ���ַ���");
		document.form2.username.focus();
		return (false);
	}

	var l2=maxlength(document.form2.password.value,1,20);
	if(!l2)
	{
		window.alert("����ĳ��ȴ���1λС��20λ");
		document.form2.password.focus();
		return (false);
	}

	var a1=document.form2.password.value;
	var a2=document.form2.repassword.value;
	if(a1!=a2)
	{
		window.alert("�������������Ӧ��ͬ");
		document.form2.repassword.focus();
		return (false);
	}

	var l3=maxlength(document.form2.name.value,1,10);
	if(!l3)
	{
		window.alert("�����ĳ��Ȳ��ܳ���5�����ֻ�10��Ӣ���ַ���");
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
��&nbsp;��&nbsp;����
</td>
<td width="249">
<input type=text name="username" size=20 maxlength="10"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
��&nbsp;&nbsp;&nbsp;&nbsp;�룺
</td>
<td width="249">
<input type="password" name="password" size=20 maxlength="20"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
����ȷ�ϣ�
</td>
<td width="249">
<input type="password" name="repassword" size=20 maxlength="20"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
��&nbsp;&nbsp;&nbsp;&nbsp;����
</td>
<td width="249">
<input type="text" name="name" size=20 maxlength="10"><font color=red>*</font>
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
��&nbsp;&nbsp;&nbsp;&nbsp;�ţ�
</td>
<td width="249">
<%=oabusyuserdept%>��
</td>
</tr>
<tr height="25">
<td width="249" bgcolor="#D4D0C8" align="right">
ְ&nbsp;&nbsp;&nbsp;&nbsp;λ��
</td>
<td width="249">
Ա��
</td>
</tr>
<tr height="25">
<td align=center colspan="2">
<input type="submit" name="submit" value="��  ��">&nbsp; <input type="reset" value="ȡ  ��">
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
