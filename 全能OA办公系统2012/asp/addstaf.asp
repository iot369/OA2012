<!--#include file="sqlstr.asp"-->
<!--#include file="opendb.asp"-->
<%
sub addstaf(href)
oabusyuserdept=request.cookies("oabusyuserdept")
if request("submit")="����" then
username=request("username")
password=request("password")
name=request("name")
userdept=oabusyuserdept
userlevel="Ա��"
'�ж��Ƿ�����������û�����ͬ��
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(username) & " and password=" & sqlstr(password)
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red>�û���Ϊ<%=username%>���û��Ѿ����ڣ���ѡ�������û���</font><br>
<a href="javascript:void(null)" onclick="history.go( -1 );return true;"><font color="blue">����</font></a>
<%
else
sql = "Insert Into userinf (username,password,name,userdept,userlevel) Values( "
sql = sql & SqlStr(username) & ", "
sql = sql & SqlStr(password) & ", "
sql = sql & SqlStr(name) & ", "
sql = sql & SqlStr(userdept) & ", "
sql = sql & SqlStr(userlevel) & ")"
conn.Execute sql
%>
<font color=red>�û���Ϊ<%=username%>���û�ע��ɹ���</font><br>
<%
end if
else
%>

<script Language="JavaScript">
 function maxlength(str,minl,maxl) {
    if(str.length <= maxl && str.length >= minl){return true;}else{return false;}
                                    }

 function form_check(){
   var l1=maxlength(document.form2.username.value,1,20);
   if(!l1){window.alert("�û����ĳ��ȴ���1λС��20λ");document.form2.username.focus();return (false);}

   var l2=maxlength(document.form2.password.value,1,20);
   if(!l2){window.alert("����ĳ��ȴ���1λС��20λ");document.form2.password.focus();return (false);}

   var a1=document.form2.password.value;
   var a2=document.form2.repassword.value;
   if(a1!=a2){window.alert("�������������Ӧ��ͬ");document.form2.repassword.focus();return (false);}

   var l3=maxlength(document.form2.name.value,1,20);
   if(!l3){window.alert("�����ĳ��ȴ���1λС��20λ");document.form2.name.focus();return (false);}

                    }

</script>
<SCRIPT language=javascript>
<!--
if (window.Event) 
��document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
��if (e.which == 2 || e.which == 3)
�� return false;
 } else if (event.button == 2 || event.button == 3) {
�� event.cancelBubble = true
�� event.returnvalue = false;
�� return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;��// for IE5+
document.onmousedown = norightclick;���� // for all others
//-->
</SCRIPT>





<form action="<%=href%>" method=post name="form2" onsubmit="return form_check();">
<table border=0 cellspacing="1" bgcolor="B0C8EA">
<tr>
<td height="25" bgcolor="D7E8F8">
��&nbsp;��&nbsp;����
<input type=text name="username" size=20><font color=red>*</font></td>
</tr>
<tr>
<td height="25" bgcolor="D7E8F8">
��&nbsp;&nbsp;&nbsp;&nbsp;�룺
<input type="password" name="password" size=20><font color=red>*</font></td>
</tr>
<tr>
<td height="25" bgcolor="D7E8F8">
����ȷ�ϣ�
  <input type="password" name="repassword" size=20><font color=red>*</font></td>
</tr>
<tr>
<td height="25" bgcolor="D7E8F8">
��&nbsp;&nbsp;&nbsp;&nbsp;����
<input type="text" name="name" size=20><font color=red>*</font></td>
</tr>
<tr>
<td height="25" bgcolor="D7E8F8">
��&nbsp;&nbsp;&nbsp;&nbsp;�ţ�<%=oabusyuserdept%></td>
</tr>
<tr>
<td height="25" bgcolor="D7E8F8">
ְ&nbsp;&nbsp;&nbsp;&nbsp;λ��Ա��
</td>
</tr>
<tr>
<td height="25" align=center bgcolor="D7E8F8">
<input type="submit" name="submit" value="����"></td>
</tr>
</table>
</form>
<%
end if
end sub
%>