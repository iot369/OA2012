<!--#include file="up.asp"--><!--#include file="md5.asp"-->
<br><br><%
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
menu=request.querystring("menu")
select case menu
case""
%><form method=POST name=kbbs action=?menu=get>
<%=t1%>�� �� �� ��<%=t2%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
    <tr>
      <td width="40%"><p style="margin: 10">����������û���</p></td>
      <td width="60%">&nbsp;<input type="text" name="name" size="20"></td>
    </tr>
    <tr>
      <td><p style="margin: 10">��������ı�������</p></td>
      <td>&nbsp;<input name="anhao" size="20"></td>
    </tr>
    <tr>
      <td colspan="2" align="center" background="pic/3.gif" height="28">
      <input type="submit" value=" �� �� " name="B1">
      <input type="reset" value=" �� �� " name="B2"></td>
    </tr>
    </table>
  </center>
</div>
</form>
<%case"get"
name=Replace(Request.Form("name"),"'","''")
anhao=Replace(Request.Form("anhao"),"'","''")
anhao=md5(anhao)
set pwd=myconn.execute("select name from user where name='"&name&"' and anhao='"&anhao&"'")
if pwd.eof then
%><%=t1%>�� �� �� Ϣ<%=t2&d1%><p style="margin: 10">������д����Ϣ����
<a href="javascript:history.go(-1)"> <img border="0" src="pic/re.gif"> �� ��</a></p><%=d2%><%else
myconn.execute("update user set password='"&anhao&"' where name='"&name&"'")
myconn.execute("update admin set password='"&anhao&"' where name='"&name&"'")
%><%=t1%>�� �� �� ��<%=t2&d1%><p style="margin: 10">��<b><%=kbbs(name)%></b> �������Ѿ�����Ϊ�������롤</p><%=d2%><%end if%><%end select%><br><!--#include file="down.asp"-->