<title>��̳����</title><link rel="stylesheet" type="text/css" href="css.css">
<!--#include file="conn.asp"--><%
lgname=Request.Cookies(cn)("lgname")
lgpwd=request.cookies(cn)("lgpwd")
set cjbz=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if cjbz.eof then
noyes="�� ½ ʧ �� ��"
mes="�㲻�ܽ����̨����<br>�������ڵ�½��̳���û��� "&lgname&" ���ǹ���Ա����"%>
<!--#include file="mes.asp"-->
<%response.end
else%>
<frameset cols="20%,*" framespacing="0" border="0" frameborder="0">
  <frame name="left" src="admin-left.asp" scrolling="auto" target="right">
  <frame name="right" src="admin-right.asp" scrolling="auto" noresize>
  <noframes>
  <body>

  <p>����ҳʹ���˿�ܣ��������������֧�ֿ�ܡ�</p>

  </body>
  </noframes>
</frameset>
<%
end if
%>