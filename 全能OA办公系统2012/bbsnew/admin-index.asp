<title>论坛管理</title><link rel="stylesheet" type="text/css" href="css.css">
<!--#include file="conn.asp"--><%
lgname=Request.Cookies(cn)("lgname")
lgpwd=request.cookies(cn)("lgpwd")
set cjbz=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if cjbz.eof then
noyes="登 陆 失 败 ！"
mes="你不能进入后台管理。<br>·你现在登陆论坛的用户名 "&lgname&" 不是管理员。·"%>
<!--#include file="mes.asp"-->
<%response.end
else%>
<frameset cols="20%,*" framespacing="0" border="0" frameborder="0">
  <frame name="left" src="admin-left.asp" scrolling="auto" target="right">
  <frame name="right" src="admin-right.asp" scrolling="auto" noresize>
  <noframes>
  <body>

  <p>此网页使用了框架，但您的浏览器不支持框架。</p>

  </body>
  </noframes>
</frameset>
<%
end if
%>