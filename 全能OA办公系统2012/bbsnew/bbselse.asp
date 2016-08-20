<!--#include file="conn.asp"--><!--#include file="md5.asp"-->
<%
menu=request.querystring("menu")
select case menu
case"vote"
function laiyuan()
laiyuan=false
come=Request.ServerVariables("HTTP_REFERER")
here=Request.ServerVariables("SERVER_NAME")
if mid(come,8,len(here))<>here then
laiyuan=false
else
laiyuan=true
end if
end function
laiyuan()
if laiyuan=false then
response.redirect"index.asp"
end if
comeurl=Request.ServerVariables("HTTP_REFERER")
type1=request.querystring("type")
id=request.querystring("id")
set hhh=myconn.execute("select*from vote where id="&id&"")
outtime=hhh("outtime")
ddd=hhh("vote")
nno=hhh("votenum")
ddd=split(ddd,"|")
nno=split(nno,"|")
nnn=ubound(ddd)
if type1=1 then
xuan1=request.form("xuan")
for i=1 to nnn
if i=cint(xuan1) then
nno(i)=nno(i)+1
end if
num=num&"|"&nno(i)
next
else
dim xuan(10)
dim num1(10)
for i=1 to nnn
xuan(i)=request.form("xuan_"&i&"")
xuan1=xuan(i)&xuan1
next
for j=1 to nnn
if xuan(j)="" then num1(j)=nno(j)
if cint(xuan(j))=j then
num1(j)=nno(j)+1
end if
num=num&"|"&num1(j)
next
end if
canvote="yes"
lgname=Request.Cookies(cn)("lgname")
lgpwd=Request.Cookies(cn)("lgpwd")
if xuan1="" then canvote="no"
if lgname="" then
canvote="no"
else
set uss=myconn.execute("select name from user where name='"&lgname&"' and password='"&lgpwd&"'")
if uss.eof then
canvote="no"
end if
set uss=nothing
set had=myconn.execute("select user from voted where user='"&lgname&"' and id="&id&"")
if not had.eof then
canvote="no"
end if
set had=nothing
end if
if now>outtime then
canvote="no"
end if
if canvote="yes" then
myconn.execute("update vote set votenum='"&num&"' where id="&id&"")
myconn.execute("update min set orders='"&now&"' where id="&id&"")
myconn.execute("insert into voted(id,user,votenum)values("&id&",'"&lgname&"','"&xuan1&"')")
end if
response.redirect comeurl
case""
comeurl=Replace(Request.Form("comeurl"),"'","''")
if comeurl="" then
comeurl=Request.ServerVariables("HTTP_REFERER")
end if
lgname=Replace(Request.Form("lgname"),"'","''")
lgpwd=Replace(Request.Form("lgpwd"),"'","''")
lgpwd=md5(lgpwd)
cook=Replace(Request.Form("cook"),"'","''")
Response.Cookies(cn)("lgname")=lgname
Response.Cookies(cn)("lgpwd")=lgpwd
if cook="j1" then
Response.Cookies(cn).Expires=date+1
Response.Cookies(cn).Expires=date+1
elseif cook="j30" then
Response.Cookies(cn).Expires=date+30
Response.Cookies(cn).Expires=date+30
elseif cook="j365" then
Response.Cookies(cn).Expires=date+365
Response.Cookies(cn).Expires=date+365
elseif cook="j0" then
Response.Cookies(cn)("lgname")=lgname
Response.Cookies(cn)("lgpwd")=lgpwd
end if
%>
<%set lg=myconn.execute("select*from user where name='"&lgname&"' and password='"&lgpwd&"'")
if lg.eof and lg.bof then
Response.Cookies(cn)("lgname")=""
Response.Cookies(cn)("lgpwd")=""
myconn.close
set myconn=nothing
%><!--#include file="up.asp"-->
<%t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
myconn.close
set myconn=nothing
response.write"<br><br><form>"&t1&"登 陆 失 败"&t2&d1&"<br><P style='MARGIN: 10px'>·你的用户名或密码错误·</p><P style='MARGIN: 10px'>·<a href='javascript:history.go(-1)'>返回重新登陆</a>·</p>"&d2&"</form>"
else
myconn.execute("update [user] set qian=qian+50,meili=meili+8,jingyan=jingyan+8 WHERE name='"&lgname&"'")
myconn.close
set myconn=nothing
%>
<!--#include file="up.asp"-->
<%t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
myconn.close
set myconn=nothing
response.write"<br><br><form>"&t1&"登 陆 成 功"&t2&d1&"<br><P style='MARGIN: 10px'>·<a href='index.asp'>进入论坛首页</a>·</p><P style='MARGIN:10px'>·<a href='"&kbbs(comeurl)&"'>"&kbbs(comeurl)&"</a>·</p>"&d2&"</form>"
end if%><br><!--#include file="conn.asp"--><!--#include file="down.asp"--><%end select%>