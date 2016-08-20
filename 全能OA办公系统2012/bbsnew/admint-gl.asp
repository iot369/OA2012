<!--#include file="conn.asp"--><!--#include file="fun.asp"-->
<link rel="stylesheet" type="text/css" href="css.css">
<%set bbs=myconn.execute("select*from bbsinfo")
sty="all"
sp=request.cookies(cn&"1")(sty)
c1=request.cookies(cn&"1")(sty&"c1")
c2=request.cookies(cn&"1")(sty&"c2")
if sp="" then sp="a"
if c1="" then c1=bbs(1)
if c2="" then c2=bbs(2)
set bbs=nothing
lgname=Request.Cookies(cn)("lgname")
lgpwd=request.cookies(cn)("lgpwd")
set cjbz=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if 1=2 then
noyes="登 陆 失 败 ！"
mes="你不能进入后台管理。<br>·你现在登陆论坛的用户名 "&lgname&" 不是管理员。·"%>
<!--#include file="mes.asp"-->
<%response.end
else
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='240' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% ><P style='MARGIN: 10px'>"
d2="</td></tr></table></center></div>"
menu=request.querystring("menu")%>
<body topmargin="0" leftmargin="0"><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=<%=c1%>>
  <tr>
    <td width="100%" height="28" background="pic/<%=sp%>3.gif" align="center">
    <b><font color="#FFFFFF">论坛后台管理系统</font></b></td>
  </tr>
  </table><br>
<%select case menu%>
<%case"bzgl"
go=request.querystring("go")
name=Replace(Request.form("name"),"'","''")
bd=Request.Form("bd")
set add=myconn.execute("SELECT*FROM user where name='"&name&"'")
if add.eof and add.bof then%>
<%=t1%>错 误 信 息<%=t2&d1%>·没有这个用户·<%=d2%>
<%
else
if go="add" then
pwd=add("password")
myconn.execute("insert into admin(name,password,bd)VALUES('"&name&"','"&pwd&"','"&bd&"')")
%>
<%=t1%>添 加 成 功<%=t2&d1%>·版主添加成功·<%=d2%><%
response.end
end if
if go="del" then
myconn.execute("delete*from admin where name='"&name&"' and bd='"&bd&"'")%>
<%=t1%>删 除 成 功<%=t2&d1%>·版主删除成功·<%=d2%><%
response.end
end if
end if
set add=nothing
%>
<%case"addbbs"%>
<%bbsbn=request.querystring("bbsbn")
set bf=myconn.execute("select*from bdinfo where key='0' order by bn")
if bf.eof then
noyes="错 误 信 息 ！"
mes="<br>·没有分类不能添加论坛！请先<a href=admin-right.asp?action=addfl>>>添加分类</a><br><br>"%>
<!--#include file="mes.asp"-->
<%
response.end
set bf=nothing
end if
%>
<%
bn=request.form("bn")
bdname=Replace(Request.Form("bdname"),"'","''")
bdinfo=Replace(Request.Form("bdinfo"),"'","''")
picurl=request.form("picurl")
key=request.form("key")
bbstype=request.form("bbstype")
if bn="" or bn="0" or bdname="" or bdinfo="" or key="" or key=0 or not isnumeric(bn) then
%>
<form method="POST">
<%response.write""&t1&"论 坛 添 加"&t2&""
%>
<%=d1%><P style='MARGIN: 5px'>论坛序号：<input type="text" name="bn" size="25"><font color=#FF0000>·</font>只能填 <b>0</b> 除外的数字
</p><P style='MARGIN: 5px'>论坛名称：<input type="text" name="bdname" size="25"><font color="#FF0000">·</font>无限制</p>
<P style='MARGIN: 5px'>标志图片：<input type="text" name="picurl" size="49">·将显示在论坛的右边，可以不填。</p>
<P style='MARGIN: 5px'>论坛介绍：<br><textarea rows="4" name="bdinfo" cols="58"></textarea><font color="#FF0000">·</font></p>
<P style='MARGIN: 5px'>属于分类：<select size="1" name="key" style="font-size: 9pt">
<%do while not bf.eof%><%if bf("bn")=bbsbn then%><option value="<%=bf("bn")%>" selected><%=bf("bdname")%></option><%else%>
<option value="<%=bf("bn")%>"><%=bf("bdname")%></option><%end if%>
<%
bf.movenext
Loop
bf.Close
set bf=nothing
%>
</select><font color="#FF0000">·</font>请选择该论坛属于哪一种分类</p><br><P style='MARGIN: 4px'>
论坛类型：(请从下面选择一个)</p><P style='MARGIN: 4px'>
<input type="radio" value="0" name="bbstype" checked>普通论坛
（用户和游客可以自由的进入该类型论坛，·推荐·）<P style='MARGIN: 4px'>
<input type="radio" value="1" name="bbstype">会员论坛
（只有注册用户才能进入该类型论坛）</p><P style='MARGIN: 4px'>
<input type="radio" value="2" name="bbstype">锁定论坛（会员和游客只能浏览帖子，不能对该论坛的帖子回复等）</p><P style='MARGIN: 4px'>
<input type="radio" value="3" name="bbstype">认证论坛
（只有版主认证的注册用户才能进入该类型论坛）</p><br>
<P style='MARGIN: 4px'><input type="submit" value=" 提 交 " name="B1"> <input type="reset" value=" 重 置 " name="B2"></p><br><%=d2%>
</form>
<%
else
set bbsyn=myconn.execute("select bn from bdinfo where bn="&bn&" and key<>'0'")
if not bbsyn.eof then
noyes="错 误 信 息 ！"
mes="<br>·论坛序号 <b>"&bn&"</b> 已经存在！·请选择别的论坛序号<br><br>"%>
<!--#include file="mes.asp"--><%
response.end
end if
set bbsyn=nothing
id=bn&"1"
if bbstype<>3 then
myconn.execute("insert into bdinfo(id,bn,bdname,bdinfo,picurl,key,type)values("&id&","&bn&",'"&bdname&"','"&bdinfo&"','"&picurl&"','"&key&"',"&bbstype&")")
noyes="添 加 成 功 ！"
mes="<br>·添加论坛成功！<br><br>"
else
myconn.execute("insert into bdinfo(id,bn,bdname,bdinfo,picurl,key,pass,type)values("&id&","&bn&",'"&bdname&"','"&bdinfo&"','"&picurl&"','"&key&"','"&lgname&"',"&bbstype&")")
noyes="添 加 成 功 ！"
mes="<br>·添加论坛成功,此论坛为认证论坛，暂时只有管理员能够进入。<br>·你可以通过 <a href=admin-gl.asp?menu=bbsgl>管理</a> 项目来添加可以进入该论坛的用户<br><br>"
end if
%><!--#include file="mes.asp"-->
<%end if%>
<%case"addadmin"
adminname=Replace(Request.Form("adminname"),"'","''")
set isadd=myconn.execute("select*from admin where name='"&adminname&"' and bd='70767766'")
if not isadd.eof then
iadd="yes"
end if
set isadd=nothing
set yon=myconn.execute("select*from user where name='"&adminname&"'")
if yon.eof or iadd="yes" then
%><%=t1%>错 误 信 息<%=t2&d1%>·该用户名已经是管理员或者还没有<a target="_blank" href="zhuce.asp">注册</a>·<%=d2%>
<%else
pwd=yon("password")
myconn.execute("insert into admin(name,password,bd)values('"&adminname&"','"&pwd&"','70767766')")
%><%=t1%>添 加 成 功<%=t2&d1%>·已经成功的添加管理员 <%=adminname%> ·<br><%=d2%><%end if
set yon=nothing%>
<%case"deladmin"
adminname=Replace(Request.Form("adminname"),"'","''")
yon=myconn.execute("select count(name) from admin where bd='70767766'")(0)
if yon<=1 then
%><%=t1%>错 误 信 息<%=t2&d1%>·论坛至少要有一个管理员·<%=d2%>
<%else
myconn.execute("delete*from admin where name='"&adminname&"' and bd='70767766'")
%><%=t1%>删 除 成 功<%=t2&d1%>·已经成功的删除管理员 <%=adminname%> ·<br><%=d2%><%end if%>
<%case"updatebbs"
id=request.querystring("id")
set old=myconn.execute("select bn from bdinfo where id="&id&"")
oldbn=old("bn")
set old=nothing
bn=Replace(Request.Form("bn"),"'","''")
set sbb=myconn.execute("select bdname from bdinfo where bn="&bn&" and id<>"&id&" and key<>'0'")
if not sbb.eof then
sb="no"
sbsb="<br>·填写的论坛序号已经被 <b>"&kbbs(sbb("bdname"))&"</b> 使用，请另外选择别的序号·"
end if
set sbb=nothing
bdname=Replace(Request.Form("bdname"),"'","''")
bdinfo=Replace(Request.Form("bdinfo"),"'","''")
picurl=request.form("picurl")
key=request.form("key")
bbstype=request.form("bbstype")
if bdname="" or bdinfo="" or bn="0" or not isnumeric(bn) or sb="no" then
%><%=t1%>错 误 信 息<%=t2&d1%>·请填写完整带有<font color="#FF0000">·</font>的项目·<br>·论坛序号必须为 <b>0</b> 除外的数字·<%=sbsb%><%=d2%><%else
if bbstype<>3 then
myconn.execute("update [bdinfo] set bdname='"&bdname&"',bdinfo='"&bdinfo&"',picurl='"&picurl&"',key='"&key&"',pass='',type="&bbstype&",bn="&bn&" where id="&id&"")
elseif bbstype="3" then
set dfdf=myconn.execute("select pass from bdinfo where id="&id&"")
if dfdf("pass")<>"" then
myconn.execute("update [bdinfo] set bdname='"&bdname&"',bdinfo='"&bdinfo&"',picurl='"&picurl&"',key='"&key&"',type=3,bn="&bn&" where id="&id&"")
else
myconn.execute("update [bdinfo] set bdname='"&bdname&"',bdinfo='"&bdinfo&"',picurl='"&picurl&"',key='"&key&"',pass='"&lgname&"',type=3,bn="&bn&" where id="&id&"")
end if
end if
myconn.execute("update min set bd="&bn&" where bd="&oldbn&"")
%><%=t1%>修 改 成 功<%=t2&d1%>·已经成功的修改了该版面的信息·<%=d2%>
<%end if%>
<%case"addpassuser"
user=Replace(Request.Form("user"),"'","''")
bn=request.querystring("bn")
myconn.execute("update bdinfo set pass='"&user&"' where bn="&bn&" and key<>'0'")
%><%=t1%>添 加 成 功<%=t2&d1%>·已经成功的添加了认证用户·<%=d2%>

<%case"deluser"%>
<%
delname=Replace(Request.form("delname"),"'","''")
set add=myconn.execute("SELECT name FROM user where name='"&delname&"'")
set isadmin=myconn.execute("select name from admin where name='"&delname&"'")
if add.eof or not isadmin.eof then
%>
<%=t1%>错 误 信 息<%=t2&d1%>不能删除，可能存在以下问题：<br><br>·没有这个用户·<br>·该用户是管理员(管理员不能删除，如果确实要把该用户删除，请先更改管理员，再把该用户删除)·<%=d2%>
<%
else
myconn.execute("delete*from user where name='"&delname&"'")
myconn.execute("delete*from min where name='"&delname&"'")
myconn.execute("delete*from admin where name='"&delname&"'")
myconn.execute("delete*from hand where tname='"&delname&"'")
%>
<%=t1%>删 除 成 功<%=t2&d1%>·已经成功的删除了用户以及这个用户的帖子和留言·<%=d2%><%end if
set isadmin=nothing
set add=nothing%>
<%case"fench"
bn=request.querystring("bn")
set fenfen=myconn.execute("select id,bdname from bdinfo where bn="&bn&" and key='0'")
id=fenfen("id")
fname=kbbs(fenfen("bdname"))
set fenfen=nothing
%><form method="POST" action="?menu=fenchok&id=<%=id%>">
<%response.write""&t1&"修 改 分 类"&t2&""&d1&""%>分类序号：<input type="text" name="xuxu" size="20" value="<%=bn%>"><br>
分类名称：<input type="text" name="fenn" size="20" value="<%=fname%>"> <input type="submit" value=" 提 交 " name="B1"> <input type="reset" value=" 重 置 " name="B2">
<%=d2%></form>
<%case"fenchok"
id=request.querystring("id")
fenn=Replace(Request.Form("fenn"),"'","''")
xuxu=Replace(Request.Form("xuxu"),"'","''")
set xo=myconn.execute("select bn from bdinfo where id="&id&"")
xox=xo("bn")
set xo=nothing
set xuyy=myconn.execute("select bn,bdname from bdinfo where bn="&xuxu&" and id<>"&id&" and key='0'")
if not xuyy.eof then
xy="<br>·该分类序号已经被 <b>"&kbbs(xuyy("bdname"))&"</b> 使用，请选用别的序号·"
end if
set xuyy=nothing
if fenn="" or xuxu="" or xuxu="0" or xy<>"" then
response.write""&t1&"修 改 失 败"&t2&""&d1&"·请输入分类名称以及正确的分类序号·"&xy&""&d2&""
else
myconn.execute("update bdinfo set key='"&xuxu&"' where key='"&xox&"'")
myconn.execute("update [bdinfo] set bdname='"&fenn&"',bn="&xuxu&" where id="&id&"")
response.write""&t1&"修 改 成 功"&t2&""&d1&"·修改分类名称成功"&d2&""
end if
%>

<%case"bbsgl"%><br>
<%response.write""&t1&"论 坛 管 理"&t2&""%>
<%=d1%>
<%
set bf=myconn.execute("select*from bdinfo where key='0' order by bn")
do while not bf.eof
%>
<table border="0" cellpadding="0" cellspacing="0" width="100%">
  <tr>
    <td width="29%" height="25"><font color=<%=c1%>><b><%=bf("bdname")%></font></b><%bbnn=bf("bn")%>　</td>
    <td width="51%"><a href="admin-gl.asp?menu=bdcon&dw=delfl&bn=<%=bbnn%>">删除此分类</a> | <a href="admin-gl.asp?menu=fench&bn=<%=bbnn%>">修改此分类</a> |
    <a href="admin-gl.asp?menu=addbbs&bbsbn=<%=bbnn%>">增加论坛</a> |</td>
    <td width="20%">分类序号：<b><font color=<%=c1%>><%=bbnn%></font></b></td>
  </tr>
  <%
set asd=myconn.execute("select*from bdinfo where key<>'0' and key='"&bbnn&"'order by bn")
do while not asd.eof
%><tr>
    <td width="29%" height="24">
·<%=asd("bdname")%></td>
    <td width="51%"><a href=admin-right.asp?action=chbbsinfo&id=<%=asd("id")%>>修改</a> | <a href="admin-gl.asp?menu=bdcon&dw=delbbs&bn=<%=asd("bn")%>">删除</a> | 
    <a href="admin-right.asp?action=delall&bn=<%=asd("bn")%>">清空帖子</a><%if asd("pass")<>"" then%> | 
    <a href="admin-right.asp?action=addpassuser&bn=<%=asd("bn")%>">添加认证用户</a><%end if%></td>
    <td width="20%">论坛序号：<b><%=asd("bn")%></b></td>
  </tr><%
asd.movenext
Loop
asd.Close
set asd=nothing
%>
</table>
<br><%
bf.movenext
Loop
bf.Close
set bf=nothing
%><br><br>
<%=d1%>说明：<br>一个分类可以包括一个或多个论坛，<br>分类与分类之间的序号不能相同，论坛与论坛之间的序号不能相同。<br>当你删除分类时，分类中的论坛也会被删除！<br>
各分类以及各论坛将按照序号显示出来，所以要放置在前面的分类或论坛的序号应该前一点<br><br><%=d2%><%=d2%>
<%case"bdcon"
dw=request.querystring("dw")
bn=request.querystring("bn")
if dw="delfl" then
myconn.execute("delete*from bdinfo where bn="&bn&" and key='0'")
myconn.execute("delete*from bdinfo where key='"&bn&"'")
%><%=t1%>删 除 成 功<%=t2&d1%>·删除论坛分类成功·<%=d2%>
<%response.end
end if
if dw="delbbs" then
myconn.execute("delete*from bdinfo where bn="&bn&" and key<>'0'")
%><%=t1%>删 除 成 功<%=t2&d1%>·删除论坛成功·<%=d2%>
<%end if%>
<%case"addlm"%>
<%
name=Replace(Request.Form("name"),"'","''")
url=Replace(Request.Form("url"),"'","''")
picurl=Replace(Request.Form("picurl"),"'","''")
if name="" or url="" or picurl="" then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·信息没有填写完整·<%=d2%>
<%else%>
<%myconn.execute("insert into lmbbs(url,picurl,name)values('"&url&"','"&picurl&"','"&name&"')")%>
<%=t1%>添 加 成 功<%=t2&d1%>·添加论坛联盟成功·<%=d2%><%end if%>
<%case"editlm"%>
<%name=Replace(Request.querystring("name"),"'","''")
url=Replace(Request.Form("url"),"'","''")
picurl=Replace(Request.Form("picurl"),"'","''")
if url="" or picurl="" then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·信息没有填写完整·<%=d2%>
<%else%>
<%myconn.execute("update [lmbbs] set url='"&url&"',picurl='"&picurl&"' where name='"&name&"'")%>
<%=t1%>编 辑 成 功<%=t2&d1%>·修改论坛联盟成功·<%=d2%><%end if%>
<%case"dellm"
name=Replace(Request.querystring("name"),"'","''")
if name="" then
%><%=t1%>错 误 信 息<%=t2&d1%>·请选择要删除的联盟名称·<%=d2%><%else
myconn.execute("delete*from lmbbs where name='"&name&"'")%>
<%=t1%>删 除 成 功<%=t2&d1%>·删除联盟成功·<%=d2%><%end if%><%case"updateuser"%>
<%
chname=Replace(Request.form("chname"),"'","''")
chqian=Replace(Request.form("chqian"),"'","''")
chmeili=Replace(Request.form("chmeili"),"'","''")
chjingyan=Replace(Request.form("chjingyan"),"'","''")
set add=myconn.execute("SELECT name FROM user where name='"&chname&"'")
if add.eof and add.bof then%>
<%=t1%>错 误 信 息<%=t2&d1%>·没有这个用户名·<%=d2%>
<%else%>
<%
if not isnumeric(chqian) or not isnumeric(chmeili) or not isnumeric(chjingyan) then%>
<%=t1%>错 误 信 息<%=t2&d1%>·金钱、魅力、经验必须为大于0的数字·<%=d2%>
<%else
myconn.execute("update [user] set qian='"&chqian&"',meili='"&chmeili&"',jingyan='"&chjingyan&"' where name='"&chname&"'")
%>
<%=t1%>修 改 成 功<%=t2&d1%>·已经成功的修改了用户的信息·<%=d2%><%end if
end if
set add=nothing
%><%case"chpwd"%><!--#include file="md5.asp"-->
<%
chaname=Replace(Request.form("chaname"),"'","''")
chapwd=Replace(Request.form("chapwd"),"'","''")
chapwd1=md5(chapwd)
set add=myconn.execute("SELECT*FROM user where name='"&chaname&"'")
if add.eof or chapwd="" then%>
<%=t1%>错 误 信 息<%=t2&d1%>·没有这个用户名或者没有填写新密码·<%=d2%>
<%else
myconn.execute("update user set password='"&chapwd1&"' where name='"&chaname&"'")
%>
<%=t1%>修 改 成 功<%=t2&d1%><%=kbbs(chaname)%> 的密码已经改为： <%=chapwd%><%=d2%><%end if
set add=nothing%>
<%case"delanymail"
daynum=request.form("daynum")
if not isnumeric(daynum) then
%><%=t1%>错 误 信 息<%=t2&d1%>·天数必须填写并且为数字·<%=d2%><%else
myconn.execute("delete*from hand where riqi<now-"&daynum&"")
%><%=t1%>删 除 成 功<%=t2&d1%>·批量删除留言成功·<%=d2%><%end if%>
<%case"delwhosemail"
ddname=Replace(Request.form("ddname"),"'","''")
if ddname="" then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·请输入用户名·<%=d2%><%else
myconn.execute("delete*from hand where tname='"&ddname&"'")
%><%=t1%>删 除 成 功<%=t2&d1%>·批量删除留言成功·<%=d2%><%end if%>
<%case"hbbbs"
frombd=request.form("frombd")
tobd=request.form("tobd")
myconn.execute("delete*from bdinfo where bn="&frombd&" and key<>'0'")
myconn.execute("update min set bd="&tobd&" where bd="&frombd&"")
%>
<%=t1%>合 并 成 功<%=t2&d1%>·论坛合并成功·<%=d2%>

<%case"delany"
daynum=request.form("daynum")
bd=request.form("bd")
if not isnumeric(daynum) then
%><%=t1%>错 误 信 息<%=t2&d1%>·天数必须填写并且为数字·<%=d2%><%else
if bd="all" then
myconn.execute("delete*from min where riqi<now-"&daynum&"")
else
myconn.execute("delete*from min where bd="&bd&" and riqi<now-"&daynum&"")
end if
%>
<%=t1%>删 除 成 功<%=t2&d1%>·批量删除帖子成功·<%=d2%><%end if%><%case"delnore"%>
<%daynum=request.form("daynum")
bd=request.form("bd")
if not isnumeric(daynum) then
%><%=t1%>错 误 信 息<%=t2&d1%>·天数必须填写并且为数字·<%=d2%><%else
if bd="all" then
myconn.execute("delete*from min where orders<now-"&daynum&"")
else
myconn.execute("delete*from min where bd="&bd&" and orders<now-"&daynum&"")
end if
%>
<%=t1%>删 除 成 功<%=t2&d1%>·批量删除帖子成功·<%=d2%><%end if%><%case"delwhose"
ddname=Replace(Request.form("ddname"),"'","''")
bd=request.form("bd")
if ddname="" then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·请输入用户名·<%=d2%><%else
if bd="all" then
myconn.execute("delete*from min where name='"&ddname&"'")
else
myconn.execute("delete*from min where bd="&bd&" and name='"&ddname&"'")
end if
%><%=t1%>删 除 成 功<%=t2&d1%>·批量删除帖子成功·<%=d2%><%end if%><%case"moveday"
daynum=request.form("daynum")
frombd=request.form("frombd")
tobd=request.form("tobd")
if not isnumeric(daynum) then
%><%=t1%>错 误 信 息<%=t2&d1%>·天数必须填写并且为数字·<%=d2%><%else
myconn.execute("update min set bd="&tobd&" where bd="&frombd&" and riqi<now-"&daynum&"")
%>
<%=t1%>移 动 成 功<%=t2&d1%>·批量移动帖子成功·<%=d2%><%end if%><%case"movename"
movename=Replace(Request.form("movename"),"'","''")
frombd=request.form("frombd")
tobd=request.form("tobd")
if movename="" then%>
<%=t1%>错 误 信 息<%=t2&d1%>·请输入用户名·<%=d2%><%else
myconn.execute("update min set bd="&tobd&" where bd="&frombd&" and name='"&movename&"'")
%><%=t1%>移 动 成 功<%=t2&d1%>·批量移动帖子成功·<%=d2%><%end if%><%case"bbs"
upnum=Replace(Request.form("upnum"),"'","''")
upsize=Replace(Request.form("upsize"),"'","''")
style=Replace(Request.form("style"),"'","''")
if not isnumeric(upnum) or not isnumeric(upsize) then
uuu="<br>·上传个数以及上传大小必须为数字·"
end if
tl=Replace(Request.form("tl"),"'","''")
c1=Replace(Request.form("c1"),"'","''")
c2=Replace(Request.form("c2"),"'","''")
topinfo=Replace(Request.form("topinfo"),"'","''")
if tl="" or c1="" or c2="" or upsize="" or upnum="" or style="" or not isnumeric(upnum) or not isnumeric(upsize) then
%><%=t1%>错 误 信 息<%=t2&d1%>·请填写完整必填项目·<%=uuu%><%=d2%>
<%else
myconn.execute("update [bbsinfo] set tl='"&tl&"',c1='"&c1&"',c2='"&c2&"',topinfo='"&topinfo&"',upnum="&upnum&",upsize="&upsize&",style='"&style&"'")%>
<%=t1%>修 改 成 功<%=t2&d1%>·论坛名称以及其他参数修改成功·<%=d2%><%end if%><%end select
end if%>