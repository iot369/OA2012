<!--#include file="up.asp"--><!--#include file="line.asp"-->
<%ggnum=5%><SCRIPT>
function showtb(tbnum)
{
whichEl = eval("tbtype" + tbnum);
if (whichEl.style.display == "none")
{
eval("tbtype" + tbnum + ".style.display=\"\";");
}
else
{
eval("tbtype" + tbnum + ".style.display=\"none\";");
}
}
</SCRIPT>
<%
'session.abandon
'Server.ScriptTimeOut=500
function opendb(DBPath,sessionname,dbsort)
dim conn
'if not isobject(session(sessionname)) then
Set conn=Server.CreateObject("ADODB.Connection")
'if dbsort="accessdsn" then conn.Open "DSN=" & DBPath
'if dbsort="access" then conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
'if dbsort="sqlserver" then conn.Open "DSN=" & DBPath & ";uid=wsw;pwd=wsw"
DBPath1=server.mappath("../db/sdoa.asa")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<%response.write"<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% ><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor=#111111 width=100% height=20><tr><td width=45% ><p style='margin: 6'>欢迎我们的新会员："%>
<% 
set newuser=myconn.execute("select top 1 name from user order by userid desc")
response.write"<a href='userinfo.asp?name="&kbbs(newuser("name"))&"'><font color="&c1&"><b>"&kbbs(newuser("name"))&"</b></font></a><br>"
set newuser=nothing      
regno=myconn.execute("select count(name)from user")(0)
response.write"注册会员：<b>"&regno&"</b><img border=0 src=pic/"&sp&"2.gif>现在时间：<b>"&FormatDateTime(now+timeset/24,4)&"</b>"
set regno=nothing
tienoa=myconn.execute("select count(*)from min where gonggao<>1")(0)
response.write"<br>总 帖 数：<b>"&tienoa&"</b><img border=0 src=pic/"&sp&"2.gif>"
set tienoa=nothing
tieno=myconn.execute("select count(*)from min where gonggao<>1 and bid=0")(0)
response.write"话题数：<b>"&tieno&"</b><br>最高日帖数：<b>"&mosttopic&"</b><img border=0 src=pic/"&sp&"2.gif>今日帖数：<font color="&c1&"><b>"&todaynum&"</b></font>"
set tieno=nothing
response.write"</td><td width=55% >"
igg=1
set gg=myconn.execute("select zhuti,id,face,bd from min where gonggao=1 and type<>5 order by id desc")
response.write"<marquee onmouseover='this.stop()' onmouseout='this.start()' scrollAmount='1' direction='up' width='100%' height='55'>"
do while not gg.eof 
response.write"<img src=pic/gonggao.gif border=0 align=absmiddle> <a href=show.asp?bd="&gg("bd")&"&id="&gg("id")&">"&kbbs(gg("zhuti"))&"</a><br>"
igg=igg+1
if igg>ggnum then exit do
gg.movenext
loop
response.write"</marquee>"
set gg=nothing
response.write"</td></tr></table></td></tr></table></center></div><br>"
if lgname="" then
response.write"<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>快 捷 登 陆</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><form method=POST action=bbselse.asp><td width=100% ><p style='margin: 4'><img border=0 src=pic/guest.gif align=absmiddle> 用户名：<input size='10' name='lgname'>&nbsp;密&nbsp; 码：<input type='password' size='10' value name='lgpwd'> Cookies：<select style='FONT-SIZE: 9pt' size='1' name='cook'><option value='j0' selected>不保存</option><option value='j1'>保存一天</option><option value='j30'>保存一月</option><option value='j365'>保存一年</option></select> <input type='submit' value=' 登 陆 ' name='B1'> <input type='reset' value=' 重 置 ' name='b2'></td></form></tr></table></center></div><br>"
end if
response.write"<div align=center><center><table height=25 border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>论 坛 列 表</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div>"
set bf=myconn.execute("select*from bdinfo where key='0' order by bn")
do while not bf.eof
bbnn=bf("bn")
response.write"<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% background=pic/"&sp&"1.gif height=25 bgcolor="&c1&">&nbsp;<a onclick=showtb(1"&bbnn&")><SPAN style='CURSOR: hand' ><img border=0 src=pic/fle.gif></span></a> <b><font color="&c1&">"&bf("bdname")&"</font></b></td></tr></table></center></div><div align=center><center><table id=tbtype1"&bbnn&" border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor='#111111' width='94%'><tr><td width='100%'>"
set asd=myconn.execute("select*from bdinfo where key<>'0' and key='"&bbnn&"'order by bn")
do while not asd.eof
set hane=myconn.execute("select riqi from min where gonggao<>1 and riqi>now-1/2 and bd="&asd("bn")&"")
if not hane.eof then
newtopic="yes"
end if
gif="on"
if asd("type")=1 then gif="member"
if asd("type")=2 then gif="lock"
if asd("type")=3 then gif="rz"
set hane=nothing
response.write"<table style='border-collapse: collapse' cellSpacing=0 cellPadding=0 border='1' bordercolor='"&c1&"' width='100%'><TBODY><TR><TD  align=middle width=48><img src='pic/"&gif&".gif'></TD><TD  vAlign=top width=*><TABLE cellSpacing=0 cellPadding=2 width='100%' border=0 style='border-collapse: collapse' bordercolor='#111111'><TBODY><TR><TD  width=*><p style='margin-left: 3; margin-top: 5'><a href=list.asp?bd="&asd("bn")&">"&asd("bdname")&"</a></TD><TD  align=middle width=40 rowSpan=2><TABLE align=left><TBODY><TR><TD>"
if asd("picurl")<>"" then
response.write"<a href=list.asp?bd="&asd("bn")&"><img border=0 src="&asd("picurl")&"></a>"
end if
response.write"</TD><TD width=20>&nbsp;&nbsp;&nbsp;</TD></TR></TBODY></TABLE></TD><TD  width='30%' rowSpan=2>"
set u1=myconn.execute("select top 1 * from min where bd="&asd("bn")&" and gonggao<>1 and gonggao<>4 and type<>5 order by id desc")
if u1.eof and u1.bof then
response.write"本版面还没有帖子！"
else
if u1("bid")=0 then
ub=u1("zhuti")
lb=kbbs(ub)
showid=u1("id")
else
ub=u1("body")
lb=kbbs(ub)
showid=u1("bid")
end if
response.write"作者：<a href='userinfo.asp?name="&kbbs(u1("name"))&"'>"&kbbs(u1("name"))&"</a> <br>时间："&u1("orders")&"<br>主题：<img align=absmiddle src=face/"&u1("face")&".gif> <a href=show.asp?id="&showid&"&bd="&asd("bn")&">"&LeftTrue(lb,30)&"</a>"
end if
set u1=nothing
response.write"</TD></TR><TR><TD width=*><p style='margin-left: 3;margin-top:3; margin-bottom:4'><img src='pic/tl.gif'> <font color=#808080>"&asd("bdinfo")&"</font></TD></TR><TR><TD class=tablebody2 width=* height=24 bgcolor="&c2&">&nbsp;"
cc=1
set cb=myconn.execute("select*from admin where bd='"&asd("bn")&"'")
if cb.eof or cb.bof then
response.write"该版面还没有斑竹！"
else
response.write"版面斑竹："
do while not cb.eof
response.write"<a href='userinfo.asp?name="&kbbs(cb("name"))&"'>"&kbbs(cb("name"))&"</a> | "
cc=cc+1
if cc>4 then exit do
cb.movenext
loop
end if
cb.Close
set cb=nothing
response.write"</TD><TD class=tablebody2 width=40 height=20 bgcolor='"&c2&"'></TD><TD class=tablebody2 vAlign=center width=200 bgcolor='"&c2&"'>"
tie=myconn.execute("select count(*)from min where gonggao<>1 and bd="&asd("bn")&"")(0)
response.write"总帖数："&tie&""
set tie=nothing
response.write"</TD></TR></TBODY></TABLE></TD></TR></TBODY></TABLE>"
asd.movenext
Loop
asd.Close
set asd=nothing
response.write"</td></tr></table></center></div>"
bf.movenext
Loop
bf.Close
set bf=nothing
response.write"<br><div align=center><center><table height=25 border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>在 线 统 计</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=100% >"&_
"<tr><td width=5% align=center height=50 rowspan=2><p><img border=0 src=pic/tj.gif></p>"&_
"</td><td height=25 width=95% >&nbsp;目前论坛总共有 <b>"&lineno&"</b> 人在线 。其中 <b>"&usno&"</b> 位会员 ， <b>"&nusno&"</b> 位游客。最高峰同时在线人数：<b>"&mostonline&"</b></td></tr><tr><td height=25 width=95% ><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor=#111111 width=100% ><tr>"
ni=1
ha=1
set useol=myconn.execute("select*from online order by ip")
do while not useol.eof
olna=useol("name")
if olna<>"" then
set fa=myconn.execute("select toupic from user where name='"&olna&"'")
mytx=fa("toupic")
set fa=nothing
iiii="<a href='userinfo.asp?name="&kbbs(olna)&"'><img align=absmiddle border=0 src='"&kbbs(mytx)&"' width='16' height='16'> "&kbbs(olna)&"</a>"
else
mytx="pic/youke.gif"
iiii="<img align=absmiddle border=0 src='"&mytx&"' width='16' height='16'> 游客"
end if
response.write"<td width='18%'>&nbsp;"&iiii&"</td>"
ha=ha+1
ni=ni+1
if ha>4 then
ha=1
response.write"</tr>"
end if
if ni>lineno then exit do
useol.movenext
Loop
useol.Close
set useol=nothing
set usno=nothing
set lineno=nothing
response.write"</table></td></tr></table></td></tr></table></td></tr></table></center></div><br>"
response.write"<div align=center><center><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=94% ><tr><div align=center><center><table height=25 border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>论 坛 联 盟</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 bordercolor="&c1&" cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=94% ><tr><td width=100% ><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=100% ><tr><td width=5% align=center height=45><img border=0 src=pic/tj.gif></td><td width=95% valign=top><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=100% ><tr>"
set lmbbs=myconn.execute("select*from lmbbs")
lmm=myconn.execute("select count(*) from lmbbs")(0)
los=1
do while not lmbbs.eof
response.write"<td width=12% > <p style='margin: 4'><a target='_blank' href='"&lmbbs("url")&"'>"&lmbbs("name")&"</a></td>"
lmbbs.movenext
los=los+1
if los>6 then 
los=1
response.write"</tr>"
end if
loop
lmbbs.close
set lmbbs=nothing
response.write"</tr></table></td></tr></table></td></tr></table></center></div><br>"&_
"<div align=center><center><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor=#111111 width=94% ><tr><td width=100% align=center><img border=0 src=pic/on.gif align=absmiddle> 普通论坛&nbsp; <img border=0 src=pic/member.gif align=absmiddle> 会员论坛&nbsp; <img border=0 src=pic/lock.gif align=absmiddle> 只读论坛&nbsp; <img border=0 src=pic/rz.gif align=absmiddle> 认证论坛</td></tr></table></center></div><br>"
%><!--#include file="down.asp"-->