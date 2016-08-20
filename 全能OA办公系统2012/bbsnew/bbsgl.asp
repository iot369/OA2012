<!--#include file="up.asp"-->
<br>
<%action=request.querystring("action")
htt=Request.ServerVariables("HTTP_REFERER")
bd=request.querystring("bd")
id=request.querystring("id")
re=request.querystring("re")
if re="yes" then
tex="·删除帖子---成功！！<br>·<a href="&htt&">回到帖子</a><meta http-equiv=refresh content='2;url="&htt&"'>"
elseif re="no" then
tex="·删除帖子---成功！！"
end if
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
if action="qk" and admin="yes" then
myconn.execute("delete*from min where type=5")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·已经清空回收站<%=d2%>
<%response.end
end if
set who=myconn.execute("select name,bd from min where id="&id&"")
bd=who("bd")
upwho=who("name")
set who=nothing
function bdlist(sename)
response.write"<select size=1 style='font-size: 9pt' name="&sename&" >"
set bf=myconn.execute("select*from bdinfo where key<>'0'")
do while not bf.eof
response.write"<option value="&bf("bn")&">"&bf("bdname")&"</option>"
bf.movenext
loop
bf.close
set bf=nothing
response.write"</select>"
end function
set del1=myconn.execute("select*from admin where name='"&lgname&"' and password='"&lgpwd&"' and (bd='70767766' or bd='"&bd&"')")%><%if not del1.eof or not del1.bof then
%>
<%if action="move" then%>
<form method="POST" action="?action=moveok&id=<%=id%>&bd=<%=bd%>"><%=t1%>移动帖子<%=t2&d1%>
<p style='margin: 15'>请选择帖子要移动到的论坛：<%bdlist("tobd")%>   <input type="submit" value=" 移 动 " name="B1"> <input type="reset" value=" 重 置 " name="B2"></p>
<%=d2%></form>
<%end if%>
<%if action="moveok" then%>
<%tobd=request.form("tobd")
set cob=myconn.execute("select*from min where bid=0 and id="&id&" and bd="&bd&"")
if cob.eof then%>
<%=t1%>操 作 失 败 ！<%=t2&d1%><p style='margin: 15'>·没有该帖子！！<%=d2%>
<%else
myconn.execute("update min set bd="&tobd&" where id="&id&" or bid="&id&"")
%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·移动帖子---成功！！<%=d2%>
<%end if
set cob=nothing%>
<%end if%>
<%if action="jh" then
myconn.execute("update min set type=1 where id="&id&"")
myconn.execute("update [user] set qian=qian+800,meili=meili+200,jingyan=jingyan+200 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·设定为精华帖子---成功！！<%=d2%>
<%elseif action="notjh" then
myconn.execute("update min set type=0 where id="&id&"")
myconn.execute("update [user] set qian=qian-500,meili=meili-100,jingyan=jingyan-100 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·取消帖子精华---成功！！<%=d2%>
<%
elseif action="top" then
myconn.execute("update [min] set gonggao='3' where id="&id&"")
myconn.execute("update [user] set qian=qian+800,meili=meili+200,jingyan=jingyan+200 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·设定为置顶帖子---成功！！<%=d2%>
<%
elseif action="lock" then
myconn.execute("update [min] set type='4' where id="&id&"")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·加锁帖子---成功！！<%=d2%>
<%
elseif action="unlock" then
myconn.execute("update [min] set type='0' where id="&id&"")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·解锁帖子---成功！！<%=d2%>
<%
elseif action="nottop" then
myconn.execute("update [min] set gonggao='0' where id="&id&"")
myconn.execute("update [user] set qian=qian-500,meili=meili-100,jingyan=jingyan-100 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·取消置顶帖子---成功！！<%=d2%>
<%
elseif action="alltop" then
myconn.execute("update [min] set gonggao='5' where id="&id&"")
myconn.execute("update [user] set qian=qian+1000,meili=meili+300,jingyan=jingyan+300 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15'>·设定为总置顶帖子---成功！！<%=d2%>

<%elseif action="del" then
myconn.execute("update min set type=5 where id="&id&"")
%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15;line-height: 150%'><%=tex%></p><%=d2%>
<%elseif action="dely" then
if admin<>"yes" then response.end
myconn.execute("delete*from min where id="&id&" or bid="&id&"")
myconn.execute("update [user] set qian=qian-250,meili=meili-40,jingyan=jingyan-40 where name='"&upwho&"'")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15;line-height: 150%'>·永久删除帖子成功·</p><%=d2%>
<%elseif action="hy" then
if admin<>"yes" then response.end
myconn.execute("update min set type=0 where id="&id&"")%>
<%=t1%>操 作 成 功 ！<%=t2&d1%><p style='margin: 15;line-height: 150%'>·还原帖子成功·</p><%=d2%>
<%end if
set del1=nothing
%>
<%else%>
<%=t1%>操 作 失 败 ！<%=t2&d1%><p style='margin: 15'>·你不是该版面的的斑竹或总斑竹，不能操作该帖！！<%=d2%>
<meta http-equiv="refresh" content="2;url=javascript:history.go(-1)">
<%end if%><br><br><!--#include file="down.asp"-->