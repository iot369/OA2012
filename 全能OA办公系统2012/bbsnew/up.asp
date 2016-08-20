<!--#include file="conn.asp"--><!--#include file="fun.asp"-->
<%response.buffer=true
timeset="0"
bd=request.querystring("bd")
set bbs=myconn.execute("select*from bbsinfo")
tl=bbs(0)
response.write"<title>"&tl&"</title>"
url=request("url")
if bd<>"" then
sty="bbs"&bd
else
sty="all"
end if
sp=request.cookies(cn&"1")(sty)
c1=request.cookies(cn&"1")(sty&"c1")
c2=request.cookies(cn&"1")(sty&"c2")
if sp="" then sp=""&bbs("style")&""
if c1="" then c1=bbs(1)
if c2="" then c2=bbs(2)
c3=bbs(3)
todaynum=bbs(4)
mostonline=bbs(6)
mosttopic=bbs(7)
topinfo=bbs(8)
upsize=bbs("upsize")
upnum=bbs("upnum")
td=FormatDateTime(now+timeset/24,2)
if td<>bbs(5) then
myconn.execute("update bbsinfo set nyr='"&td&"'")
myconn.execute("update bbsinfo set todaynum='0'")
end if
set bbs=nothing
id=request.querystring("id")
lgname=Request.Cookies(cn)("lgname")
lgpwd=request.cookies(cn)("lgpwd")
st=Request.Cookies(cn)("style")
if st="" then st="a"
if lgname="" then
lolo="会员登陆"
else
lolo="重新登陆"
mailnewno=myconn.execute("select count(tname) from hand where tname='"&lgname&"' and isnew='0'")(0)
if mailnewno>0 then
tishi="<a href=bbsmail.asp?id=0><img border=0 src=pic/newmail.gif> 新留言</a>"
mailnewno="<font color="&c1&">"&mailnewno&"</font>"
end if
info="<a onmouseover=ShowMenu(userhelp,130)><SPAN style='CURSOR: hand' >用户助手</span></a><img border=0 src=pic/"&sp&"2.gif>"
liu="<a href=bbsmail.asp?id=0>留言板 <b>"&mailnewno&"</b></a>"
ex="<a href=login.asp?action=exit>退出论坛</a><img border=0 src=pic/"&sp&"2.gif>"
end if
set cjbz=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if not cjbz.eof then
admin="yes" 
gl="<a href=admin-index.asp>论坛管理</a><img border=0 src=pic/"&sp&"2.gif><a href=elselist.asp?action=hsz>回 收 站</a><img border=0 src=pic/"&sp&"2.gif>"
else
end if
set cjbz=nothing
ip=request.servervariables("remote_addr")
%>
<SCRIPT language=JavaScript>
 var h;
 var w;
 var l;
 var t;
 var topMar = 1;
 var leftMar = -2;
 var space = 1;
 var isvisible;
 var global = window.document
 global.fo_currentMenu = null
 global.fo_shadows = new Array
function HideMenu() 
{
 var mX;
 var mY;
 var vDiv;
 var mDiv;
	if (isvisible == true)
{
		vDiv = document.all("menuDiv");
		mX = window.event.clientX + document.body.scrollLeft;
		mY = window.event.clientY + document.body.scrollTop;
		if ((mX < parseInt(vDiv.style.left)) || (mX > parseInt(vDiv.style.left)+vDiv.offsetWidth) || (mY < parseInt(vDiv.style.top)-h) || (mY > parseInt(vDiv.style.top)+vDiv.offsetHeight)){
			vDiv.style.visibility = "hidden";
			isvisible = false;
		}
}
}

function ShowMenu(vMnuCode,tWidth) {
	vSrc = window.event.srcElement;
	vMnuCode = "<table id='submenu' cellspacing=1 cellpadding=3 style='width:"+tWidth+"' class=tableborder onmouseout='HideMenu()'><tr height=23><td nowrap align=left class=tablebody>" + vMnuCode + "</td></tr></table>";

	h = vSrc.offsetHeight;
	w = vSrc.offsetWidth;
	l = vSrc.offsetLeft + leftMar+4;
	t = vSrc.offsetTop + topMar + h + space-2;
	vParent = vSrc.offsetParent;
	while (vParent.tagName.toUpperCase() != "BODY")
	{
		l += vParent.offsetLeft;
		t += vParent.offsetTop;
		vParent = vParent.offsetParent;
	}

	menuDiv.innerHTML = vMnuCode;
	menuDiv.style.top = t;
	menuDiv.style.left = l;
	menuDiv.style.visibility = "visible";
	isvisible = true;
}
var stylelist='<img src=pic/fl.gif> <a href=style.asp?skin=a&bd=<%=bd%>>蓝色幻想</a><br><img src=pic/fl.gif> <a href=style.asp?skin=b&bd=<%=bd%>>绿色世界</a><br><img src=pic/fl.gif> <a href=style.asp?skin=c&bd=<%=bd%>>红色天空</a><br><img src=pic/fl.gif> <a href=style.asp?skin=d&bd=<%=bd%>>金黄稻田</a><br><img src=pic/fl.gif> <a href=style.asp?skin=e&bd=<%=bd%>>银灰酷色</a><br>'
var userhelp= '<img src=pic/fl.gif> <%=liu%><br><img src=pic/fl.gif> <a href=myinfo.asp>修改资料</a><br><img src=pic/fl.gif> <a href=elselist.asp?action=mytop>我发表的帖子</a><br><img src=pic/fl.gif> <a href=elselist.asp?action=withmetop>我参与的帖子</a><br><img src=pic/fl.gif> IP：<%=ip%></a>'
var bbsmenu= '<img src=pic/fl.gif> <a href=elselist.asp?action=jh>精 华 区</a><br><img src=pic/fl.gif> <a href=elselist.asp?action=new>查看新帖</a><br><img src=pic/fl.gif> <a href=userlist.asp>用户列表</a><br><img src=pic/fl.gif> <a href=paihang.asp>论坛排行</a><br>'
</SCRIPT>
<%response.write"<DIV id=menuDiv style='Z-INDEX: 2; VISIBILITY: hidden;  POSITION: absolute;'></DIV><body onmousemove=HideMenu() topmargin=0 background='pic/"&sp&"bg.gif'><link rel=stylesheet type=text/css href=css.css><style><!--body{SCROLLBAR-FACE-COLOR:"&c1&";SCROLLBAR-HIGHLIGHT-COLOR: "&c2&"; SCROLLBAR-SHADOW-COLOR:"&c1&"; SCROLLBAR-3DLIGHT-COLOR: "&c1&"; SCROLLBAR-ARROW-COLOR:"&c2&"; SCROLLBAR-TRACK-COLOR:"&c2&"; FONT-FAMILY: 宋体; SCROLLBAR-DARKSHADOW-COLOR:"&c1&"}a:hover{text-decoration:none;color:"&c1&"}TD.TableBody {BACKGROUND-COLOR: "&c2&"}.tableBorder {BORDER-RIGHT: 1px; BORDER-TOP: 1px; BORDER-LEFT: 1px; WIDTH: 80%; BORDER-BOTTOM: 1px; BACKGROUND-COLOR: "&c1&"}</STYLE>--></style>"&_
"<div><center><table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% height=4 bgcolor="&c1&"></td></tr><tr><td style='border: 1px solid "&c1&"' width=100% height=25 bgcolor="&c2&" background=pic/"&sp&"1.gif>&nbsp; <a href=login.asp>"&lolo&"</a><img border=0 src=pic/"&sp&"2.gif><a href=zhuce.asp>会员注册</a><img border=0 src=pic/"&sp&"2.gif>"&_
"<a onmouseover=ShowMenu(bbsmenu,90)><SPAN style='CURSOR: hand' >论坛菜单</span></a><img border=0 src=pic/"&sp&"2.gif><a onmouseover=ShowMenu(stylelist,90)><SPAN style='CURSOR: hand' >论坛样式</span></a><img border=0 src=pic/"&sp&"2.gif>"&info&""&ex&""&gl&""&tishi&"</td></tr></table></center></div>"&_
"<div><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% ><p style='margin: 7'></td></tr></table></center></div>"
response.write"<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% height=26 background=pic/4.gif><tr><td style='border: 1px solid "&c1&"' width=100% >&nbsp;<img border=0 src=pic/gonggao.gif align=absmiddle><b>你的位置：</b><a href=index.asp>"&tl&"</a>"
if bd<>"" then
set wei=myconn.execute("select bdname,pass,type from bdinfo where bn="&bd&" and key<>'0'")
bdtype=wei("type")
set cjbz1=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='"&bd&"'")
if not cjbz1.eof then
canre="yes"
end if
set cjbz1=nothing
function bdlogin(nt)
if bdtype=0 then
exit function
else
select case bdtype
case"1"
if lgname="" then
userin="no" 
noyes="进 入 失 败 ！"
mes="<font color="&c1&"><b>你不能成功的进入该版面，可能存在以下问题：</b></font><br>● 该版面为只有注册会员可以进入！<br> ● 你还没有<a href=login.asp>登陆</a>！<br><br>"
end if
case"2"
if nt=1 then
exit function
else
if canre="yes" or admin="yes" then
exit function
else
noyes="操 作 失 败 ！"
mes="<font color="&c1&"><b>你不能"&canre&"成功的操作该版面，可能存在以下问题：</b></font><br>● 该版面为只读版面，只有管理员或斑竹能够操作！<br><br>"
userin="no" 
end if
end if
case"3"
if isnull(pass) or pass="" then
userin="ok"
else
user=split(pass,",")
for i = 0 to ubound(user)
if lgname=trim(user(i)) and lgname<>"" then
userin="ok"
exit for
else
userin="no"
end if
next
end if
if userin="no" then
noyes="进 入 失 败 ！"
mes="<font color="&c1&"><b>你不能成功的进入该版面，可能存在以下问题：</b></font><br>● 该版面为认证论坛，你还没有得版主的认证！<br> ● 你还没有<a href=login.asp>登陆</a>！<br><br>"
else
end if
end select
if userin="no" then%><br><!--#include file="mes.asp"--><br><!--#include file="down.asp"--><%
response.end
end if
end if
end function
pass=wei("pass")
response.write"→ <a href=list.asp?bd="&bd&">"
wz=wei("bdname")
response.write""&wz&"</a>"
end if
set wei=nothing
if instr(Request("url"),"elselist.asp")>0 and request.querystring("action")="hsz" then
response.write"→ <a href='elselist.asp?action=hsz'>回 收 站</a>"
end if
if instr(Request("url"),"bbsmail.asp")>0 or instr(Request("url"),"mailcon.asp")>0 then
response.write"→ <a href='bbsmail.asp?id=0'>个人留言板</a>"
else
if id<>"" then
set w1=myconn.execute("select*from min where id="&id&"")
wzhuti=kbbs(w1("zhuti"))
wbody=w1("body")
wname=w1("name")
wriqi=w1("riqi")
wface=w1("face")
ggtype=w1("gonggao")
bbtype=w1("type")
isvote=w1("isvote")
if isvote=2 then wface="vote"
if bbtype=4 then wface="lock"
if bbtype=1 then wface="jing"
if ggtype=3 then wface="top"
if ggtype=5 then wface="alltop"
whits=w1("hits")
response.write"→ 浏览帖子："&LeftTrue(wzhuti,44)&""
set w1=nothing
end if
end if
response.write"</td></tr></table></center></div>"
%>