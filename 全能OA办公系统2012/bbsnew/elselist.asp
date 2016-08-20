<!--#include file="ubbcode.asp"-->
<%
action=request.querystring("action")
pai=request.querystring("pai")
if pai="" then
pai="orders"
end if
%>
<!--#include file="up.asp"-->
<script language="javascript">
function Check(){var Name=document.form.topage.value;
document.location='?bd=<%=bd%>&topage='+Name+'';
}
</script>
<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="94%" height="40">
    <tr>
      <td width="100%"> · <%if lgname<>"" and action<>"hsz" then%><a href="elselist.asp?action=mytop">我发表的主题</a> · <a href="elselist.asp?action=withmetop">我参与的主题</a> · <%end if%><%if action<>"hsz" then%><a href="elselist.asp?action=new">查看新帖</a> ·<%end if%><%if action="hsz" and admin="yes" then%><a href="bbsgl.asp?action=qk">清空回收站</a> ·<%end if%></td>
    </tr>
  </table>
  </center>
</div>
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
select case action
case"hsz"
if admin<>"yes" then 
noyes="操 作 失 败"
mes="<p style='margin: 9'>你不是管理员，不能浏览回收站！</p>"%>
<!--#include file="mes.asp"-->
<%
response.end
end if
href1="<a href=?topage=1&action=hsz>"
sql = "select * from min where type=5 order by "&pai&" desc"
count=myconn.execute("select count(*)from min where bid=0 and type=5")(0)
case"jh"
href1="<a href=?topage=1&action=jh>"
sql = "select * from min where bid=0 and gonggao<>1 and gonggao<>4 and type=1 order by "&pai&" desc"
count=myconn.execute("select count(*)from min where bid=0 and gonggao<>1 and gonggao<>4 and type=1")(0)
case"new"
href1="<a href=?topage=1&action=new>"
sql = "select * from min where bid=0 and gonggao<>1 and gonggao<>4 and type<>5 order by "&pai&" desc"
count=myconn.execute("select count(*)from min where bid=0 and gonggao<>1 and gonggao<>4 and type<>5")(0)
case"mytop"
href1="<a href=?topage=1&action=mytop>"
sql = "select * from min where bid=0 and gonggao<>1 and gonggao<>4 and name='"&lgname&"' and type<>5 order by gonggao desc,"&pai&" desc"
count=myconn.execute("select count(*)from min where bid=0 and gonggao<>1 and gonggao<>4 and type<>5 and name='"&lgname&"'")(0)
case"withmetop"
href1="<a href=?topage=1&action=withmetop>"
sql = "select * from min where type<>5 and (id in (select bid from min where gonggao<>1 and gonggao<>4 and name='"&lgname&"') or (bid=0 and gonggao<>1 and gonggao<>4 and name='"&lgname&"')) order by gonggao desc,"&pai&" desc"
count=myconn.execute( "select count(*) from min where type<>5 and (id in (select bid from min where gonggao<>1 and gonggao<>4 and name='"&lgname&"') or (bid=0 and gonggao<>1 and gonggao<>4 and name='"&lgname&"'))")(0)
end select
on error resume next
pagesetup=20
rs.Open sql,myConn,1
If Count/pagesetup > (Count\pagesetup) then
TotalPage=(Count\pagesetup)+1
else TotalPage=(Count\pagesetup)
End If
PageCount= 0
RS.MoveFirst
if Request.QueryString("ToPage")<>"" then PageCount = cint(Request.QueryString("ToPage"))
if PageCount <=0 then PageCount = 1
if PageCount > TotalPage then PageCount = TotalPage
RS.Move (PageCount-1) * pagesetup
%>
<div align=center><center>
<table bgcolor=<%=c1%> border=1 cellpadding=0 cellspacing=0 style="border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px" bordercolor=<%=c1%> width="94%" id=AutoNumber5 height="27">
<tr>
  <td width=4% align=center background="pic/<%=sp%>3.gif">
  <b><font color="#FFFFFF">状态</font></b></td>
  <td width=42% align=center background="pic/<%=sp%>3.gif">
  <b><font color="#FFFFFF">主 题</font><font color="#FFFFFF"> ( 点小图标在新窗口浏览 )</font></b></td>
<td width=9% align=center background="pic/<%=sp%>3.gif"><b><font color="#FFFFFF">作 者</font></b></td>
  <td width=9% align=center background="pic/<%=sp%>3.gif">
  <b><font color="#FFFFFF">回复/</font><font color="#FFFFFF">人气</font></b></td>
<td width=16% align=center background="pic/<%=sp%>3.gif">
<b><font color="#FFFFFF">最后更新时间</font></b></td>
  <td width=9% align=center background="pic/<%=sp%>3.gif">
  <b><font color="#FFFFFF">最后回复</font></b></td></tr></table></center></div>
<%
i=1
do while not rs.eof
rrzz=kbbs(rs("zhuti"))
if action="hsz" then
if rs("bid")<>0 then
set las=myconn.execute("select zhuti from min where id="&rs("bid")&"")
rrzz="RE:"&las("zhuti")
set las=nothing
end if
end if
set last=myconn.execute("select body,name from min where bid="&rs("id")&" order by id desc")
if last.eof then
zui="-----"
rb=rs("body")
tltltl="帖子内容："&kbbs(rb)&""
else
lb=last("body")
zui="<a href=userinfo.asp?name="&last("name")&">"&last("name")&"</a>"
tltltl="最后跟帖："&kbbs(lb)&""
set last=nothing
end if
fface=rs("face")
if rs("type")=1 then fface="jing"
response.write"<div align=center><center><table width='94%' border=1 cellpadding=0 cellspacing=0 style='TABLE-LAYOUT: fixed; WORD-BREAK: break-all;border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px' bordercolor="&c1&"  id=AutoNumber5><tr><td width=4% height=23 align=center><a target=_blank href=show.asp?id="&rs("id")&"&bd="&rs("bd")&"><img border=0 src=face/"&fface&".gif></a></td><td width=42% height=27 align=left onmouseover=javascript:this.bgColor='"&c2&"' onmouseout=javascript:this.bgColor=''>&nbsp;<a  target=_self href=show.asp?id="&rs("id")&"&bd="&rs("bd")&" title="&LeftTrue(tltltl,25)&">"&LeftTrue(rrzz,44)&"</a>"
bno1=myconn.execute("select count(name)from min where bid="&rs("id")&"")(0)
if bno1>10 then
response.write"<img align=absmiddle border=0 src=pic/hot.gif>"
end if
response.write"</td><td width=9% height=23 align=center><a href='userinfo.asp?name="&kbbs(rs("name"))&"'>"&kbbs(rs("name"))&"</a></td><td width=9% height=23 align=center>"&bno1&"/"&rs("hits")&"</td><td width=16% height=23 align=center>"&rs("orders")&"</td><td width=9% height=23 align=center>"&zui&"</td></tr></table></center></div>"
i=i+1
if i>pagesetup then exit do
rs.movenext
loop
rs.Close
%>
  <div align="center">
    <center>
<TABLE bgcolor=<%=c1%> borderColor=<%=c1%> cellSpacing=0 cellPadding=0 width="94%" border=1 style="border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px">
<TBODY>
<TR height=25>
<TD height=2>
<TABLE cellSpacing=0 cellPadding=3 width="100%" border=0 background="pic/<%=sp%>3.gif" style="border-collapse: collapse; border-left-width:0; border-top-width:0; border-bottom-width:0">
<TBODY>
<TR>
<TD><b><font color="#FFFFFF">&nbsp;<img border="0" src="pic/fl.gif"> 本论坛共有</font><font color="#00FFFF"> <%=TotalPage%> </font><font color="#FFFFFF">页,<font color="#00FFFF"> <%=count%> </font>个话题，每页有<font color="#00FFFF"> <%=pagesetup%> </font> 张贴子 >> [
<%
ii=PageCount-5
iii=PageCount+5
if ii < 1 then
ii=1
end if
if iii > TotalPage then
iii=TotalPage
end if

if PageCount > 6 then
Response.Write ""&href1&"<font color=yellow>1</font></a> ... "
end if

for i=ii to iii
If i<>PageCount then
if action="new" then
href2="<a href=?topage="& i &"&action=new>"
href3="<a href=?topage="&TotalPage&"&action=new>"
end if
if action="hsz" then
href2="<a href=?topage="& i &"&action=hsz>"
href3="<a href=?topage="&TotalPage&"&action=hsz>"
end if
if action="jh" then
href2="<a href=?topage="& i &"&action=jh>"
href3="<a href=?topage="&TotalPage&"&action=jh>"
end if
if action="mytop" then
href2="<a href=?topage="& i &"&action=mytop>"
href3="<a href=?topage="&TotalPage&"&action=mytop>"
end if
if action="withmetop" then
href2="<a href=?topage="& i &"&action=withmetop>"
href3="<a href=?topage="&TotalPage&"&action=withmetop>"
end if

Response.Write ""&href2&"<font color=yellow>" & i & "</font></a> "
else
Response.Write " <font color=red><b>"&i&"</b></font> "
end if
next

if TotalPage > PageCount+5 then
Response.Write " ... "&href3&"<font color=yellow>"&TotalPage&"</font></a>"
end if
%> ]</font></b></TD>
<form name=form method="POST" action=javascript:Check()>
<TD height=2 align="right"><font color="#FFFFFF">页码：<input style=FONT-SIZE:9pt maxLength="6" size="6" name="topage" value="<%=PageCount%>">
<input style=FONT-SIZE:9pt value="GO!" type="submit"></font></TD></form>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</center>
  </div>
<br><!--#include file="line.asp"--><div align=center>
  <center>
<table bgcolor=<%=c1%> border=0 cellpadding=0 cellspacing=0 style="border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px" bordercolor=<%=c1%> width="94%" height="27" >
<tr>
<td width=100% height="27" style="border: 1px solid <%=c5%>" background="pic/<%=sp%>3.gif">
<font color="#FFFFFF">&nbsp;<b><img border="0" src="pic/tj.gif" align="absmiddle"> 
在线统计：</b>目前论坛总共有 <b><%=lineno%></b> 人在线。其中有 <b><%=usno%></b> 位会员， <b><%=nusno%></b> 位游客。 </font></td>
  </tr>
</table>
  </center>
</div>
<br><!--#include file="down.asp"-->