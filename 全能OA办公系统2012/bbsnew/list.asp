<%find=Replace(Request.Form("find"),"'","''")
if find<>"" then
response.cookies("find")=find
end if
search=request.querystring("search")
mybbs=request.querystring("mybbs")
pai=request.querystring("pai")
if pai="" then
pai="orders"
end if%>
<!--#include file="up.asp"-->
<%
response.write"<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style><script language='javascript'>function Check(){var Name=document.form.topage.value;document.location='?bd="&bd&"&topage='+Name+'';}</script>"
bdlogin(1)
if bdtype<>2 or canre="yes" or admin="yes" then 
fbt="<a target='_self' href='say.asp?bd="&bd&"&re=no'><img border=0 src=pic/fabiao.gif align='middle'></a>"
end if
response.write"<div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='94%' height='42'><tr><td width='68%'>"&fbt&"</td><td width='32%' align='right' valign='bottom'><p style='margin-right: 4; margin-bottom: 3'><a href='?bd="&bd&"&mybbs="&lgname&"'>查看我的帖子</a></td></tr></table></center></div>"&_
"<div align='center'><center><table width='94%' height=26 cellpadding=0 style='border-collapse: collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px' bordercolor='"&c1&"' cellspacing='0'><tr><td width='76%' height='21' style='border-left-style: solid; border-left-width: 1; border-top-style: solid; border-top-width: 1'><p style='margin-left: 5'><img align=absmiddle src=pic/gonggao.gif> <b>公告：</b>"
goga=1
set showgg=myconn.execute("select*from min where bd="&bd&" and gonggao=4 and type<>5 order by id desc")
if showgg.eof and showgg.bof then
else
response.write"<marquee onmouseover=this.stop() onmouseout=this.start() scrollAmount=2  width='86%'>"
do while not showgg.eof
response.write"<img src='pic/tl.gif' align='absmiddle'> <a href=show.asp?id="&showgg("id")&"&bd="&bd&"><font color="&c1&">"&kbbs(showgg("name"))&"</font> 说："&kbbs(showgg("zhuti"))&"</a>"
goga=goga+1
if goga>10 then exit do
showgg.movenext
loop
response.write"</marquee>"
end if
showgg.close
set showgg=nothing
response.write"</td><td style='border-right-style:solid; border-right-width:1px; border-top-style:solid; border-top-width:1px' bordercolor="&c1&" align='right' width=$1><p style='margin-left: 4'><p style='margin-right: 3; margin-top: 1; margin-bottom: 1'><img border='0' src='pic/fl.gif'> "
q1=1
set body1=myconn.execute("SELECT*FROM admin where bd='"&bd&"'")
if body1.eof and body1.bof then
response.write"本版诚聘版主！"
else
response.write"<select size=1 name=D1 style='font-size: 9pt'><option>论坛版主</option><option>———</option>"
do while not body1.eof
response.write"<option>"&kbbs(body1("name"))&"</option>"
q1=q1+1
if q1>4 then exit do
body1.movenext
Loop
response.write"</select>"
end if
body1.Close
set body1=nothing
response.write"</td></tr></table></center></div>"
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
if search="" and mybbs="" then
href1="<a href=?topage=1&bd="&bd&">"
sql = "select * from min where ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5 order by gonggao desc,"&pai&" desc"
count=myconn.execute("select count(*)from min where ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5")(0)
end if
if search="yes" and mybbs="" then
href1="<a href=?topage=1&search=yes&bd="&bd&">"
find=request.cookies("find")
sql = "select * from min where (zhuti like '%" &find& "%' or name like '%" &find& "%' or body like '%" &find& "%') and ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5 order by riqi desc"
count=myconn.execute("select count(*) from min where (zhuti like '%" &find& "%' or name like '%" &find& "%' or body like '%" &find& "%') and ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5")(0)
end if
if mybbs<>"" and search="" then
href1="<a href=?topage=1&mybbs="&mybbs&"bd="&bd&">"
sql = "select * from min where name='"&mybbs&"'and ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5 order by gonggao desc,"&pai&" desc"
count=myconn.execute("select count(*) from min where name='"&mybbs&"'and ((bid=0 and bd="&bd&" and gonggao<>1 and gonggao<>4) or gonggao=5) and type<>5")(0)
end if
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
response.write"<div align=center><center><table background=pic/"&sp&"3.gif bgcolor="&c1&" border=1 cellpadding=0 cellspacing=0 style='border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px' bordercolor="&c1&" width='94%' id=AutoNumber5 height='27'><tr><td width=4% align=center><b><font color='#FFFFFF'>状态</font></b></td><td width=42% align=center ><b><a target=_self href=list.asp?bd="&bd&"&pai=zhuti><font color='#FFFFFF'>主 题</font></a><font color='#FFFFFF'> ( 点小图标在新窗口浏览 )</font></b></td>"&_
"<td width=9% align=center ><b><a target=_self href=list.asp?bd="&bd&"&pai=name><font color='#FFFFFF'>作 者</font></a></b></td><td width=9% align=center ><b><font color='#FFFFFF'>回复/</font><a target=_self href=list.asp?bd="&bd&"&pai=hits><font color='#FFFFFF'>人气</font></a></b></td><td width=16% align=center ><b><a target=_self href=list.asp?bd="&bd&"><font color='#FFFFFF'>最后更新时间</font></a></b></td><td width=9% align=center ><b><font color='#FFFFFF'>最后回复</font></b></td></tr></table></center></div>"
i=1
do while not rs.eof
rrzz=kbbs(rs("zhuti"))
set last=myconn.execute("select body,name from min where bid="&rs("id")&" order by id desc")
if last.eof then
zui="-----"
rb=rs("body")
tltltl="帖子内容："&kbbs(rb)&""
else
lb=last("body")
zui="<a href='userinfo.asp?name="&kbbs(last("name"))&"'>"&kbbs(last("name"))&"</a>"
tltltl="最后跟帖："&kbbs(lb)&""
set last=nothing
end if
bno1=myconn.execute("select count(name)from min where bid="&rs("id")&"")(0)
ggtype=rs("gonggao")
bbtype=rs("type")
isvote=rs("isvote")
face=rs("face")
if isvote=2 then face="vote"
if bbtype=4 then face="lock"
if bbtype=1 then face="jing"
if ggtype=3 then face="top"
if ggtype=5 then face="alltop"
response.write"<div align=center><center><table width='94%' border=1 cellpadding=0 cellspacing=0 style='TABLE-LAYOUT: fixed; WORD-BREAK: break-all;border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px' bordercolor="&c1&"  id=AutoNumber5><tr><td width=4% height=23 align=center><a target=_blank href=show.asp?id="&rs("id")&"&bd="&bd&"><img border=0 src=face/"&face&".gif></a></td><td width=42% height=27 align=left onmouseover=javascript:this.bgColor='"&c2&"' onmouseout=javascript:this.bgColor=''>&nbsp;<a target=_self href=show.asp?id="&rs("id")&"&bd="&bd&" title="&LeftTrue(tltltl,25)&">"&LeftTrue(rrzz,44)&"</a>"

remy=10
if bno1/remy>(bno1\remy) then
repage=(bno1\remy)+1
else
repage=bno1\remy
end if
if repage>1 then
response.write" <img align=absmiddle border=0 src=pic/hot.gif><font color="&c1&">[</font><b> "
if repage<=4 then
for nnre=1 to repage
response.write"<a target=_self href=show.asp?id="&rs("id")&"&bd="&bd&"&topage="&nnre&"><font color="&c1&">"&nnre&"</font></a> "
next
else
for nnnre=1 to 3
response.write"<a target=_self href=show.asp?id="&rs("id")&"&bd="&bd&"&topage="&nnnre&"><font color="&c1&">"&nnnre&"</font></a> "
next
response.write"... <a target=_self href=show.asp?id="&rs("id")&"&bd="&bd&"&topage="&repage&"><font color="&c1&">"&repage&"</font></a> "
end if
response.write" </b><font color="&c1&">]</font>"
end if


response.write"</td><td width=9% height=23 align=center><a href='userinfo.asp?name="&kbbs(rs("name"))&"'>"&kbbs(rs("name"))&"</a></td><td width=9% height=23 align=center>"&bno1&"/"&rs("hits")&"</td><td width=16% height=23 align=center>"&rs("orders")&"</td><td width=9% height=23 align=center>"&zui&"</td></tr></table></center></div>"
i=i+1
if i>pagesetup then exit do
rs.movenext
loop
rs.Close
response.write"<div align='center'><center><TABLE bgcolor="&c1&" borderColor="&c1&" cellSpacing=0 cellPadding=0 width='94%' border=1 style='border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px'><TBODY><TR height=25><TD height=2><TABLE cellSpacing=0 cellPadding=3 width='100%' border=0 background='pic/"&sp&"3.gif' style='border-collapse: collapse; border-left-width:0; border-top-width:0; border-bottom-width:0'><TBODY><TR><TD><b><font color='#FFFFFF'><img border='0' src='pic/fl.gif'> 本论坛共有</font><font color='#00FFFF'> "&TotalPage&" </font><font color='#FFFFFF'>页,<font color='#00FFFF'> "&count&" </font>个话题，每页有<font color='#00FFFF'> "&pagesetup&" </font> 张贴子 >> ["
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
if search="" and mybbs="" then
href2="<a href=?topage="& i &"&bd="&bd&">"
href3="<a href=?topage="&TotalPage&"&bd="&bd&">"
end if
if search="yes" and mybbs="" then
href2="<a href=?topage="& i &"&search=yes&bd="&bd&">"
href3="<a href=?topage="&TotalPage&"&search=yes&bd="&bd&">"
end if
if mybbs<>"" and search="" then
href2="<a href=?topage="& i &"&mybbs="&mybbs&"&bd="&bd&">"
href3="<a href=?topage="&TotalPage&"&mybbs="&mybbs&"&bd="&bd&">"
end if
Response.Write ""&href2&"<font color=yellow>" & i & "</font></a> "
else
Response.Write " <font color=red><b>"&i&"</b></font> "
end if
next
if TotalPage > PageCount+5 then
Response.Write " ... "&href3&"<font color=yellow>"&TotalPage&"</font></a>"
end if
response.write" ]</font></b></TD><form name=form method='POST' action=javascript:Check()><TD height=2 align='right'><font color='#FFFFFF'>页码：<input style=FONT-SIZE:9pt maxLength='6' size='6' name='topage' value='"&PageCount&"'><input style=FONT-SIZE:9pt value='GO!' type='submit'></font></TD></form></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></center></div>"
response.write"<div align=center><center><table border=0 cellspacing=1 style='border-collapse: collapse' bordercolor=#111111 width=94% ><tr><td width=100% align=right height=35><select onchange=if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;} style='font-size: 9pt'><option selected>跳转论坛至...</option>"
set shsh=myconn.execute("select bdname,bn from bdinfo where key='0'")
do while not shsh.eof
response.write"<option>+"&shsh("bdname")&"</option>"
set fen=myconn.execute("select bdname,bn from bdinfo where key='"&shsh("bn")&"'")
do while not fen.eof
response.write"<option value='list.asp?bd="&fen("bn")&"'>-"&fen("bdname")&"</option>"
fen.movenext
loop
set fen=nothing
shsh.movenext
loop
set shsh=nothing
response.write"</select></td></tr></table></center></div>"
%>
<!--#include file="line.asp"-->
<%
response.write"<div align=center><center><table bgcolor="&c1&" border=1 cellpadding=0 cellspacing=0 style='border-collapse:collapse; border-left-width:1px; border-right-width:1px; border-top-width:1px' bordercolor="&c1&" width='94%' height='27' ><tr><td width=100% height='27' background='pic/"&sp&"3.gif'><font color='#FFFFFF'>&nbsp;<b><img border='0' src='pic/tj.gif' align='absmiddle'> 在线统计：</b>目前论坛总共有 <b>"&lineno&"</b> 人在线。其中有 <b>"&usno&"</b> 位会员， <b>"&nusno&"</b> 位游客。 </font></td></tr></table></center></div>"&_
"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>论坛帖子图例</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% height=40>&nbsp;<img border=0 src=face/alltop.gif> 总置顶帖子 &nbsp;&nbsp;&nbsp;  <img border=0 src=face/top.gif> 置顶帖子 &nbsp;&nbsp;&nbsp;  <img border=0 src=face/jing.gif> 精华帖子 &nbsp;&nbsp;&nbsp;  <img border=0 src=face/lock.gif> 锁定帖子 &nbsp;&nbsp;&nbsp;  <img border=0 src=face/vote.gif> 投票帖子 &nbsp;&nbsp;&nbsp;  <img border=0 src=pic/hot.gif> 热门帖子</td></tr></table></center></div>"&_
"<br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>论坛快速搜索</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><form method='POST' action='list.asp?bd="&bd&"&search=yes'><td width=100% height=30>&nbsp;查询字符：<INPUT name=find size=20> <input type=submit value=' 搜 索 ' name=B1> <input type=reset value=' 重 置 ' name=B2></td></form></tr></table></center></div>"
%><br><!--#include file="down.asp"-->