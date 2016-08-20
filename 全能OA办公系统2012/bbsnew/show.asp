<!--#include file="up.asp"--><!--#include file="ubbcode.asp"-->
<%zai=request.querystring("topage")
if zai="" then
zai=1
end if
response.write"<script language='javascript'>function Check(){var Name=document.form.topage.value;document.location='?bd="&bd&"&id="&id&"&topage='+Name+'';}</script>"
bdlogin(1)
myconn.execute("update min set hits=hits+1 WHERE id="&id&"")
count=myconn.execute("select count(riqi) from min where bid=" &id& "")(0)
pagesetup=10
renum=count+1
If renum/pagesetup > (renum\pagesetup) then
yenum=(renum\pagesetup)+1
else yenum=(renum\pagesetup)
End If
del="<a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=del&re=no'>删除</a>"
if bbtype=5 then 
del="<a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=hy&re=no'>还原</a><img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=dely&re=no'>永久删除</a>"
if admin<>"yes" then
noyes="操 作 失 败"
mes="<p style='margin: 14'>该帖子已经删除或被锁定</p>"
%>
<!--#include file="mes.asp"-->
<%response.end
end if
end if
if bbtype=4 then
testl="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=unlock'>解锁</a>"
else
testl="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=lock'>加锁</a>"
end if
fabiao="yes"
backbt="yes"
if bbtype=4 then
fabiao="yes"
backbt="no"
end if
if bdtype=2 then 
fabiao="no"
backbt="no"
end if
if canre="yes" or admin="yes" then
fabiao="yes"
backbt="yes"
end if
if fabiao="yes" then fbt="<a target='_self' href='say.asp?bd="&bd&"&re=no'><img border=0 src=pic/fabiao.gif align='middle'></a>"
if backbt="yes" then bbt=" <a href='say.asp?bd="&bd&"&id="&id&"&re=yes&pagenum="&yenum&"'><img border='0' src='pic/back.gif' align='absmiddle'></a>"
response.write"<div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='94%' height='42'><tr><td width='70%'>"&fbt&""&bbt&"</td><td width='30%' align='right' valign='bottom'><img border='0' src='pic/tl.gif'> 你是本帖的第 <b>"&whits+1&"</b> 位阅读者</td></tr></table></center></div>"
if isvote=2 then
set hhh=myconn.execute("select*from vote where id="&id&"")
outtime=hhh("outtime")
votetype=hhh("type")
ddd=hhh("vote")
nno=hhh("votenum")
ddd=split(ddd,"|")
nno=split(nno,"|")
nnn=ubound(ddd)
allvote=0
for j=1 to nnn
allvote=allvote+nno(j)
next
if allvote=0 then allvote=1
response.write"<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='WORD-BREAK: break-all; border-collapse:collapse' bordercolor="&c1&" width=94% ><form method=POST action=bbselse.asp?menu=vote&id="&id&"&type="&votetype&"><tr><td width=100% background=pic/"&sp&"3.gif height=25 colspan=2>&nbsp;<img border=0 src=face/vote.gif> <font color=#FFFFFF><b>投票选项</b></font></td></tr>"
for i=1 to nnn 
width=nno(i)/allvote*85
if votetype=1 then
form1="<input type=radio value="&i&" name=xuan>"
else
form1="<input type=checkbox name='xuan_"&i&"' value="&i&">"
end if
response.write"<tr><td width=58% height=25 >&nbsp;·"&i&" "&form1&""&kbbs(ddd(i))&"</td><td width=42% >&nbsp;<img border=0 height=8 width=2 src=pic/line.gif><img border=0 height=8 width="&width&"% src=pic/line.gif> <b>"&nno(i)&"</b> 票 </td></tr>"
next
canvote="yes"
if lgname="" then
info1="你还没有登陆，不能进行投票。"
canvote="no"
else
set had=myconn.execute("select user from voted where user='"&lgname&"' and id="&id&"")
if not had.eof then
info1="你已经投过票了,不能再投票了。"
canvote="no"
end if
set had=nothing
end if
if now>outtime then
info1="该投票已经过期，不能进行投票。"
canvote="no"
end if
if canvote="yes" then
buto="<input type=submit value=' 投 票 ' name=B1>"
end if
response.write"<tr><td width=100% height=25 colspan=2>&nbsp;<img border=0 src=pic/tl.gif> "&info1&""&buto&" [ 截止时间："&outtime&" ] </td></tr></form></table></center></div><br>"
end if

set huati=myconn.execute("select*from user where name='"&wname&"'")
response.write"<div align='center'><center><table bgcolor="&c1&" border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse; ' width='94%'><tr><td width='100%' background='pic/"&sp&"3.gif' height='26'><font color='#FFFFFF'>&nbsp;<b>『 帖子主题 』："&LeftTrue(wzhuti,50)&"</b></font></td></tr></table></center></div>"&_
"<div align='center'><center><table border='1' cellpadding='0' cellspacing='0' style='TABLE-LAYOUT: fixed; WORD-BREAK: break-all;border-collapse:collapse' bordercolor='"&c1&"' width='94%' height='140'><tr><td width='20%' valign='top'><div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111'><tr><td width='100%'>"&_
"<div align='center'><center><br><table width=$1 style='filter:glow(color="&c1&", strength=1); border-collapse:collapse' bordercolor='#111111' cellpadding='0' cellspacing='0'><font color=black>"&kbbs(wname)&"</font></table><br><img border='0' src='"&kbbs(huati("toupic"))&"' width="&huati("ku")&" height="&huati("ch")&"></center></div><br>金钱："
q1=huati("qian")
response.write""&q1&"<br>经验："
j1=huati("jingyan")
response.write""&j1&"<br>魅力："
m1=huati("meili")
response.write""&m1&"<br>帖数："&myconn.execute("select count(name)from min where name='"&wname&"'")(0)&"<br>"
myname=wname
sqltype="my"
response.write"等级："%><!--#include file="upji.asp"-->
<%
response.write"<b>"&dj&"</b><br>『 "&dd&" 』<br></td></tr></table></center></div></td><td width='80%' valign='top'> "
kw=kbbs(wname)
response.write"<table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='97%' height='25'><tr><td width='81%' valign='bottom'>&nbsp;<a href='userinfo.asp?name="&kw&"' title='查看"&kw&"的资料'><img src='pic/info.gif' border='0'> 信 息</a>&nbsp; <a href='mailcon.asp?menu=write&towho="&kw&"' title='给"&kw&"留言'><img src='pic/newmail.gif' border='0'> 留 言</a>&nbsp; <a title='"&kw&"的OICQ号码："&huati("qq")&"'><SPAN style='CURSOR: hand' ><img border='0' src='pic/oicq.gif'> OICQ</span></a>&nbsp; <a href='mailto:"&huati("email")&"' title='发邮件给"&kw&"'><img border='0' src='pic/mail.gif'> 邮 箱</a>&nbsp; <a href='"&kbbs(huati("home"))&"' title='访问"&kw&"的主页'><img border='0' src='pic/home.gif'> 主 页</a>&nbsp; <a href=edit.asp?id="&id&"&ed=1&reid="&id&"&bd="&bd&"><img border='0' src='pic/bmp.gif'> 编 辑</a> &nbsp;<a href=say.asp?bd="&bd&"&id="&id&"&re=yes&quoteid="&id&"><img src='pic/xie.gif' border='0'> 引 用</a></td>"
gxqm=huati("gxqm")
set huati=nothing
response.write"<td width='19%' valign='bottom' align='right'>楼 &nbsp; 顶</td></tr></table><hr color='"&c1&"' width='98%' size='1'><blockquote><p style='line-height:15pt'><img src='face/"&wface&".gif'> <b>"&wzhuti&"</b><br>"
response.write""&ubb(wbody)&"<p></p>"
if gxqm<>"" then
response.write"<br><div align=right>————————————————————<br>"&ubb(gxqm)&"</div>"
end if
response.write"</blockquote></td></tr><tr><td height='26' align='center'>状态："
set onoff=myconn.execute("select name from online where name='"&wname&"'")
if onoff.eof then
ooo="offline"
else
ooo="online"
end if
set onoff=nothing
response.write"<img border=0 align=absmiddle src=pic/"&ooo&".gif></td><td><div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='#111111' width='98%'><tr><td width='30%'><img border='0' src='pic/xie.gif'> "&wriqi&"</td><td width='70%' align='right'>"
if ggtype=3 then 
testg1="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=nottop'>解除置顶</a>"
else
testg1="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=top'>置顶</a>"
end if
if ggtype=5 then
testg2="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=nottop'>解除总置顶</a>"
else
testg2="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=alltop'>总置顶</a>"
end if
if bbtype=1 then 
testj="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=notjh'>取消精华</a>"
else
testj="<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=jh'>精华</a>"
end if
response.write"<img border=0 src=pic/tl.gif> 帖子管理："&testg2&""&testg1&""&testj&""&testl&"<img border=0 src=pic/"&sp&"2.gif>"&del&"<img border=0 src=pic/"&sp&"2.gif><a href='bbsgl.asp?bd="&bd&"&id="&id&"&action=move'>移动</a><img border=0 src=pic/"&sp&"2.gif></td></tr></table></center></div></td></tr></table></center></div>"
set huati=nothing
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select*from min where bid=" &id& " and type<>5 order by id"
on error resume next
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
ai=1
do while not rs.eof
rnn=rs("name")
set back=myconn.execute("select*from user where name='"&rnn&"'")
lou=pagesetup-ai
if lou/2>lou\2 then
bgcl=""&c2&""
else
bgcl="white"
end if
response.write"<div align='center'><center><table width='94%' bordercolor='#111111' style='border-collapse: collapse; border-width: 0' cellpadding='0' cellspacing='0'><tr><td width=100% height='2'></td></tr></table></center></div>"&_
"<div align='center'><center><table border='1' cellpadding='0' cellspacing='0' style='TABLE-LAYOUT: fixed; WORD-BREAK: break-all; border-collapse:collapse;background-color:"&bgcl&"' bordercolor='"&c1&"' width='94%' height='150'><tr><td width='20%' valign='top'><div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse;background-color:"&bgcl&"' bordercolor='#111111'><tr><td>"
response.write"<div align='center'><center><br><table width=$1 style='filter:glow(color="&c1&", strength=1); border-collapse:collapse' bordercolor='#111111' cellpadding='0' cellspacing='0'><font color=black>"&kbbs(rnn)&"</font></table><br><img border='0' src='"&kbbs(back("toupic"))&"' width="&back("ku")&" height="&back("ch")&"></center></div><br>金钱："
q1=back("qian")
response.write""&q1&"<br>经验："
j1=back("jingyan")
response.write""&j1&"<br>魅力："
m1=back("meili")
response.write""&m1&"<br>帖数："&myconn.execute("select count(name)from min where name='"&rnn&"'")(0)&"<br>"
myname=rnn
sqltype="my"
response.write"等级："
%><!--#include file="upji.asp"-->
<%
response.write"<b>"&dj&"</b><br>『 "&dd&" 』<br></td></tr></table></center></div>"&_
"</td><td width='80%' valign='top'><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse;background-color:"&bgcl&"' bordercolor='#111111' width='97%' height='25'><tr><td width='81%' valign='bottom'>&nbsp;<a href='userinfo.asp?name="&rnn&"' title='查看"&rnn&"的资料'><img src='pic/info.gif' border='0'> 信 息</a>&nbsp; <a href='mailcon.asp?menu=write&towho="&rnn&"' title='给"&rnn&"留言'><img src='pic/newmail.gif' border='0'> 留 言</a>&nbsp; <a title='"&rnn&"的OICQ号码："&kbbs(back("qq"))&"'><SPAN style='CURSOR: hand' ><img border='0' src='pic/oicq.gif'> OICQ</span></a>&nbsp; <a href='mailto:"&back("email")&"' title='发邮件给"&rnn&"'><img border='0' src='pic/mail.gif'> 邮 箱</a>&nbsp; <a href='"&kbbs(back("home"))&"' title='访问"&rnn&"的主页'><img border='0' src='pic/home.gif'> 主 页</a>&nbsp; <a href='edit.asp?id="&rs("id")&"&ed=2&reid="&id&"&bd="&bd&"'><img border='0' src='pic/bmp.gif'> 编 辑</a> &nbsp;<a href=say.asp?bd="&bd&"&id="&id&"&re=yes&quoteid="&rs("id")&"><img src='pic/xie.gif' border='0'> 引 用</a></td>"
gxqm=back("gxqm")
set back=nothing
response.write"</td><td width='19%' valign='bottom' align='right'>"
if lou>0 then
response.write"<font color='"&c1&"'>"&CHR(64+zai)&"</font> 栋<font color='"&c1&"'> "&lou&"</font> 楼"
else
response.write"<font color='"&c1&"'>"&CHR(64+zai)&"</font> 栋 楼 下"
end if
response.write"</td></tr></table><hr color='"&c1&"' width='98%' size='1'><blockquote><p style='line-height:15pt'><img src='face/"&rs("face")&".gif'>&nbsp;<br>"&ubb(rs("body"))&"<p></p>"
if gxqm<>"" then
response.write"<br><div align=right>————————————————————<br>"&ubb(gxqm)&"</div>"
end if
response.write"</blockquote></td></tr><tr><td height='26' align='center'>状态："
set onoff=myconn.execute("select name from online where name='"&rnn&"'")
if onoff.eof then
ooo="offline"
else
ooo="online"
end if
set onoff=nothing
response.write"<img border=0 align=absmiddle src=pic/"&ooo&".gif></td><td><div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse;background-color:"&bgcl&"' bordercolor='#111111' width='98%'><tr><td width='50%'><img border='0' src='pic/xie.gif'> "&rs("riqi")&"</td><td width='50%' align='right'>"&_
"<a href='bbsgl.asp?bd="&bd&"&id="&rs("id")&"&action=del&re=yes'><img border=0 src=pic/"&sp&"2.gif>删除<img border=0 src=pic/"&sp&"2.gif></a></td></tr></table></center></div></td></tr></table></center></div>"
ai=ai+1
if ai>pagesetup then exit do
rs.movenext
loop
rs.Close
response.write"<div align='center'><center><TABLE bgcolor="&c1&" borderColor="&c1&" cellSpacing=0 cellPadding=0 width='94%' border=1 style='border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px'><TBODY><TR height=25><TD height=2><TABLE cellSpacing=0 cellPadding=3 width='100%' border=0 background='pic/"&sp&"3.gif' style='border-collapse: collapse; border-left-width:0; border-top-width:0; border-bottom-width:0'><TBODY><TR><TD><b><font color='#FFFFFF'><img border='0' src='pic/fl.gif'> 共有</font><font color='#00FFFF'> "&TotalPage&" </font><font color='#FFFFFF'>页,<font color='#00FFFF'> "&count&" </font><font color='#FFFFFF'>张回复帖,每页有<font color='#00FFFF'> "&pagesetup&" </font> 张贴子 >> ["
ii=PageCount-5
iii=PageCount+5
if ii < 1 then
ii=1
end if
if iii > TotalPage then
iii=TotalPage
end if
if PageCount > 6 then
Response.Write "<a href=?topage=1&bd="&bd&"&id="&id&"><font color=yellow>1</font></a> ... "
end if

for i=ii to iii
If i<>PageCount then
Response.Write " <a href=?topage="& i &"&bd="&bd&"&id="&id&"><font color=yellow>" & i & "</font></a> "
else
Response.Write " <font color=red><b>"&i&"</b></font> "
end if
next
if TotalPage > PageCount+5 then
Response.Write " ... <a href=?topage="&TotalPage&"&bd="&bd&"&id="&id&"><font color=yellow>"&TotalPage&"</font></a>"
end if
response.write" ]</font></b></TD><form name=form method='POST' action=javascript:Check()><TD height=2 align='right'><font color='#FFFFFF'>页码：<input style=FONT-SIZE:9pt maxLength='6' size='6' name='topage' value='"&PageCount&"'><input style=FONT-SIZE:9pt value='GO!' type='submit'></font></TD></form></TR></TBODY></TABLE></TD></TR></TBODY></TABLE></center></div>"
response.write"<div align=center><center><table border=0 cellspacing=1 style='border-collapse: collapse' bordercolor=#111111 width=94% ><tr><td width=100% align=right valign=bottom height=30><select onchange=if(this.options[this.selectedIndex].value!=''){location=this.options[this.selectedIndex].value;} style='font-size: 9pt'><option selected>跳转论坛至...</option>"
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
response.write"</select></td></tr></table></center></div><SCRIPT>var i=0;function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('帖子正在发出，请耐心等待！');return false;}this.document.re.submit();}}</SCRIPT>"
if backbt="yes" then
response.write"<form method='POST' name=re action='save.asp?bd="&bd&"&id="&id&"&re=yes&pagenum="&yenum&"'><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>快 速 回 复</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-style:solid; border-width:0; border-collapse: collapse' bordercolor="&c1&" width='94%'><tr><td width='100%'><div align='center'><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='98%'><tr><td width='17%' height='28'><p style='margin: 4'>用户信息：</td><td width='83%'>用户名：<input type='text' name='name' size='19' value='"&lgname&"'> <a href=zhuce.asp>没有注册？</a><img border='0' src='pic/"&sp&"2.gif'>密码：<input type='password' name='password' size='20' value='"&lgpwd&"'> <a href=getpwd.asp>忘记密码？</a></td></tr><tr><td width='17%'><p style='margin: 4; line-height:150%'><b>帖子内容：</b><BR>·HTML标签： 不可用<br>·UBB标签： 可用<br>·贴图标签： 可用<br>·Flash标签：不可用<br>·表情字符转换：不可用<br>·最多15KB</p></td><td width='83%' valign='top'><textarea onkeydown=presskey(); rows=6 name='body' cols='92' Title='按 Ctrl+Enter 直接发送'></textarea><p><input type='submit' value='OK_好了！回复帖子' name='submit1' tabindex='4'> <input type='reset' value='NO_不行！我要重写' name='B2'> [按 Ctrl+Enter 直接发送]</p></td></tr></table></center></div></td></tr></table></center></div></form>"
end if
response.write"<br>"%><!--#include file="down.asp"-->