<%function laiyuan()
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
end if%><!--#include file="up.asp"-->
<%bdlogin(2)
pagenum=request.querystring("pagenum")
re=request.querystring("re")
riqi=now+timeset/24
name=Replace(Request.Form("name"),"'","''")
password=Replace(Request.Form("password"),"'","''")
if lgname="" then%><!--#include file="md5.asp"--><%
password=md5(password)
end if
body=Replace(Request.Form("body"),"'","''")
bodyok=Replace(Request.Form("body")," ","")
face=request.form("face")
if face="" then
face="re"
end if
ls=session("lasttime")
if ls+1/8640>now() then
noyes="�� �� ʧ �� ��"
mes="<meta http-equiv=refresh content=4;url=javascript:history.go(-1)><font color="&c1&">&nbsp;<b>�Բ����㲻�ܳɹ��ط������ӣ�����</b></font><br>������̳Ϊ�˷�ֹ��ˮ��������ͬһ�˷�����ʱ����Ϊ <b>10</b> �룡<br><br>"
else
select case re
case"no"
gonggao=request.form("gonggao")
zhuti=Replace(Request.Form("zhuti"),"'","''")
zhutiok=Replace(Request.Form("zhuti")," ","")
set rs=myconn.execute("SELECT*FROM user where name='"&name&"'and password='"&password&"'")
if rs.eof and rs.bof or zhutiok="" or bodyok="" then
noyes="�� �� ʧ �� ��"
mes="<meta http-equiv=refresh content=4;url=javascript:history.go(-1)><font color="&c1&">&nbsp;<b>�Բ����㲻�ܳɹ��ط������ӣ��������ܴ����������⣺</b></font><br>�� �㲢û����д�������Ҫ���ݣ�<br>�� ����д�����ֻ��������<br>�� ����㻹û��ע��һ���û�����<a href=zhuce.asp><font color=#000080>����ע��</font></a>��<br><br>"
else
select case gonggao
case"4"
if rs("qian")<1000 or rs("jingyan")<200 or rs("meili")<200 then
mes="<meta http-equiv=refresh content=4;url=javascript:history.go(-1)><font color="&c1&">&nbsp;<b>�Բ����㲻�ܳɹ��ط������ӣ�</b></font><br>�� ��Ľ�Ǯ����������������ĳһ������㹻���������"
cangg="no"
else
cangg="yes"
if admin<>"yes" then
myconn.execute("update [user] set qian=qian-1000,meili=meili-200,jingyan=jingyan-200 WHERE name='"&name&"'")
end if
end if
case"0"
cangg="yes"
case"1"
if admin="yes" then
cangg="yes"
else
cangg="no"
end if
end select
isvote="0"

voteyn=request.form("voteyn")
if voteyn=1 then
function checkStr(str)
	if isnull(str) then
		checkStr = ""
		exit function 
	end if
	checkStr=replace(str,"'","''")
end function
votetype=request.form("votetype")
vote=Checkstr(trim(replace(request.Form("vote"),"|","")))
vote=split(vote,chr(13)&chr(10))
nnn1=ubound(vote)
if nnn1>9 then nnn1=9
for i=0 to nnn1
if not (vote(i)="" or vote(i)=" ") then
bodyv=bodyv&"|"&vote(i)
num=num&"|0"
end if
next
if bodyv="" then
isvote="0"
else
isvote="2"
set lastid=myconn.execute("select top 1 id from min order by id desc")
idno=lastid("id")+1
set lastid=nothing
outtime=request.form("outtime")
outt=now+outtime
myconn.execute("insert into vote(id,vote,votenum,type,outtime) values("&idno&",'"&bodyv&"','"&num&"',"&votetype&",'"&outt&"')")
end if
cangg="yes"
end if




if cangg="yes" then
myconn.execute("insert into min(zhuti,name,body,riqi,face,bd,orders,gonggao,type,isvote)VALUES('"&zhuti&"','"&name&"','"&body&"','"&riqi&"','"&face&"',"&bd&",'"&riqi&"','"&gonggao&"',0,'"&isvote&"')")
myconn.execute("update [user] set qian=qian+200,meili=meili+30,jingyan=jingyan+30 WHERE name='"&name&"'")
set seeme=myconn.execute("select top 1 id from min order by id desc")
fid=seeme("id")
set seeme=nothing
noyes="�� �� �� �� ��"
mes="<meta http-equiv=refresh content=3;url=list.asp?bd="&bd&"><font color="&c1&"><b>&nbsp;����ɹ�----����㲻�����������ӣ����� 3 ����Զ���ת�� "&wz&"��</b></font><br>�� <a href=show.asp?id="&fid&"&bd="&bd&">�ص�������������ҳ�棡</a><br>�� <a href=list.asp?bd="&bd&">"&wz&"</a><br>�� <a href=index.asp>"&tl&"</a><br><br>"
end if
end if
set rs=nothing
case"yes"

set lock=myconn.execute("select type from min where id="&id&"")
if lock("type")=4 or lock("type")=5 then
canre="no"
end if
set lock=nothing
if admin="yes" then
canre="yes"
end if
set cjbz=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='"&bd&"'")
if not cjbz.eof then
canre="yes"
end if
set cjbz=nothing
if canre="no" then
noyes="�� �� ʧ ��"
mes="<p style='margin: 15'>�������Ѿ�ɾ��������</p>"
%><!--#include file="mes.asp"-->
<%
response.end
end if


set rs=myconn.execute("SELECT*FROM user where name='"&name&"'and password='"&password&"'")
if rs.eof or rs.bof or bodyok="" then
noyes="�� �� ʧ �� ��"
mes="<meta http-equiv=refresh content=4;url=javascript:history.go(-1)><font color="&c1&">&nbsp;<b>�Բ����㲻�ܳɹ��ػظ����ӣ��������ܴ����������⣺</b></font><br>�� �㲢û����д��Ҫ���ݣ�<br>�� ����д�����ֻ��������<br>�� ����㻹û��ע��һ���û�����<a href=zhuce.asp><font color=#000080>����ע��</font></a>��<br><br>"
else
id=request.querystring("id")
set upid=myconn.execute("select name from min where id="&id&"")
upname=upid("name")
set upid=nothing
myconn.execute("insert into min(name,body,riqi,bd,orders,bid,face,type)VALUES('"&name&"','"&body&"','"&riqi&"',"&bd&",'"&riqi&"',"&id&",'"&face&"',0)")
myconn.execute("update min set orders='"&riqi&"' where id="&id&"")
myconn.execute("update [user] set qian=qian+100,meili=meili+15,jingyan=jingyan+25 WHERE name='"&name&"'")
myconn.execute("update [user] set qian=qian+20,meili=meili+5 WHERE name='"&upname&"'")
noyes="�� �� �� �� ��"
mes="<meta http-equiv=refresh content=3;url=list.asp?bd="&bd&"><font color="&c1&"><b>&nbsp;�ظ��ɹ�----����㲻�����������ӣ����� 3 ����Զ���ת�� "&wz&"��</b></font><br>�� <a href=show.asp?id="&id&"&bd="&bd&"&topage="&pagenum&">�ص������ظ�������ҳ�棡</a><br>�� <a href=list.asp?bd="&bd&">"&wz&"</a><br>�� <a href=index.asp>"&tl&"</a><br><br>"
end if
end select
set ty=myconn.execute("select nyr from bbsinfo")
myconn.execute("update bbsinfo set todaynum=todaynum+1")
if todaynum+1>mosttopic then
myconn.execute("update bbsinfo set mosttopic=todaynum")
end if
session("lasttime")=Now()
end if
%><br><!--#include file="mes.asp"--><br><!--#include file="down.asp"-->