<!--#include file="conn.asp"--><!--#include file="fun.asp"-->
<%
function opendb(DBPath,sessionname,dbsort)
dim conn
Set conn=Server.CreateObject("ADODB.Connection")
DBPath1=server.mappath("../db/sdoa.asa")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
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
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='" & oabusyusername&"'"
rs.open sql,conn,1
cook_allow_control_all_user=rs("allow_control_all_user")     
conn.close
set conn=nothing
set rs=nothing
if cook_allow_control_all_user="no" then
response.write("<font color=red size=""+1"">�Բ�����û�����Ȩ�ޣ�</font>")
	response.end
	end if
%>
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
noyes="�� ½ ʧ �� ��"
mes="�㲻�ܽ����̨����<br>�������ڵ�½��̳���û��� "&lgname&" ���ǹ���Ա����"%>
<!--#include file="mes.asp"-->
<%response.end
else
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='240' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
function bdlist(sename,n)
response.write"<select size=1 name="&sename&" style='font-size: 9pt; '>"
if n=1 then
response.write"<option value=all selected>������̳</option>"
end if
set bf=myconn.execute("select*from bdinfo where key<>'0'")
do while not bf.eof
response.write"<option value="&bf("bn")&">"&bf("bdname")&"</option>"
bf.movenext
loop
bf.close
set bf=nothing
response.write"</select>"
end function

%>
<body topmargin="0" leftmargin="0"><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>

<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
  <tr>
    <td width="100%" height="28" background="pic/<%=sp%>3.gif" align="center">
    <b><font color="#FFFFFF">��̳��̨����ϵͳ</font></b></td>
  </tr>
  </table><br>
<%
action=request.querystring("action")
action="chadmin"
select case action%>
<%case"addfl"
noyes="�� �� �� ��"
mes="<P style='MARGIN: 5px'>������ţ�<input type=text name=bn size=20><font color=#FF0000>��</font>ֻ���� <b>0</b> ���������</p><P style='MARGIN: 5px'>�������ƣ�<input type=text name=bdname size=20><font color=#FF0000>��</font>��������ƣ�������</p><P style='MARGIN: 5px'><input type=submit value=' �� �� ' name=B1> <input type=reset value=' �� �� ' name=B2></p>"
%>
<%
bn=request.form("bn")
bdname=Replace(Request.Form("bdname"),"'","''")
if bn="" or bn="0" or bdname="" or not isnumeric(bn) then
%>
<%else
set flyn=myconn.execute("select bn,bdname from bdinfo where bn="&bn&" and key='0'")
if not flyn.eof then
mes="<br>��������� <b>"&bn&"</b> �Ѿ����ڣ�����ѡ���ķ������<br><br>"
else
id=bn&"0"
myconn.execute("insert into bdinfo(id,bn,bdname,key)values("&id&","&bn&",'"&bdname&"','0')")
mes="<br>�������̳����ɹ���<br><br>"
end if
set flyn=nothing
end if
%><form method=POST>
<!--#include file="mes.asp"--></form>
<%case"bzgl"
bz=request.querystring("bz")
if bz="add" then
bt="����°���"
go="add"
put="&nbsp;��&nbsp;��&nbsp;"
elseif bz="del" then
bt="ɾ������"
go="del"
put="&nbsp;ɾ&nbsp;��&nbsp; "
end if
%>
<form action="admint-gl.asp?menu=bzgl&go=<%=go%>" method="POST">
<%=t1&bt&t2&d1%>
<P style='MARGIN: 5px'>�������ƣ�<input type="text" name="name" size="20"></p><P style='MARGIN: 5px'>������̳��<%=bdlist("bd",0)%>
</p><P style='MARGIN: 5px'><input type="submit" value=<%=put%> name="B1">&nbsp;
<input type="reset" value=" �� �� " name="B2"></p>

<%=d2%>
</form>
<%case"chadmin"%>
<%=t1%>���й���Ա<%=t2&d1%><P style="MARGIN: 5px">���й���Ա���ƣ�<br><%
set sho=myconn.execute("select name from admin where bd='70767766'")
do while not sho.eof
%><%=sho("name")%>&nbsp;&nbsp;&nbsp;<%sho.movenext
loop
set sho=nothing%>
<P style="MARGIN: 5px">�����û����ƣ�<br><%
set sho=myconn.execute("select name from user")
do while not sho.eof
%><%=sho("name")%>&nbsp;&nbsp;&nbsp;<%sho.movenext
loop
set sho=nothing%><%=d2%>

<%=d2%>
<form action="admint-gl.asp?menu=addadmin" method="POST">
<%=t1%>��ӹ���Ա<%=t2&d1%><P style="MARGIN: 5px">��ӹ���Ա���ƣ�( �����Ʊ����Ѿ�����̳��ע�� ) <input type="text" name="adminname" size="20"> 
<input type="submit" value=" �� �� " name="B1">
<input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<form action="admint-gl.asp?menu=deladmin" method="POST">
<%=t1%>ɾ������Ա<%=t2&d1%><P style="MARGIN: 5px">ɾ������Ա���ƣ�<input type="text" name="adminname" size="20"> 
<input type="submit" value=" ɾ �� " name="B1">
<input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<form action="admint-gl.asp?menu=deluser" method="POST">
<%=t1%>ɾ���û�<%=t2&d1%><P style="MARGIN: 5px">�û�����<input type="text" name="delname" size="20"> 
<input type="submit" value=" ɾ �� " name="B1">
<input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<form action="admint-gl.asp?menu=chpwd" method="POST">
<%=t1%>�����û�����<%=t2&d1%><P style="MARGIN: 5px">�û�����<input type="text" name="chaname" size="20"> �����룺<input type="text" name="chapwd" size="20"> 

<input type="submit" value=" �� �� " name="B1">
        <input type="reset" value=" �� �� " name="B2"></p><%=d2%> 
</form>

</form>
<%case"deluser"%>
<form action="admint-gl.asp?menu=deluser" method="POST">
<%=t1%>ɾ���û�<%=t2&d1%><P style="MARGIN: 5px">�û�����<input type="text" name="delname" size="20"> 
<input type="submit" value=" ɾ �� " name="B1">
<input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<%case"addpassuser"
bn=request.querystring("bn")
set showps=myconn.execute("select pass from bdinfo where bn="&bn&" and key<>'0'")
%>
<form method="POST" action="admint-gl.asp?menu=addpassuser&bn=<%=bn%>">
<%=t1%>�޸���̳��֤�û�<%=t2&d1%>
<P style="MARGIN: 5px">�������Ѿ�ͨ����֤���û���Ҫ����������д�����û�֮���á�,��������</p><P style="MARGIN: 5px"><b>
<font color="#FF0000">ע�⣺</font></b>��д������һ������ʹ�ûس�</p><P style="MARGIN: 5px">
<textarea name="user" cols=90 rows="15"><%=showps("pass")%></textarea></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="b1"> <input type="reset" value=" �� �� " name="b2">
</p><%set showps=nothing%>
</form>
<%case"bbs"
set bbs=myconn.execute("select*from bbsinfo")
%>
<form method="POST" action="admint-gl.asp?menu=bbs">
<%=t1%>��̳��������<%=t2&d1%><hr>
<P style="MARGIN: 5px"><b>��̳�������ã�</b></p>
<P style="MARGIN: 5px">��̳����:<input type="text" name="tl" size="37" value="<%=bbs("tl")%>"><font color="#FF0000">��</font>(�����̳������)</p>
<P style="MARGIN: 5px">��̳�������:������ʹ��html��</p>
<P style="MARGIN: 5px"><textarea rows="5" name="topinfo" cols="69"><%=bbs("topinfo")%></textarea></p><hr>
<P style="MARGIN: 5px"><b>��̳�ϴ����ã�</b></p>
<P style="MARGIN: 5px">ÿ���ϴ�������<input type="text" name="upnum" size="5" value="<%=bbs("upnum")%>"> ��<font color="#FF0000">��</font></p>
<P style="MARGIN: 5px">����ϴ���С��<input type="text" name="upsize" size="5" value="<%=bbs("upsize")%>"> KB<font color="#FF0000">��</font></p>
<hr><P style="MARGIN: 5px"><b>��̳Ĭ����ʽ���ã�</b></p>
<P style="MARGIN: 5px">Ĭ�ϱ߿���ɫ��<input type="text" name="c1" size="20" value="<%=bbs("c1")%>"><font color="#FF0000">��</font>(���߿����ɫ)</p>
<P style="MARGIN: 5px">Ĭ����̳��ɫ��<input type="text" name="c2" size="20" value="<%=bbs("c2")%>"><font color="#FF0000">��</font>(һЩ���ĵ�ɫ)</p>
<P style="MARGIN: 5px">Ĭ����̳��ʽ��<input type="text" name="style" size="10" value="<%=bbs("style")%>"><font color="#FF0000">��</font>(���Բ���������д)</p>
<P style="MARGIN: 5px">��̳��ʽ���գ�����ɫ��a ��ɫ��b ��ɫ��c ��ɫ��d ��ɫ��e��</p>
<hr>
<P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1"> <input type="reset" value=" �� �� " name="B2"></p>
<%=d2%>
</form>
<%set bbs=nothing%>
<%case"chbbsinfo"
id=request.querystring("id")
set bbsinfo=myconn.execute("select*from bdinfo where id="&id&"")
%><form action="admint-gl.asp?menu=updatebbs&id=<%=id%>" method="POST">
<%=t1%>�޸���̳��Ϣ<%=t2&d1%>
<P style='MARGIN: 5px'>��̳��ţ�<input type="text" name="bn" size="25" value="<%=bbsinfo("bn")%>"><font color="#FF0000">��</font><P style='MARGIN: 5px'>��̳���ƣ�<input type="text" name="bdname" size="25" value="<%=bbsinfo("bdname")%>"><font color="#FF0000">��</font>������</p>
<P style='MARGIN: 5px'>��־ͼƬ��<input type="text" name="picurl" size="49" value="<%=bbsinfo("picurl")%>">������ʾ����̳���ұߣ����Բ��</p>
<P style='MARGIN: 5px'>��̳���ܣ�<br><textarea rows="4" name="bdinfo" cols="58"><%=bbsinfo("bdinfo")%><%
shuyu=cint(bbsinfo("key"))
btype=bbsinfo("type")
set bbsinfo=nothing%></textarea><font color="#FF0000">��</font></p>
<P style='MARGIN: 5px'>���ڷ��ࣺ<select size="1" name="key" style="font-size: 9pt">
<%set bf=myconn.execute("select*from bdinfo where key='0' order by bn")
do while not bf.eof%><%if shuyu=bf("bn") then%><option value="<%=bf("bn")%>" selected><%=bf("bdname")%></option><%else%>
<option value="<%=bf("bn")%>"><%=bf("bdname")%></option><%end if%>
<%
bf.movenext
Loop
bf.Close
set bf=nothing
%>
</select><font color="#FF0000">��</font>��ѡ�����̳������һ�ַ���</p><br><P style='MARGIN: 4px'>
��̳���ͣ�(������������ѡ��һ��)<font color="#FF0000">��</font></p><P style='MARGIN: 4px'><input type="radio" value="0" name="bbstype" <%if btype=0 then%>checked<%end if%>>��ͨ��̳
��ע���û����οͿ������ɵĽ����������̳�����Ƽ�����
<P style='MARGIN: 4px'><input type="radio" value="1" name="bbstype" <%if btype=1 then%>checked<%end if%>>��Ա��̳��ֻ��ע���û��������ɵĽ����������̳��
<P style='MARGIN: 4px'><input type="radio" value="2" name="bbstype" <%if btype=2 then%>checked<%end if%>>ֻ����̳����ͨ��Ա���ܷ����ȣ�ֻ�������
<P style='MARGIN: 4px'><input type="radio" value="3" name="bbstype" <%if btype=3 then%>checked<%end if%>>��֤��̳
��ֻ�а�����֤��ע���û����ܽ����������̳��</p><br>
<P style='MARGIN: 4px'><input type="submit" value=" �� �� " name="B1"> <input type="reset" value=" �� �� " name="B2"></p><br><%=d2%>
</form>
<%case"delall"
bn=request.querystring("bn")
myconn.execute("delete*from min where bd="&bn&"")
%><%=t1%>ɾ �� �� ��<%=t2&d1%><P style="MARGIN: 5px"><br>���Ѿ��ɹ���ɾ���˸ð�����������ӡ�</p><br><%=d2%>

<%case"updateuser"%>
<form action="admint-gl.asp?menu=updateuser" method="POST">
<%=t1%>�����û�����<%=t2&d1%>
<P style="MARGIN: 5px">�û�����<input type="text" name="chname" size="19"></p>
<P style="MARGIN: 5px">��&nbsp; Ǯ��<input type="text" name="chqian" size="10"> 
  ��&nbsp; ����<input type="text" name="chmeili" size="10">
  ��&nbsp; �飺<input type="text" name="chjingyan" size="10"></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<%case"chpwd"%>
<form action="admint-gl.asp?menu=chpwd" method="POST">
<%=t1%>�����û�����<%=t2&d1%><P style="MARGIN: 5px">�û�����<input type="text" name="chaname" size="20"> �����룺<input type="text" name="chapwd" size="20"> 

<input type="submit" value=" �� �� " name="B1">
        <input type="reset" value=" �� �� " name="B2"></p><%=d2%> 
</form>
<%case"delany"%>
<center><font color="#FF0000">ע�⣺ʹ�ô˹��ܽ�ɾ��ָ���������Լ����棬ɾ�����ָܻ���������ʹ�ã���</font></center>
<form action="admint-gl.asp?menu=delany" method="POST">
<%=t1%>ɾ��ָ�������ڵ�����<%=t2&d1%>
<P style="MARGIN: 5px">ɾ��������ǰ�����ӣ�( ������ ) 
<input type="text" name="daynum" size="19"></p>
<P style="MARGIN: 5px">ɾ���������ڵ���̳��( ��ѡ�� ) <%=bdlist("bd",1)%></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<form action="admint-gl.asp?menu=delnore" method="POST">
<%=t1%>ɾ��ָ��������û�лظ�������<%=t2&d1%>
<P style="MARGIN: 5px">ɾ��������ǰ�����ӣ�( ������ ) 
<input type="text" name="daynum" size="19"></p>
<P style="MARGIN: 5px">ɾ���������ڵ���̳��( ��ѡ�� ) <%=bdlist("bd",1)%></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form><form action="admint-gl.asp?menu=delwhose" method="POST">
<%=t1%>ɾ��ָ���û�����������<%=t2&d1%>
<P style="MARGIN: 5px">ɾ��ָ���û������ӣ�( �û��� ) 
<input type="text" name="ddname" size="19"></p>
<P style="MARGIN: 5px">ɾ���������ڵ���̳��( ��ѡ�� ) <%=bdlist("bd",1)%></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<%case"bbsmail"%>
<center><font color="#FF0000">ע�⣺ʹ�ô˹��ܽ�ɾ��ָ�����ʼ���ɾ�����ָܻ���������ʹ�ã���</font></center>
<form action="admint-gl.asp?menu=delanymail" method="POST">
<%=t1%>ɾ��ָ�������ڵ�����<%=t2&d1%>
<P style="MARGIN: 5px">ɾ��������ǰ�����ԣ�( ������ ) 
<input type="text" name="daynum" size="19"></p>
<P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<form action="admint-gl.asp?menu=delwhosemail" method="POST">
<%=t1%>ɾ��ָ���û�����������<%=t2&d1%>
<P style="MARGIN: 5px">ɾ��ָ���û������ԣ�( �û��� ) 
<input type="text" name="ddname" size="19"></p>
<P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>

<%case"seepwd"%>
<form action="admin.gl.asp?menu=lookpwd" method="POST">
<%=t1%>�鿴�û�����<%=t2&d1%><P style="MARGIN: 5px">�û�����<input type="text" name="chaname" size="20"> 
<input type="submit" value=" �� �� " name="B1">
        <input type="reset" value=" �� �� " name="B2"></p><%=d2%> 
</form>
<%case"hbbbs"%>
<form action="admint-gl.asp?menu=hbbbs" method="POST">
<%=t1%>�ϲ���̳<%=t2&d1%><P style="MARGIN: 10px">����̳ <%=bdlist("frombd",0)%> �ϲ��� <%=bdlist("tobd",0)%>
<input type="submit" value=" ��  �� " name="B1"><br><br><font color="#FF0000">ע�⣺�ϲ��󣬱��ϲ��������̳��ȥ����̳����ɾ��<%=d2%> </font>
</form>
<%case"moveany"%>
<form action="admint-gl.asp?menu=moveday" method="POST">
<%=t1%>��ָ�������ƶ�����<%=t2&d1%>
<P style="MARGIN: 5px">�ƶ�������ǰ�����ӣ�( ������ ) 
<input type="text" name="daynum" size="19"></p>
<P style="MARGIN: 5px">����ԭ�����ڵ���̳��( ��ѡ�� ) 
<%=bdlist("frombd",0)%></p>
<P style="MARGIN: 5px">����Ҫ�ƶ�������̳��( ��ѡ�� ) 
<%=bdlist("tobd",0)%></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<form action="admint-gl.asp?menu=movename" method="POST">
<%=t1%>��ָ���û��ƶ�����<%=t2&d1%>
<P style="MARGIN: 5px">�ƶ�ָ���û������ӣ�( �û��� ) 
<input type="text" name="movename" size="19"></p>
<P style="MARGIN: 5px">����ԭ�����ڵ���̳��( ��ѡ�� ) 
<%=bdlist("frombd",0)%></p>
<P style="MARGIN: 5px">����Ҫ�ƶ�������̳��( ��ѡ�� ) 
<%=bdlist("tobd",0)%></p><P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1">&nbsp; 
<input type="reset" value=" �� �� " name="B2">&nbsp;<%=d2%></form>
<%case"lm"%>
<form action="admint-gl.asp?menu=addlm" method="POST">
<%=t1%>�����̳����<%=t2&d1%><P style="MARGIN: 5px">��̳���ƣ�<input type="text" name="name" size="20"></p>
<P style="MARGIN: 5px">��̳��ַ��<input type="text" name="url" size="38"></p>
<P style="MARGIN: 5px">��̳ͼƬ��<input type="text" name="picurl" size="38"></p> 
<P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1"> <input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<%=t1%>������̳����<%=t2&d1%><P style="MARGIN: 5px">
<%set slm=myconn.execute("select*from lmbbs")
do while not slm.eof
ha=slm("name")
response.write"<table border=0 cellpadding=0 cellspacing=0 style='border-collapse: collapse' width=100% height=18><tr><td width=33% >"&kbbs(ha)&"</td>"
response.write"<td width='33%'><a href='admint-gl.asp?menu=dellm&name="&kbbs(ha)&"'>ɾ��</a></td><td width='34%'><a href='?action=editlm&name="&kbbs(ha)&"'>�༭</a></td></tr></table>"
slm.movenext
loop
set slm=nothing
%></p><%=d2%>
<%case"editlm"
name=Replace(Request.querystring("name"),"'","''")
set elm=myconn.execute("select*from lmbbs where name='"&name&"'")
%>
<form action="admint-gl.asp?menu=editlm&name=<%=kbbs(elm("name"))%>" method="POST">
<%=t1%>�༭��̳����<%=t2&d1%><P style="MARGIN: 5px">��̳���ƣ�<%=kbbs(elm("name"))%></p>
<P style="MARGIN: 5px">��̳��ַ��<input type="text" name="url" size="38" value="<%=elm("url")%>"></p>
<P style="MARGIN: 5px">��̳ͼƬ��<input type="text" name="picurl" size="38" value="<%=elm("picurl")%>"></p> 
<P style="MARGIN: 5px"><input type="submit" value=" �� �� " name="B1"> <input type="reset" value=" �� �� " name="B2"></p><%=d2%>
</form>
<%set elm=nothing%>
<%end select
end if%>
