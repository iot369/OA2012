<!--#include file="up.asp"--><%
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
id=request.querystring("id")
%><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<!--#include file="fun.asp"--><!--#include file="ubbcode.asp"--><br><br>
<%set mail=myconn.execute("select name from user where name='"&lgname&"' and password='"&lgpwd&"'")
if mail.eof then%>
<%=t1%>�� �� �� Ϣ<%=t2&d1%><P style='MARGIN: 10px'>��������ʧ�ܣ����ܴ����������⣺</p><P style='MARGIN: 10px'>���㻹û��<a href="login.asp">��½</a>��</p><P style='MARGIN: 10px'>������û������������</p>
<%
response.end
end if
%>
<%menu=request.querystring("menu")
select case menu
%>
<%case "write"
towho=request.querystring("towho")
%>
<form method="POST" action="?menu=save">
<%=t1%>��������<%=t2&d1%>
<p style="margin: 10">���Զ���<input type="text" name="tname" size="82" value="<%=towho%>"></p>
<p style="margin: 10">�������ݣ�[ ����ʹ��UBB�����Լ���̳���ӵ����й��ܣ������ϴ������棩 ]</p>
<p style="margin: 10"><textarea rows="12" name="neirong" cols="92"></textarea></p>
<p style="margin: 10"><input type="submit" value=" �� �� " name="B1"> <input type="reset" value=" �� �� " name="B2"></p>
<%=d2%>
</form>
<%case"save"
tname=Replace(Request.Form("tname"),"'","''")
neirong=Replace(Request.Form("neirong"),"'","''")
if tname="" or neirong="" then
%><%=t1%>�� �� �� Ϣ<%=t2&d1%><p style="margin: 10">������ʧ�ܣ����Զ�����������ݲ������ա�</p><%=d2%>
<%else
set isha=myconn.execute("select name from user where name='"&tname&"'")
if isha.eof then%><%=t1%>�� �� �� Ϣ<%=t2&d1%><p style="margin: 10">������ʧ�ܣ���̳�в����ڸ����Զ���</p><%=d2%>
<%else
myconn.execute("insert into hand(name,neirong,riqi,tname)values('"&lgname&"','"&neirong&"',now,'"&tname&"')")%>
<%=t1%>�� �� �� ��<%=t2&d1%><p style="margin: 10">���Ѿ��ɹ��ĸ� <b><%=kbbs(tname)%></b> ���ԡ�</p><%=d2%>
<%end if
set isha=nothing
end if%>
<%end select%>