<!--#include file="up.asp"--><%
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
%><!--#include file="fun.asp"--><!--#include file="ubbcode.asp"--><br><br>
<%set mail=myconn.execute("select name from user where name='"&lgname&"' and password='"&lgpwd&"'")
if mail.eof then%>
<%=t1%>�� �� �� Ϣ<%=t2&d1%><P style='MARGIN: 10px'>��������ʧ�ܣ����ܴ����������⣺</p><P style='MARGIN: 10px'>���㻹û��<a href="login.asp">��½</a>��</p><P style='MARGIN: 10px'>������û������������</p><%=d2%>
<%
response.end
end if
set mail=nothing%>
<%menu=request.querystring("menu")
select case menu
case""%>
<%=t1%>�������԰�<%=t2%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse;" bordercolor="<%=c1%>" width="94%">
    <tr>
      <td width="100%" height="28" colspan="5" style="border-left-style: solid; border-left-width: 1; border-right-style: solid; border-right-width: 1; border-top-width: 1; border-bottom-width: 1" bordercolor="<%=c1%>">&nbsp;<img border="0" src="pic/xie.gif"> <span lang="zh-cn">
      <a href="mailcon.asp?menu=write">��������</a></span>&nbsp;
      <img border="0" src="pic/del.gif"> <a href="bbsmail.asp?menu=delall">������԰�</a>&nbsp;&nbsp;&nbsp;&nbsp; 
      ��������԰干��<%mailno=myconn.execute("select count(name) from hand where tname='"&lgname&"'")(0)%> <b><%=mailno%></b> �����ԡ�&nbsp;&nbsp;&nbsp; <marquee scrollamount="2" width="25%">����ϧÿһ��ռ䣬�鷳�㼰ʱɾ�����õ�������Ϣ��лл��</marquee></td>
    </tr>
</table>
  </center>
</div><script language='javascript'>function Check(){var Name=document.form.topage.value;document.location='?id=0&topage='+Name+'';}</script>
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select * from hand where tname='"&lgname&"' order by id desc"
count=myconn.execute("select count(name)from hand where tname='"&lgname&"'")(0)
on error resume next
pagesetup=10
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
aa=1
do while not rs.eof%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%" height="150">
    <tr>
      <td width="20%" valign="top">
      <div align="center">
        <center>
        <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="90%">
          <tr>
            <td width="100%" align="center"><br>
            <div align="center">
              <center><TABLE style="FILTER: glow(color=<%=c1%>, strength=1); BORDER-COLLAPSE: collapse" borderColor=#111111 cellSpacing=0 cellPadding=0 width=*><FONT color=black><%=kbbs(rs("name"))%></FONT></TABLE>
              </center>
            </div><br>
<%set gh=myconn.execute("select top 1 toupic,ch,ku from user where name='"&rs("name")&"'")%><img src=<%=kbbs(gh("toupic"))%> border="0" width="<%=gh("ku")%>" height="<%=gh("ch")%>"></td>
          </tr>
        </table>
        </center>
      </div><br>
      </td>
      <td width="80%" valign="top">
      <table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" width="100%" height="25"><tr>
        <td width="100%" valign="bottom">&nbsp;<a href='userinfo.asp?name=<%=kbbs(rs("name"))%>'><img border="0" src="pic/info.gif"> �� Ϣ</a>
        <a href="mailcon.asp?menu=write&towho=<%=kbbs(rs("name"))%>">
        <img border="0" src="pic/xie.gif"> �� ��</a>
        <a href="?id=<%=rs("id")%>&menu=del"> <img border="0" src="pic/del.gif"> ɾ ��</a></td></tr></table>
      <hr color=<%=c1%> width="98%" size="1">
      <blockquote><img src="pic/tl.gif" border="0"> <%rrr=rs("neirong")%><%=ubb(rrr)%><p></p><div align=right>
        <img src="pic/xie.gif" border="0"> <%=rs("riqi")%></div></blockquote>
      </td>
    </tr>
  </table>
  </center>
</div><table cellspacing=0 border=0><tr><td height=2></td></tr></table>
<%
aa=aa+1
if aa>pagesetup then exit do
rs.movenext
loop
rs.Close
%>
<div align="center">
<center>
<TABLE borderColor=<%=c1%> cellSpacing=0 cellPadding=0 width="94%" border=1 style="border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px">
<TBODY>
<TR height=25>
<TD height=2>
<TABLE cellSpacing=0 cellPadding=3 width="100%" border=0 background="pic/<%=sp%>3.gif" style="border-collapse: collapse; border-left-width:0; border-top-width:0; border-bottom-width:0" bgcolor="<%=c1%>">
<TBODY>
<TR>
<TD><b><font color="#FFFFFF"><img border="0" src="pic/fl.gif"> ����</font><font color="#00FFFF"> <%=TotalPage%> </font><font color="#FFFFFF">ҳ,<font color="#00FFFF"> <%=count%> </font>
�����ԣ�ÿҳ��<font color="#00FFFF"> <%=pagesetup%> </font> ������ >> [
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
Response.Write "<a href=?topage=1&id=0><font color=yellow>1</font></a> ... "
end if

for i=ii to iii
If i<>PageCount then
Response.Write "<a href=?topage="& i &"&id=0><font color=yellow>" & i & "</font></a> "
else
Response.Write " <font color=red><b>"&i&"</b></font> "
end if
next

if TotalPage > PageCount+5 then
Response.Write " ... <a href=?topage="&TotalPage&"&id=0><font color=yellow>"&TotalPage&"</font></a>"
end if
%> ]</font></b></TD>
<form name=form method="POST" action=javascript:Check()>
<TD height=2 align="right"><font color="#FFFFFF">ҳ�룺<input style=FONT-SIZE:9pt maxLength="6" size="6" name="topage" value="<%=PageCount%>">
<input style=FONT-SIZE:9pt value="GO!" type="submit"></font></TD></form>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</center>
  </div>
<%case"del"
myconn.execute("delete*from hand where id="&id&" and tname='"&lgname&"'")
%><%=t1%>ɾ �� �� ��<%=t2&d1%><p style="margin: 10">������ɾ���ɹ���</p><%=d2%>
<%case"delall"
myconn.execute("delete*from hand where tname='"&lgname&"'")
%><%=t1%>�� �� �� ��<%=t2&d1%><p style="margin: 10">��������ճɹ���</p><%=d2%>
<%end select
myconn.execute("update hand set isnew='1' where tname='"&lgname&"'")
%>
<br><!--#include file="down.asp"-->