<!--#include file="up.asp"--><!--#include file="fun.asp"-->
<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<br><br>
<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>用 户 列 表</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px" bordercolor="<%=c1%>" width="94%">
    <tr>
      <td width="22%" align="center" height="25" bgcolor="<%=c2%>"><b>用户名</b></td>
      <td width="10%" align="center"><b>性别</b></td>
      <td width="10%" align="center" bgcolor="<%=c2%>"><b>E-mail</b></td>
      <td width="14%" align="center"><b>QQ号码</b></td>
      <td width="10%" align="center" bgcolor="<%=c2%>"><b>主页</b></td>
      <td width="10%" align="center"><b>发帖数</b></td>
      <td width="25%" align="center" bgcolor="<%=c2%>"><b>等级</b></td>
    </tr>
  </table>
  </center>
</div>
<%
dim rs
dim sql
set rs = server.createobject("adodb.recordset")
sql = "select * from user order by userid desc"
count=myconn.execute("select count(name)from user")(0)
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
i=1
do while not rs.eof
response.write"<div align=center><center><table border=1 cellpadding=0 cellspacing=0 bordercolor="&c1&" width=94% ><tr><td width='22%' align=center height=25 bgcolor="&c2&">"
myname=kbbs(rs("name"))
response.write"<a href='userinfo.asp?name="&myname&"'>"&myname&"</a></td><td width=10% align=center>"
if rs("sex")=1 then
sex1="男"
else
sex1="女"
end if
response.write""&sex1&"　</td><td width=10% align=center bgcolor="&c2&"><a href='mailto:"&rs("email")&"'><img src='pic/mail.gif' border='0'></a>　</td><td width='14%' align=center>"&rs("qq")&"　</td><td width=10% align=center bgcolor="&c2&">"
if rs("home")="" then
response.write"<img border='0' src='pic/home.gif'>"
else
response.write"<a target='_blank' href='"&rs("home")&"'><img border='0' src='pic/home.gif'></a>"
end if
response.write"</td><td width=10% align=center>"&myconn.execute("select count(riqi) from min where name='"&rs("name")&"'")(0)&"　</td><td width='25%' align=center bgcolor="&c2&">"
q1=rs("qian")%><%m1=rs("meili")%><%j1=rs("jingyan")%><%sqltype="my"%><!--#include file="upji.asp"-->
<%
response.write"<b>"&dj&"</b>『 "&dd&" 』</td></tr></table></center></div>"
%>

<%i=i+1
if i>pagesetup then exit do
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
<TD><b><font color="#FFFFFF"><img border="0" src="pic/fl.gif"> 注册用户共有</font><font color="#00FFFF"> <%=TotalPage%> </font><font color="#FFFFFF">页,<font color="#00FFFF"> <%=count%> </font>位，每页有<font color="#00FFFF"> <%=pagesetup%> </font> 位用户 >> [
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
Response.Write "<a href=?topage=1><font color=yellow>1</font></a> ... "
end if

for i=ii to iii
If i<>PageCount then
Response.Write "<a href=?topage="& i &"><font color=yellow>" & i & "</font></a> "
else
Response.Write " <font color=red><b>"&i&"</b></font> "
end if
next

if TotalPage > PageCount+5 then
Response.Write " ... <a href=?topage="&TotalPage&"><font color=yellow>"&TotalPage&"</font></a>"
end if
%> ]</font></b></TD>
<form name=form method="POST" action=javascript:Check()>
<TD height=2 align="right"><font color="#FFFFFF">页码：<input style=FONT-SIZE:9pt maxLength="6" size="6" name="topage" value="<%=PageCount%>">
<input style=FONT-SIZE:9pt value="GO!" type="submit"></font></TD></form>
</TR></TBODY></TABLE></TD></TR></TBODY></TABLE>
</center>
  </div><br><!--#include file="down.asp"-->