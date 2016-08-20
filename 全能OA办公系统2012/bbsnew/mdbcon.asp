<!--#include file="conn.asp"--><%set bbs=myconn.execute("select*from bbsinfo")
sty="all"
sp=request.cookies(cn&"1")(sty)
c1=request.cookies(cn&"1")(sty&"c1")
c2=request.cookies(cn&"1")(sty&"c2")
if sp="" then sp="a"
if c1="" then c1=bbs(1)
if c2="" then c2=bbs(2)
set bbs=nothing
%>
<body topmargin="0" leftmargin="0"><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>

<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=<%=c1%>>
  <tr>
    <td width="100%" height="28" background="pic/<%=sp%>3.gif" align="center">
    <b><font color="#FFFFFF">论坛后台管理系统</font></b></td>
  </tr>
  </table><br>
<link rel="stylesheet" type="text/css" href="css.css">
<%
t1="<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% background=pic/"&sp&"3.gif height=25 bgcolor="&c1&">&nbsp;<img border=0 src=pic/fl.gif> <font color=#FFFFFF><b>"
t2="</b></font></td></tr>"
d1="<tr><td width=100% >"
d2="</td></tr></table></center></div>"
lgname=Request.Cookies(cn)("lgname")
lgpwd=Request.Cookies(cn)("lgpwd")
set can=myconn.execute("select*from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if can.eof or can.bof then
%>

<br><br><%=t1%>权限不够！<%=t2&d1%>你没有权限访问该页！！<br>你正在论坛使用的用户：<b><%=lgname%></b>没有斑竹的权限，或者密码错误！<br>请使用具有斑竹权限的用户<a target="_self" href="login.asp"><font color="#0033CC">登陆本论坛</font></a>
<%=d2%>
<%else%>
<%myconn.close
set myconn=nothing%><%
select case Request("menu")
case "bakbf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("yl"))&"",""&Server.MapPath(Request("bf"))&""
%><%=t1%>备份成功！<%=t2&d1%><P style='MARGIN: 5px'>·备份成功·</p><%=d2%>
<%case "bakhf"
set MyFileObject=Server.CreateOBject("Scripting.FileSystemObject")
MyFileObject.CopyFile ""&Server.MapPath(Request("bf"))&"",""&Server.MapPath(Request("yl"))&""
%><%=t1%>恢复成功！<%=t2&d1%><P style='MARGIN: 5px'>·恢复成功·</p><%=d2%>

<%case "yasuo"%>
<%
if instr(Request.ServerVariables("http_referer"),""&Request.ServerVariables("server_name")&"") = 0 then
response.write ""&t1&"来源错误！"&t2&""&d1&"<P style='MARGIN: 5px'>来源错误！<a href=javascript:history.go(-1)>返回</a></p>"&d2&""
response.end
end if
Const JET_3X = 4
Function CompactDB(dbPath, boolIs97)
Dim fso, Engine, strDBPath
strDBPath = left(dbPath,instrrev(DBPath,"\"))
Set fso = CreateObject("Scripting.FileSystemObject")
If fso.FileExists(dbPath) Then
Set Engine = CreateObject("JRO.JetEngine")
Engine.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & dbpath, _
"Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & strDBPath & "temp.mdb"
fso.CopyFile strDBPath & "temp.mdb",dbpath
fso.DeleteFile(strDBPath & "temp.mdb")
Set fso = nothing
Set Engine = nothing
CompactDB = ""&t1&"压缩成功！"&t2&""&d1&"<P style='MARGIN: 5px'>压缩成功！<a href=javascript:history.go(-1)>返回</a></p>"&d2&""
Else
CompactDB = ""&t1&"压缩失败！"&t2&""&d1&"<P style='MARGIN: 5px'>找不到数据库！请检查数据库路径是否输入错误！<a href=javascript:history.go(-1)>返回</a></p>"&d2&""
End If
End Function
Dim dbpath,boolIs97
dbpath = request("dbpath")
boolIs97 = request("boolIs97")
If dbpath <> "" Then
dbpath = server.mappath(dbpath)
response.write(CompactDB(dbpath,boolIs97))
End If
%>
<%
end select
%>
<form action=?menu=yasuo method="POST">
<%=t1%>压缩数据库<%=t2&d1%><P style='MARGIN: 5px'>压缩的数据库路径： 
<input size="30" name="dbpath" value="db\〓6k〓.mdb"></p><P style='MARGIN: 5px'><input type="submit" value=" 压 缩 " name="Submit"></p><%=d2%>
</form><br>
<form action=?menu=bakbf method="POST">
<%=t1%>备份数据库<%=t2&d1%><P style='MARGIN: 5px'>原来的数据库路径：
<input size="30" value="db\〓6k〓.mdb" name="yl"></p>
  <P style='MARGIN: 5px'>备份的数据库路径： <input size="30" value="db\bak6k.mdb" name="bf"></p>
  <P style='MARGIN: 5px'><input type="submit" value=" 备 份 " name="Submit1"></p>
    <%=d2%>
</form>
<br>
<form action=?menu=bakhf method="POST">
<%=t1%>恢复数据库<%=t2&d1%><P style='MARGIN: 5px'>备份的数据库路径：
<input size="30" value="db\bak6k.mdb" name="bf"> </p>
  <P style='MARGIN: 5px'>原来的数据库路径： <input size="30" value="db\〓6k〓.mdb" name="yl"></p>
  <P style='MARGIN: 5px'><input type="submit" value=" 恢 复 " name="Submit"></p>
    <%=d2%>
</form><%end if%>