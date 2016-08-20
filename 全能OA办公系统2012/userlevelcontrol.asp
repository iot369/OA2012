<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->

<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>oa办公系统</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
              <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">用户设置</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <br>
  <table>
<tr>
<td>职位名称管理&nbsp;&nbsp;&nbsp;&nbsp;</td>
<form method="post" action="adduser.asp">
<td>
<input type="submit" name="submit" value="增加用户">
</td>
</form>
<form method="post" action="deptcontrol.asp">
<td>
<input type="submit" name="submit" value="部门名称管理">
</td>
</form>
</tr>
</table>
</center>

<br>
<center>
<%
userlevel=request("userlevel")
olduserlevel=request("olduserlevel")
id=request("id")
'-----------------------------------------------
if request("submit")="增加" and userlevel<>"" then

'判断是否有与增加的职位相同的
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userlevel where userlevel=" & sqlstr(userlevel)
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>的职位已经存在，请选择其他职位名</font><br>
<%

else
sql = "Insert Into userlevel (userlevel) Values( " & sqlstr(userlevel) & ")"
conn.Execute sql
%>
<font color=red><%=userlevel%>增加成功！</font>
<%
end if
end if
'---------------------------------------------------
if request("submit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete * from userlevel where userlevel=" & sqlstr(olduserlevel)
conn.Execute sql
%>
<font color=red><%=userlevel%>删除成功！</font>
<%
end if
'---------------------------------------------------
if request("submit")="修改" and userlevel<>"" then

'判断是否有与修改的职位相同的
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userlevel where userlevel=" & sqlstr(userlevel) & " and id<>" & id
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>的职位已经存在，不能修改为此名称！</font><br>
<%
else
sql = "update userlevel set userlevel=" & sqlstr(userlevel) & " where id=" & id
conn.Execute sql

sql = "update userinf set userlevel=" & sqlstr(userlevel) & " where userlevel=" & sqlstr(olduserlevel)
conn.Execute sql

%>
<font color=red>修改成功！</font>
<%
end if
end if
%>
<table border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userlevel"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr>
<form method="post" action="userlevelcontrol.asp">
<td bgcolor="#FFFFFF">
<input type="submit" name="submit" value="删除"></td><td bgcolor="#FFFFFF"><input type="hidden" name="olduserlevel" value="<%=rs("userlevel")%>"><input type="hidden" name="id" value=<%=rs("id")%>><input type="text" name="userlevel" value="<%=rs("userlevel")%>"></td><td bgcolor="#FFFFFF"><input type="submit" name="submit" value="修改"></td>
</form>
</tr>
<%
rs.movenext
wend
%>
</table>
<form method="post" action="userlevelcontrol.asp">
<input type="text" name="userlevel"><input type="submit" name="submit" value="增加">
</form>
</center>
<%

%>
</body>
</html>