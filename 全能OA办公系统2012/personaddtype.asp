<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
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
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统</title>
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
                    <td class="style7">个人通讯录</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
<tr>
<td>个人通讯录类别管理&nbsp;&nbsp;&nbsp;&nbsp;</td>
<form method="post" action="personlist.asp">
<td>
<input type="submit" name="submit" value="返回">
</td>
</form>
</tr>
<tr>
<td>
<font color="#dd0000">注意：删除个人通讯录名称时将删除该通讯录中的所有资料！</font>
</td>
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
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from persontype where typename='" & userlevel&"' and username='"&oabusyusername&"'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>已经存在，请换名重试！</font><br>
<%

else
sql = "Insert Into persontype (username,typename) Values('" & oabusyusername &"','"&userlevel& "')"
conn.Execute sql
%>
<font color=red><%=userlevel%>增加成功！</font>
<%
end if
end if
'---------------------------------------------------
if request("submit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete from personrecord where recordtype=" &id
conn.Execute sql
sql="delete from persontype where id="&id
conn.execute sql
%>
<font color=red><%=userlevel%>删除成功！</font>
<%
end if
'---------------------------------------------------
if request("submit")="修改" and userlevel<>"" then

'判断是否有与修改的职位相同的
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from persontype where typename='" & userlevel&"' and username='"&oabusyusername&"' and id<>"&id
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<font color=red><%=userlevel%>已经存在，请换名重试！</font><br>
<%
else
sql = "update persontype set typename='" & userlevel & "' where id=" & id
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
sql="select * from persontype where username='"&oabusyusername&"'"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr>
<form method="post" action="personaddtype.asp">
<td>
<input type="submit" name="submit" value="删除"></td><td bgcolor="D7E8F8"><input type="hidden" name="olduserlevel" value="<%=rs("typename")%>"><input type="hidden" name="id" value=<%=rs("id")%>><input type="text" name="userlevel" value="<%=rs("typename")%>" maxlentgh="25"></td><td bgcolor="D7E8F8"><input type="submit" name="submit" value="修改"></td>
</form>
</tr>
<%
rs.movenext
wend
%>
</table>
<form method="post" action="personaddtype.asp">
<input type="text" name="userlevel"><input type="submit" name="submit" value="增加">
</form>
</center>

</body>
</html>


</body>
</html>
