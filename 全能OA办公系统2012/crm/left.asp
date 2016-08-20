<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
set conn=dbconn("conn")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="120">
    <tr>
      <td width="13"><img border="0" src="images/openfld.gif" align="absmiddle" width="16" height="16"></td>
    </center>
    <td width="103">
      <p align="left"><img border="0" src="images/treeline1.gif" align="absmiddle" width="18" height="16">企业名录</td>
  </tr>
  <center>
  <tr>
    <td width="13"><img border="0" src="images/midopenedfolder.gif" align="top" width="16" height="22"></td>
    <td width="103"><img border="0" src="images/menu_admin.gif" align="absmiddle" width="16" height="16"> 
      按省份分类</td>
  </tr>
<%
set rs=server.createobject("adodb.recordset")
sql="select * from diqu"
rs.open sql,conn,1
j=1
do while not rs.eof
%>
  <tr>
    <td width="13"><img border="0" src="images/vertline.gif" align="top" width="16" height="22"></td>
    <td width="103">
<%
	if j<rs.recordcount then
%>
	<img border="0" src="images/midnodeline.gif" align="absmiddle" width="16" height="22">
<%
	else
%>
	<img border="0" src="images/lastnodeline.gif" align="absmiddle" width="16" height="22">
<%
	end if
%>
	<img border="0" src="images/menu_about.gif" align="absmiddle" width="16" height="16"><a href="dispinfo.asp?page=1&typenumber=1&lookstr=<%=rs("diqu")%>" target="mainFrame"><%=server.htmlencode(trim(rs("diqu")))%></a></td>
  </tr>
<%
	rs.movenext
	j=j+1
loop
set rs=nothing
set rs=server.createobject("adodb.recordset")
sql="select * from fenlei"
rs.open sql,conn,1
%>
  <tr>
    <td width="13"><img border="0" src="images/midopenedfolder.gif" align="top" width="16" height="22"></td>
    <td width="103"><img border="0" src="images/menu_admin.gif" align="absmiddle" width="16" height="16"> 
      按行业分类</td>
  </tr>
<%
j=1
do while not rs.eof
%>
  <tr>
    <td width="13"><img border="0" src="images/vertline.gif" align="top" width="16" height="22"></td>
    <td width="103">
<%
	if j<rs.recordcount then
%>	
	<img border="0" src="images/midnodeline.gif" align="absmiddle" width="16" height="22">
<%
	else
%>
	<img border="0" src="images/lastnodeline.gif" align="absmiddle" width="16" height="22">
<%
	end if
%>
	<img border="0" src="images/menu_about.gif" align="absmiddle" width="16" height="16"><a href="dispinfo.asp?page=1&typenumber=2&lookstr=<%=rs("leibie")%>" target="mainFrame"><%=server.htmlencode(trim(rs("leibie")))%></a></td>
  </tr>
<%
	rs.movenext
	j=j+1
loop
conn.close
set rs=nothing
set conn=nothing
%>
  <tr>
    <td width="13"><img border="0" src="images/midopenedfolder.gif" align="top" width="16" height="22"></td>
    <td width="103"><img border="0" src="images/menu_admin.gif" align="absmiddle" width="16" height="16">&nbsp;<a href="findinfo.asp" target="mainFrame">企业查询</a></td>
  </tr>
  <tr>
    <td width="13"><img border="0" src="images/midopenedfolder.gif" align="top" width="16" height="22"></td>
    <td width="103"><img border="0" src="images/menu_admin.gif" align="absmiddle" width="16" height="16">&nbsp;<a href="addtype.asp" target="mainFrame">企业类别管理</a></td>
  </tr>
  <tr>
    <td width="13"><img border="0" src="images/lastnodeline.gif" width="16" height="22"></td>
    <td width="103"><img border="0" src="images/menu_admin.gif" align="absmiddle" width="16" height="16">&nbsp;<a href="inputinfo.asp" target="mainFrame">增加企业</a></td>
  </tr>
  <tr>
    <td width="13"></td>
    <td width="103"></td>
  </tr>
  </table>
  </center>
</div>

</body>

</html>
