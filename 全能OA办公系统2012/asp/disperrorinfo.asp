<%response.expires=0%>
<!--#include file="sqlstr.asp"-->
<!--#include file="opendb.asp"-->
<!--#include file="bgsub.asp"-->
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>伴江行办公系统</title>
</head>
<body  topmargin="5" leftmargin="5">

<center>
<table>
<tr>
<td><font color="#ee0000" size="+2">出错了！</font>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td>
<input type="button" name="button" value="返回" onclick="history.go(-1);">
</td>
</tr>
</table>
</center>

<br>
<div align="center">
<table border="0" cellpadding="0" cellspacing="0" width="80%">
  <tr>
    <td width="12%"><img src="../image/errorico.gif" border="0">
</td>
    <td width="88%" height="200"><font color="#ee0000" size="+2"><%=request("errorinfo")%></font></td>
  </tr>
</table>
</div>
<%

%>
</body>
</html>
