<%@ LANGUAGE = VBScript %>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->

<%
oabusyusername=request.cookies("oabusyusername")
oabusyuserid=request.cookies("oabusyuserid")
if request("submit")="我已经看了此通告" then
	id=request("id")
	set conn=opendb("oabusy","conn","accessdsn")
	sql = "update newnotice set readuserid=readuserid+'("&oabusyuserid&")' where ID="&id
	conn.Execute sql
	conn.close
%>
<SCRIPT language=JavaScript>                   
window.close();
opener.location.reload();
</script> 
<%
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<title>查看公司最新通告</title>
<style type="text/css">
<!--
.style6 {color: #098abb}
.style7 {color: #0d79b3}
.z14 {font-size: 14px;
	font-weight: bold;
	color: #098abb;
}
-->
</style>
</head>
<bgsound src="xbmsg.wav" loop="1">
<body  topmargin="5" leftmargin="5" onunload="opener.location.reload();">
<%

set conn=opendb("oabusy","conn","accessdsn")
Set rs=Server.CreateObject("ADODB.recordset")
sql="select * from newnotice where readuserid NOT LIKE '%("&oabusyuserid&")%' and sendusername<>'"&oabusyusername&"'"
rs.open sql,conn,1
%>
<center>
<table width="550"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="30"><div align="center" class="z14">通告标题：<%=rs("title")%></div></td>
  </tr>
  <tr>
    <td height="1" bgcolor="6FC6E7"></td>
  </tr>
  <tr>
    <td height="15" bgcolor="F0FAFF"><table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td>&nbsp;</td>
        </tr>
        <tr>
          <td><span class="style7"><%=keepformat(rs("content"))%>
                <%

%>
          </span></td>
        </tr>
    </table></td>
  </tr>
  <tr>
    <td height="30" bgcolor="F0FAFF"><div align="right"><span class="style6">[发布日期：<%=rs("noticedate")%>] </span>　</div></td>
  </tr>
</table>
</center>

&nbsp;<br>
<center>
<form method=post action="popnotice.asp">
<input type="hidden" value="<%=rs("ID")%>" name="id">
<input type="submit" name="submit" value="我已经看了此通告">
</form>
</center>
<%

%> 
</body>
</html>