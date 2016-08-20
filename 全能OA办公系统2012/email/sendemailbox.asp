<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<%
'session.abandon
'Server.ScriptTimeOut=500
function opendb(DBPath,sessionname,dbsort)
dim conn
'if not isobject(session(sessionname)) then
Set conn=Server.CreateObject("ADODB.Connection")
'if dbsort="accessdsn" then conn.Open "DSN=" & DBPath
'if dbsort="access" then conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
'if dbsort="sqlserver" then conn.Open "DSN=" & DBPath & ";uid=wsw;pwd=wsw"
DBPath1=server.mappath("../db/lmtof.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<!--#include file="../asp/bgsub.asp"-->
<!--#include file="publiclist.asp"-->
<!--#include file="delemail.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>oa办公系统</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0" bgcolor="#F9F9FF">
<SCRIPT language=javascript>
<!--
if (window.Event) 
　document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
　if (e.which == 2 || e.which == 3)
　 return false;
 } else if (event.button == 2 || event.button == 3) {
　 event.cancelBubble = true
　 event.returnvalue = false;
　 return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;　// for IE5+
document.onmousedown = norightclick;　　 // for all others
//-->
</SCRIPT>
<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
              <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">个人邮件</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
<tr>
<td>发件箱</td>
</tr>
</table>
</center>

<form method="post" name="form1" action="sendemailbox.asp">
<%
if request("delbutton")="永久删除" then
	call delemail(2)
end if
sql="select * from sendemailtable where senduserid="&cstr(oabusyuserid)&" order by emaildate desc"
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		conn.close
		set rs=nothing
		response.write("<p align=""center""><font color=""#dd0000"">您的发件箱中没有邮件！</font></p>")
	else
%>
<script language="javascript">
function lookemail(recordid)
{
	win=window.open('resendemail.asp?id='+recordid,'win'+recordid,'toolbar=no,scrollbars=yes,resizable=0,menubar=no,width=550,height=500');	
}
</script>
<p align="center">
共<%=cstr(rs.recordcount)%>条已发送邮件
（<font color="#336699">单击邮件“主题”可重发或转发该邮件！</font>）
</p>
<div align="center">
  <center>
  <table border="1" width="540" cellspacing="0" cellpadding="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
    <tr bgcolor="D7E8F8">
      <td width="35" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><span class="style8">选择</span></td>
      <td width="113" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><span class="style8">收件人</span></td>
      <td width="276" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><span class="style8">主题</span></td>
      <td width="125" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><span class="style8">日期</span></td>
    </tr>
<%
	do while not rs.eof
%>
    <tr bgcolor="#ffffff">
      <td width="35" align="center">
  	  	<input type="checkbox" name="selectnumber" value="<%=cstr(rs("autoid"))%>">
	  </td>
      <td width="113" align="center"><%=server.htmlencode(rs("explain"))%></td>
      <td width="275" align="center"><a href="#" onclick="javscript:lookemail('<%=cstr(rs("autoid"))%>')"><font color="#336699"><%=server.htmlencode(rs("emailtitle"))%></font></a></td>
      <td width="125" align="center"><%=cstr(rs("emaildate"))%></td>
    </tr>
<%
	rs.movenext
	loop
%>
  </table>
  </center>
</div>
<%
	end if
%>
<br>
<div align="right">
<input type="submit" name="delbutton" value="永久删除">
</div>
</form>
<%

%>
</body>
</html>
