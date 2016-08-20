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
<!--#include file="publiclist.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("alert(""您已经过期，请重新登录！"");")
	response.write("window.close();")
	response.write("</script>")
	response.end
end if
id=request.querystring("id")
if id="" then
	response.write("<script language=""javascript"">")
	response.write("alert(""数据传送出错，将关闭窗口！"");")
	response.write("window.close();")
	response.write("</script>")
	response.end
end if	
if request("submit")="发送" then
	emailtitle=rtrim(request("title"))
	emailcontent=rtrim(request("content"))
	adduser=trim(request("getuser"))
	hidevalue=trim(request("hidevalue"))
	call sendemailsub("add")
	response.write("<p align=""center""><font color=""#dd0000"">窗口2秒钟自动关闭！</font></p>")
	response.write("<script language=""javascript"">")
	response.write("setTimeout('window.close()',2000);")
	response.write("</script>")
	response.end
end if
set conn=opendb("oabusy","conn","accessdsn")
on error resume next
conn.begintrans
sql="update getemailtable set readflag=true where autoid="&id
conn.execute(sql)
if err.number<>0 then
	conn.rollbacktrans
	response.write("<script language=""javascript"">")
	response.write("alert(""更新数据库出错，将关闭窗口！"");")
	response.write("window.close();")
	response.write("</script>")
	response.end
else
	conn.committrans
end if
set rs=server.createobject("adodb.recordset")
sql="select * from userinf,getemailtable where getemailtable.autoid="&id&" and getemailtable.senduserid=userinf.id"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
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
<script language="javascript">
function addgetemailuser()
{
	win=window.open('addgetuser.asp','adduserwin','toolbar=no,scrollbars=yes,resizable=0,menubar=no,width=450,height=440');
}
function checkform()
{
	if (document.form1.getuser.value.length<1)
		{
			alert("收件人不能为空，请按“增加”按钮增加收件人！");
			document.form1.adduser.focus();
			return (false);
		}
	if (document.form1.title.value.length<1)
		{
			alert("邮件标题不能空！");
			document.form1.title.focus();
			return (false);
		}
	return (true);
}
</script>
<title><%=server.htmlencode(rs("name"))%>给您的邮件！</title>
</head>
<body>
<center>
<p><font size="+1" color="#dd0000"><%=server.htmlencode(rs("name"))%></font>给您的邮件</p>
  <table border="1" width="500" cellspacing="0" cellpadding="0" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
    <tr>
      <td bgcolor="D8F7FF" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" align="left" width="488" height=30>
      <font color="#336699"><%=server.htmlencode(rs("emailtitle"))%>（<%=server.htmlencode(rs("name"))%>&nbsp;&nbsp;<%=cstr(rs("emaildate"))%>）</font>	  </td>
    </tr>
  <center>
    <tr>
      <td width="488" align="left">
<%
	content=server.htmlencode(rs("emailcontent"))
	content=replace(content,chr(13)&chr(10),"<br>")
	content=replace(content," ","&nbsp;")
	response.write(content)
%>		
	</td>
    </tr>
  </table>
<br><center><font color="#dd0000" size="+1">回复邮件</font></center>
<form method="post" name="form1" action="lookemail.asp?id=<%=id%>" onsubmit="return checkform();"> 
<table width="500" border="0" cellpadding="0" cellspacing="1" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="6FECFF"> 
<tr> 
<td bgcolor="D8F7FF"> 
收 件 人：
  <input type="text" name="getuser" size="50" value="<%=server.htmlencode(rs("name"))&"|"%>" onfocus="document.form1.title.focus();"><font color=red>*<input type="button" value="增加" name="adduser" onclick="addgetemailuser();"></font>
<input type="hidden" name="hidevalue" value="<%=rs("id")&"|"%>"></td>
</tr>
<tr>
<td bgcolor="D8F7FF">
邮件标题：
  <input type="text" name="title" size=50><font color=red>*</font></td>
</tr>
<tr>
<td bgcolor="D8F7FF">
<center>邮件内容</center><textarea rows="9" name="content" cols="60"></textarea></td>
</tr>
</table>
<br>
<input type="submit" name="submit" value="发送">&nbsp;&nbsp;<input type="button" value="关闭" onclick="window.close()">
</form>
</center>	
</body>
</html>
<%
conn.close
end if
%>
