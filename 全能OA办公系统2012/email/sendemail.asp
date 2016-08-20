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
<body  topmargin="0" leftmargin="0" bgcolor="#F9F9FF">

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
<td>发送邮件</td>
</tr>
</table>
</center>

<br> 
<center>
<%
if request("submit")="发送" then
	emailtitle=rtrim(request("title"))
	emailcontent=rtrim(request("content"))
	adduser=trim(request("getuser"))
	hidevalue=trim(request("hidevalue"))
	'call sendemailsub(emailtitle,emailcontent,adduser,hidevalue)
	call sendemailsub("add")
end if
%>
<form method="post" name="form1" action="sendemail.asp" onsubmit="return checkform();"> 
<table border=0 cellpadding="0" cellspacing="0"> 
<tr>
  <td height="1" bgcolor="B0C8EA"></td>
</tr>
<tr> 
<td height="35"> 
收 件 人：
  <input type="text" name="getuser" size="50" onfocus="document.form1.title.focus();"><font color=red>*<input type="button" value="增加" name="adduser" onclick="addgetemailuser();"></font>
<input type="hidden" name="hidevalue"></td>
</tr>
<tr>
  <td height="1" bgcolor="B0C8EA"></td>
</tr>
<tr>
<td height="35">
邮件标题：
  <input type="text" name="title" size=50><font color=red>*</font></td>
</tr>
<tr>
  <td height="1" bgcolor="B0C8EA"></td>
</tr>
<tr>
<td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="14%">邮件内容：</td>
    <td width="86%"><textarea rows="9" name="content" cols="50"></textarea></td>
  </tr>
</table></td>
</tr>
</table>
<br>
<input type="submit" name="submit" value="发送">
</form>
</center>
<%

%>
</body>
</html>
