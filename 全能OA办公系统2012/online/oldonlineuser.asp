<%response.expires=0%>
<%
oabusyusername=request.cookies("oabusyusername")
%>
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
DBPath1=server.mappath("../db/jzud-oa.asa")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<%
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select id,username,name,userdept,userlevel from userinf"
rs.open sql,conn,1
%>


<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="expires" content="no-cache">
<title>在线短信</title>
<link rel="stylesheet" type="text/css" href="css/css.css">
<script language="javascript">
function calluser(ip,username)
{
	parent.NetMeeting1.CallTo(ip);
	parent("temp").location.href="sendinfo.asp?username="+username+"&ip="+ip;
}
</script>
<script language="javascript">
  function OpenSmallWindows(strURL)
  {
     window.open (strURL,"_blank",
     "status=no,resizable=0,toolbar=no,menubar=no,scrollbars=no,width=320,height=170,left=200,top=150");
  }
//OpenSmallWindows('asp/sendinfo.asp')
</script>
</head>
<body background="images/bgonline.jpg" leftmargin="1" topmargin="2">
<div align="center">
  <center>
    <p><font color="#FF0000" size="+2"><strong>在线短消息</strong></font></p>
    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">

<%
dim onlineuser
application.lock
onlineuser=application("onlineuser")
number=ubound(onlineuser)
for i=0 to number
	if onlineuser(i)<>"" then
		dimstr=split(onlineuser(i),"$")
		if 1 then
%>
  <tr>
        <td width="50%" height="25" bgcolor="#99CC33"> 
          <input type="button" style="height:20px" name="call" value="在线" onclick="OpenSmallWindows('sendinfo.asp?receiveuser=<%=dimstr(0)%>');">
          <%=dimstr(1)%> </td>

  </tr>
<%
		end if
	end if
next
application.unlock
%>

<%
i=1
do while not rs.eof 
	if 1 then
%>
<%

application.lock
onlineuser=application("onlineuser")
number=ubound(onlineuser)
for i=0 to number
	if onlineuser(i)<>"" then
		dimstr=split(onlineuser(i),"$")
		if dimstr(1)<>rs("name") then
%>

  <tr>
        <td width="50%" height="25" bgcolor="#99CC33"> 
          <input type="button" style="height:20px" name="call" value="离线" onclick="OpenSmallWindows('sendinfo.asp?receiveuser=<%=server.htmlencode(rs("username"))%>');">
          <%=server.htmlencode(rs("name"))%> </td>

  </tr>
<%
		end if
	end if
next
application.unlock
%>
<%
	end if
	rs.movenext
	loop
%>
</table>
<p align="center"><input type="button" name="reload" value="刷新" onclick="location.reload();"></p>
  </center>
</div>
</body>
</html>
