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
DBPath1=server.mappath("../db/lmtof.mdb")
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
<title>企业内部短信</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css" type="text/css">
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body text="#cccccc"  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
    <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
          <td class="style7">在线用户</td>
        </tr>
    </table></td>
    <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
  </tr>
</table>
<table id=table1 width="100%" border="0" cellspacing="0" cellpadding="0" class="borderon"> 
<tr> 
 <td height="9" style="padding-top: 3px"><nobr>&nbsp;</nobr><nobr><font color="#000000">企业内部短信</font></nobr></td>
</tr> 
</table>
<div style="position:relative; width:100%; height:expression(body.offsetHeight-table1.offsetHeight-2); z-index:1; left: 0px; top: 0px; overflow: auto"> 
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
     "status=no,resizable=0,toolbar=no,menubar=no,scrollbars=no,width=320,height=165,left=200,top=150");
  }
//OpenSmallWindows('asp/sendinfo.asp')
</script>
<div align="center">
<center>
    <div align="left">
      <table id=table1 width="100%" border="0" cellspacing="0" cellpadding="0" class="borderon">
        <tr>
          <td width="100%" style="padding-top: 3px"><font color="#000000">在线用户:</font></td>
        </tr>
      </table>
 </div>
    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#D4D0C8" bordercolordark="#D4D0C8" bgcolor="#C0C0C0">

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
        <td width="50%" height="25" bgcolor="#FFFFFF"> 
          <font color="#666666"> 
          <input type="button" style="height:20px" name="call" value="呼叫" onclick="OpenSmallWindows('sendinfo.asp?receiveuser=<%=dimstr(0)%>');">
          <%=dimstr(1)%> </font> </td> 

  </tr>
<%
		end if
	end if
next
application.unlock
%>
</table>
<div align="center">
<center>
    <div align="left">
      <table id=table1 width="100%" border="0" cellspacing="0" cellpadding="0" class="borderon">
        <tr>
          <td width="100%" style="padding-top: 3px"><font color="#000000">所有用户:</font></td>
        </tr>
      </table>
 </div>
    <table border="1" cellpadding="0" cellspacing="0" width="100%" bordercolorlight="#D4D0C8" bordercolordark="#D4D0C8">
<%
i=1
do while not rs.eof 
	if 1 then
%>
<%

application.lock
onlineuser=application("onlineuser")
number=ubound(onlineuser)
	if onlineuser(i)<>"" then
		dimstr=split(onlineuser(i),"$")
		if 1 then
%>

  <tr>
        <td width="50%" height="25" bgcolor="#FFFFFF"> 
          <font color="#666666"> 
          <input type="button" style="height:20px" name="call" value="留言" onclick="OpenSmallWindows('sendinfo.asp?receiveuser=<%=server.htmlencode(rs("username"))%>');">
          <%=server.htmlencode(rs("name"))%> </font> </td> 

  </tr>
<%
		end if
	end if
application.unlock
%>
<%
	end if
	rs.movenext
	loop
%>
</table>
<p align="center"><input type="button" name="reload" value="刷新在线用户" onclick="location.reload();"></p>
  </center>
</div>
</div>
<p align="center">　
</body>
</html>



