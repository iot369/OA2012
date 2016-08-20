<script language="javascript">
  function OpenSmallWindows(strURL)
  {
     window.open (strURL,"_blank",
     "status=no,resizable=0,toolbar=no,menubar=no,scrollbars=no,width=300,height=160,left=200,top=150");
  }
</script>
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
DBPath1=server.mappath("../db/jzud-oa.asa")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<%
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if request("yesbutton")="确定" then
	dim namestr,numberstr
	on error resume next
	namestr=""
	numberstr=""
	for i=1 to request.form("addnumber").count
		if request.form("addnumber")(i)<>"" then
			strdim=split(request.form("addnumber")(i),"|")
			numberstr=numberstr&strdim(0)&"|"
			namestr=namestr&strdim(1)&"|"
		end if
	next
	response.write("<script language=""javascript"">")
	response.write("OpenSmallWindows('sendinfo.asp');")
	response.write("</script>")
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select id,name,userdept,userlevel from userinf"
rs.open sql,conn,1
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<title>办公人员列表</title>
</head>

<body bgcolor="#EFEFEF">

<p align="center">办公人员列表</p>
<div align="center">
  <center>
<form method="post" name="form1" action="listall.asp">
  <table border="1" width="389" cellspacing="0" cellpadding="0">
    <tr>
      <td width="128" bgcolor="#C0C0C0" align="center">姓名</td>
      <td width="128" bgcolor="#C0C0C0" align="center">部门</td>
      <td width="129" bgcolor="#C0C0C0" align="center">职务</td>
    </tr>
<%
i=1
do while not rs.eof 
	if 1 then
%>
    <tr>
      <td width="128">
	  	<input type="checkbox" name="addnumber" value="<%=cstr(rs("id"))&"|"&rs("name")%>">
	  	<%=server.htmlencode(rs("name"))%>
	  </td>
      <td width="128"><%=server.htmlencode(rs("userdept"))%></td>
      <td width="129"><%=server.htmlencode(rs("userlevel"))%></td>
    </tr>
<%
	end if
	rs.movenext
	loop
%>
  </table>
  <br>

<input type="button" value="关闭" name="closebutton" onclick="window.close();">
  </form>
  </center>
</div>
<p align="center">
</p>

</body>

</html>
