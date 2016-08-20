<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/usercontrol.asp"-->

<%
function strlength(inputstr)
	dim length,i
	length=0
	for i=1 to len(inputstr)
		if asc(mid(inputstr,i,1))<0 then
			length=length+2
		else
			length=length+1
		end if
	next
	strlength=length
end function
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
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='" & oabusyusername&"'"
rs.open sql,conn,1
cook_allow_control_all_user=rs("allow_control_all_user")     
conn.close
set conn=nothing
set rs=nothing
if cook_allow_control_all_user="no" then
response.write("<font color=red size=""+1"">对不起，您没有这个权限！</font>")
	response.end
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
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
                  <td class="style7">用户设置</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br>
<br>
单位名称维护
</center>

<center>
<br>
<%
companyname=trim(request.form("companyname"))
if companyname<>"" then
	if strlength(companyname)>11 then
		response.write("<font color=""#ee0000"" size=""+1"">部门名称最多只能输入5个汉字！</font>")
	else
		set fs=createobject("scripting.filesystemobject")
		set fp=fs.createtextfile(server.mappath("db/companyname.asp"),true)
		fp.writeline(companyname)
		fp.close
		set fp=nothing
		set fs=nothing
	end if
end if
set fs=createobject("scripting.filesystemobject")
set fp=fs.opentextfile(server.mappath("db/companyname.asp"),,true)
if not fp.AtEndOfStream then
	companyname=fp.readline
end if
%>
<form method="post" name="theForm" action="companymanager.asp">
单位名称：<input type="text" name="companyname" maxlength="11" value="<%=companyname%>">
<br><br>
<input type="submit" name="submit" value="确定">
</form>
</center>
<%
fp.close
set fp=nothing
set fs=nothing

%>
</body>
</html>