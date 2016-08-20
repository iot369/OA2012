<%@ LANGUAGE = VBScript %>
<%
set file=server.createobject("scripting.FileSystemObject")
addr1=server.mappath("top1.asp")
addr2=server.mappath("top1.asp")
If Not file.FileExists(addr1) or Not file.FileExists(addr2) Then
response.write "<script LANGUAGE='javascript'>alert('系统发生严重错误即将关闭！！！');window.close();</script>"
End If
%>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="kq/conn.asp"-->
<%
function find_online_user(oabusyusername)
	dim i,dimsums,siteinfo,findok
	findok=0
	Application.Lock
	onlineuser=application("onlineuser")
	dimsums=ubound(onlineuser)
	for i=0 to dimsums
		siteinfo=onlineuser(i)
		if siteinfo<>"" then
			if instr(siteinfo,oabusyusername&"$")>0 then
				findok=1
				exit for
			end if
		end if
	next
	Application.UnLock
	find_online_user=findok
end function
sub write_online_user(username,name,dept)
	dim userinfo
	Application.Lock
	userinfo=username&"$"&name&"$"&dept&"$"&request.servervariables("REMOTE_ADDR")
	onlineuser=application("onlineuser")
	dimsums=ubound(onlineuser)
	redim preserve onlineuser(dimsums+1)
	onlineuser(dimsums+1)=userinfo
	application("onlineuser")=onlineuser
	Application.UnLock
end sub
sub checkkqdatabase()
	set fileobject=server.createobject("Scripting.FileSystemObject")
	if not fileobject.FileExists(server.mappath("kq/"&cstr(year(date()))&".mdb")) then
		file1=server.mappath("kq\backup\new.mdb")
		file2=server.mappath("kq\"&cstr(year(date()))&".mdb")
		fileobject.copyfile file1,file2
	end if
	set fileobject=nothing
end sub

call checkkqdatabase
username=request.form("username")
password=request.form("password")
if username<>"" and password<>"" then
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
on error resume next
sql="select * from userinf where username=" & sqlstr(username) & " and password=" & sqlstr(password) & " and forbid='no'"
rs.open sql,conn,1,1
'如果有此用户就进入系统
if not rs.eof and not rs.bof then
	response.cookies("oabusyname")=rs("name")
	response.cookies("oabusyuserid")=cstr(rs("ID"))
	response.cookies("oabusyusername")=rs("username")
	response.cookies("oabusyuserdept")=rs("userdept")
	response.cookies("oabusyuserlevel")=rs("userlevel")

	response.cookies("cook_allow_see_all_workrep")=rs("allow_see_all_workrep")
	response.cookies("cook_allow_see_dept_workrep")=rs("allow_see_dept_workrep")

	response.cookies("cook_allow_control_dept_user")=rs("allow_control_dept_user")
	response.cookies("cook_allow_control_all_user")=rs("allow_control_all_user")

	response.cookies("cook_allow_send_note")=rs("allow_send_note")
	response.cookies("cook_allow_control_note")=rs("allow_control_note")

	response.cookies("cook_allow_control_file")=rs("allow_control_file")
	response.cookies("cook_allow_send_file")=rs("allow_send_file")
	response.cookies("allow_transmit_file")=rs("allow_transmit_file")

	response.cookies("cook_allow_control_level")=rs("allow_control_level")
	response.cookies("allow_check_resource_requirement")=rs("allow_check_resource_requirement")
	response.cookies("allow_auditing_workthings")=rs("allow_auditing_workthings")
	response.cookies("allow_manage_workthings")=rs("allow_manage_workthings")
	response.cookies("allow_lookallinfo_workthings")=rs("allow_lookallinfo_workthings")
	response.cookies("allow_look_all_kq_info")=rs("allow_look_all_kq_info")
	response.cookies("allow_edit_help")=rs("allow_edit_help")
			
	application.lock
	onlineuserdim=application("onlineuser")
	if isempty(onlineuserdim) then
			dim onlineuserdim(0)
			dim netmeetinginfodim(0)
			application("onlineuser")=onlineuserdim
			application("netmeetinginfo")=netmeetinginfodim
	end if
	application.unlock
	if find_online_user(rs("username"))=0 then
		call write_online_user(rs("username"),rs("name"),rs("userdept"))
	end if
	conn.close
	set conn=nothing
	response.redirect "main.asp"
	response.end
end if
end if
%>
<html>
<head>
<title>全能通用OA办公系统</title>
<META HTTP-EQUIV="Content-Type" CONTENT="text/html; charset=gb2312"> 
<link rel="stylesheet" href="inc/style.css" type="text/css">
<style type="text/css">
<!--
.sr {
	height: 17px;
	width: 140px;
}
.style1 {color: #3c4d82}
body {
	background-image: url(images/bg_index.gif);
}
-->
</style>
</head>
<body bgcolor="menu" text="#000000" leftmargin="1" topmargin="1" scroll="no" style="border:0px;">
<table width="100%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="50">&nbsp;</td>
  </tr>
  <tr>
    <td><table width="640" height="447"  border="0" align="center" cellpadding="0" cellspacing="0" background="images/login_bg.gif">
      <tr>
        <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td width="65%">&nbsp;</td>
            <td width="35%" height="145">&nbsp;</td>
          </tr>
          <tr>
            <td>&nbsp;</td>
            <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <form name="form1" method="post" action="default.asp">
			  <tr>
                <td>&nbsp;</td>
                <td valign="bottom"><span class="style1">请输入用户名</span></td>
			  </tr>
              <tr>
                <td>&nbsp;</td>
                <td><input name="username" type="text" class="sr" size="15" maxlength="50"></td>
              </tr>
              <tr>
                <td>&nbsp;</td>
                <td height="25" valign="bottom"><span class="style1">请输入密码</span></td>
              </tr>
              <tr>
                <td width="7">&nbsp;</td>
                <td><input name="password" type="password" class="sr" size="15" maxlength="50"></td>
              </tr>
              <tr>
                <td height="50" colspan="2"><INPUT name="S1" type=image
            src="images/bt_login.gif" width="98" height="37" ></td>
              </tr>
			   </form>
            </table></td>
          </tr>
        </table></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
<div align="center"><script src="count/mystat.asp"></script></div>
</body>
</html>
