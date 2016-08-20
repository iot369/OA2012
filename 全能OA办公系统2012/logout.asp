<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="kq/conn.asp"-->
<%
closeflag=request("closeflag")
response.cookies("oabusyname")=""
response.cookies("oabusyuserid")=""
response.cookies("oabusyusername")=""
response.cookies("oabusyuserdept")=""
response.cookies("oabusyuserlevel")=""
response.cookies("cook_allow_see_all_workrep")=""
response.cookies("cook_allow_see_dept_workrep")=""
response.cookies("cook_allow_control_dept_user")=""
response.cookies("cook_allow_control_all_user")=""
response.cookies("cook_allow_send_note")=""
response.cookies("cook_allow_control_note")=""
response.cookies("cook_allow_control_file")=""
response.cookies("cook_allow_control_level")=""
response.write("<script language=""javascript"">")
if closeflag="0" then
	response.write("opener.top.location.href='default.asp';")
end if
response.write("window.close();")
response.write("</script>")
response.end
%>