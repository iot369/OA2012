<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<%
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserid=request.cookies("oabusyuserid")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if
'查看是否有新邮件
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select autoid from getemailtable where getuserid="&cstr(oabusyuserid)&" and readflag=false and deleteflag=false"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<Object ID=MyCGI ClassID=CLSID:D45FD31B-5C6E-11D1-9EC1-00C04FD7081F></Object>
<Script> 
//var MerlinID; 
//var MerlinACS; 
//MyCGI.Connected = true; 
//MerlinLoaded = LoadLocalAgent(MerlinID, MerlinACS); 
//Merlin = MyCGI.Characters.Character(MerlinID); 
//Merlin.Show(); 
//Merlin.Speak("您好，您现在有<%=cstr(rs.recordcount)%>封新邮件！"); 
//Merlin.Play("GestureLeft"); 
//Merlin.Speak("请单击邮件连接查看新邮件！"); 
//Merlin.Play("GestureLeft"); 
//Merlin.Speak("谢谢使用，再见！"); 
//Merlin.Play("GestureLeft"); 
//Merlin.Hide(); 
//function LoadLocalAgent(CharID, CharACS) {
//LoadReq = MyCGI.Characters.Load(CharID, CharACS);
//return(true);
//} 
emailwindowvar=window.open('','emailwindow','left=50,top=300,toolbar=no,scrollbars=no,resizable=0,menubar=no,width=152,height=153');
emailwindowvar.location.href="asp/msg_page.asp?info=您好，您现在有<%=cstr(rs.recordcount)%>封新邮件，请尽快查收！&title=新邮件";
</Script>
<%
end if
'查看是否有新的资源预约待审核
auditingflag=request.cookies("allow_check_resource_requirement")
if auditingflag="yes" then
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select ID from booking where auditing=0"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
%>
<script Language="JavaScript">
auditingwindowvar=window.open('','auditingwindow','left=400,top=300,toolbar=no,scrollbars=no,resizable=0,menubar=no,width=152,height=153');
auditingwindowvar.location.href="asp/msg_page.asp?info=您好，有最新资源预约等您审核！&title=资源预约审核";
</script>

<%
end if
set rs=nothing
end if
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="expires" content="no-cache">
<script language="javascript">
setTimeout("location.href='top1.asp'",20000);
</script>
<title>自动刷新页面</title>
</head>
<body style="BACKGROUND-COLOR: transparent">
</body>
</html>
