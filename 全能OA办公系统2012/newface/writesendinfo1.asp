<%@ Language=vbScript %> 
<!--#include file="../qqprg/function.asp"-->
<script language="javascript" runat="server">
function bm(inputstr)
{
return escape(inputstr);
}
</script>
<%
'on error resume next
response.expiresabsolute=dateadd("s",1,now())
if isempty(session("username")) or session("username")="" then
		application.lock
		application("online")=application("online")+1
		if application("online")>1000 then
			application("online")=1
		end if
		application.unlock
		session("username")="·Ã¿Í"&application("online")
end if
if find_online_user(session("siteid"))=0 then
	call write_online_user(session("siteid"),1)
	response.write("<script language=""javascript"">")
	response.write("parent.refflag.value=1;")
	Response.Write("parent.opener.refflag.value=1;")
	response.write("</script>")
end if
backinfo=Request.QueryString("headinfo")
sendinfo=Request.QueryString("info")
'Response.Write(sendinfo)
if sendinfo<>"" and backinfo<>"" then
	sendinfo=bm(sendinfo)
	application.lock
	writestr=backinfo&"$"&sendinfo&"$"&now()&"|"
	application("info")=application("info")&writestr
	onlineuser=application("onlineuser"&session("siteid"))
	nowtime=now()
	for i=0 to ubound(onlineuser)
			if left(onlineuser(i),instr(onlineuser(i),"$"))=cstr(session.sessionid)&"$" then
				number=instrrev(onlineuser(i),"$")
				frontstr=left(onlineuser(i),number-1)
				number1=instrrev(frontstr,"$")
				frontstr=left(frontstr,number1)
				backstr=right(onlineuser(i),len(onlineuser(i))-number)
				onlineuser(i)=frontstr&nowtime&"$"&backstr
				exit for
			end if
	next
	application("onlineuser"&session("siteid"))=onlineuser
	if session("manager")="1" then
		onlinemanager=application("onlinemanager")
		for i=0 to ubound(onlinemanager)
				if 	left(onlinemanager(i),instr(onlinemanager(i),"$"))=cstr(session.sessionid)&"$" then
					number=instrrev(onlinemanager(i),"$")
					frontstr=left(onlinemanager(i),number-1)
					number1=instrrev(frontstr,"$")
					frontstr=left(frontstr,number1)
					backstr=right(onlinemanager(i),len(onlinemanager(i))-number)
					onlinemanager(i)=frontstr&nowtime&"$"&backstr
					exit for
				end if
		next
		application("onlinemanager")=onlinemanager
	end if
	application.unlock
end if
%>