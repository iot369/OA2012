<%response.expires=0%>
<%
on error resume next
response.cookies("qqdir")("onqqcookie")="yes"
response.cookies("onqqcookie").expires=dateadd("s",15,now())
response.cookies("onqqcookie").path="/qqdir"
response.expiresabsolute=dateadd("s",15,now())
sitenumber=request("sitenumber")
sub lost_user(flag)
	dim delflag,dimsums,onlinesums,num
	Application.Lock
	if flag=1 then
		onlineuser=application("onlineuser"&session("siteid"))
	else
		onlineuser=application("onlinemanager")
	end if
	dimsums=ubound(onlineuser)
	onlinesums=dimsums
	num=0
	for i=0 to dimsums
		delflag=0
		if onlineuser(i)="" then
			delflag=1
		else
			sj=left(onlineuser(i),instrrev(onlineuser(i),"$")-1)
			if isdate(right(sj,len(sj)-instrrev(sj,"$"))) then
				sj=cdate(right(sj,len(sj)-instrrev(sj,"$")))
				if datediff("s",sj,now())>420 then
					delflag=1
				end if
			else
				delflag=1
			end if
		end if
		if delflag=0 then
			if num<i then
				onlineuser(num)=onlineuser(i)
			end if
			num=num+1
		else
			if num<i then
				onlineuser(num)=onlineuser(i)
			end if
			onlinesums=onlinesums-1
		end if
	next
	redim preserve onlineuser(onlinesums)
	if flag=1 then
		application("onlineuser"&session("siteid"))=onlineuser
	else
		application("onlinemanager")=onlineuser
	end if
	Application.UnLock
end sub



'查找当前用户是否在线
function find_online_user(site_id)
	dim i,dimsums,findok,middle
	findok=0
	Application.Lock
	onlineuser=application("onlineuser"&site_id)
	dimsums=ubound(onlineuser)
	if dimsums>=0 then
		middle=fix(dimsums/2)
		for i=0 to middle
			findok=0
			if onlineuser(i)<>"" then
				if instr(onlineuser(i),session.SessionID)>0 then
					findok=1
					exit for
				end if
			end if
			if onlineuser(dimsums-i)<>"" then
				if instr(onlineuser(dimsums-i),session.SessionID)>0 then
					findok=1
					exit for
				end if
			end if
		next
	end if
	Application.UnLock
	find_online_user=findok
end function
%>
<html>
<head>
<title>本站在线用户</title>
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="expires" content="web,26 Feb 1960 08:21:57 GMT">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="style.css" type="text/css">
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
</head>
<body bgcolor="#8482c6" text="#000000" leftmargin="0" topmargin="2">
<%
Response.Write("<script language=""javascript"">")
if datediff("s",application("reftime"&session("siteid")),now())>120 then
	application("reftime"&session("siteid"))=now()
	call lost_user(1)
	call lost_user(0)
end if
siteinfo=get_site_info(sitenumber)
historyinfo=""
if application("info")<>"" then
	application.lock
	msgstr=split(application("info"),"|")
	application("info")=""
	number=ubound(msgstr)
	if number>0 then
		number=number-1
	end if
	for i=0 to number
		msg=split(msgstr(i),"$")
		if isdate(msg(6)) then
			if msg(0)=cstr(session.sessionID) then
				historyinfo=historyinfo&msgstr(i)&"|"
			elseif datediff("s",cdate(msg(6)),now())<120 then
				application("info")=application("info")&msgstr(i)&"|"
			end if
		end if
	next
	application.unlock
end if
siteid=session("siteid")
Application.Lock 
onlinemanager=application("onlinemanager")
Application.UnLock 
number1=ubound(onlinemanager)
onlineuserstr=""
for i=number1 to 0 step -1
	if onlinemanager(i)<>"" then
		onlineuserstr=onlineuserstr&onlinemanager(i)&"|"
	end if
next
if find_online_user(siteid)=1 then
	Response.Write("parent.refflag.value=1;")
	Response.Write("parent.opener.refflag.value=1;")
else
	Response.Write("parent.opener.refflag.value=0;")
	Response.Write("parent.refflag.value=0;")
end if
Response.write("parent.onlineuser.value="&chr(34)&onlineuserstr&chr(34)&";")
Response.Write("parent.getinfo.value="&chr(34)&historyinfo&chr(34)&";")
Response.Write("parent.listflag.value=""1"";")
Response.Write("</script>")
%>
</body>
</html>