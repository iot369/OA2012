<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<%
on error resume next
oabusyusername=request.cookies("oabusyusername")
if oabusyusername="" then
	response.write("<script language=""javascript"">")
	response.write("window.open('lostuser.asp','lostwin','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');")	
	response.write("parent.window.close();")
	response.write("</script>")
	response.end
end if
sitenumber=1
sub lost_user()
	dim delflag,dimsums,onlinesums,num
	Application.Lock
	onlineuser=application("onlineuser"&session("siteid"))
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
				if datediff("s",sj,now())>3600 then
					delflag=1
				end if
			else
				delflag=1
			end if
		end if
		if num<i then
			onlineuser(num)=onlineuser(i)
		end if
		num=num+1
	next
	redim preserve onlineuser(onlinesums)
	application("onlineuser"&session("siteid"))=onlineuser
	Application.UnLock
end sub

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
<body bgcolor="#0099FF" text="#000000" leftmargin="0" topmargin="2">
<%
Response.Write("<script language=""javascript"">")
if datediff("s",application("reftime"&session("siteid")),now())>120 then
	application("reftime"&session("siteid"))=now()
	call lost_user()
end if
'siteinfo=get_site_info(sitenumber)
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
onlineuser=application("onlineuser"&siteid)
Application.UnLock 
number1=ubound(onlineuser)
onlineuserstr=""
for i=number1 to 0 step -1
	if onlineuser(i)<>"" then
		onlineuserstr=onlineuserstr&onlineuser(i)&"|"
	end if
next
Response.write("parent.onlineuser.value="&chr(34)&onlineuserstr&chr(34)&";")
Response.Write("parent.getinfo.value="&chr(34)&historyinfo&chr(34)&";")
Response.Write("parent.listflag.value=""1"";")
Response.Write("</script>")
%>
</body>
</html>
