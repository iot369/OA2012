
<%@ LANGUAGE = VBScript %>
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/check_resource.asp"-->
<!--#include file="asp/sendeventemail.asp"-->
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
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("alert(""对不起，您已经过期，请重新登录！"");")
	response.write("</script>")
	response.end
end if
checkflag=check_resource_setting(oabusyusername,1)
if checkflag<>0 then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"对不起，您不能发表审核意见！"
	response.end
end if
id=request.form("id")
ideavalue=request.form("R1")
explaintext=trim(request.form("explain"))
if strlength(explaintext)>100 then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"审核意见说明不能超过50个汉字！"
	response.end
end if
if oabusyname="" then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"审核用户名为空！"
	response.end
end if
on error resume next
set conn=opendb("oabusy","conn","accessdsn")
conn.begintrans
set rs=server.createobject("adodb.recordset")
sql="select username,equipment from booking where ID="&id
rs.open sql,conn,1
if rs.eof or rs.bof then
	set rs=nothing
	conn.close
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"对不起，该条预约信息可能已经被删除！"
	response.end
else
	getusername=rs("username")
	getequipment=rs("equipment")
	set rs=nothing
end if
if ideavalue=1 then
	sql="update booking set auditing="&ideavalue&",auditing_explain='"&explaintext&"',auditing_user='"&oabusyname&"',auditing_time=#"&cdate(cstr(date())&" "&cstr(time()))&"# where ID="&id
	conn.execute(sql)
else
	sql="delete from booking where ID="&id
	conn.execute(sql)
end if
if err.number<>0 then
	conn.rollbacktrans
	conn.close
	response.write(err.description)
	response.end
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"对不起，审核资源预约出错！"
	response.end
else
	conn.committrans
	set rs=server.createobject("adodb.recordset")
	sql="select ID from userinf where  username='"&getusername&"'"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		if ideavalue=1 then 
			emailtitle="您好，您的"&getequipment&"预约已由"&oabusyname&"审核！[同意]"
			emailcontent="您好，您的"&getequipment&"预约已由"&oabusyname&"审核！审核意见：[同意]  审核时间：["&cstr(date())&" "&cstr(time())&"]  审核意见说明：["&explaintext&"]"
		elseif ideavalue=2 then
emailtitle="您好，您的"&getequipment&"预约已由"&oabusyname&"审核！[不同意，您的预约已被删除]"
			emailcontent="您好，您的"&getequipment&"预约已由"&oabusyname&"审核！审核意见：[不同意]  审核时间：["&cstr(date())&" "&cstr(time())&"]  审核意见说明：["&explaintext&"]"
		end if
				errstr="对不起，系统自动发送您的审核意见出错，请手动发送邮件通知对方！"
		errinfo=send_event_email(emailtitle,oabusyuserid,rs("ID"),emailcontent,errstr)
		if errinfo<>"" then
			set rs=nothing
			conn.close
			response.redirect "asp/disperrorinfo.asp?errorinfo="&errinfo
			response.end
		end if
	else
		set rs=nothing
		conn.close
		response.redirect "asp/disperrorinfo.asp?errorinfo="&errstr
		response.end
	end if
	conn.close
	response.redirect "booking.asp"
end if
%>