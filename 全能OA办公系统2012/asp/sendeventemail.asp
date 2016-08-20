<%
'向某人发送事件邮件
'errstr:如果自动发送邮件出现错误时输出错误提示
function send_event_email(emailtitle,senduserid,getuserid,emailcontent,errstr)
	dim errorstr
	errorstr=""
	if emailtitle="" then
		errorstr=errorstr&"电子邮件标题为空！<br>"
	end if
	if senduserid="" then
		errorstr=errorstr&"发送用户ID号为空！<br>"
	end if
	if getuserid="" then
		errorstr=errorstr&"接收用户ID号为空！<br>"
	end if
	on error resume next
	set conn1=opendb("oabusy","conn1","accessdsn")	
	conn1.begintrans
	sql="insert into getemailtable (senduserid,getuserid,emailtitle,emailcontent)"
	sql=sql&"  values("&senduserid&","&getuserid&",'"&emailtitle&"','"&emailcontent&"')"
	conn1.execute(sql)
	if err.number<>0 then
		conn1.rollbacktrans
		conn1.close
		set conn1=nothing
		errorstr=errorstr&errstr
	else
		conn1.committrans
		conn1.close
		set conn1=nothing
	end if
	send_event_email=errorstr
end function
%>