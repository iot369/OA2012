<%
'��ĳ�˷����¼��ʼ�
'errstr:����Զ������ʼ����ִ���ʱ���������ʾ
function send_event_email(emailtitle,senduserid,getuserid,emailcontent,errstr)
	dim errorstr
	errorstr=""
	if emailtitle="" then
		errorstr=errorstr&"�����ʼ�����Ϊ�գ�<br>"
	end if
	if senduserid="" then
		errorstr=errorstr&"�����û�ID��Ϊ�գ�<br>"
	end if
	if getuserid="" then
		errorstr=errorstr&"�����û�ID��Ϊ�գ�<br>"
	end if
	on error resume next
	set conn1=opendb("oabusy","conn1","accessdsn")	
	conn1.begintrans
	sql="insert into getemailtable (senduserid,getuserid,emailtitle,emailcontent)"
	sql=sql&"  values("&senduserid&","&getuserid&","&sqlstr(emailtitle)&","&sqlstr(emailcontent)&")"
	conn1.execute(sql)
	if err.number<>0 then
		conn1.rollbacktrans
		conn1.close
		set conn1=nothing
		errorstr=errorstr&errstr&err.description
	else
		conn1.committrans
		conn1.close
		set conn1=nothing
	end if
	send_event_email=errorstr
end function
%>
