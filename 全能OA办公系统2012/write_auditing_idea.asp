
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
	response.write("alert(""�Բ������Ѿ����ڣ������µ�¼��"");")
	response.write("</script>")
	response.end
end if
checkflag=check_resource_setting(oabusyusername,1)
if checkflag<>0 then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"�Բ��������ܷ�����������"
	response.end
end if
id=request.form("id")
ideavalue=request.form("R1")
explaintext=trim(request.form("explain"))
if strlength(explaintext)>100 then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"������˵�����ܳ���50�����֣�"
	response.end
end if
if oabusyname="" then
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"����û���Ϊ�գ�"
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
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"�Բ��𣬸���ԤԼ��Ϣ�����Ѿ���ɾ����"
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
	response.redirect "asp/disperrorinfo.asp?errorinfo="&"�Բ��������ԴԤԼ����"
	response.end
else
	conn.committrans
	set rs=server.createobject("adodb.recordset")
	sql="select ID from userinf where  username='"&getusername&"'"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		if ideavalue=1 then 
			emailtitle="���ã�����"&getequipment&"ԤԼ����"&oabusyname&"��ˣ�[ͬ��]"
			emailcontent="���ã�����"&getequipment&"ԤԼ����"&oabusyname&"��ˣ���������[ͬ��]  ���ʱ�䣺["&cstr(date())&" "&cstr(time())&"]  ������˵����["&explaintext&"]"
		elseif ideavalue=2 then
emailtitle="���ã�����"&getequipment&"ԤԼ����"&oabusyname&"��ˣ�[��ͬ�⣬����ԤԼ�ѱ�ɾ��]"
			emailcontent="���ã�����"&getequipment&"ԤԼ����"&oabusyname&"��ˣ���������[��ͬ��]  ���ʱ�䣺["&cstr(date())&" "&cstr(time())&"]  ������˵����["&explaintext&"]"
		end if
				errstr="�Բ���ϵͳ�Զ����������������������ֶ������ʼ�֪ͨ�Է���"
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