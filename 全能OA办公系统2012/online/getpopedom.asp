
<%
function getpopedom(popedomtype,username)
	dim errinfo
	errinfo=""
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from userinf where username='" &username&"'"
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		errinfo="�Բ���û������û���"
	elseif rs(popedomtype)="no" then
		errinfo="�Բ���������ִ��������ܣ�"
	end if
	set rs=nothing
	conn.close
	set conn=nothing
	getpopedom=errinfo
end function
function getpopedom1(popedomtype,username)
	dim popedomvalue
	popedomvalue=""
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from userinf where username='" &username&"'"
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		popedomvalue="no"
	else
		popedomvalue=trim(rs(popedomtype))
	end if
	set rs=nothing
	conn.close
	set conn=nothing
	getpopedom1=popedomvalue
end function
%>
