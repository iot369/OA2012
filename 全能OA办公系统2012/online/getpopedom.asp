
<%
function getpopedom(popedomtype,username)
	dim errinfo
	errinfo=""
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from userinf where username='" &username&"'"
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		errinfo="对不起，没有这个用户！"
	elseif rs(popedomtype)="no" then
		errinfo="对不起，您不能执行这个功能！"
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
