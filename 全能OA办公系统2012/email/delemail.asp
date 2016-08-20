
<%
'delflag=0:将getemailtable表中的邮件放入废件箱
'delflag=1:将getemailtable表中的邮件彻底删除
'delflag=2:将sendemailtable表中的邮件彻底删除
'delflag=3:恢复getemailtable表中的邮件
sub delemail(delflag)
	set conn=opendb("oabusy","conn","accessdsn")
	on error resume next
	err.clear
	conn.begintrans
	for i=1 to request.form("selectnumber").count
		sql=""
		if request.form("selectnumber")(i)<>"" then
			select case delflag
				case 0
					sql="update getemailtable set deleteflag=true where autoid="&cstr(request.form("selectnumber")(i))
					conn.execute(sql)
				case 1
					sql="delete from getemailtable where autoid="&cstr(request.form("selectnumber")(i))
					conn.execute(sql)
				case 2
					sql="delete from sendemailtable where autoid="&cstr(request.form("selectnumber")(i))
					conn.execute(sql)
				case 3
					sql="update getemailtable set deleteflag=false where autoid="&cstr(request.form("selectnumber")(i))
					conn.execute(sql)
			end select
		end if
	next
	if err.number<>0 then
		conn.rollbacktrans
		response.write("<center><font color=""#dd0000"">删除邮件出现错误！</font></center><br>")
	else
		conn.committrans
	end if
end sub
%>