
<%
'delflag=0:��getemailtable���е��ʼ�����ϼ���
'delflag=1:��getemailtable���е��ʼ�����ɾ��
'delflag=2:��sendemailtable���е��ʼ�����ɾ��
'delflag=3:�ָ�getemailtable���е��ʼ�
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
		response.write("<center><font color=""#dd0000"">ɾ���ʼ����ִ���</font></center><br>")
	else
		conn.committrans
	end if
end sub
%>