<!--#include file="sqlstr.asp"-->
<!--#include file="checked.asp"-->
<!--#include file="opendb.asp"-->
<%
sub usercontrol(href)
	On Error Resume Next
	oabusyuserdept=request.cookies("oabusyuserdept")
	if request.form("detel")="ɾ��" and request.form("delid")<>"" then
		set conn=opendb("oabusy","conn","accessdsn")
		conn.begintrans
		count=0
		condition=""
		condition1=""
		for each idno in request.form("delid")
			count=count+1
			condition=condition+"username=" & sqlstr(idno)
			condition1=condition1+"Doc_UserName="&sqlstr(idno)
			if count<request.form("delid").count then
				condition=condition+" or "
				condition1=condition1+" or "
			end if
		next
		'ɾ�����ݿ��еļ�¼
		sql = "delete * from userinf where " & condition
		conn.Execute(sql)
		if err.number<>0 then
			conn.rollbacktrans
			call DispErrorInfo1("�Բ���ɾ���û���������ԭ��"&err.description&"--"&keepformat(sql)&"--"&keepformat(sql1))
			conn.close
			set conn=nothing
			call bgsub()
			response.end
		else
			conn.committrans
		end if
	end if
	if request.form("forbid")="����" then
		set conn=opendb("oabusy","conn","accessdsn")
		conn.beginTrans
		count=0
		count1=0
		allid=""
		condition=""
		notcondition=""
		for each id in request.form("allid")
			count1=count1+1
			allid=allid & id
			if count1<request.form("allid").count then
				allid=allid+","
			end if
		next
		for each idno in request("forbidid")
			count=count+1
			condition=condition+"id=" & idno
			notcondition=notcondition+"id<>" & idno
			if count<request("forbidid").count then
				condition=condition+" or "
				notcondition=notcondition+" and "
			end if
		next
		if condition<>"" then
			sql = "update userinf set forbid='yes' where " & condition
			conn.Execute sql
		end if
		if notcondition<>"" then
			sql = "update userinf set forbid='no' where id in (" & allid & ") and " & notcondition
		else
			sql = "update userinf set forbid='no' where id in (" & allid & ")"
		end if
		conn.Execute sql
		if err.number<>0 then
			conn.rollbacktrans
			call DispErrorInfo1("�Բ���ɾ���û���������ԭ��"&err.description&"--"&keepformat(sql)&"--"&keepformat(sql1))
			conn.close
			set conn=nothing
			call bgsub()
			response.end
		else
			conn.committrans
		end if
	end if
	'�����ݿ���ʾ������oabustuserdept���û�
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from userinf"
	rs.open sql,conn,1
%>
<form action="<%=href%>" method=post>
<table border="1" cellspacing="0" cellpadding="5" width="90%" bordercolorlight="#808080" bordercolordark="#D6D5CB" bgcolor="#ffffff">
<tr bgcolor="#eeeeee">
<td width="30" align="center" bgcolor="#D4D0C8"><input type="submit" value="ɾ��" name="detel"></td>
<td align=center bgcolor="#D4D0C8">����</td>
<td align=center bgcolor="#D4D0C8">�û���</td>
<td align=center bgcolor="#D4D0C8">����</td>
<td align=center bgcolor="#D4D0C8">����</td>
<td align=center bgcolor="#D4D0C8">����</td>
<td width="30" align="center" bgcolor="#D4D0C8"><input type="submit" value="����" name="forbid"></td>
</tr>
<%
while not rs.eof and not rs.bof
%>
<tr>
<td align="center"><input type="checkbox" name="delid" value="<%=rs("username")%>"></td>
<td align="center"><a href="edituserinf.asp?username=<%=rs("username")%>"><%=keepformat(rs("name"))%></a></td>
<td align="center"><%=keepformat(rs("username"))%></td>
<td align="center"><%=keepformat(rs("password"))%></td>
<td align="center"><%=keepformat(rs("userdept"))%></td>
<td align="center"><%=keepformat(rs("userlevel"))%></td>
<td align="center"><input type="checkbox" name="forbidid" value="<%=rs("id")%>"<%=checked(rs("forbid"),"yes")%>><input type="hidden" name="allid" value="<%=rs("id")%>">
</td>
</tr>
<%
	rs.movenext
wend
%>
</table>
</form>
<%
end sub
%>
