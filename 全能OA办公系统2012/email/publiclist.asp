
<%
sub listemail(sql,errorstr,recordtype)
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		conn.close
		set rs=nothing
		response.write("<p align=""center""><font color=""#dd0000"">"&errorstr&"</font></p>")
	else
%>
<script language="javascript">
function lookemail(recordid)
{
	win=window.open('lookemail.asp?id='+recordid,'win'+recordid,'toolbar=no,scrollbars=yes,resizable=0,menubar=no,width=550,height=500');	
}
</script>
<p align="center">
共<%=cstr(rs.recordcount)%>条邮件
（<font color="#336699"><img src="../images/newmail.gif" border="0">：新邮件&nbsp;&nbsp;&nbsp;<img src="../images/readmail.gif" border="0">：已读邮件&nbsp;&nbsp;&nbsp;<img src="../images/delmail.gif" border="0">：已删除邮件</font>）
</p>
<div align="center">
  <center>
  <table border="1" width="540" cellspacing="0" cellpadding="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
    <tr bgcolor="D7E8F8">
      <td width="35" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">选择</font></td>
      <td width="34" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">状态</font></td>
      <td width="78" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">发件人</font></td>
      <td width="276" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">主题</font></td>
      <td width="125" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">日期</font></td>
    </tr>
<%
	do while not rs.eof
%>
    <tr bgcolor="#ffffff">
      <td width="35" align="center">
  	  	<input type="checkbox" name="selectnumber" value="<%=cstr(rs("autoid"))%>">
	  </td>
      <td width="35" align="center">
<%
	select case recordtype	
		case "new"
			imgstr="../images/newmail.gif"
			explainstr="新邮件"
		case "newandread"
			if rs("readflag") then
				imgstr="../images/readmail.gif"
				explainstr="已读邮件"
			else
				imgstr="../images/newmail.gif"
				explainstr="新邮件"
			end if
		case "delete"
			imgstr="../images/delmail.gif"
			explainstr="已删除邮件"
	end select
%>
		<img src="<%=imgstr%>" border="0" title="<%=explainstr%>">
	  </td>
      <td width="78" align="center"><%=server.htmlencode(rs("name"))%></td>
      <td width="275" align="center"><a href="#" onclick="javscript:lookemail('<%=cstr(rs("autoid"))%>')"><font color="#336699"><%=server.htmlencode(rs("emailtitle"))%></font></a></td>
      <td width="125" align="center"><%=cstr(rs("emaildate"))%></td>
    </tr>
<%
	rs.movenext
	loop
%>
  </table>
  </center>
</div>
<%
	end if
end sub
%>
<%
'inputstr="add":新发邮件，在sendemailtable中增加一条记录
'inputstr="change":表示转发邮件，修改sendemailtable中的对应记录
sub sendemailsub(inputstr)
	set conn=opendb("oabusy","conn","accessdsn")
	on error resume next
	errorstr=""
	if emailtitle="" then
		errorstr=errorstr&"邮件标题不能为空！"
	end if
	if adduser="" or hidevalue="" then
		errorstr=errorstr&"未选择接收邮件人！"
	end if
	if errorstr<>"" then
		conn.close
		response.write("<script language=""javascript"">")
		response.write("alert("&chr(34)&errorstr&chr(34)&");")
		response.write("history.go(-1);")
		response.write("</script>")
		response.end
	end if
	conn.begintrans
	if inputstr="add" then
		sql="insert into sendemailtable(senduserid,emailtitle,emailcontent,explain,explain1) "
		sql=sql&" values("&oabusyuserid&",'"&emailtitle&"','"&emailcontent&"','"&adduser&"','"&hidevalue&"')"
		conn.execute(sql)
	elseif inputstr="change" then
		sql="update sendemailtable set emailtitle='"&emailtitle&"',emailcontent='"&emailcontent&"',explain='"&adduser&"',explain1='"&hidevalue&"' where autoid="&id
		conn.execute(sql)
	end if
	numberdim=split(request("hidevalue"),"|")
	for i=0 to ubound(numberdim)
		if numberdim(i)<>"" then
			sql="insert into getemailtable (senduserid,getuserid,emailtitle,emailcontent)"
			sql=sql&"  values("&oabusyuserid&","&numberdim(i)&",'"&emailtitle&"','"&emailcontent&"')"
			conn.execute(sql)
		end if
	next
	if err.number<>0 then
		conn.rollbacktrans
		response.write("<script language=""javascript"">")
		response.write("alert(""发送邮件未成功，请返回重试！"");")
		response.write("history.go(-1);")
		response.write("</script>")
		response.end
	else
		conn.committrans
		response.write("<center><font color=""#dd0000"">成功发送邮件！</font><br><br>")
	end if
end sub
%>