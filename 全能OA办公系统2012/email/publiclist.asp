
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
��<%=cstr(rs.recordcount)%>���ʼ�
��<font color="#336699"><img src="../images/newmail.gif" border="0">�����ʼ�&nbsp;&nbsp;&nbsp;<img src="../images/readmail.gif" border="0">���Ѷ��ʼ�&nbsp;&nbsp;&nbsp;<img src="../images/delmail.gif" border="0">����ɾ���ʼ�</font>��
</p>
<div align="center">
  <center>
  <table border="1" width="540" cellspacing="0" cellpadding="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
    <tr bgcolor="D7E8F8">
      <td width="35" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">ѡ��</font></td>
      <td width="34" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">״̬</font></td>
      <td width="78" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">������</font></td>
      <td width="276" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">����</font></td>
      <td width="125" height="25" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><font color="#2b486a">����</font></td>
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
			explainstr="���ʼ�"
		case "newandread"
			if rs("readflag") then
				imgstr="../images/readmail.gif"
				explainstr="�Ѷ��ʼ�"
			else
				imgstr="../images/newmail.gif"
				explainstr="���ʼ�"
			end if
		case "delete"
			imgstr="../images/delmail.gif"
			explainstr="��ɾ���ʼ�"
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
'inputstr="add":�·��ʼ�����sendemailtable������һ����¼
'inputstr="change":��ʾת���ʼ����޸�sendemailtable�еĶ�Ӧ��¼
sub sendemailsub(inputstr)
	set conn=opendb("oabusy","conn","accessdsn")
	on error resume next
	errorstr=""
	if emailtitle="" then
		errorstr=errorstr&"�ʼ����ⲻ��Ϊ�գ�"
	end if
	if adduser="" or hidevalue="" then
		errorstr=errorstr&"δѡ������ʼ��ˣ�"
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
		response.write("alert(""�����ʼ�δ�ɹ����뷵�����ԣ�"");")
		response.write("history.go(-1);")
		response.write("</script>")
		response.end
	else
		conn.committrans
		response.write("<center><font color=""#dd0000"">�ɹ������ʼ���</font><br><br>")
	end if
end sub
%>