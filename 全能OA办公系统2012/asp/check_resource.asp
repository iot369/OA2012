<%
'ȡ���û���Դ�����Ȩ��
function check_resource_setting(inputusername,checkflag)
	dim resultvalue
	resultvalue=0
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select allow_add_resource,allow_check_resource_requirement from userinf where  username='"&inputusername&"'"
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		conn.close
		set rs=nothing
		resultvalue=1
	else
		select case checkflag
			case 0
				if rs("allow_add_resource")<>"yes" then
					conn.close
					set rs=nothing
					resultvalue=1
				end if
			case 1
				if rs("allow_check_resource_requirement")<>"yes" then
					conn.close
					set rs=nothing
					resultvalue=1
				end if
		end select
	end if
	check_resource_setting=resultvalue
end function
'������������
sub writeidea(doprogram,auditingname,hideid)
%>
<p align="center">������</p>
<form method="POST" action="<%=doprogram%>" name="ideaform">
  <div align="center">
    <center>
    <table width="540" border="0" cellpadding="0" cellspacing="1" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="B0C8EA">
      <tr>
        <td height="25" bgcolor="#FFFFFF">����ˣ�<%=auditingname%></td>
      </tr>
      <tr>
        <td height="25" bgcolor="#FFFFFF">��������
          <input type="radio" value="1" checked name="R1">ͬ��<input type="radio" name="R1" value="2">��ͬ��</td>
      </tr>
      <tr>
        <td height="25" bgcolor="#FFFFFF">������˵����<br>
          <textarea rows="5" name="explain" cols="60"></textarea>
		  <input type="hidden" name="id" value="<%=hideid%>">		  </td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="�ύ������" name="ok"></p>
</form>
<%
end sub
%>