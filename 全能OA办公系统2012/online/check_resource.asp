<%
'取得用户资源管理的权限
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
'发表审核意见表单
sub writeidea(doprogram,auditingname,hideid)
%>
<p align="center"><b>审核意见</b></p>
<form method="POST" action="<%=doprogram%>" name="ideaform">
  <div align="center">
    <center>
    <table border="1" cellpadding="5" cellspacing="0" width="90%" bordercolorlight="#808080" bordercolordark="#D4D0C8">
      <tr>
        <td height="25" bgcolor="#D4D0C8"><b>审核人：</b><%=auditingname%></td>
      </tr>
      <tr>
        <td height="25"><b>审核意见：</b><input type="radio" value="1" checked name="R1">同意<input type="radio" name="R1" value="2">不同意</td>
      </tr>
      <tr>
        <td height="25"><b>审核意见说明：</b><br>
          <textarea rows="5" name="explain" cols="60"></textarea>
		  <input type="hidden" name="id" value="<%=hideid%>"> 
		  </td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="提交审核意见" name="ok"></p>
</form>
<%
end sub
%>
