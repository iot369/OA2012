<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title>�����б�</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style>
<!--
body      { font-size: 9pt; background-color: <%=m_bg2%> }
td        { font-size: 9pt }
A:link    { text-decoration: none;color: #000000 }
A:visited { text-decoration: none;color: #FFFFFF }
A:active  { text-decoration: none;color: #FFFFFF }
A:hover   { text-decoration: none;color: #FFFFFF }
-->
</style>
<script language="JavaScript">
<!--
function selectact(list){
	parent.frm_input.document.forms[0].saystemp.value=list;
	parent.frm_input.document.forms[0].saystemp.focus();
	return true;
}
function viewact(list,act){
     if (window.confirm(list)) {
	parent.frm_input.document.forms[0].saystemp.value=act;
	parent.frm_input.document.forms[0].saystemp.focus();
	parent.frm_input.document.forms[0].sub.click();
	return true;
     }
}

//-->
</script>
</head>

<body>
  <table bgcolor="<%=m_bg3%>" cellpadding="5" cellspacing="1" width="100%" align="right">
  <tr bgcolor="<%=m_bg1%>"> 
    <td align="center" style="color:<%=m_text1%>">�����б�</td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td style="color:<%=m_text2%>">
    <%    set rs=server.createobject("adodb.recordset")
	  sql="select * from "&dbtable_function
	  set rs=my_conn.execute(sql)
	  if not rs.eof then
	  do while not rs.eof %>
    <a href="#" onclick="viewact('��ȷ��ͬʱ�ύ���¶���ָ���ȡ�����ύ\n\nָ��: <%=rs("cmd")%>    <%=rs("show")%>\n��ʾ: <%=rs("xiang")%>��\n\nע: var_who��ʾ���Լ���var_to��ʾ�Է�','<%=rs("cmd")%>');" onmouseover="selectact('<%=rs("cmd")%>');" title="<%=rs("xiang")%>"><%=rs("cmd")%></a><br>
    <%    rs.movenext
          loop
          else %>
    ��û�ж����أ�
    <%    end if
          rs.close
          set rs=nothing
          my_conn.close
          set my_conn=nothing %>
              <center><a href="onlineuser.asp">����������</a></center> 
     </td>
    </tr>
  </table>
</body>
</html> 
