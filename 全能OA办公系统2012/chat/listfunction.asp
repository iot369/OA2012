<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title>动作列表</title>
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
    <td align="center" style="color:<%=m_text1%>">动作列表</td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td style="color:<%=m_text2%>">
    <%    set rs=server.createobject("adodb.recordset")
	  sql="select * from "&dbtable_function
	  set rs=my_conn.execute(sql)
	  if not rs.eof then
	  do while not rs.eof %>
    <a href="#" onclick="viewact('按确定同时提交以下动作指令，按取消不提交\n\n指令: <%=rs("cmd")%>    <%=rs("show")%>\n显示: <%=rs("xiang")%>”\n\n注: var_who表示你自己，var_to表示对方','<%=rs("cmd")%>');" onmouseover="selectact('<%=rs("cmd")%>');" title="<%=rs("xiang")%>"><%=rs("cmd")%></a><br>
    <%    rs.movenext
          loop
          else %>
    还没有动作呢！
    <%    end if
          rs.close
          set rs=nothing
          my_conn.close
          set my_conn=nothing %>
              <center><a href="onlineuser.asp">回在线名单</a></center> 
     </td>
    </tr>
  </table>
</body>
</html> 
