<!--#INCLUDE FILE="inc_config.asp"-->
<% response.expires=0 %>
<html>
<head>
<title>在线人员列表</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv=refresh content='18;url=onlineuser.asp'>
<style>
<!--
body      { font-size: 9pt; background-color: "#cccccc" }
td        { font-size: 9pt }
A:link    { text-decoration: none;color: #000000 }
A:visited { text-decoration: none;color: #FFFFFF }
A:active  { text-decoration: none;color: #FFFFFF }
A:hover   { text-decoration: none;color: #FFFFFF }
-->
</style>
<script language="JavaScript">
<!--
 function selectwho(list)
  { parent.frm_input.document.forms[0].towho.text=list;
    parent.frm_input.document.forms[0].towho.value=list;
    parent.frm_input.document.forms[0].saystemp.focus();
    parent.overselectenable=false; }
-->
</script>
</head>

<body>

<table bgcolor="<%=m_bg3%>" cellpadding="5" cellspacing="1" width="100%" align="right">
  <tr bgcolor="<%=m_bg1%>"> 
    <td align="center" style="color:<%=m_text1%>">在线名单</td>
  </tr>
  <tr bgcolor="#cccccc"> 
    <td style="color:<%=m_text2%>"><% if session("user")="" then %> 
           <hr><p align="center">你已离开<br>请<a href="index.asp" target="_top">重新进入</a></p><hr>
           <%
           response.end
		   end if %>
           <a href="javascript:selectwho('大家');" title="选取大家作为谈话对象">大家</a><br>
           <% men=0
              for i=1 to 100
		      if application("user"&i)<>"" then %>
                 <a href="javascript:selectwho('<%=application("user"&i)%>');" title="选取 <%=application("user"&i)%> 作为谈话对象">
                 <%=application("sex"&i)%> <%=application("user"&i)%>
<%if session("userlevel")>2 then%>[<%=application("userlevel"&i)%>]
<%end if%></a><br>
                 <% men=men+1 
              end if
			  next %> 
              <center>共 <% =men %> 人</center> 
     </td>
    </tr>
  </table>
</body>
</html> 
