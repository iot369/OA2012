
<!--#INCLUDE FILE="inc_config.asp" -->
<!--#INCLUDE FILE="inc_dbconn.asp" -->
<html>
<head>
<title><%=r_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body      { font-size: 10.4pt ;}
td        { font-size: 10.4pt ;}
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>

<body bgcolor="<%=m_bg%>" text="<%=m_text%>">
<%if session("userlevel")>7 then%>
<br>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <form method="POST" action="lowmanager.asp">
    <tr bgcolor="<%=m_bg1%>"> 
      <td height="3"></td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td><br>
        <font color="#0000FF"><%=session("user")%>�������</font><br>
        <br>
      </td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="last">
        ���100�˷��ʼ�¼</td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="kick">
        ���˼�¼</td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="rate">
        �ӷּ�¼ </td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="book">
        �鿴���� </td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="chak">
        ������Ϊ 
        <input type="text" name="ji" size="2">
        �����û�<br>
      </td>
    </tr>
<tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="delly">
        ɾ�����ԣ� 
        <input type="text" name="delid" size="12">
        (����������) </td>
    </tr> 
       
    <tr> 
      <td bgcolor="#cccccc" align="center"> 
        <input type="submit" value="ȷ��" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
        <input type="button" value="�ر�" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()">
      </td>
    </tr>
    <tr bgcolor="<%=m_bg1%>"> 
      <td height="3"></td>
    </tr>
  </form>
</table>
<br>
<%
 dowhat=trim(request("dowhat"))
 if dowhat="last" then
    set rs=my_conn.execute("select * from "&dbtable_user&" order by "&dbfield_user_lasttime&" desc")
    i=1
    do while not rs.eof
       if i=100 then exit do
       response.write  rs(dbfield_user_username)&"  "&rs(dbfield_user_rate)&" "&rs(dbfield_user_ip)&" "&rs(dbfield_user_lasttime)&"<br>"
       i=i+1
    rs.movenext
    loop
    rs.close
    set rs=nothing
	
end if



if dowhat="rate" then
set rs=my_conn.execute("select * from "&dbtable_change&" order by "&dbfield_change_id&" desc")
   if not rs.eof then
      do while not rs.eof
         response.write rs(dbfield_change_id)&" "&rs(dbfield_change_change)&"<br>"
      rs.movenext
      loop
   end if
   rs.close
   set rs=nothing
end if
if dowhat="book" then
set rs=my_conn.execute("select * from "&dbtable_gbook&" order by "&dbfield_gbook_id&" desc")
   if not rs.eof then
      do while not rs.eof
         response.write rs(dbfield_gbook_id)&" "&rs(dbfield_gbook_lyname)&" "&rs(dbfield_gbook_message)&"<br>"
      rs.movenext
      loop
   end if
   rs.close
   set rs=nothing
end if
if dowhat="kick" then
set rs=server.createobject("adodb.recordset")
rs.open "select * from "&dbtable_kill&" order by "&dbfield_kill_id&" desc",my_conn,1,1
   if not rs.eof then
      do while not rs.eof
         response.write rs(dbfield_kill_id)&" "&rs(dbfield_kill_kill)&"<br>"
      rs.movenext
      loop
   end if
   rs.close
   set rs=nothing
end if

if dowhat="delly" then
xxxmmm = trim(request("delid"))
my_conn.execute("delete * from "&dbtable_gbook&" where "&dbfield_gbook_lyname&"='"&xxxmmm&"'")
response.write "<center>�������</center>"
end if

if dowhat="chak" then
   ji=trim(request("ji"))
   if ji="" then ji="1"
   if ji="1" or ji="2" or ji="3" or ji="4" or ji="5" or ji="6" or ji="7" or ji="8" or ji="9" then
      ji=cint(ji)
      set rs=my_conn.execute("select * from "&dbtable_user)
          if not rs.eof then
             do while not rs.eof
                leve=1
                if rs(dbfield_user_rate)>level2rate then leve=2
                if rs(dbfield_user_rate)>level3rate then leve=3
                if rs(dbfield_user_rate)>level4rate then leve=4
                if rs(dbfield_user_rate)>level5rate then leve=5
                if rs(dbfield_user_rate)>level6rate then leve=6
                if rs(dbfield_user_rate)>level7rate then leve=7
                if rs(dbfield_user_manager)=1 then leve=8
                if rs(dbfield_user_manager)=2 then leve=9
                if leve=ji then %>
 <%=rs(dbfield_user_username)%>&nbsp;<%=leve%>�� <a href="moduser.asp?dowhat=mod&id=<%=rs(dbfield_user_id)%>">�޸�</a> <a href="moduser.asp?dowhat=del&id=<%=rs(dbfield_user_id)%>" onclick="return confirm('��ȷ��Ҫɾ������û���������')">ɾ��</a><br>
<%
                end if
                rs.movenext
              loop
          end if
      rs.close
      set rs=nothing
    end if
end if
   my_conn.close
   set my_conn=nothing
%>
<%end if%>


</body>
</html>