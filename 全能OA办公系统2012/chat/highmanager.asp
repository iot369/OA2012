
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
<%if session("userlevel")>8 then%>
<br>
<table width="400" border="0" cellspacing="0" cellpadding="0" align="center">
  <form method="POST" action="highmanager.asp">
    <tr bgcolor="<%=m_bg1%>"> 
      <td height="3"></td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td><br>
        <%=session("user")%>�߼��������<br>
        <br>
      </td>
    </tr>
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="tjgl">
        ��ӹ���Ա�� 
        <input type="text" name="gly" size="12">
        (����) </td>
    </tr>
    
    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="jfgl">
        ���ֹ���:
         <input type="text" name="xm" size="12">
         (����)</td>
    </tr>

<tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="delyh">
        ɾ���û��� 
        <input type="text" name="yhdxm" size="12">
        (����) </td>
    </tr>

<tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="delly">
        ɾ�����ԣ� 
        <input type="text" name="delid" size="12">
        (����������) </td>
    </tr>      

    <tr bgcolor="#cccccc" align="center"> 
      <td> 
        <input type="radio" name="dowhat" value="last" >
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
        <br>
      </td>
    </tr>
<tr bgcolor="#cccccc" align="center">
           
      <td> <font color="#FF0000">�û�����</font><br>
        ����: 
        <select name="dowhat1">
          <option>��ѡ��</option>
          <option value="aaa" >һ��</option>
          <option value="bbb" >����</option>
          <option value="ccc" >����</option>
          <option value="ddd" >�ļ�</option>
          <option value="eee" >�弶</option>
          <option value="fff" >����</option>
          <option value="ggg" >�߼�</option>
          </select>
        &nbsp;����:<input type="text" name="jfxm" size="12">
        ע:�Ӽ��ʿɶ�ֹ</td>
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
       response.write rs(dbfield_user_id)&" "&rs(dbfield_user_username)&" "&rs(dbfield_user_password)&" "&rs(dbfield_user_rate)&" "&rs(dbfield_user_ip)&" "&rs(dbfield_user_lasttime)&"<br>"
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

dowhat2=trim(request("dowhat1"))
if dowhat2="aaa" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=1 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="bbb" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=300 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="ccc" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=600 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="ddd" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=1200 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="eee" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=2400 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="fff" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=3600 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat2="ggg" then
xxmm = trim(request("jfxm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=6000 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat="tjgl" then
zsx = trim(request("gly"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_manager&"=1  where "&dbfield_user_username&"='"&zsx&"'")
response.write "<center>�������</center>"
end if

if dowhat="jfgl" then
xxmm = trim(request("xm"))
my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"=10000 where "&dbfield_user_username&"='"&xxmm&"'")
response.write "<center>�������</center>"
end if

if dowhat="delyh" then
mmxx = trim(request("yhdxm"))
my_conn.execute("delete * from "&dbtable_user&" where "&dbfield_user_username&"='"&mmxx&"'")
response.write "<center>�������</center>"
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
                   <%=rs(dbfield_user_username)%>&nbsp;<%=leve%>�� <a href="moduser.asp?dowhat=mod&id=<%=rs(dbfield_user_id)%>">�޸�</a> <a href="moduser.asp?dowhat=del&id=<%=rs(dbfield_user_id)%>" onclick="return confirm('��ȷ��Ҫɾ������û���������')">ɾ��</a><br><%
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