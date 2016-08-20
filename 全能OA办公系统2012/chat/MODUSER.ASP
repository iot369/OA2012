<!--#INCLUDE FILE="inc_config.asp"-->
<!-- #include file="inc_dbconn.asp" -->
<html>
<head>
<title>修改用户资料</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
body      {  font-size: 10.4pt }
td        {  font-size: 10.4pt }
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>

<body bgcolor="#cccccc">
<%if session("userlevel")>7 then
     id=trim(request("id"))
     dowhat=trim(request("dowhat"))
     if id="" or dowhat="" then %>
               <% response.end
     end if
     set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_id&"="&id)
         if rs.eof then %>
            <br><br><br><center><strong>无此人！</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
            <% response.end
         else
            if rs(dbfield_user_manager)<>0 and session("userlevel")<9 then %>
               <br><br><br><center><strong>无权限！</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
               <% response.end
            end if
         end if
         changeuser=rs(dbfield_user_username)
         rs.close
         set rs=nothing
         if dowhat="del" then
            my_conn.execute("delete from "&dbtable_user&" where "&dbfield_user_id&"="&id) %>
            <br><br><br><center><strong>删除用户成功！</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
            <% response.end
         end if
         if dowhat="mod" then
            set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_id&"="&id)
            leve=1
            if rs(dbfield_user_rate)>level2rate then leve=2
            if rs(dbfield_user_rate)>level3rate then leve=3
            if rs(dbfield_user_rate)>level4rate then leve=4
            if rs(dbfield_user_rate)>level5rate then leve=5
            if rs(dbfield_user_rate)>level6rate then leve=6
            if rs(dbfield_user_rate)>level7rate then leve=7
            if rs(dbfield_user_manager)=1 then leve=8
            if rs(dbfield_user_manager)=2 then leve=9 %>
            <center><form action="moduser.asp" method="post">
            修改用户：<%=rs(dbfield_user_username)%> 的等级<select name="selectrate">
            <option value="1" <%if leve=1 then response.write "selected"%>>1</option>
            <option value="2" <%if leve=2 then response.write "selected"%>>2</option>
            <option value="3" <%if leve=3 then response.write "selected"%>>3</option>
            <option value="4" <%if leve=4 then response.write "selected"%>>4</option>
            <option value="5" <%if leve=5 then response.write "selected"%>>5</option>
            <option value="6" <%if leve=6 then response.write "selected"%>>6</option>
            <option value="7" <%if leve=7 then response.write "selected"%>>7</option>
            <%if session("userlevel")>8 then%>
                 <option value="8" <%if leve=8 then response.write "selected"%>>8</option>
                 <option value="9" <%if leve=9 then response.write "selected"%>>9</option>
            <%end if%>
            </select>
            <input type="hidden" name="dowhat" value="rat">
            <input type="hidden" name="id" value="<%=rs(dbfield_user_id)%>">
            <input type="submit" value="更改" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
            <input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()" onclick="window.close();">
            </form></center>
            <%rs.close
              set rs=nothing
          end if
          if dowhat="rat" then
             selectrate=trim(request("selectrate"))
             if selectrate<>"" and selectrate<8 then
                if selectrate="1" then changerate=1
                if selectrate="2" then changerate=level2rate+1
                if selectrate="3" then changerate=level3rate+1
                if selectrate="4" then changerate=level4rate+1
                if selectrate="5" then changerate=level5rate+1
                if selectrate="6" then changerate=level6rate+1
                if selectrate="7" then changerate=level7rate+1
                my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"="&changerate&" where "&dbfield_user_id&"="&id)
                my_conn.execute("update "&dbtable_user&" set "&dbfield_user_manager&"=0 where "&dbfield_user_id&"="&id)
		        my_conn.execute("insert into "&dbtable_change&" ("&dbfield_change_change&") values ('"&now&" "&session("user")&"修改"&changeuser&"的等级为"&selectrate&"级')")
 %>
                <br><br><br><center><strong>修改用户 <%=changeuser%> 等级为 <%=selectrate%> 成功!</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>				
             <% end if
             if trim(request("selectrate"))=8 and session("userlevel")>8 then
                changerate=level8rate+1
                my_conn.execute("update "&dbtable_user&" set "&dbfield_user_rate&"="&changerate&","&dbfield_user_manager&"=1 where "&dbfield_user_id&"="&id)
		        my_conn.execute("insert into "&dbtable_change&" ("&dbfield_change_change&") values ('"&now&" "&session("user")&"修改"&changeuser&"的等级为8级')") %>
	            <br><br><br><center><strong>修改用户 <%=changeuser%> 等级为 8 成功!</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
             <% end if
             if trim(request("selectrate"))=9 and session("userlevel")>8 then
                my_conn.execute("update "&dbtable_user&" set "&dbfield_user_manager&"=2 where "&dbfield_user_id&"="&id)
		        my_conn.execute("insert into "&dbtable_change&" ("&dbfield_change_change&") values ('"&now&" "&session("user")&"修改"&changeuser&"的等级为9级')") %>
		        <br><br><br><center><strong>修改用户 <%=changeuser%> 等级为 9 成功!</strong><br><br><input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()"></center>
             <% end if
          end if
  end if 
  my_conn.close
  set my_conn=nothing%>
</body>
</html>