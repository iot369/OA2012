<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<html>
<head>
<title>聊天室踢人功能</title>
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

<body bgcolor="#cccccc" text="<%=m_text2%>">
<br>
<%
name=request("name")
if session("user")<>"" and session("userlevel")>6 and name<>"" then
   if name=session("user") then
      response.write "<br><center>你不能踢自已呀!<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'></center></body></html>"
      response.end
   end if
   for i=1 to 100
       if application("user" & i)=name then
          set rs=my_conn.execute("select * from "&dbtable_user&" where username ='" & name & "'")
              uleve=1
              if rs(dbfield_user_rate)>level2rate then uleve=2
              if rs(dbfield_user_rate)>level3rate then uleve=3
              if rs(dbfield_user_rate)>level4rate then uleve=4
              if rs(dbfield_user_rate)>level5rate then uleve=5
              if rs(dbfield_user_rate)>level6rate then uleve=6
              if rs(dbfield_user_rate)>level7rate then uleve=7
              if rs(dbfield_user_rate)>level8rate and rs(dbfield_user_manager)=1 then uleve=8
              if rs(dbfield_user_manager)=2 then uleve=9
              if session("userlevel")<=uleve then
                 response.write "<br><center>权限不够!<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'></center></body></html>"
                 response.end
              end if
              application("user" & i)=""
          exit for
       end if
   next
rs.close
set rs=nothing


sz = "'" & now &" " & request("name") &" 被 " & session("user") & " 踢出，因为 "&request("why")& "'"  
my_conn.execute("insert into "&dbtable_kill&" ("&dbfield_kill_kill&") values (" & sz & ")")
my_conn.close
set my_conn=nothing

application("killname")=name
if request("lo")="l" then application("iskillip")="yes"
application.lock
for i =1 to 39
	      application("systemsay" & i)  = application("systemsay" & i+1)
          application("sexsay" & i)     = application("sexsay" &i+1)
          application("me" & i)         = application("me" & i+1)
          application("color" & i)      = application("color" & i+1)
          application("showip" & i)     = application("showip" & i+1)
          application("mood" & i)       = application("mood" & i+1)
          application("to" & i)         = application("to" & i+1)
          application("siliao" & i)     = application("siliao" & i+1)
          application("zixiao" & i)     = application("zixiao" & i+1)
          application("page" & i)       = application("page" & i+1)
          application("tuan" & i)       = application("tuan" & i+1)
          application("tietu" & i)      = application("tietu" & i+1)
          application("saytime" & i)    = application("saytime" & i+1)
next
		  application("systemsay" & 40) = "yes"
          application("me" & 40)        = session("user")
          application("sexsay" & 40)    = session("sex")
          application("color" & 40)     = "#FF0000"
          application("showip" & 40)    = session("ip")
          application("mood" & 40)      = ""
          application("to" & 40)        = ""
          application("siliao" & 40)    = ""
          application("zixiao" & 40)    = ""
		  if trim(request("why"))<>"" then
          application("page" & 40)      = " 一脚将 "&application("killname")&" 踢出聊天室，因为“"&request("why")&"”。"
          else
          application("page" & 40)      = " 一脚将 "&application("killname")&" 踢出聊天室，因为“不需要理由”。"
          end if
          application("tuan" & 40)      = ""
          application("tietu" & 40)     = ""
          application("saytime" & 40)   = time
application.unlock
response.write "<br><center>"& name & "被踢出聊天室<br><br><input type='button' value=' 返 回 ' class='normal' onmouseover='this.className="&chr(34)&"over"&chr(34)&"' onmousedown='this.className="&chr(34)&"down"&chr(34)&"' onmouseout='this.className="&chr(34)&"normal"&chr(34)&"' onclick='javascript:history.go(-1)'></center></body></html>"
response.end 
end if %>

<%if session("user")<>"" and session("userlevel")>6 then%>
<table border="1" align="center" cellpadding="3" cellspacing="0" width="300">
  <form method="POST" action="kill.asp">
    <tr> 
     <td align="center" bgcolor="<%=m_bg1%>" style="color:<%=m_text1%>">我 要 踢 人</td>
    </tr>
    <tr>
      <td style="padding-left:20pt">踢谁： 
        <select name="name">
     <% for i=1 to 100
        if application("user" & i) <> "" then %> 
     <option value='<% =application("user" & i) %>'> <% =application("user" & i) %> 
     <% end if
        next %> 
     </select>
        <br>
        <% if session("userlevel")>7 then %> 踢法：
        <input type="radio" name="lo" value="o" checked>不踢IP
     <input type="radio" name="lo" value="l">踢IP
     <% end if %> 
     <br>原因： <input type="text" name="why" size="20" maxlength=100>
     <br><input type="submit" value="  踢  " class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'">
     <input type="button" value="关  闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onClick="window.close()">
     </td>
    </tr>
   </form>
  </table>
  <%end if%>
 </body>
</html>