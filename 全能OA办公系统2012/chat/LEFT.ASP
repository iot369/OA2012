<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp" -->
<% response.expires=0 %>
<%
if session("user")=application("killname") then
   session("user")=""
   response.redirect "index.asp"
end if

if session("user") = "" then response.redirect "index.asp" 
   for i =1 to 100
       if session("user")=application("user" & i ) then j=i
   next
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
          application("color" & 40)     = "#007700"
          application("showip" & 40)    = session("ip")
          application("mood" & 40)      = ""
          application("to" & 40)        = ""
          application("siliao" & 40)    = ""
          application("zixiao" & 40)    = ""
          application("page" & 40)      = " 悄悄地离开了聊天室！"
          application("tuan" & 40)      = ""
          application("tietu" & 40)     = ""
          application("saytime" & 40)   = time
   application.unlock
set rs=my_conn.execute("select * from "&dbtable_user&" where "&dbfield_user_username&"='" & session("user") & "'")
oldtime=cdate(application("intime"&j))
if application("intime"&j)="" then
   oldtime=now
end if
ti=int(datediff("s",oldtime,now())/60)
if rs(dbfield_user_rate)=session("oldrate") then
   my_conn.execute ("update "&dbtable_user&" set "&dbfield_user_rate&"='" & session("oldrate")+ti & "' where "&dbfield_user_username&"='" & session("user") & "'")
end if
if session("havebook")=1 then my_conn.execute ("delete from "&dbtable_gbook&" where "&dbfield_gbook_name&"='" & session("user") & "'")
my_conn.close
set my_conn=nothing

for i =1 to 100
    if session("user")=application("user" & i ) then
       application("user" & i) = ""
       application("sex" & i) = ""
       application("intime" & i)=""
	   application("userlevel" & i)=""
       application("haveuser" & i)=""
       application("last" & i)=""
    end if 
next
            session("oldrate")=""
            session("ip")=""
			session("id")=""
            session("user")=""
            session("sex")=""
			session("userlevel")=""
%>

<html>
<head>
<title><%=r_title%></title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="Refresh" content="1:url="http://office.11k.net">
</head>

<body bgcolor="#cccccc" text="<%=m_text2%>">
<br><br>
<center>
<strong>
您已离开会议室，欢迎您再次光临!
三秒钟后转到首页,欢迎再来!<br><br>
</strong>
</center>
</body>
</html>