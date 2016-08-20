<!--#INCLUDE FILE="inc_config.asp"-->
<!--#INCLUDE FILE="inc_dbconn.asp"-->
<%
response.expires=0
if session("user")<>"" and application("killname")=session("user") then
   if application("iskillip")="yes" then
      application("killip")=session("ip")
      response.cookies("killtime")=now
      response.cookies("killtime").expires= date + 30
      application("iskillip")=""
   end if %>
   	  <html>
      <head>
      <title>错误</title>
      </head>
      <body bgcolor="#cccccc" text="<%=m_text2%>">
	  <Script LANGUAGE="JavaScript">
	  <!--
        setTimeout('window.parent.close();',3000);
      -->
	  </script>
      <br><br><br><center><strong>你被踢出聊天室！三十分钟后再进来!</strong></center>
      </body>
      </html>
   <% response.end
end if

if application("user"&session("id"))<>session("user") then
'   response.write application("user"&i)&"<br>"
'   response.write session("id")&"<br>"
'   response.write session("user")&"<br>" %>
   	  <html>
      <head>
      <title>错误</title>
      </head>
      <body bgcolor="#cccccc" text="<%=m_text2%>">
      <br><br><br><center><strong>由于聊天室无法验证您的身份，请按刷新进行再次验证!</strong></center>
      </body>
      </html>
   <% response.end
end if %>

<html>
<head>
<title>聊天室显示区</title>
</head>
<meta http-equiv=refresh content='10;url=content.asp<%if r_refresh="down" then%>#dian<%end if%>'>
<style type="text/css">
<!--
body      { font-size: 10.4pt ; line-height:16pt ; background-color: <%=m_bg%>}
A:link    { text-decoration: none ; color: #0000ff}
A:visited { text-decoration: none ; color: #0000ff}
A:active  { text-decoration: underline ; color: #0088ff}
A:hover   { text-decoration: underline ; color: #0088ff}
-->
</style>
<script language="JavaScript">
<!--
function selectwho(list)
{
parent.frm_input.document.forms[0].towho.text=list;
parent.frm_input.document.forms[0].towho.value=list;
parent.frm_input.document.forms[0].saystemp.focus();
parent.overselectenable=false;  
}
//-->
</script>
</head>

<body topmargin="20">
<%
if trim(request("says"))<>"" and session("user")<>"" then
   if session("userlevel")<7 then
      inpm=Server.HTMLEncode(trim(request("says")))
   else
      inpm=trim(request("says"))
   end if
   if trim(request("fs"))<>"" and session("userlevel")>4 then zixiao=trim(request("fs"))
   if trim(request("addsign"))<>"" and session("userlevel")>1 then tuan=Server.HTMLEncode(trim(request("addsign")))
   if trim(request("addpic"))<>"" and session("userlevel")>2 then tietu=Server.HTMLEncode(trim(request("addpic")))
   
   if left(inpm,1)="/" and replace(inpm," ","")=inpm and len(inpm)<10 and session("userlevel")>3 then
      set rs=server.createobject("adodb.recordset")
	  sql="select * from "&dbtable_function&" where "&dbfield_function_cmd&"='"&inpm&"'"
	  set rs=my_conn.execute(sql)
	  if not rs.eof then
	     inpm=rs(dbfield_function_xiang)
         systemsay="function"
	 end if
	 rs.close
	 set rs=nothing
	 my_conn.close
	 set my_conn=nothing
   end if
   
   if tietu<>"" then
      tietu="<img src='images/"&tietu&"'>"
   else
      tietu=""
   end if
   if zixiao="1" then inpm=""&inpm&""
   if zixiao="2" then inpm="<i>"&inpm&"</i>"
   if zixiao="3" then inpm="<font size=-1>"&inpm&"</font>"
   if zixiao="4" then inpm="<big><big>"&inpm&"</big></big>"
   if zixiao="5" then inpm="<marquee border='0' width='70%'>"&inpm&"</marquee>"
   if zixiao="6" then inpm="<marquee direction='right' border='0' width='70%'>"&inpm&"</marquee>"
   if zixiao="7" then inpm="<marquee behavior='alternate' border='0' width='70%'>"&inpm&"</marquee>"
   if zixiao="8" then inpm="<u>"&inpm&"</u>"
   
   'addsing=trim(request("addsign"))           '##图案
   'addsays=trim(request("addsays"))           '##表情
   'addpic=trim(request("addpic"))             '##贴图
   'fs=trim(request("fs"))                     '##字效
   'addwordcolor=trim(request("addwordcolor")) '##颜色
   'towhoway=trim(request("towhoway"))         '##私聊


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
	  if systemsay="function" then
          application("systemsay" & 40) = "function"
	  else
	      application("systemsay" & 40) = ""
	  end if
	  systemsay=""
          application("me" & 40)        = session("user")
          application("sexsay" & 40)    = session("sex")
          application("color" & 40)     = Server.HTMLEncode(trim(request("addwordcolor")))
          application("showip" & 40)    = session("ip")
          application("mood" & 40)      = Server.HTMLEncode(trim(request("addsays")))
          application("to" & 40)        = Server.HTMLEncode(trim(request("towho")))
          application("siliao" & 40)    = trim(request("towhoway"))
          application("zixiao" & 40)    = zixiao
          application("page" & 40)      = inpm
          application("tuan" & 40)      = tuan
          application("tietu" & 40)     = tietu
          application("saytime" & 40)   = time
          application.unlock
   end if

   r_title=replace(r_title," ","")

if application("title")<>"" then  response.write "<center>"&application("title")&"</center><br>"
if r_refresh="down" then
   for i=1 to 40 step 1
   if application("systemsay"&i)<>"" or application("me"&i)<>"" or application("to"&i)<>"" or application("page"&i)<>"" then
       if application("systemsay"&i)="yes" then %>
          <font color=<%=application("color"&i)%>><%=r_title%>公告：<a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
          <%if application("sexsay"&i)="boy" then%>
          #3366FF>Boy
          <%else%>
          #FF0000>Girl
          <%end if%> <%=application("me"&i)%></font></a>
		  <%if session("userlevel")>5 then%>
		  <font size=-3>[<%=application("showip"&i)%>]</font>
		  <%end if%>
          <%=application("page"&i)%></font><font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
       <% elseif application("systemsay"&i)="function" then%>
	      <font color=<%=application("color"&i)%>>
          <%if application("sexsay"&i)="boy" then
		       functioncolor="<font color=#3366FF>"
			else
			   functioncolor="<font color=#FF6060>"
			end if
			functionuser=application("me"&i)
			if session("userlevel")>5 then
			   functionip="<font size=-3>["&application("showip"&i)&"]</font>"
			   else
			   functionip=""
			   end if
			   if application("me"&i)=session("user") then
			      changefunction="〖<a href='javascript:selectwho("&chr(34) & functionuser & chr(34)&");' title='选取"&functionuser&" 作为谈话对象'>"&functioncolor&functionuser&"</font></a>〗"&functionip&" "
               else
			     changefunction="<a href='javascript:selectwho("&chr(34) & functionuser & chr(34)&");' title='选取"&functionuser&" 作为谈话对象'>"&functioncolor&functionuser&"</font></a>"&functionip&" "
                 end if
  		   response.write replace(replace(application("page"&i),"var_who",changefunction),"var_to",application("to"&i))%></font><font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
	   <%
	   else
	   if application("to"&i)="大家" then %>
	            <font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对 大家 说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
                <%
	   else
          if application("siliao"&i)="yes" then
             if application("me"&i)=session("user") or application("to"&i)=session("user") then%>
                →<font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对<%if application("to"&i)=session("user") then response.write "〖"%><%=application("to"&i)%><%if application("to"&i)=session("user") then response.write "〗"%>悄悄说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
                <%end if%>
             <%else%>
                <font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对<%if application("to"&i)=session("user") then response.write "〖"%><%=application("to"&i)%><%if application("to"&i)=session("user") then response.write "〗"%>说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
             <%end if
		  end if
       end if
   else
   end if
   next
else
   for i=40 to 1 step -1
   if application("systemsay"&i)<>"" or application("me"&i)<>"" or application("to"&i)<>"" or application("page"&i)<>"" then
       if application("systemsay"&i)="yes" then %>
          <font color=<%=application("color"&i)%>><%=r_title%>公告：<a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
          <%if application("sexsay"&i)="boy" then%>
          #3366FF>Boy
          <%else%>
          #FF6060>Girl
          <%end if%> <%=application("me"&i)%></font></a>
		  <%if session("userlevel")>5 then%>
		  <font size=-3>[<%=application("showip"&i)%>]</font>
		  <%end if%>
          <%=application("page"&i)%></font><font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
       <% elseif application("systemsay"&i)="function" then%>
	      <font color=<%=application("color"&i)%>>
          <%if application("sexsay"&i)="boy" then
		       functioncolor="<font color=#3366FF>"
			else
			   functioncolor="<font color=#FF6060>"
			end if
			functionuser=application("me"&i)
			if session("userlevel")>5 then
			   functionip="<font size=-3>["&application("showip"&i)&"]</font>"
			   else
			   functionip=""
			   end if
			   if application("me"&i)=session("user") then
			      changefunction="〖<a href='javascript:selectwho("&chr(34) & functionuser & chr(34)&");' title='选取"&functionuser&" 作为谈话对象'>"&functioncolor&functionuser&"</font></a>〗"&functionip&" "
               else
			     changefunction="<a href='javascript:selectwho("&chr(34) & functionuser & chr(34)&");' title='选取"&functionuser&" 作为谈话对象'>"&functioncolor&functionuser&"</font></a>"&functionip&" "
                 end if
  		   response.write replace(replace(application("page"&i),"var_who",changefunction),"var_to",application("to"&i))%></font><font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
	   <%
	   else
	   if application("to"&i)="大家" then %>
	            <font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对 大家 说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
                <%
	   else
          if application("siliao"&i)="yes" then
             if application("me"&i)=session("user") or application("to"&i)=session("user") then%>
                →<font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对<%if application("to"&i)=session("user") then response.write "〖"%><%=application("to"&i)%><%if application("to"&i)=session("user") then response.write "〗"%>悄悄说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
                <%end if%>
             <%else%>
                <font color=<%=application("color"&i)%>><%if application("me"&i)=session("user") then response.write "〖"%><a href="javascript:selectwho('<%=application("me"&i)%>');" title="选取 <%=application("me"&i)%> 作为谈话对象"><font color=
                <%if application("sexsay"&i)="boy" then%>
                #3366FF><%else%>#FF6060><%end if%><%=application("me"&i)%></font></a><%if application("me"&i)=session("user") then response.write "〗"%><%if session("userlevel")>5 then%><font size=-3>[<%=application("showip"&i)%>]</font>
                <%end if%><%=application("mood"&i)%>对<%if application("to"&i)=session("user") then response.write "〖"%><%=application("to"&i)%><%if application("to"&i)=session("user") then response.write "〗"%>说：
                <%=application("page"&i)%><font face=wingdings><%=application("tuan"&i)%></font><%=application("tietu"&i)%></font> <font size=-3 color=black>(<%=application("saytime"&i)%>)</font><br>
             <%end if
		  end if
       end if
   else
   end if
   next
end if
response.write r_link
 %>
   <a name="dian"></a>
</body>
</html>