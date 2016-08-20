<!--#INCLUDE FILE="inc_config.asp" -->
<% if session("user")<>"" and session("userlevel")>7 then
      application.lock
        for x =1 to 39
	      application("systemsay" & x)  = ""
          application("sexsay" & x)     = ""
          application("me" & x)         = ""
          application("color" & x)      = ""
          application("showip" & x)     = ""
          application("mood" & x)       = ""
          application("to" & x)         = ""
          application("siliao" & x)     = ""
          application("zixiao" & x)     = ""
          application("page" & x)       = ""
          application("tuan" & x)       = ""
          application("tietu" & x)      = ""
          application("saytime" & x)    = ""
        next
		  application("systemsay" & 40) = "yes"
          application("me" & 40)        = session("user")
          application("sexsay" & 40)    = session("sex")
          application("color" & 40)     = "#0000FF"
          application("showip" & 40)    = session("ip")
          application("mood" & 40)      = ""
          application("to" & 40)        = ""
          application("siliao" & 40)    = ""
          application("zixiao" & 40)    = ""
          application("page" & 40)      = " 清空了一下聊天室的信息！"
          application("tuan" & 40)      = ""
          application("tietu" & 40)     = ""
          application("saytime" & 40)   = time
          application.unlock
%>
<html>
<head>
<title>清空信息</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
<!--
.normal   {  <%=m_button%> ; font-size: 10.4pt}
.over     {  <%=m_buttonover%> ; font-size: 10.4pt}
.down     {  <%=m_buttondown%> ; font-size: 10.4pt}
-->
</style>
</head>
<body bgcolor="#cccccc" text="<%=m_text2%>" style="font-size:10.4pt;border:0">
<br><center>聊天室的界面已经清空！<br><br>
<input type="button" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" value="关闭" onclick="javascript:window.close();"></center>
</body>
</html>
<% end if %>