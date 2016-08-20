<!--#INCLUDE FILE="inc_config.asp"-->
<html>
<head>
<title>点歌台</title>
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

<body text="<%=m_text2%>" bgcolor="#cccccc">
<%
if request("midi")="" then response.end
if request("midi")="noname" then midin="无名"
if request("midi")="lowheart" then midin="心太软"
if request("midi")="tian" then midin="甜蜜蜜"
if request("midi")="love2000" then midin="恋曲2000"
if request("midi")="crysea" then midin="泪海"
if request("midi")="huanzu" then midin="还珠"
if request("midi")="wind" then midin="风中奇缘"
if request("midi")="heart" then midin="动心"
if request("midi")="donotask" then midin="别问我是谁"
if request("midi")="think" then midin="想说爱你不容易"
if request("midi")="thank" then midin="谢谢你的爱"
if request("midi")="little" then midin="有一点点动心"
if request("midi")="dream" then midin="半梦半醒"
if request("midi")="hero" then midin="真心英雄"
if request("midi")="birthday" then midin="祝你生日快乐"
if request("midi")="foryou" then midin="特别的爱给特别的你"
if request("midi")="letme" then midin="让我欢喜让我忧"
if request("midi")="promise" then midin="诺言（李翊君）"
if request("midi")="nextlife" then midin="来生缘"
if request("midi")="choose" then midin="选择"
if request("midi")="yihan" then midin="遗憾"
if request("midi")="realy" then midin="真的有你"
if request("midi")="know" then midin="知不知"

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
	      application("systemsay" & 40) = "function"
          application("me" & 40)        = session("user")
          application("sexsay" & 40)    = session("sex")
          application("color" & 40)     = "#5F417C"
          application("showip" & 40)    = session("ip")
          application("mood" & 40)      = ""
          application("to" & 40)        = Server.HTMLEncode(trim(request("name")))
          application("siliao" & 40)    = trim(request("towhoway"))
          application("zixiao" & 40)    = ""
          application("page" & 40)      = "var_who点了首『"& midin &"』给 var_to，您是否要收听？ 〖<a href='can.asp?midi="& request("midi") &"' target='frm_can'>收听</a>〗"
          application("tuan" & 40)      = ""
          application("tietu" & 40)     = ""
          application("saytime" & 40)   = time
application.unlock
%>
<br><br>
<center><strong>您的歌曲已放送，谢谢您！</strong><br><br>
<input type="button" value="返回" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:history.go(-1)"> 
<input type="button" value="关闭" class="normal" onmouseover="this.className='over'" onmousedown="this.className='down'" onmouseout="this.className='normal'" onclick="javascript:window.close()">
</center>
</body>
</html>