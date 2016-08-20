<!--#INCLUDE FILE="inc_config.asp"-->
<html>
<head>
<title>功能区</title>
<script language="JavaScript">
<!--
 function openwin(win,url,w,h)
  { parameter = "toolbar=no,location=no,directories=no,status=no,menubar=no,scrollbars=yes,resizable=no,copyhistory=no,width="+w+",height="+h+"" ;  
    window.open(url,win,parameter); }

 function turnon(obj1)
  { obj1.style.background="<%=m_bg3%>";
    obj1.style.color="<%=m_text3%>";
	obj1.style.cursor="hand" }

 function turnoff(obj1)
  { obj1.style.background="#cccccc";
    obj1.style.color="<%=m_text2%>" }

<%if session("userlevel")<9 then%>
function click() {
if (event.button==2) {
var char_=navigator.appName
if(char_=="Netscape"){
window.parent.resizeTo(3,1)
window.parent.moveTo(100,30)
}
else{
window.parent.resizeTo(4,1)
window.parent.moveTo(100,30)
}
alert("请不要按右键，现在请在任务栏最大化聊天室的窗口，谢谢");
}
}
document.onmousedown=click
<%end if%>
if (self==top)
{top.location="index.htm";}
//-->
</script>
<style type="text/css">
<!--
body      { font-size: 9pt; background-color: "#cccccc" }
td        { font-size: 9pt; color: <%=m_text%> }
A:link    { text-decoration: none; color: <%=m_text%> }
A:visited { text-decoration: none; color: <%=m_text%> }
A:active  { text-decoration: none; color: <%=m_text%> }
A:hover   { text-decoration: none; color: <%=m_text%> }
-->
</style>
<% if request("midi")<>"" then %>
<bgsound src="midi/<%=request("midi")%>.mid" loop="1">
<% end if %>
</head>

<body bgcolor="#cccccc" text="<%=m_text2%>">
 <table border="0" width="80%" cellspacing="1" cellpadding="3" bgcolor="<%=m_bg3%>" align="center">
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="parent.frm_input.location.href='input.asp'">更新名单</td>
  </tr>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="parent.frm_onlineuser.location.href='onlineuser.asp'">谁在线上</td>
  </tr>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('userinfo','userinfo.asp',350,370)">人员信息</td>
  </tr> 
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('changeinfo','changeinfo.asp',350,320)">修改资料</td>
  </tr> 
  <% if session("userlevel")>8 then %>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('highmanager','highmanager.asp',440,500)">高级管理</td>
  </tr> 
  <% end if %>
  <% if session("userlevel")>7 then %>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('lowmanager','lowmanager.asp',440,500)">网管专用</td>
  </tr> 
  <% end if %>
  <% if session("userlevel")>8 then %>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('title','title.asp',350,220)">修改标题</td>
  </tr> 
  <% end if %>
  <% if session("userlevel")>6 then %>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('kill','kill.asp',350,180)">赶出会议</td>
  </tr> 
  <% end if %>
  <% if session("userlevel")>7 then %>
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="javascript:openwin('clear','clear.asp',200,110)">清信息区</td>
  </tr> 
  <% end if %>
  </tr> 
  <tr bgcolor="#cccccc" align="center">
   <td onmouseover="turnon(this);" onmouseout="turnoff(this);" onclick="parent.location.href='left.asp'">离开聊天</td>
  </tr> 
</body>
</html>