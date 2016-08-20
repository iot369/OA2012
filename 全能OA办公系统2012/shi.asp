<%@ LANGUAGE = VBScript %>
<html>
<head>
<title>伴江行OA办公系统</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css" type="text/css">
<SCRIPT language=javascript>
<!--
if (window.Event) 
　document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
　if (e.which == 2 || e.which == 3)
　 return false;
 } else if (event.button == 2 || event.button == 3) {
　 event.cancelBubble = true
　 event.returnvalue = false;
　 return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;　// for IE5+
document.onmousedown = norightclick;　　 // for all others
//-->
</SCRIPT>
<SCRIPT LANGUAGE="javascript">
	windowWidth = window.screen.availWidth;
	windowHeight = window.screen.availHeight;
	window.moveTo(100,80);
	window.resizeTo(840,600);
</SCRIPT>

<script src="qqprg/init.asp"></script>
<script>
function tick() {
var hours, minutes, seconds, ap;
var intHours, intMinutes, intSeconds;
var today;
today = new Date();
intHours = today.getHours();
intMinutes = today.getMinutes();
intSeconds = today.getSeconds();
if (intHours == 0) {
hours = "12:";
ap = "Midnight";
} else if (intHours < 12) { 
hours = intHours+":";
ap = "A.M.";
} else if (intHours == 12) {
hours = "12:";
ap = "Noon";
} else {
hours = intHours + ":";
ap = "P.M.";
}
if (intMinutes < 10) {
minutes = "0"+intMinutes+":";
} else {
minutes = intMinutes+":";
}
if (intSeconds < 10) {
seconds = "0"+intSeconds+" ";
} else {
seconds = intSeconds+" ";
} 
timeString = hours+minutes+seconds+ap;
Clock.innerHTML = timeString;
window.setTimeout("tick();", 1000);
}
window.onload = tick;
</script>

<script language="javascript">
//单击"注销"连接时，弹出对话框是否要求退出系统
function closesystem()
{
	window.open('logout.asp?closeflag=1','closesystem','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
}

</script>

</head>
<body>
<table width="100%" align="center" height="190" class="borderon" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>
      <table  width="100%" height="384" border="0" align="center" cellpadding="0" cellspacing="0" bordercolor="#808080" >
        <tr> 
          <td colspan="2" height="25"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="10%">&nbsp;</td>
              <td width="56%">&nbsp;</td>
              <td width="34%">&nbsp;</td>
            </tr>
             <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
%>
			<tr>
              <td>&nbsp;</td>
              <td><span style="font-size:11px">登录用户：[<%=oabusyname%>] &nbsp;&nbsp;部门：[<%=oabusyuserdept%>] &nbsp;&nbsp;职位：[<%=oabusyuserlevel%>] &nbsp;&nbsp;时间：[
                  <script language="JavaScript">
<!---
//取得日期和星期
   today=new Date();
   function initArray(){
	 this.length=initArray.arguments.length
	 for(var i=0;i<this.length;i++)
	 this[i+1]=initArray.arguments[i]  }
	 
  var d=new initArray("<font color=RED>星期日","<font color=black>星期一","<font color=black>星期二","<font color=black>星期三","<font color=black>星期四","<font color=black>星期五","<font color=GREEN>星期六"); 
document.write(today.getYear(),"年",today.getMonth()+1,"月",today.getDate(),"日",d[today.getDay()+1]);  
//-->
                  </script>
]</span></td>
              <td><img src="images/group.gif" width="16" height="16"><a href="" class="colin2" onclick="history.go(0)">MSG</a>&nbsp;&nbsp; &nbsp;&nbsp; <img src="images/refresh.gif" width="16" height="16"><a href="" class="colin2" onclick="history.go(0)" >刷新</a> &nbsp;&nbsp;&nbsp;&nbsp;<img src="images/bhome.gif" width="16" height="15"><a href="about.htm" class="colin2" target=main_wanglongdai>关于</a> &nbsp;&nbsp;&nbsp;&nbsp;<img src="images/m1.gif" width="16" height="16"> <a href="oareg.asp" class="colin2" target=main_wanglongdai>注册</a>&nbsp;&nbsp;</td>
            </tr>
          </table></td>
        </tr>
        <tr> 
          <td  valign="top" > <iframe name=main marginwidth=0 marginheight=0 src="leftoa.asp" frameborder=0 scrolling="auto"  width=100% height=480></iframe> 
          </td>
          <td  valign="top" width="80%" > 
            <table  width="99%" border="0" cellspacing="0" cellpadding="0" class="borderon" align="center">
              <tr> 
                <td height="20" > <iframe name=main_wanglongdai marginwidth=0 marginheight=0 src="desk.asp" frameborder=0 scrolling="auto"  width=100% height=480></iframe></td>
              </tr>
            </table>
          </td>
        </tr>
      </table>
    </td>
  </tr>

</table>

</body>
</html>

