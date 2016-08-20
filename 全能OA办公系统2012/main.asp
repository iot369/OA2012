<%@ LANGUAGE = VBScript %>
<%
set file=server.createobject("scripting.FileSystemObject")
addr1=server.mappath("top1.asp")
addr2=server.mappath("top1.asp")
If Not file.FileExists(addr1) or Not file.FileExists(addr2) Then
response.write "<script LANGUAGE='javascript'>alert('系统发生严重错误即将关闭！！！');window.close();</script>"
End If
%>
<html>
<head>
<title>全能通用OA办工系统</title>
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
<SCRIPT LANGUAGE="JavaScript">
<!--

<!-- Hide
function killErrors() {
return true;
}
window.onerror = killErrors;
// -->
//-->
</SCRIPT>
<style type="text/css">
<!--
body {
	background-color: #335e91;
}
.style2 {color: #2d4865}
.style3 {color: #334d66}
.style4 {color: #2e4869}
-->
</style>
</head>
<body topmargin="0" leftmargin="0" onmouseover="self.status='欢迎进入温岭市三艾机械厂OA智能办公自动化系统';return true">
<table width="1003" height="741"  border="0" cellpadding="0" cellspacing="0">
                      <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
%>
  <tr>
    <td><table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td width="14"><img src="images/main/l.gif" width="14" height="700"></td>
        <td valign="top" background="images/main/bg_m.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="96"><div align="center">
              <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td><img src="images/logo_fir.gif" width="312" height="96"></td>
                  <td>　</td>
                  <td><div align="right"><img src="images/to_r.gif" width="56" height="96"></div></td>
                </tr>
              </table>
              </div></td>
          </tr>
          <tr>
            <td height="28"><table width="951" height="28"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="A3B4C4">
              <tr>
                <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="1" cellpadding="0">
                    <tr>
                      <td height="24" background="images/main/bg_t1.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td width="28%"><div align="center" class="style3">全能通用OA智能办公自动化系统</div></td>
                          <td width="33%"><span class="style3">欢迎您,<span style="font-size:11px"><%=oabusyname%></span></span></td>
                          <td width="39%">　</td>
                        </tr>
                      </table></td>
                    </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="951" height="32"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="4B779C">
              <tr>
                <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="1"><img src="images/main/l1.gif" width="1" height="30"></td>
                    <td background="images/main/m1.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td><div align="center" class="style3"><a href="desk.asp" target="main_body"><font color="#334d66" >我的办公桌面(D)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="dayrep.asp" target="main_body"><font color="#334d66" >我的个人计划(P)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="personlist.asp" target="main_body"><font color="#334d66" >我的通讯录(A)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="online/onlineuser.asp" target="main_body"><font color="#334d66" >在线员工(O)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="javascript:history.back(-1)"><font color="#334d66" >返回(B)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="" onclick="history.go(0)"><font color="#334d66" >刷新(R)</font></a></div></td>
                        <td><div align="center" class="style3"><a href="" onclick="closesystem();"><font color="#334d66" >安全退出(Q)</font></a></div></td>
                        <td>　</td>
                      </tr>
                    </table></td>
                    <td width="1"><img src="images/main/r1.gif" width="1" height="30"></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
          <tr>
            <td><table width="951" height="541"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="9EB4C9">
              <tr>
                <td valign="top" bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="155" height="539" valign="top"><table width="100%" height="540"  border="0" cellpadding="0" cellspacing="1" bgcolor="4F769D">
                      <tr>
                        <td valign="top" bgcolor="#FFFFFF"><iframe name=main marginwidth=0 marginheight=0 src="leftbutton.asp" frameborder=0 scrolling="auto"  width=153 height=540></iframe></td>
                      </tr>
                    </table></td>
                    <td width="583" valign="top"><iframe name=main_body marginwidth=0 marginheight=0 src="desk.asp" frameborder=0 scrolling="auto"  width=583 height=542 ></iframe></td>
                    <td valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="1" bgcolor="50769D"></td>
                        <td height="542" valign="top"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td height="25"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
                                <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                  <tr>
                                    <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                                    <td><span class="style2">常用信息</span></td>
                                  </tr>
                                </table></td>
                                <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="6"></td>
                          </tr>
                          <tr>
                            <td><div align="center">
                              <table width="200"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
                                <tr>
                                  <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                        <tr>
                                          <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                                          <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                            <tr>
                                              <td width="10">　</td>
                                              <td><span class="style4">登陆信息</span></td>
                                            </tr>
                                          </table></td>
                                          <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
                                        </tr>
                                      </table></td>
                                    </tr>
                                    <tr>
                                      <td><table width="92%"  border="0" align="center" cellpadding="0" cellspacing="0">
                                        <tr>
                                          <td height="10"></td>
                                        </tr>
                                        <tr>
                                          <td height="20"><span style="font-size:11px">用户：<%=oabusyname%> &nbsp;</span></td>
                                        </tr>
                                        <tr>
                                          <td height="20"><span style="font-size:11px">部门：<span style="font-size:11px"><%=oabusyuserdept%></span> </span></td>
                                        </tr>
                                        <tr>
                                          <td height="20"><span style="font-size:11px">职位：<%=oabusyuserlevel%> </span></td>
                                        </tr>
                                        <tr>
                                          <td height="20"><span style="font-size:11px">时间：
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
                                          </span></td>
                                        </tr>
                                        <tr>
                                          <td><iframe name=fir1 marginwidth=0 marginheight=0 src="top0.asp" frameborder=0 scrolling="no"  width=100% height=5 allowTransparency="true"></iframe></td>
                                        </tr>
                                        <tr>
                                          <td><iframe name=fir2 marginwidth=0 marginheight=0 src="top1.asp" frameborder=0 scrolling="no"  width=100% height=5 allowTransparency="true"></iframe></td>
                                        </tr>
                                      </table></td>
                                    </tr>
                                  </table></td>
                                </tr>
                              </table>
                            </div></td>
                          </tr>
                          <tr>
                            <td height="8"></td>
                          </tr>
                          <tr>
                            <td><table width="200"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
                              <tr>
                                <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                          <tr>
                                            <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                                            <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td width="10">　</td>
                                                  <td><span class="style4">办公常用信息</span></td>
                                                </tr>
                                            </table></td>
                                            <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
                                          </tr>
                                      </table></td>
                                    </tr>
                                    <tr>
                                      <td><table width="92%"  border="0" align="center" cellpadding="0" cellspacing="0">
                                          <tr>
                                            <td height="5" colspan="2"></td>
                                          </tr>
                                          <tr>
                                            <td width="50%" height="20"><div align="center"><a href="cy/link.asp" target="main_body">常用网址</a></div></td>
                                            <td height="20"><div align="center"><a href="cy/links.asp" target="main_body">网址管理</a></div></td>
                                          </tr>
                                          <tr>
                                            <td height="20"><div align="center"><a href="rl/cal.htm" target="main_body">万 年 历</a></div></td>
                                            <td height="20"><div align="center"><a href="ip/index.asp" target="main_body">手机及ip</a></div></td>
                                          </tr>
                                          <tr>
                                            <td height="20"><div align="center"><a href="youbian/index.asp" target="main_body">邮编区号</a></div></td>
                                            <td height="20"><div align="center"><a href="http://www.cma.gov.cn/netcenter_news/qxyb/city/index.php?city=&#21271;&#20140;&province=&#21271;&#20140;&area=&#21326;&#21271;" target="main_body">天气查询</a></div></td>
                                          </tr>
                                          <tr>
                                            <td height="20"><div align="center"><a href="http://www.hao123.com" target="main_body">网址大全</a></div></td>
                                            <td height="20"><div align="center"><a href="http://www.cngoto.com/tr/chaxun.htm" target="main_body">列车时刻</a></div></td>
                                          </tr>
                                          <tr>
                                            <td height="20"><div align="center"><a href="http://www.yoee.com/?src=hao123h" target="main_body">航班查询</a></div></td>
                                            <td height="20"><div align="center"><a href="http://www.hao123.com/ss/fy.htm" target="main_body">在线翻译</a></div></td>
                                          </tr>
                                          <tr>
                                            <td height="20"><div align="center"><a href="http://www.hao123.com/soft/default.htm" target="main_body">常用软件</a></div></td>
                                            <td height="20"><div align="center"><a href="http://map.baidu.com/" target="main_body">电子地图</a></div></td>
                                          </tr>
                                      </table></td>
                                    </tr>
                                </table></td>
                              </tr>
                            </table></td>
                          </tr>
                          <tr>
                            <td height="8"></td>
                          </tr>
                          <tr>
                            <td><table width="200"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
                              <tr>
                                <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                    <tr>
                                      <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                          <tr>
                                            <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                                            <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                                                <tr>
                                                  <td width="10">　</td>
                                                  <td><span class="style4">今日天气</span></td>
                                                </tr>
                                            </table></td>
                                            <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
                                          </tr>
                                      </table></td>
                                    </tr>
                                    <tr>
                                      <td height="3"></td>
                                    </tr>
                                    <tr>
                                      <td><IFRAME ID='ifm2' WIDTH='189' HEIGHT='190' ALIGN='CENTER' MARGINWIDTH='0' MARGINHEIGHT='0' HSPACE='0' VSPACE='0' FRAMEBORDER='0' SCROLLING='NO' SRC='http://weather.qq.com/inc/ss133.htm'></IFRAME>
</td>
                                    </tr>
                                </table></td>
                              </tr>
                            </table></td>
                          </tr>
                        </table></td>
                      </tr>
                    </table></td>
                  </tr>
                </table></td>
              </tr>
            </table></td>
          </tr>
        </table></td>
        <td width="18"><img src="images/main/r.gif" width="18" height="700"></td>
      </tr>
    </table></td>
  </tr>
  <tr>
    <td height="44" background="images/main/d.gif">　</td>
  </tr>
</table>
<div align="center"><script src="count/mystat.asp"></script></div>
</body>
</html>

