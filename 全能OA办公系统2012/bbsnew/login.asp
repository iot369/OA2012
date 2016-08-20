<!--#include file="conn.asp"--><%comeurl=Request.ServerVariables("HTTP_REFERER")
action=request.querystring("action")
select case action
case"exit"
myconn.execute("delete*from online where name='"&request.cookies(cn)("lgname")&"'")
Response.Cookies(cn)("lgname")=""
Response.Cookies(cn)("lgpwd")=""
myconn.close
set myconn=nothing%>
<!--#include file="up.asp"-->
<%response.write"<br><br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>退 出 论 坛</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% ><P style='MARGIN: 10px'>·已经成功的退出论坛·</p><P style='MARGIN: 10px'>·<a href=index.asp>进入论坛首页</a>·</p></td></tr></table></center></div>"
case""%><!--#include file="up.asp"-->
<%response.write"<br><br><form method='POST' name='login' action='bbselse.asp'><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>论 坛 登 陆</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div>"&_
"<div align='center'><center><table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='"&c1&"' width='94%'><tr><td width='30%' height='30'>&nbsp;请输入您的用户名</td><td width='70%'> &nbsp;<input type=text name=lgname size=20> <a href=zhuce.asp>没有注册？</a></td></tr><tr><td height='30'>&nbsp;请输入您的密码</td><td>&nbsp;<input type=password name=lgpwd size=20> <a href=getpwd.asp> 忘记密码？</a></td></tr>"&_
"<tr><td height='80'>&nbsp;<font color=#000000><b>Cookie 选项</b><br>&nbsp;请选择你的 Cookie 保存时间，<br>&nbsp;下次访问可以方便输入。</font></td><td><font color=#000000><input type=radio CHECKED value=j0 name=Cook>不保存，关闭浏览器就失效<br><input type=radio value=j1 name=Cook>保存一天<br><input type=radio value=j30 name=Cook>保存一月<br><input type=radio value=j365 name=Cook>保存一年<input type='hidden' name='comeurl' size='30' value='"&comeurl&"'></font></td></tr><tr><td bgcolor="&c1&" colspan='2' background='pic/"&sp&"3.gif' height='26' align='center'> <input type='submit' value=' 登  陆 ' name='B1'></td></tr></table></center></div></form>"%>
<%end select%><br><!--#include file="down.asp"-->