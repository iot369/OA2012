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
<%response.write"<br><br><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>�� �� �� ̳</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% ><P style='MARGIN: 10px'>���Ѿ��ɹ����˳���̳��</p><P style='MARGIN: 10px'>��<a href=index.asp>������̳��ҳ</a>��</p></td></tr></table></center></div>"
case""%><!--#include file="up.asp"-->
<%response.write"<br><br><form method='POST' name='login' action='bbselse.asp'><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>�� ̳ �� ½</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div>"&_
"<div align='center'><center><table border='1' cellpadding='0' cellspacing='0' style='border-collapse: collapse' bordercolor='"&c1&"' width='94%'><tr><td width='30%' height='30'>&nbsp;�����������û���</td><td width='70%'> &nbsp;<input type=text name=lgname size=20> <a href=zhuce.asp>û��ע�᣿</a></td></tr><tr><td height='30'>&nbsp;��������������</td><td>&nbsp;<input type=password name=lgpwd size=20> <a href=getpwd.asp> �������룿</a></td></tr>"&_
"<tr><td height='80'>&nbsp;<font color=#000000><b>Cookie ѡ��</b><br>&nbsp;��ѡ����� Cookie ����ʱ�䣬<br>&nbsp;�´η��ʿ��Է������롣</font></td><td><font color=#000000><input type=radio CHECKED value=j0 name=Cook>�����棬�ر��������ʧЧ<br><input type=radio value=j1 name=Cook>����һ��<br><input type=radio value=j30 name=Cook>����һ��<br><input type=radio value=j365 name=Cook>����һ��<input type='hidden' name='comeurl' size='30' value='"&comeurl&"'></font></td></tr><tr><td bgcolor="&c1&" colspan='2' background='pic/"&sp&"3.gif' height='26' align='center'> <input type='submit' value=' ��  ½ ' name='B1'></td></tr></table></center></div></form>"%>
<%end select%><br><!--#include file="down.asp"-->