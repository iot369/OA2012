<!--#include file="up.asp"--><!--#include file="fun.asp"-->
<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%myname=request.querystring("name")
set rs=myconn.execute("SELECT*FROM user where name='"&myname&"'")
%>
<br><br>
<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>�� �� �� Ϣ</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse; border-left-width: 1px; border-right-width: 1px; border-top-width: 1px" bordercolor="<%=c1%>" width="94%">
    <tr>
      <td width="923" colspan="4" height="51">
      <p style="margin: 4">
<img align=absmiddle border="0" src="<%=rs("toupic")%>" width="<%=rs("ku")%>" height="<%=rs("ch")%>"> <a title='�����<%=kbbs(myname)%>����' href=mailcon.asp?menu=write&towho=<%=kbbs(myname)%>><img align=absmiddle border="0" src="pic/newmail.gif"> <b><%=kbbs(myname)%></b></a></td>
    </tr>
    <tr>
      <td width="460" height="25" colspan="2" background="pic/4.gif">
      <font color="<%=c1%>"><b>&nbsp;�������ϣ�</b></font></td>
      <td width="462" height="25" colspan="2" background="pic/4.gif">&nbsp;<b><font color="<%=c1%>">��̳���ϣ�</font></b></td>
    </tr>
    <tr>
      <td width="10%" align="right" height="30" bgcolor="<%=c2%>">�� ��</td>
      <td width="40%" height="30" bgcolor="<%=c2%>">��<%if rs("sex")=1 then
      sex1="��"
      else
      sex1="Ů"
      end if
      %><%=sex1%>
</td>
      <td width="10%" height="30" align="right" bgcolor="<%=c2%>">�� Ǯ����</td>
      <td width="40%" height="30" bgcolor="<%=c2%>">��<%q1=rs("qian")%><%=q1%></td>
    </tr>
    <tr>
      <td align="right" height="30" width="10%">�� �գ�</td>
      <td height="30" width="40%">��<%=rs("burn")%></td>
      <td height="30" align="right" width="10%">�� ������</td>
      <td height="30" width="40%">��<%m1=rs("meili")%><%=m1%></td>
    </tr>
    <tr>
      <td align="right" height="30" bgcolor="<%=c2%>" width="10%">E-mail��</td>
      <td height="30" bgcolor="<%=c2%>" width="40%">��<a href="mailto:<%=rs("email")%>"><%=rs("email")%></a></td>
      <td height="30" align="right" bgcolor="<%=c2%>" width="10%">�� �飺��</td>
      <td height="30" bgcolor="<%=c2%>" width="40%">��<%j1=rs("jingyan")%><%=j1%></td>
    </tr>
    <tr>
      <td align="right" height="30" width="10%">QQ���룺</td>
      <td height="30" width="40%">��<%=rs("qq")%></td>
      <td height="30" align="right" width="10%">�����⣺��</td>
      <td height="30" width="40%">��<%=myconn.execute("select count(riqi) from min where bid=0 and name='"&myname&"'")(0)%></td>
    </tr>
    <tr>
      <td align="right" height="30" bgcolor="<%=c2%>" width="10%">�� ҳ��</td>
      <td height="30" bgcolor="<%=c2%>" width="40%">��<a href="<%=rs("home")%>"><%=rs("home")%></a></td>
      <td height="30" align="right" bgcolor="<%=c2%>" width="10%">�ظ���������</td>
      <td height="30" bgcolor="<%=c2%>" width="40%">��<%=myconn.execute("select count(riqi) from min where bid<>0 and name='"&myname&"'")(0)%></td>
    </tr>
    <tr>
      <td align="right" height="30" width="10%">��̳�ȼ���</td>
      <td height="30" width="40%">��<%sqltype="my"%><!--#include file="upji.asp"--><b><%=dj%></b>�� <%=dd%> ��</td>
      <td height="30" align="right" width="10%">�������ӣ���</td>
      <td height="30" width="40%">��<%=myconn.execute("select count(riqi) from min where face='jing' and name='"&myname&"'")(0)%></td>
    </tr>
    <tr>
      <td height="25" colspan="4" background="pic/4.gif">��</td>
    </tr>
  </table>
  </center>
</div><p></p><!--#include file="down.asp"-->