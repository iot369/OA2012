<!--#include file="conn.asp"--><%set bbs=myconn.execute("select*from bbsinfo")
sty="all"
sp=request.cookies(cn&"1")(sty)
c1=request.cookies(cn&"1")(sty&"c1")
c2=request.cookies(cn&"1")(sty&"c2")
if sp="" then sp="a"
if c1="" then c1=bbs(1)
if c2="" then c2=bbs(2)
set bbs=nothing
myconn.close
set myconn=nothing
%>
<link rel="stylesheet" type="text/css" href="css.css">
<base target="right">
<body topmargin="0" leftmargin="0">
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%">
  <tr>
    <td width="100%" height="28" background="pic/<%=sp%>3.gif" align="center">
    <a target="main" href="index.asp">
    <img border="0" src="pic/home.gif"> <b> <font color="#FFFFFF">������̳��ҳ</font></b></a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>��̳����</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=addfl">�����̳����</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-gl.asp?menu=addbbs">��̳���</a>��<a href="admin-gl.asp?menu=bbsgl">����</a>��</td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=hbbbs">��̳�ϲ�</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bzgl&bz=add">�������</a>��<a href="admin-right.asp?action=bzgl&bz=del">ɾ��</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=lm">��̳���˹���</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=chadmin">�༭����Ա</a></td>
  </tr>
  </table><br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>�û�����</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=deluser">ɾ���û�</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=updateuser">�����û�����</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=chpwd">�����û�����</a></td>
  </tr>

</table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>���ӹ���</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=delany">����ɾ������</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=moveany">�����ƶ�����</a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>��̳����</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bbs">��̳��������</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bbsmail">��̳���Թ���</a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fl.gif"> <font color="#FFFFFF"><b>���ݴ���</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="mdbcon.asp">���ݿ����</a></td>
  </tr>
  </table>