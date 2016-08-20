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
    <img border="0" src="pic/home.gif"> <b> <font color="#FFFFFF">返回论坛首页</font></b></a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>论坛管理</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=addfl">添加论坛分类</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-gl.asp?menu=addbbs">论坛添加</a>·<a href="admin-gl.asp?menu=bbsgl">管理</a>　</td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=hbbbs">论坛合并</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bzgl&bz=add">版主添加</a>·<a href="admin-right.asp?action=bzgl&bz=del">删除</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=lm">论坛联盟管理</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=chadmin">编辑管理员</a></td>
  </tr>
  </table><br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>用户管理</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=deluser">删除用户</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=updateuser">更改用户资料</a></td>
  </tr>
  <tr>
    <td width="100%" height="27"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=chpwd">更改用户密码</a></td>
  </tr>

</table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>帖子管理</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=delany">批量删除帖子</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=moveany">批量移动帖子</a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fle.gif"> <font color="#FFFFFF"><b>论坛其它</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bbs">论坛参数设置</a></td>
  </tr>
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="admin-right.asp?action=bbsmail">论坛留言管理</a></td>
  </tr>
  </table>
<br>
<table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#3A51B8>
  <tr>
    <td width="100%" height="22" background="pic/<%=sp%>3.gif">
    &nbsp;<img border="0" src="pic/fl.gif"> <font color="#FFFFFF"><b>数据处理</b></font></td>
  </tr>
</table>
<table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" width="100%" bordercolor=#F7F8FD height="9">
  <tr>
    <td width="100%" height="26"><img border="0" src="pic/fl.gif">
    <a href="mdbcon.asp">数据库操作</a></td>
  </tr>
  </table>