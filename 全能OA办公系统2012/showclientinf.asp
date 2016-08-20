<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/maillink.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

'打开数据库，读出编辑通讯录权限
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
allow_edit_client_addressinf=rs("allow_edit_client_addressinf")

id=request("id")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统.边缘特别版</title>
<style type="text/css">
<!--
.style1 {color: #0d79b3}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                <td class="style7">客户资源</td>
              </tr>
          </table></td>
          <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
        </tr>
      </table>
    <font color="0D79B3"></font></div></td>
  </tr>
</table>
<center>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td> <br>
<table align="center">
<tr>
<td><span class="style1">客户详细资料</span>&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%
if allow_edit_client_addressinf="yes" then
%>
<form method="post" action="editclientinf.asp">
<td><input type="submit" value="编辑"><input type="hidden" name="project" value="<%=request("project")%>">
<input type="hidden" name="id" value="<%=request("id")%>">
</td>
</form>
<%
end if
%>
<form method="post" action="clientinf.asp">
<td><input type="submit" value="返回"><input type="hidden" name="project" value="<%=request("project")%>">
</td>
</form>
</tr>
</table>


<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from clientinf where id=" & id
rs.open sql,conn,1
%>
<br><br>
<center>
   <table width="550"  border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
        </table><table width="550" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="B0C8EA">
    <tr>
      <td height=2 colspan="4" align=center 　　　　　　　　></td>
      </tr>
    <tr bgcolor="#FFFFFF">
      <td width="100" height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">客户姓名</span></td>
      <td width="170" 　　　　　　　　>　<%=checked3(rs("name"))%></td>
      <td width="100" align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">客户性别</span></td>
      <td width="181" 　　　　　　　　>　<%=checked3(rs("sex"))%></td>
    </tr>
     <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">职&nbsp;&nbsp;&nbsp;&nbsp;位</span></td>
      <td 　　　　　　　　>　<%=checked3(rs("position"))%></td>
      <td align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">业务项目</span></td>
      <td >　<%=checked3(rs("project"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">所在单位</span></td>
      <td colspan="3" >　<%=checked3(rs("company"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">部&nbsp;&nbsp;&nbsp;&nbsp;门</span></td>
      <td 　　　　　　　　>　<%=checked3(rs("dept"))%></td>
      <td align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">邮政编码</span></td>
      <td >　<%=checked3(rs("postcard"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">公司地址</span></td>
      <td colspan="3" >　<%=checked3(rs("address"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">传真号码</span></td>
      <td colspan="3" >　<%=checked3(rs("fax"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">联系电话</span></td>
      <td 　　　　　　　　>　<%=checked3(rs("tel"))%></td>
      <td align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">手&nbsp;&nbsp;&nbsp;&nbsp;机</span></td>
      <td >　<%=checked3(rs("handset"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">M　S　N</span></td>
      <td 　　　　　　　　>　<%=checked3(rs("callno"))%></td>
      <td align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">电子邮箱</span></td>
      <td >　<%=maillink(rs("email"))%></td>
    </tr>
    <tr bgcolor="#FFFFFF">
      <td height=25 align=center bgcolor="D7E8F8" 　　　　　　　　><span class="style8">备注说明</span></td>
      <td colspan="3" 　　　　　　　　>　<%=keepformat(checked3(rs("remark")))%></td>
    </tr>
  </table>
</td>
    </tr>
  </table>
 

</body>
</html>










