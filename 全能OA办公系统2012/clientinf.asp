<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
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
.style4 {color: #0d79b3}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
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

  <tr>
    <td><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
        <center>
          <table>
            <tr>
              <td><span class="style4">客户资源列表&nbsp;</span>&nbsp;&nbsp;&nbsp;</td>
              <%
if allow_edit_client_addressinf="yes" then
%>
              <form method=post action="addclientinf.asp">
                <td>
                  <input name="submit" type=submit value="增加">
                </td>
              </form>
              <%
end if
%>
              <%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT project from clientinf"
rs.open sql,conn,1
if not rs.eof and not rs.bof then project=rs("project")
if request("project")<>"" then project=request("project")
%>
              <form method="post" action="clientinf.asp">
                <td>
                  <select size=1 name="project">
                    <%
while not rs.eof and not rs.bof
%>
                    <option value="<%=rs("project")%>"<%=selected(project,rs("project"))%>><%=rs("project")%></option>
                    <%
rs.movenext
wend
%>
                  </select>
                  <input name="submit2" type="submit" value="查询">
                </td>
              </form>
            </tr>
          </table>
        </center>
        <br>
        <br>
        <center>
          <table width="550"  border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
          </table><table width="550" height="64" border="0" align="center" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
            <tr>
              <td height="2" colspan="7" align=center></td>
            </tr>
            <tr bgcolor="D7E8F8">
              <td width="11%" height="20" align=center><span class="style8">姓名</span></td>
              <td width="11%" height="20" align=center><span class="style8">姓别</span></td>
              <td width="22%" height="20" align=center><span class="style8">单位名称</span></td>
              <td width="11%" height="20" align=center><span class="style8">部门</span></td>
              <td width="11%" height="20" align=center><span class="style8">职位</span></td>
              <td width="22%" height="20" align=center><span class="style8">业务项目</span></td>
              <td width="12%" height="20" align=center><span class="style8">电话</span></td>
            </tr>
            <%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from clientinf where project=" & sqlstr(project)
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
            <tr bgcolor="#FFFFFF">
              <td width="11%" height=20 align=center><a href="showclientinf.asp?project=<%=request("project")%>&id=<%=rs("id")%>"><%=checked3(rs("name"))%></a></td>
              <td width="11%" align=center><%=checked3(rs("sex"))%></td>
              <td width="22%" align=center><%=checked3(rs("company"))%></td>
              <td width="11%"><%=checked3(rs("dept"))%></td>
              <td width="11%" align=center><%=checked3(rs("position"))%></td>
              <td width="22%" align=center><%=checked3(rs("project"))%></td>
              <td width="12%" align=center><%=checked3(rs("tel"))%></td>
            </tr>
            <%
rs.movenext
wend
%>
          </table>
        </center>
        <br>
        <br>
        <br>
        <center>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>