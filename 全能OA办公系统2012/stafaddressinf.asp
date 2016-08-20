<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/maillink.asp"-->
<%
'-----------------------------------------
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

'--------------------------------------
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
.style2 {	color: #0d79b3;
	font-weight: bold;
}
.style4 {color: #0d79b3}
.style5 {color: #2d4865}
.style6 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center"><span class="style2"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                <td><span class="style5">通讯助理</span></td>
              </tr>
          </table></td>
          <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
        </tr>
      </table>
      <font color="0D79B3"></font></div>
    </span></div></td>
  </tr>
  <tr>
    <td><div align="center">
        <center>
          <br>
          <center>
            <table>
              <tr>
                <td><span class="style4">本单位员工通讯录</span></td>
                <%
'打开数据库，读出部门
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
%>
                <form method="post" action="stafaddressinf.asp">
                  <td>
                    <select size=1 name="userdept">
                      <%
if not rs.eof and not rs.bof then userdept=rs("userdept")
if request("userdept")<>"" then userdept=request("userdept")
while not rs.eof and not rs.bof
%>
                      <option value="<%=rs("userdept")%>"<%=selected(userdept,rs("userdept"))%>><%=rs("userdept")%></option>
                      <%
rs.movenext
wend
%>
                    </select>
                    <input name="submit" type="submit" value="查看">
                  </td>
                </form>
              </tr>
            </table>
            <table width="10%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>
          </center>
          <center>
            <table border="0"  cellspacing="0" cellpadding="0" width="95%" height=10>
            </table>
            <span class="style4"><%=userdept%>通讯录
            </span>            
            <table width="10%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td>&nbsp;</td>
              </tr>
            </table>                       
              <table width="550"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
          </table><table border="0"  cellspacing="1" cellpadding="0" width="550" bgcolor="B0C8EA" align="center">
              <tr bgcolor="B0C8EA">
                <td height=2 colspan="5" align=center ></td>
              </tr>
              <tr bgcolor="D7E8F8">
                <td height=30 align=center ><span class="style6">姓名</span></td>
                <td align=center ><span class="style6">部门</span></td>
                <td align=center ><span class="style6">职位</span></td>
                <td align=center ><span class="style6">公司电话或分机</span></td>
                <td align=center ><span class="style6">Email</span></td>
              </tr>
              <%
'显示职员表
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where userdept=" & sqlstr(userdept)
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
              <tr bgcolor="#FFFFFF">
                <td height="25" align=center><a href="showstafaddressinf.asp?userdept=<%=userdept%>&username=<%=rs("username")%>&name=<%=rs("name")%>&userlevel=<%=rs("userlevel")%>"><%=rs("name")%></a></td>
                <td align=center><%=rs("userdept")%></td>
                <td align=center><%=rs("userlevel")%></td>
                <%
'打开数据库显示通讯信息
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from stafaddressinf where username=" & sqlstr(rs("username"))
rs1.open sql,conn,1
if not rs1.eof and not rs1.bof then
companytel=rs1("companytel")
email=rs1("email")
else
companytel=""
email=""
end if
%>
                <td align=center><%=checked3(companytel)%></td>
                <td align=center><%=maillink(email)%></td>
              </tr>
              <%
rs.movenext
wend
%>
            </table>
          </center>
          <br>
          <br>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>