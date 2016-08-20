<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
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
'打开数据库，读出通讯录类型
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from persontype where username='"&oabusyusername&"'"
rs.open sql,conn,1
if rs.eof or rs.bof then
	conn.close
	set rs=nothing
	response.redirect "personaddtype.asp"
	response.end
end if
%>
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<style type="text/css">
<!--
.style1 {color: #098abb}
-->
body {SCROLLBAR-FACE-COLOR:#3499D0;
SCROLLBAR-HIGHLIGHT-COLOR:#CCFFFF;
SCROLLBAR-SHADOW-COLOR:#2587C3;
SCROLLBAR-ARROW-COLOR:#CCFFFF;
SCROLLBAR-BASE-COLOR:#1068A4; 
SCROLLBAR-DARK-SHADOW-COLOR:#3499D0} 
</style>
<script language="javascript">
function gllb()
{
	location.href="personaddtype.asp";
}
</script>
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
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">个人通讯录</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<center>
  <br>
  <table>
    <tr> 
      <td> 
        <input type="button" value="增加" onclick="location.href='personaddrecord.asp'">
        <span class="style1">个人通讯录</span></td>
      <form method="post" action="personlist.asp" name="form1">
        <td> 
          <select size=1 name="userdept">
            <%
if not rs.eof and not rs.bof then 
	userdept=rs("id")
	lbname=rs("typename")
end if
if request("userdept")<>"" then userdept=request("userdept")
while not rs.eof and not rs.bof
%>
            <option value="<%=cstr(rs("id"))%>"><%=rs("typename")%></option>
            <%
if rs("id")=clng(userdept) then
	lbname=rs("typename")
end if
if cstr(rs("id"))=cstr(userdept) then
	response.write("<script language=""javascript"">")
	response.write("form1.userdept.value="&chr(34)&cstr(userdept)&chr(34)&";")
	response.write("</script>")
end if
rs.movenext
wend
%>
          </select>
          <input type="submit" value="查看">
          &nbsp; 
          <input type="button" onclick="gllb()" value="管理通讯录类别">
        </td>
      </form>
    </tr>
  </table>
</center>


<center>
<table border="0"  cellspacing="0" cellpadding="0" width="95%" height=10>
<tr><td></td></td></tr></table>
个人通讯录（<%=lbname%>） 
  <table border="0"  cellspacing="0" cellpadding="0" width="95%" height=10>
<tr><td></td></td></tr></table>
  <table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table width="540" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
    <tr>
  <td height="2" colspan="5" align=center ></td>
  </tr>
	<tr bgcolor="D7E8F8"> 
      <td height=24 align=center><span class="style8">姓名</span></td>
      <td align=center><span class="style8">单位</span></td>
      <td align=center><span class="style8">职务</span></td>
      <td align=center><span class="style8">电话或分机</span></td>
      <td align=center><span class="style8">Email</span></td>
    </tr>
    <%
'显示通讯录简表
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personrecord where thisinfousername='" & oabusyusername &"' and recordtype="&userdept
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
    <tr bgcolor="#FFFFFF"> 
      <td height="24" align=center bgcolor="#FFFFFF"><a href="persondispinfo.asp?id=<%=rs("id")%>"><%=rs("xm")%></a></td>
      <td align=center><%=rs("company")%></td>
      <td align=center><%=rs("userzw")%></td>
      <%
companytel=rs("companytel")
email=rs("email")
%>
      <td align=center><%=checked3(companytel)%></td>
      <td align=center><%=maillink(email)%></td>
    </tr>
    <%
rs.movenext
wend
conn.close
set rs=nothing
%>
  </table>
</center>
<br>
<%

%>

</body>
</html>