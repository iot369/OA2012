<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/keepformat.asp"-->
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
'打开数据库，读出权限
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select allow_edit_all_checkinf from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
cook_allow_edit_all_checkinf=rs("allow_edit_all_checkinf")
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>oa办公系统</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
              <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">员工管理</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
<tr>
      <td> 员工考核档案&nbsp;&nbsp; </td>
<%
if cook_allow_edit_all_checkinf="yes" then
%>
<form action="checkinf.asp" method=get name="form1">
<td>
<select name="userdept" size=1 onChange="document.form1.submit();">
<%
'打开数据库读出部门
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
if not rs.eof and not rs.bof then firstdept=rs("userdept")
if request("userdept")<>"" then firstdept=request("userdept")
while not rs.eof and not rs.bof
%>
<option value="<%=rs("userdept")%>"<%=selected(firstdept,rs("userdept"))%>><%=rs("userdept")%></option>
<%
rs.movenext
wend
%>
</select>
</td>
</form>
<%
else
firstdept=oabusyuserdept
end if
%>
<form action="checkinf.asp" method=get name="form2">
<td>
<input type="hidden" name="userdept" value="<%=firstdept%>">
<select name="username" size=1>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name,username from userinf where userdept=" & sqlstr(firstdept) & " and userlevel<>'总管' and forbid='no'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then username=rs("username")
if request("username")<>"" then username=request("username")
while not rs.eof and not rs.bof
%>
<option value="<%=rs("username")%>"<%=selected(username,rs("username"))%>><%=rs("name")%></option>
<%
rs.movenext
wend
%>
</select>
<td>
<input type="submit" name="submit" value="查询">
</td>
</form>
<form method="post" action="addcheckinf.asp">
<td>
<input type="hidden" name="userdept" value="<%=firstdept%>">
<input type="hidden" name="username" value="<%=username%>">
<input type="submit" value="增加档案">
</td>
</form>
</tr>
</table>
</center>

<%
'打开数据库读出员工姓名
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
name=rs("name")
%>
<br>
<center>
  <%=name%>的考核档案
</center>
<br>
<%
'打开数据库，读出员工考核数据库
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from checkinf where username=" & sqlstr(username)
rs.open sql,conn,1
while not rs.eof and not rs.bof
checkname=rs("checkname")
checkdate=rs("checkdate")
checkcommment=rs("checkcommment")
checksort=rs("checksort")
checktype=rs("checktype")
checkresult=rs("checkresult")
remark=rs("remark")
recname=rs("recname")
recdate=rs("recdate")
updatename=rs("updatename")
updatedate=rs("updatedate")
id=rs("id")
%>
<center>
  <table border="0" cellpadding="4" cellspacing="0" width="540">
    <tr> 
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">员工姓名</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=name%> 
      </td>
    </tr>
    <tr> 
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">考核名称</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(checkname)%> 
      </td>
      <td width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">考核时间</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(checkdate)%> 
      </td>
    </tr>
    <tr> 
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">考核类型</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(checksort)%> 
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">考核方式</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(checktype)%> 
      </td>
    </tr>
    <tr> 
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">考核评语</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(checkcommment))%> 
      </td>
    </tr>
    <tr> 
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">考核结果</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(checkresult))%> 
      </td>
    </tr>
    <tr> 
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">备注说明</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(remark))%> 
      </td>
    </tr>
    <tr> 
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">记 
        录 人</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(recname)%> 
      </td>
      <td width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">记录时间</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=recdate%> 
      </td>
    </tr>
    <tr> 
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" width="15%" align="center">更 
        改 人</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><%=updatename%> 
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" width="15%" align="center">更改时间</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><%=updatedate%> 
      </td>
    </tr>
  </table>
  <table border="0" cellpadding="0" cellspacing="0" width="540">
    <tr><form method="post" action="editcheckinf.asp">
      <td align=right><input type="submit" value="编辑">
<input type="hidden" name="userdept" value="<%=firstdept%>">
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="id" value=<%=id%>>
</td>
        </form>
    </tr>
  </table>
<br>
</center>
<%
rs.movenext
wend 
%>


</body>
</html>










