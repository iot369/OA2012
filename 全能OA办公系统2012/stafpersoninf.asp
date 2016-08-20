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
sql="select allow_see_all_personinf from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
cook_allow_see_all_personinf=rs("allow_see_all_personinf")
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
.style1 {color: #098abb}
-->
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
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
      <td> <br>
      员工基本档案&nbsp;&nbsp; </td>
      <%
if cook_allow_see_all_personinf="yes" then
%>
      <form action="stafpersoninf.asp" method=get name="form1">
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
      <form action="stafpersoninf.asp" method=get name="form2">
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
    </tr>
  </table>
</center>

<%
'读出姓名、部门、职位
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name,userdept,userlevel from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
name=rs("name")
userdept=rs("userdept")
userlevel=rs("userlevel")
'打开数据库，读出个人档案
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personinf where username=" & sqlstr(username)
rs.open sql,conn,1
dim a(33)
if not rs.eof and not rs.bof then
for i=1 to 33
a(i)=rs("a" & i)
next
inputdate=rs("inputdate")
updatedate=rs("updatedate")
havephoto=rs("havephoto")
id=rs("id")
else
for i=1 to 33
a(i)=""
next
inputdate=""
updatedate=""
havephoto="no"
end if
%>
<center>
<br>
  <table border="0" cellpadding="0" cellspacing="0" width="95%">
    <tr> 
      <td width="30%">员工编号：<%=a(1)%></td>
      <td width="35%">录入时间：<%=inputdate%></td>
      <td align="right">修改时间：<%=updatedate%></td>
    </tr>
  </table>        
    
  <table border="0" cellpadding="0" cellspacing="0" width="540">
    <tr> 
      <td width="15%" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">姓&nbsp;&nbsp;&nbsp;&nbsp;名</td>
      <td width="35%" style="border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=name%></td>
      <td width="15%" align="center" style="border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">曾&nbsp;用&nbsp;名</td>
      <td width="100%" style="border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(2))%></td>
      <td width="80" height="100" rowspan="5" align="center" valign=center style="border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <%
if havephoto="yes" then
%>
        <img src="showphoto.asp?id=<%=id%>" width="80" height="100" border="0"> 
        <%
else
%>
        <table width=80>
          <tr> 
            <td align=center> 无<br>
              照<br>
              片 </td>
          </tr>
        </table>
        <%
end if
%>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">性&nbsp;&nbsp;&nbsp;&nbsp;别</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(3))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">民&nbsp;&nbsp;&nbsp;&nbsp;族</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(4))%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">所属部门</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=userdept%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">职&nbsp;&nbsp;&nbsp;&nbsp;务</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=userlevel%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">职&nbsp;&nbsp;&nbsp;&nbsp;称</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(5))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">出生日期</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(6))%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">政治面貌</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(7))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">健康状况</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(8))%></td>
    </tr>
    <tr> 
      <td align="center" width="15%" height="20" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">籍&nbsp;&nbsp;&nbsp;&nbsp;贯</td>
      <td width="35%" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(9))%></td>
      <td width="15%" align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">体&nbsp;&nbsp;&nbsp;&nbsp;重</td>
      <td width="35%" colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(10))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">身份证号</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(11))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">身&nbsp;&nbsp;&nbsp;&nbsp;高</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(12))%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">婚姻状况</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(13))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">毕业院校</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(14))%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">本人成分</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(15))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">专&nbsp;&nbsp;&nbsp;&nbsp;业</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(16))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">工&nbsp;&nbsp;&nbsp;&nbsp;龄</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(17))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">学&nbsp;&nbsp;&nbsp;&nbsp;历</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(18))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">外语语种</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(19))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">外语水平</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(20))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">普通话程度</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(21))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">粤语程度</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(22))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">计算机能力</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(23))%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">户口所在地</td>
      <td colspan="2" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(24))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">现&nbsp;住&nbsp;址</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(25))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">档案存放地</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(26))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">个人专长<br>
        以及爱好</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(a(27))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">本人曾受<br>
        过何种奖<br>
        励和处分</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(a(28)))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">工作经历</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(a(29)))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">家庭情况</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(a(30)))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">本&nbsp;&nbsp;&nbsp;&nbsp;人<br>
        联系方式</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(a(31)))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">发生意外<br>
        紧急情况<br>
        通知何人</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=keepformat(checked3(a(32)))%></td>
    </tr>
    <tr> 
      <td height="20" align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">备&nbsp;&nbsp;&nbsp;&nbsp;注</td>
      <td colspan="4" style="border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><%=keepformat(checked3(a(33)))%></td>
    </tr>
  </table>        

<br>

</center>

</body>
</html>










