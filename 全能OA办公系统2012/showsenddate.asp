<%@ LANGUAGE = VBScript %>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->

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

%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<title>公文传阅</title>
<style type="text/css">
<!--
.style4 {color: #2e4869}
.style6 {color: #FF0000}
.style7 {font-weight: bold}
-->
</style>
</head>
<body  topmargin="0" leftmargin="5" bgcolor="#F9F9FF">

<br><table width="550"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="10">&nbsp;</td>
                      <td><span class="style4">公文传阅</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><%
set conn=opendb("oabusy","conn","accessdsn")
Set rs=Server.CreateObject("ADODB.recordset")
sql="select * from senddate,texttype where id=" & request("id")&" and senddate.documenttype=texttype.number"
rs.open sql,conn,1
%>
<center>
<table>
<tr>
<td align=center>
<span class="style7"><font size="+1"><%=rs("title")%></font></span>
<br>
<span class="style6">（<%=server.htmlencode(rs("typename"))%>）</font>
</span></td>
<td rowspan="2">
<%
if rs("filename")<>"" then
%>
<a href="upload/<%=rs("filename")%>" target="_blank"><img src="images/attach.gif" width=30 height=30 border=0></a>
<%
end if
%>
</td>
<tr>
<td>
[发送日期：<%=rs("inputdate")%>][接收人所在部门：<%=rs("recipientuserdept")%>]

<%
if rs("recipientusername")="所有人" then
%>
[接收者：<%=rs("recipientusername")%>]
<%
else
set conn=opendb("oabusy","conn","accessdsn")
Set rs1=Server.CreateObject("ADODB.recordset")
sql="select name from userinf where username=" & sqlstr(rs("recipientusername"))
rs1.open sql,conn,1
if not rs1.eof and not rs1.bof then
%>
[接收者：<%=rs1("name")%>]
<%
end if
end if
%>
</td>
</tr>
</table>
</center>

&nbsp;
<div align="center"><br>
  <%=keepformat(rs("content"))%> <br>
  <%
'打开数据库读出回复
set conn=opendb("oabusy","conn","accessdsn")
Set rs2=Server.CreateObject("ADODB.recordset")
sql="select * from senddate where reid=" & request("id") & " order by id desc"
rs2.open sql,conn,1
while not rs2.bof and not rs2.eof
%>
  <br>
  <br>
  -----------------------------------<br>
  <%=rs2("title")%><br>
  [回复时间：<%=rs2("inputdate")%>] 
  <%
set conn=opendb("oabusy","conn","accessdsn")
Set rs3=Server.CreateObject("ADODB.recordset")
sql="select userdept,name from userinf where username=" & sqlstr(rs2("sender"))
rs3.open sql,conn,1
if not rs3.eof and not rs3.bof then
%>
  [回复部门：<%=rs3("userdept")%>][回复者：<%=rs3("name")%>]<br>
  <br>
  <%=keepformat(rs2("content"))%> 
  <%
end if
rs2.movenext
wend
%>
  <%

%>
</div></td>
        </tr>
    </table></td>
  </tr>
</table>

</body>
</html>