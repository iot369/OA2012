<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
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
if request("submit")="Ӧ��" then
allow_edit_all_jobchanginf=request("allow_edit_all_jobchanginf")
if allow_edit_all_jobchanginf="" then allow_edit_all_jobchanginf="no"

allow_edit_dept_jobchanginf=request("allow_edit_dept_jobchanginf")
if allow_edit_dept_jobchanginf="" then allow_edit_dept_jobchanginf="no"

set conn=opendb("oabusy","conn","accessdsn")
sql="update userinf set "
sql=sql & "allow_edit_dept_jobchanginf=" & sqlstr(allow_edit_dept_jobchanginf) & ","
sql=sql & "allow_edit_all_jobchanginf=" & sqlstr(allow_edit_all_jobchanginf) & " where id=" & request("id")
conn.Execute sql

end if
%>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='" & oabusyusername&"'"
rs.open sql,conn,1
cook_allow_control_all_user=rs("allow_control_all_user")     
conn.close
set conn=nothing
set rs=nothing
if cook_allow_control_all_user="no" then
response.write("<font color=red size=""+1"">�Բ�����û�����Ȩ�ޣ�</font>")
	response.end
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>oa�칫ϵͳ</title>
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
                    <td class="style7">�û�����</td>
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
<td>�༭Ա��ְ��䶯����Ȩ������&nbsp;&nbsp;&nbsp;&nbsp;</td>
<%
'�����ݿ⣬��������
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
%>
<form method="post" action="jobchanginfpopedom.asp">
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
</select><input type="submit" value="�鿴">
</td>
</form>
</tr>
</table>
  <br>
</center>

<br>
<center>
<table border="1"  cellspacing="0" cellpadding="0" width="95%" bgcolor="#FFFFFF" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
<tr bgcolor="D7E8F8"><td height=25 align=center>����</td>
<td align=center>����</td>
<td align=center>ְλ</td>
<td align=center>�ɱ༭����Ա��ְ�񵵰�</td>
<td align=center>�ɱ༭������Ա��ְ�񵵰�</td>
<td>&nbsp;</td>
</tr>
<%
'��ʾ�û���
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where userdept=" & sqlstr(userdept)
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<form method="post" action="jobchanginfpopedom.asp">
<tr>
<td align=center><%=rs("name")%></td><td align=center><%=rs("userdept")%></td><td align=center><%=rs("userlevel")%></td>
<td align=center><input type="checkbox" name="allow_edit_all_jobchanginf" value="yes"<%=checked(rs("allow_edit_all_jobchanginf"),"yes")%>></td>
<td align=center><input type="checkbox" name="allow_edit_dept_jobchanginf" value="yes"<%=checked(rs("allow_edit_dept_jobchanginf"),"yes")%>></td>
<td align=center><input type="submit" name="submit" value="Ӧ��"></td>
</tr>
<input type="hidden" name="id" value=<%=rs("id")%>>
<input type="hidden" name="userdept" value=<%=userdept%>>
</form>
<%
rs.movenext
wend
%>
</table>
</center>
<br>
<%

%>

</body>
</html>