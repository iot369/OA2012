<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/check_resource.asp"-->
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
checkflag=check_resource_setting(oabusyusername,1)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA�칫ϵͳ.��Ե�ر��</title>
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
                    <td class="style7">������Դ</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
</table>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td> <table width="1%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
        <div align="center">��ԴԤԼ��ϸ��Ϣ
          </center>

        </div>
        <center>
<br><br>
<%
id=request("id")
if request("submit")="ɾ��" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete * from booking where id=" & id
conn.Execute sql
%>
<br>�ɹ�ɾ��ԤԼ��<br><br>
<form method="post" action="booking.asp">
<input type="submit" value="����">
</form>
<%
else
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from booking where id=" & id
rs1.open sql,conn,1
if not rs1.bof and not rs1.eof then
%>
<table width=540 border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<tr bgcolor="#FFFFFF">
<td width=126 height="30" bgcolor="D7E8F8"><div align="center">ԤԼʹ����Դ���ƣ�</div></td>
<td width="411" height="30">��<%=rs1("equipment")%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td width="126" height="30" bgcolor="D7E8F8"><div align="center">Ԥ�ƿ�ʼʹ��ʱ�䣺</div></td>
<td width="411" height="30">��<%=rs1("starttime")%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td width="126" height="30" bgcolor="D7E8F8"><div align="center">Ԥ�ƽ���ʹ��ʱ�䣺</div></td>
<td width="411" height="30">��<%=rs1("endtime")%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height="30" colspan="2"> ��
  <%
'������Ա����
set conn=opendb("oabusy","conn","accessdsn")
set rs3=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(rs1("username"))
rs3.open sql,conn,1
if not rs3.eof and not rs3.bof then
if oabusyusername=rs1("username") then response.write "<font color=red>"
%>
ԤԼ�ˣ�<%=rs3("name")%>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;���ڲ��ţ�<%=rs3("userdept")%>
<%
if oabusyusername=rs1("username") then response.write "</font>"
end if
%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height=100 colspan="2" valign=top><br>
��ʹ��˵����<br> ��
<%=keepformat(rs1("purpose"))%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height="30" colspan="2"> ��
  <%
select case rs1("auditing")
	case 1
%>
��������<font color="#ff0000">ͬ��</font>&nbsp;&nbsp;&nbsp;����ˣ�<%=rs1("auditing_user")%>&nbsp;&nbsp;&nbsp;���ʱ�䣺<%=cstr(rs1("auditing_time"))%>
<%
	case 2
%>
��������<font color="#ff0000">��ͬ��</font>&nbsp;&nbsp;&nbsp;����ˣ�<%=rs1("auditing_user")%>&nbsp;&nbsp;&nbsp;���ʱ�䣺<%=cstr(rs1("auditing_time"))%>
<%
	case 0
%>
<font color="#ff0000">δ��ˣ�</font>
<%
end select
%></td>
</tr>
<tr bgcolor="#FFFFFF">
<td colspan="2">
<br>��������˵����<br> ��
<%
if rs1("auditing_explain")<>"" then
	response.write server.htmlencode(rs1("auditing_explain"))
end if
%><br></td>
</tr>
</table>
<form method="post" action="editbooking.asp">
<input type="button" value="����" onclick="window.location.href='booking.asp'">
<%
if oabusyusername=rs1("username") then
%>
<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('�����Ҫɾ������ԤԼ��')">
<input type="hidden" name="id" value=<%=id%>>
<%
end if
%>
</form>
<%
end if
if checkflag=0 and rs1("auditing")=0 then
	call writeidea("write_auditing_idea.asp",oabusyname,id)
end if
end if
%>
</center>
</td>
    </tr>
</table>
 
</body>
</html>




