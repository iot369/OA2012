<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->

<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/checked.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
cook_allow_see_all_workrep=request.cookies("cook_allow_see_all_workrep")
cook_allow_see_dept_workrep=request.cookies("cook_allow_see_dept_workrep")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

'--------------------------
function havefinished(value)
if value="yes" then
havefinished="�����"
else
havefinished="<font color=red>δ���</font>"
end if
end function
'---------------------------
function impdegree(value)
if value="yes" then
impdegree="<font color=red>��Ҫ</font>"
else
impdegree="һ��"
end if
end function
'---------------------------
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<title>��������</title>
<style type="text/css">
<!--
.style1 {color: #0d79b3}
.style5 {color: #2b486a}
.style4 {color: #2e4869}
-->
</style>
</head>
<body  topmargin="5" leftmargin="5">
<center>
  <table width="550"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
    <tr>
      <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                  <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="10">&nbsp;</td>
                        <td><span class="style4">��������</span></td>
                      </tr>
                  </table></td>
                  <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
                </tr>
            </table></td>
          </tr>
          <tr>
            <td><%
username=request("username")
superior=request("superior")
recdate=request("recdate")
'�����ݿ�����û�����
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
if not rs.eof and not rs.bof then stafname=rs("name")
%>
                <center>
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="1" height="5"></td>
                    </tr>
                  </table>
                  <table border="0"  cellspacing="0" cellpadding="0">
                    <tr>
                      <td><span class="style1"><%=stafname%>�Ĺ�������<%=recdate%>��</span></td>
                      <form action="addworkrep.asp" method=post name="form1">
                        <%
if (username=oabusyusername) or ( username<>oabusyusername and (cook_allow_see_all_workrep="yes" or cook_allow_see_dept_workrep="yes")) then
%>
                        <td><input type="submit" name="addworkrep" value="����"></td>
                        <%
end if
%>
                        <input type="hidden" name="username" value="<%=username%>">
                        <input type="hidden" name="superior" value="<%=superior%>">
                        <input type="hidden" name="recdate" value="<%=recdate%>">
                      </form>
                    </tr>
                  </table>
                </center>
                <br>
                <center>
                  <%
'�����ݿ⣬��������Ϊrecdate���û���Ϊusername�ļ�¼
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from workrep where username=" & sqlstr(username) & " and recdate=" & "#" & recdate & "#"
rs.open sql,conn,1
while not rs.eof and not rs.bof
'��������ί��������
	if rs("superior")<>"" then
		set conn=opendb("oabusy","conn","accessdsn")
		set rs1=server.createobject("adodb.recordset")
		sql="select name from userinf where username=" & sqlstr(rs("superior"))
		rs1.open sql,conn,1
		superiorname=rs1("name")
	else
			superiorname="���˰���"
	end if
%>
                  <table width="98%" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
                    <tr>
                      <form method="post" name="a1" action="editworkrep.asp">
                        <td width=60 align=center bgcolor="D7E8F8"> <span class="style5">
                          <%
if (username=oabusyusername) or (superior=oabusyusername) then
%>
                          <input type="submit" value="�༭" name="submit">
                          <input type="hidden" name="username" value="<%=username%>">
                          <input type="hidden" name="superior" value="<%=superior%>">
                          <input type="hidden" name="recdate" value="<%=recdate%>">
                          <input type="hidden" name="id" value=<%=rs("id")%>>
                          <%
else
%>
                      ���ɱ༭</span> <span class="style5">
                      <%
end if
%>
                    </span></td>
                      </form>
                      <td width=60 height="30" bgcolor="D7E8F8"><div align="center" class="style5">������</div></td>
                      <td bgcolor="#FFFFFF"><div align="center" class="style5"><%=havefinished(rs("finished"))%></div></td>
                      <td width=60 bgcolor="D7E8F8"><div align="center" class="style5">��Ҫ�̶�</div></td>
                      <td bgcolor="#FFFFFF"><div align="center" class="style5"><%=impdegree(rs("imp"))%></div></td>
                      <td width=70 bgcolor="D7E8F8"><div align="center" class="style5">ί��������</div></td>
                      <td width=60 bgcolor="#FFFFFF"><div align="center" class="style5"><%=checked3(superiorname)%></div></td>
                    </tr>
                    <tr>
                      <td height="30" align=center bgcolor="D7E8F8"><span class="style5">��Ҫ����</span></td>
                      <td colspan="6" bgcolor="#FFFFFF"><span class="style5">��<%=rs("title")%></span></td>
                    </tr>
                    <tr>
                      <td height="30" align=center bgcolor="D7E8F8"><span class="style5">��ϸ˵��</span></td>
                      <td colspan="6" bgcolor="#FFFFFF"><span class="style5">��<%=checked3(keepformat(rs("remark")))%></span></td>
                    </tr>
                  </table>
                  <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="1" height="5"></td>
                    </tr>
                  </table>
                  <%


rs.movenext
wend
%>
                </center>
                <%

%></td>
          </tr>
      </table></td>
    </tr>
  </table>
</center>

</body>
</html>