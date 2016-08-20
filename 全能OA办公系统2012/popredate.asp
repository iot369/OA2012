<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->

<%
oabusyusername=request.cookies("oabusyusername")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("alert('对不起，您已经过期，请重新登录！');")
	response.write("window.close();")
	response.write("</script>")
	response.end
end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<meta http-equiv="expires" content="no-cache">
<title>公文回复提示</title>
<script language="javascript">
;
</script>
<style type="text/css">
<!--
.style4 {color: #2e4869}
.style6 {color: #FF0000}
.style7 {font-weight: bold}
-->
</style>
</head>
<bgsound src="xbmsg.wav" loop="1">
<body bgcolor="#ffffff" topmargin="5" leftmargin="5" >
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
                      <td><span class="style4">公文传阅</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><%

if request.form("submit")="我已经看过" then
	id=request.form("id")
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from senddate where ID="&id
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		sql="Insert into seesenddate (senddateid,username,havesee) values ("
		sql=sql & id & ","
		sql=sql & sqlstr(oabusyusername) & ","
		sql=sql & sqlstr("yes") & ")"
		conn.Execute sql
%>
<SCRIPT language=JavaScript>                   
	window.close();
</script> 

<%
		response.end
	end if
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from senddate where id=" & request("id")
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<center><br>
<%=rs("title")%>
<br>[回复时间：<%=rs("inputdate")%>]
[回复人所在部门：
<%
	set rs1=Server.CreateObject("ADODB.recordset")
	sql="select userdept,name from userinf where username=" & sqlstr(rs("sender"))
	rs1.open sql,conn,1
%>
<%=rs1("userdept")%>
][回复者：<%=rs1("name")%>]
</center>

&nbsp;
<div align="center"><br>
  <%=keepformat(rs("content"))%>
</div>
<center>
<form method="post" name="form1" action="popredate.asp?id=<%=request("id")%>">
<input type="hidden" name="id" value="<%=rs("id")%>">
<input type="submit" name="submit" value="我已经看过">
</form>
</center>
<div align="center"><br>
  -----------------------------------<br>
</div>
<div align="center">原公文标题:
  <%
	'打开数据库，显示id=rs("reid")的记录
	set rs2=Server.CreateObject("ADODB.recordset")
	sql="select * from senddate where id=" & rs("reid")
	rs2.open sql,conn,1
	if not rs2.eof and not rs2.bof then
		response.write(rs2("title"))
       <!-- <%
if rs2("filename")<>"" then
%>
          <%

else
%>
        &nbsp; 
          <%
end if
%>
          <!--#include file="showfile.asp"-->
  <br>
  <%
		response.write(keepformat(rs2("content")))
	end if
else
%>
</div>
<table width="100%"><tr><td></td></tr></table>
<div align="center">
  <%
	
	response.write("<center><br><br><font color=""#ee0000"" size=""+1"">对不起，该公文已被删除！</font><br><br>")
	response.write("<input type=""button"" value=""关闭"" onclick=""window.close()""></center>")
end if
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