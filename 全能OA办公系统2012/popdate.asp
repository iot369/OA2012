<%response.expires=0%>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->

<%
oabusyusername=request.cookies("oabusyusername")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("alert('�Բ������Ѿ����ڣ������µ�¼��');")
	response.write("window.close();")
	response.write("</script>")
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<meta http-equiv="pragma" content="no-cache">
<script language="javascript">
;
function redotopprg()
{
	
//	a=opener.parent("banner2").doflag.value;
//	alert(a);
}
</script>
<title>���Ľ�����ʾ</title>
<style type="text/css">
<!--
.style4 {color: #2e4869}
.style6 {color: #FF0000}
.style7 {font-weight: bold}
-->
</style>
</head>
<bgsound src="xbmsg.wav" loop="1">
<body bgcolor="#F9F9FF" topmargin="5" leftmargin="5" onunload="redotopprg();">		
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
                      <td><span class="style4">���Ĵ���</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><%

if request.form("submit")="�ظ�" then
	sender=request.form("sender")
	recipientusername=request.form("recipientusername")
	reid=request.form("reid")
	title=request.form("title")
	content=request.form("content")
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from senddate,texttype where id=" & request("id")&" and senddate.documenttype=texttype.number"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		sql="Insert into senddate (title,content,sender,recipientusername,reid) values ("
		sql=sql & sqlstr(title) & ","
		sql=sql & sqlstr(content) & ","
		sql=sql & sqlstr(sender) & ","
		sql=sql & sqlstr(recipientusername) & ","
		sql=sql & reid & ")"
		conn.Execute sql
		conn.close
		set conn=nothing
%>
<SCRIPT language=JavaScript>                   
	
	window.close();
</script> 
<%
		response.end
	else
%>
<table width="100%"><tr><td></td></tr></table>
<%

		response.write("<center><br><br><font color=""#ee0000"" size=""+1"">�Բ��𣬸ù����ѱ�ɾ�������ڲ��ܻظ���</font><br><br>")
		response.write("<input type=""button"" value=""�ر�"" onclick=""window.close()""></center>")
		%>
    
<%
		conn.close
		set conn=nothing
		response.end
	end if
end if
set conn=opendb("oabusy","conn","accessdsn")
Set rs=Server.CreateObject("ADODB.recordset")
sql="select * from senddate,texttype where id=" & request("id")&" and senddate.documenttype=texttype.number"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<center>
<table>
<tr>
<td align=center>
<span class="style7"><font size="+1"><%=keepformat(rs("title"))%></font></span>
<br>
<span class="style6">��<%=server.htmlencode(rs("typename"))%>��</font>
</span></td>
<tr>
<td>[���ڣ�<%=rs("inputdate")%>][���������ڲ��ţ�
<%
set rs1=Server.CreateObject("ADODB.recordset")
sql="select userdept,name from userinf where username=" & sqlstr(rs("sender"))
rs1.open sql,conn,1
if not rs1.eof and not rs1.bof then
response.write(rs1("userdept"))
%>
][�����ߣ�<%=rs1("name")%>]
<%end if%>
</td>
</tr>
</table>
</center>

&nbsp;
<div align="center">
  <!--#include file="showfile.asp"-->
  <br>
</div>
<div align="center"><br>
  <%=keepformat(rs("content"))%> 
</div>
<center>
<form method="post" name="form1" action="popdate.asp?id=<%=request("id")%>">
<input type="hidden" name="title" value="Re:<%=server.htmlencode(rs("title"))%>">
<input type="hidden" name="sender" value="<%=oabusyusername%>">
<input type="hidden" name="recipientusername" value="<%=rs("sender")%>">
<input type="hidden" name="reid" value="<%=rs("id")%>">
<textarea name="content" rows="15" cols="50"></textarea><br>
<input type="submit" name="submit" value="�ظ�">
</form>
</center>
<%
else
%>
<table width="100%"><tr><td></td></tr></table>
<%
	
	response.write("<center><br><br><font color=""#ee0000"" size=""+1"">�Բ��𣬸ù����ѱ�ɾ����</font><br><br>")
	response.write("<input type=""button"" value=""�ر�"" onclick=""window.close()""></center>")
end if
%>
<%
conn.close
set conn=nothing
%></td>
        </tr>
    </table></td>
  </tr>
</table>

</body>
</html>
