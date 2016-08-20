<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<%
oabusyusername=request.cookies("oabusyusername")
oabusyuserid=request.cookies("oabusyuserid")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserid=request.cookies("oabusyuserid")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if
'打开数据库，读出此用户是否看过此通告
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
set rs1=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs1.open sql,conn,1
joindate=rs1("joindate")
sql="select * from newnotice where readuserid NOT LIKE '%("&oabusyuserid&")%' and sendusername<>'"&oabusyusername&"'"
rs.open sql,conn,1
if rs.recordcount>0 then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>OA办公系统</title>
</head>
<body STYLE="background-color:transparent">
<SCRIPT language=JavaScript>                   
window.open('popnotice.asp','NewWin1','scrollbars=yes,width=640,height=400');
</script> 
</body>
</html>
<%
	set rs=nothing
	conn.close
	set conn=nothing
	response.end
end if
'--------------------------------------------------
'所接收的文件不是回复文件时，如果没回复就弹出窗口
'打开公文数据库，读出接收人是本人的或本部门所有人的并reid为0而且公文发布时间比用户建立时间晚的公文记录
set rs=nothing
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from senddate where reid=0 and (recipientusername=" & sqlstr(oabusyusername) & " or (recipientusername='所有人' and recipientuserdept=" & sqlstr(oabusyuserdept) & ")) and inputdate>#" & joindate & "#"
rs.open sql,conn,1
while not rs.bof and not rs.eof
'打开公文数据库，读出发送人是本人并reid等于接收公文的id的记录
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select * from senddate where sender=" & sqlstr(oabusyusername) & " and reid=" & rs("id")
	rs1.open sql,conn,1
'如果无记录就弹出窗口并终止程序
	if rs1.eof or rs1.bof then
%>
<html>a
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>New Page 1</title>
</head>
<body STYLE="background-color:transparent">
<SCRIPT language=JavaScript>                   
window.open('popdate.asp?id=<%=rs("id")%>','NewWin1','scrollbars=yes,width=640,height=500');       </script> 
</body>
</html>
<%
		response.end
	end if
'循环
	rs.movenext
wend
'-------------------------------------------------
'所接收的文件是回复文件时，如果没看就弹出窗口
'打开公文数据库，读出接收人是本人的或本部门所有人的并reid不为0的公文记录
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from senddate where reid<>0 and (recipientusername=" & sqlstr(oabusyusername) & " or (recipientusername='所有人' and recipientuserdept=" & sqlstr(oabusyuserdept) & "))"
rs.open sql,conn,1
while not rs.bof and not rs.eof
'打开已看公文数据库，读出用户名是本人并senddateid等于接收公文的id,且havesee为“yes”的记录
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select * from seesenddate where havesee='yes' and username=" & sqlstr(oabusyusername) & " and senddateid=" & rs("id")
	rs1.open sql,conn,1
'如果无记录就弹出窗口并终止程序
	if rs1.eof or rs1.bof then
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>New Page 1</title>
</head>
<body STYLE="background-color:transparent">
<SCRIPT language=JavaScript>
window.open("popredate.asp?id=<%=rs("id")%>",'NewWin1','scrollbars=yes,width=640,height=400');     
</script> 
</body>
</html>
<%
		response.end
	end if
'循环
	rs.movenext
wend
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="expires" content="no-cache">
<script language="javascript">
setTimeout("location.href='top0.asp'",30000);
</script>
<title>自动刷新页面</title>
</head>
<body STYLE="background-color:transparent">
</body>
</html>

  