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
'�����ݿ⣬�������û��Ƿ񿴹���ͨ��
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
<title>OA�칫ϵͳ</title>
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
'�����յ��ļ����ǻظ��ļ�ʱ�����û�ظ��͵�������
'�򿪹������ݿ⣬�����������Ǳ��˵Ļ򱾲��������˵Ĳ�reidΪ0���ҹ��ķ���ʱ����û�����ʱ����Ĺ��ļ�¼
set rs=nothing
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from senddate where reid=0 and (recipientusername=" & sqlstr(oabusyusername) & " or (recipientusername='������' and recipientuserdept=" & sqlstr(oabusyuserdept) & ")) and inputdate>#" & joindate & "#"
rs.open sql,conn,1
while not rs.bof and not rs.eof
'�򿪹������ݿ⣬�����������Ǳ��˲�reid���ڽ��չ��ĵ�id�ļ�¼
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select * from senddate where sender=" & sqlstr(oabusyusername) & " and reid=" & rs("id")
	rs1.open sql,conn,1
'����޼�¼�͵������ڲ���ֹ����
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
'ѭ��
	rs.movenext
wend
'-------------------------------------------------
'�����յ��ļ��ǻظ��ļ�ʱ�����û���͵�������
'�򿪹������ݿ⣬�����������Ǳ��˵Ļ򱾲��������˵Ĳ�reid��Ϊ0�Ĺ��ļ�¼
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from senddate where reid<>0 and (recipientusername=" & sqlstr(oabusyusername) & " or (recipientusername='������' and recipientuserdept=" & sqlstr(oabusyuserdept) & "))"
rs.open sql,conn,1
while not rs.bof and not rs.eof
'���ѿ��������ݿ⣬�����û����Ǳ��˲�senddateid���ڽ��չ��ĵ�id,��haveseeΪ��yes���ļ�¼
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select * from seesenddate where havesee='yes' and username=" & sqlstr(oabusyusername) & " and senddateid=" & rs("id")
	rs1.open sql,conn,1
'����޼�¼�͵������ڲ���ֹ����
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
'ѭ��
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
<title>�Զ�ˢ��ҳ��</title>
</head>
<body STYLE="background-color:transparent">
</body>
</html>

  