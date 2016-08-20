            <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
function opendb(DBPath,sessionname,dbsort)
dim conn
Set conn=Server.CreateObject("ADODB.Connection")
DBPath1=server.mappath("../db/sdoa.asa")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
set opendb=session(sessionname)
end function
'打开数据库，读出权限
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
'set conn=opendb("oabusy","conn","accessdsn")
'set rs=server.createobject("adodb.recordset")
'on error resume next
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='"&oabusyusername&"'"
rs.open sql,session("conn"),1,1
cook_allow_gwlz_manage=rs("allow_gwlz_manage")
cook_allow_send_file=rs("allow_send_file")
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
response.write("<font color=red size=""+1"">对不起，您没有这个权限！</font>")
	response.end
	end if
%>       