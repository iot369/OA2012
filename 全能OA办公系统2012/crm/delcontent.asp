<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
id=request("id")
printflag=request("printflag")
if id="" then
	response.write("<script language=""javascript"">")
	response.write("alert(""�Բ������ݴ������"");")
	response.write("window.close();")
	response.write("</script>")
else
	set conn=dbconn("conn")
	set rs=server.createobject("adodb.recordset")
	sql="delete * from qiye where id="&id
conn.execute sql
		response.write("ɾ���ɹ���")
		conn.close
		set rs=nothing
		set conn=nothing

end if
%>

<input type="button" value=" �� �� " onclick="javascript:window.close();">