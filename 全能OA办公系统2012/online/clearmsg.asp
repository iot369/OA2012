<!--#INCLUDE FILE="conn.asp" -->
<%
set rs=server.createobject("ADODB.recordset") 
rs.open "select * from msg where send='"&session("Uid")&"' or receive='"&session("Uid")&"' order by id",conn,1,3
if not rs.eof then
do while not (rs.eof or rs.bof)
rs.Delete        
rs.movenext 
loop 
end if
rs.close
set rs=nothing
conn.Close
set conn = nothing
Response.Redirect ("sendinfo.asp?receiveuser="&session("receiveuser")&"&id="&session("receive"))

%>

