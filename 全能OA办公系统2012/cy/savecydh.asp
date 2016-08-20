<!--#include file="cyconn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('网络超时或您还没有登陆！');window.location.href='login.asp';</script>"
response.End
else
if session("flag")<5 then
response.Write "<p align=center><font color=red>您没有此项目管理权限！</font></p>"
response.End
end if
end if
 dim action,id
id=request.QueryString("id")
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from tel",conn,1,3
rs.AddNew
rs("cytel")=trim(request("cytel1"))
rs("cyname")=trim(request("cyname1"))
rs("gstel")=trim(request("gstel1"))
rs("gsname")=trim(request("gsname1"))
rs("khtel")=trim(request("khtel1"))
rs("khname")=trim(request("khname1"))
rs("idorder")=int(request("idorder1"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "admin_cydh.asp"
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from tel where id="&id,conn,1,3
rs("cytel")=trim(request("cytel"))
rs("cyname")=trim(request("cyname"))
rs("gstel")=trim(request("gstel"))
rs("gsname")=trim(request("gsname"))
rs("khtel")=trim(request("khtel"))
rs("khname")=trim(request("khname"))
rs("idorder")=int(request("idorder"))
rs.update
rs.close
response.Redirect "admin_cydh.asp"
set rs=nothing
case "del"
conn.execute "delete from tel where id="&id
response.Redirect "admin_cydh.asp"

end select

%>