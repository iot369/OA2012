<!--#include file="cyconn.asp"-->
<%
 dim action,linkid
linkid=request.QueryString("id")
action=request.QueryString("action")
select case action
case "add"
set rs=server.CreateObject("adodb.recordset")
rs.Open "select * from url",conn,1,3
rs.AddNew
rs("linkname")=trim(request("linkname1"))
rs("linkurl")=trim(request("linkurl1"))
rs("linkidorder")=int(request("linkidorder1"))
rs.Update
rs.Close
set rs=nothing
response.Redirect "links.asp"
case "edit"
set rs=server.CreateObject("adodb.recordset")
rs.open "select * from url where linkid="&linkid,conn,1,3
rs("linkname")=request("linkname")
rs("linkurl")=trim(request("linkurl"))
rs("linkidorder")=int(request("linkidorder"))
rs.update
rs.close
response.Redirect "links.asp"
set rs=nothing
case "del"
conn.execute "delete from url where linkid="&linkid
response.Redirect "links.asp"

end select

%>