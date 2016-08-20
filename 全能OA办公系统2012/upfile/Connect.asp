<%
Dim conn
Dim connstr
Dim ServerPath

ServerPath=Server.MapPath("data.mdb")
connstr="Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & ServerPath
Set conn=Server.CreateObject("Adodb.Connection")
conn.Open connstr
%>