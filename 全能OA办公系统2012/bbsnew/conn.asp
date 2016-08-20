<%cn="6k"
Server.ScriptTimeout="10"
Set myconn = Server.CreateObject("ADODB.Connection")
connstr="driver={Microsoft Access Driver (*.mdb)};dbq=" & Server.MapPath("db\bbs.mdb")
myconn.open connstr
%>