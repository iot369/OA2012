<%
function opendb(DBPath,sessionname,dbsort)

	dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	DBPath1=server.mappath("db/lmtof.mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
	set session(sessionname)=conn
	set opendb=session(sessionname)
end function
%>
