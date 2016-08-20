<%
function openconn(sessionname)
	dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	DBPath1=server.mappath("../kq/"&cstr(year(date()))&".mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
	set session(sessionname)=conn
	set openconn=session(sessionname)
end function
function opennewdb(sessionname,yearvalue)
	dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	DBPath1=server.mappath("../kq/"&yearvalue&".mdb")
	conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
	set session(sessionname)=conn
	set opennewdb=session(sessionname)
end function

%>