
<%
function dbconn(sessionname)
	dim conn
	Set conn=Server.CreateObject("ADODB.Connection")
	conn.Open "Driver={Microsoft Access Driver (*.mdb)}; DBQ=" & Server.MapPath("data/data.mdb")
	set session(sessionname)=conn
	set dbconn=session(sessionname)
end function
%>