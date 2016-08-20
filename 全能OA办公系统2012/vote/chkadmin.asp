<%
If session("MoonDowner_Poll") <> "MoonDowner_Poll" Then
	Response.redirect "login.asp"
End If
%>