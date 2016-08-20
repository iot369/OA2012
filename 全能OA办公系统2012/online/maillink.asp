
<%
function maillink(a)
	if a<>"" then
		maillink="<a href=mailto:" & a & ">" & a & "</a>"
	else
		maillink="нч"
	end if
end function
%>
