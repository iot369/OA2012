<%
function keepformat(content)
	if typename(content)="Null" then
		keepformat=""
	else
		str=replace(content," ","&nbsp;")
		str=replace(str,"<","&lt")
		str=replace(str,">","&gt")
		str=replace(str,chr(13)+chr(10),"<br>")
		keepformat=str
	end if
end function
%>
