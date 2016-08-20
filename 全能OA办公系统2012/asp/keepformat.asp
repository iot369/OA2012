<%
function keepformat(content)
if typename(content)="Null" then
keepformat=""
else
str=replace(content," ","&nbsp;&nbsp;")
keepformat=replace(str,chr(13)+chr(10),"<br>")
end if
end function
%>