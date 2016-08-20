<%
if trim(rs("filename"))<>empty then
  kobeoutput=split(rs("filename"),",")
  for i=0 to Ubound(kobeoutput)
  response.write "<a target=_blank href=upload/" & trim(kobeoutput(i)) & ">¸½¼þ" & i+1 & "</a>&nbsp;"
  next
end if
%>