<%dim ThisKey
ThisKey = "@"
%>

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" type="text/css" href="../CSS/main.css">
</head>
<%
if request("cmdUp")<>"" then
dim sentences(60)
for i=0 to ubound(sentences)
sentences(i)=""
next
Application("Room1sentences")=sentences
Application("Room2sentences")=sentences
Application("Room3sentences")=sentences
Application("Room4sentences")=sentences
response.write "<br><br><p align=center><font color=red>�Ѿ��ɹ������</font></p>"
response.end
end if
%>
<body>
<BR><BR><BR><BR>
<div align=center>
������°�Ť����������л����ҵĻ����¼��
<form name="form1" method="post" action=""> 
<INPUT TYPE="submit" name="cmdUp" class="font9boldwhite" value="������л����¼" onclick="javascript:return confirm('ȷ�������');">
</form>
</div>
</body>
</html>
