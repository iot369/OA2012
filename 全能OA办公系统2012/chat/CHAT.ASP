<%Response.Expires=0%>

<html>
<head>
<meta http-equiv="Content-Type" content='text/html; charset=gb2312'>
<title><%=Session("CRNAME")%></title>
<frameset cols='*,100' border=0 frameborder=0 framespacing=0>
 <frameset rows='*,80' border=0 frameborder=0 framespacing=0>
  <frame src='f1.asp#bottom' name='f1'>
  <frame src='f2.asp' name='f2' scrolling='no'>
 </frameset>
 <frameset rows='*,90' border=0 frameborder=0 framespacing=0 >
  <frame src='f3.asp' name='f3' marginheight=10 marginwidth=5>
  <frame src='f4.asp' name='f4' marginwidth=0 scrolling='no'>
 </frameset>
</frameset>
</head>
</html>
