<!--#include file="cyconn.asp"-->
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../admin_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from url order by linkidorder",conn,1,1
		dim i
		i=rs.recordcount%>
<table width="96%" align="center" cellpadding="2" cellspacing="1" border="0" class="border">
<tr> 
<td colspan="4" align="center"  class="topbg"><b>常用网址管理</font></b></td>
</tr>
<tr class="tdbg"> 
<td width="30%" align="center" class="tdbg"><strong>网站名称</strong></td>
<td width="30%" align="center" class="tdbg"><strong>网站地址</strong></td>
</tr>
			<%if rs.eof and rs.bof then
			response.write "还没有数据，请添加！"
			else
			do while not rs.eof%>
<tr class="tdbg"> 
<td align="center" class="tdbg"><a href="http://<%=trim(rs("linkurl"))%>" target="_blank"><%=trim(rs("linkname"))%></a>
</td>
<td align="center" class="tdbg"><a href="http://<%=trim(rs("linkurl"))%>" target="_blank"><%=trim(rs("linkurl"))%></a>
</td>
</tr>
<%rs.movenext
		  loop
		  end if
		  rs.close
		  set rs=nothing%>
</table>
<br>
</body>
</html>