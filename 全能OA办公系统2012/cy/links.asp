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
<td width="20%" align="center" class="tdbg"><strong>排 序</strong></td>
<td width="20%" align="center" class="tdbg"><strong>操 作</strong></td>
</tr>
			<%if rs.eof and rs.bof then
			response.write "还没有数据，请添加！"
			else
			do while not rs.eof%>
<tr class="tdbg"> 
<form name="form1" method="post" action="savelinks.asp?action=edit&id=<%=rs("linkid")%>">
<td align="center" class="tdbg"><input name="linkname" type="text" id="linkname" value="<%=trim(rs("linkname"))%>" size="16">
</td>
<td align="center" class="tdbg">
<input name="linkurl" type="text" id="linkurl" value="<%=trim(rs("linkurl"))%>" size="26">
</td>
<td align="center" class="tdbg">
<input name="linkidorder" type="text" id="linkidorder" value=<%=rs("linkidorder")%> size="3">
</td>
<td align="center" class="tdbg">
<input type="submit" name="Submit" value="修 改">
&nbsp;<a href=savelinks.asp?action=del&id=<%=rs("linkid")%>><font color="#FF0000">删除</font></a> 
</td>
</form>
</tr>
<%rs.movenext
		  loop
		  end if
		  rs.close
		  set rs=nothing%>
</table>
<br>
<table width="96%" align="center" cellpadding="2" cellspacing="1" border="0" class="border">
<form name="form2" method="post" action="savelinks.asp?action=add">
<tr>
<td colspan="4" align="center"  class="topbg"><b>添加网站地址</font></b></td>
</tr>
<tr bgcolor="#CCCCCC"> 
<td width="30%" align="center" class="tdbg"><strong>网站名称 </strong></td>
<td width="30%" align="center" class="tdbg"><strong>网站地址 </strong></td>
<td width="20%" align="center" class="tdbg"><strong>排 序 </strong></td>
<td width="20%" align="center" class="tdbg"><strong>操 作 </strong></td>
</tr>
<tr class="tdbg"> 
<td align="center" class="tdbg">
<input name="linkname1" type="text" id="linkname1" size="16">
</td>
<td align="center" class="tdbg">
<input name="linkurl1" type="text" id="linkurl1" size="26">
</td>
<td align="center" class="tdbg">
<input name="linkidorder1" type="text" id="linkidorder1" value=<%=i+1%> size="3">
</td>
<td align="center" class="tdbg">
<input type="submit" name="Submit2" value="添加">
</td>
</tr>
</form>
</table>
</body>
</html>