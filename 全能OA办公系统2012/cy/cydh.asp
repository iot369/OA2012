<!--#include file="cyconn.asp"-->
<html><head><title>Untitled Document</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link href="../admin_style.css" rel="stylesheet" type="text/css">
</head>
<body>
<%set rs=server.CreateObject("adodb.recordset")
		rs.open "select * from tel order by idorder",conn,1,1
		dim i
		i=rs.recordcount
		%>
<table width="96%" align="center" cellpadding="2" cellspacing="1" border="0" class="border">
  <tr align="center" valign="middle">
    <td height="20" colspan="8" align="center" class="topbg"><b>���õ绰����</b></td>
  </tr>
  <tr align="center" class="tdbg">
    <td class="tdbg"  align="center"><strong>��������</strong></td>
    <td class="tdbg"  align="center"><strong>���õ绰</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ�����</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ��绰</strong></td>
	<td class="tdbg"  align="center"><strong>ͬ������</strong></td>
    <td class="tdbg"  align="center"><strong>��˾�绰</strong></td>
  </tr>
  <%if rs.eof and rs.bof then
			response.write "��û�����ݣ�����ӣ�"
			else
			do while not rs.eof
			%>
			
  <tr align="center"> 
    <td  align="center" class="tdbg"><%=trim(rs("cyname"))%></td>
    <td  align="center" class="tdbg"><%=trim(rs("cytel"))%></td>
    <td class="tdbg"  align="center"><%=trim(rs("khname"))%></td>
    <td class="tdbg"  align="center"><%=trim(rs("khtel"))%></td>
	<td class="tdbg"  align="center"><%=trim(rs("gsname"))%></td>
    <td class="tdbg"  align="center"><%=trim(rs("gstel"))%></td>
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