<!--#include file="cyconn.asp"-->
<%if session("admin")="" then
response.Write "<script language='javascript'>alert('���糬ʱ������û�е�½��');window.location.href='login.asp';</script>"
response.End
else
if session("flag")<9 then
response.Write "<p align=center><font color=red>��û�д���Ŀ����Ȩ�ޣ�</font></p>"
response.End
end if
end if
%>
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
    <td class="tdbg"  align="center"><strong><strong>����</strong>����</strong></td>
    <td class="tdbg"  align="center"><strong>���õ绰</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ�����</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ��绰</strong></td>
	<td class="tdbg"  align="center"><strong>ͬ������</strong></td>
    <td class="tdbg"  align="center"><strong>��˾�绰</strong></td>
    <td class="tdbg"  align="center"><strong>����</strong></td>
    <td class="tdbg"  align="center"><strong>����</strong></td>
  </tr>
  <%if rs.eof and rs.bof then
			response.write "��û�����ݣ�����ӣ�"
			else
			do while not rs.eof
			%>
			
  <tr align="center">
  <form name="form1" method="post" action="savecydh.asp?action=edit&id=<%=rs("id")%>">
  
    <td  align="center" class="tdbg"><input name="cyname" type="text" id="cyname" value="<%=trim(rs("cyname"))%>" size="10"></td>
    <td  align="center" class="tdbg"><input name="cytel" type="text" id="cytel" value="<%=trim(rs("cytel"))%>" size="16"></td>
    <td class="tdbg"  align="center"><input name="khname" type="text" id="khname" value="<%=trim(rs("khname"))%>" size="10"></td>
    <td class="tdbg"  align="center"><input name="khtel" type="text" id="khtel" value="<%=trim(rs("khtel"))%>" size="16"></td>
	<td class="tdbg"  align="center"><input name="gsname" type="text" id="gsname" value="<%=trim(rs("gsname"))%>" size="10"></td>
    <td class="tdbg"  align="center"><input name="gstel" type="text" id="gstel" value="<%=trim(rs("gstel"))%>" size="16"></td>
    <td class="tdbg"  align="center"><input name="idorder" type="text" id="idorder" value=<%=rs("idorder")%> size="3"></td>
    <td class="tdbg"  align="center"><input type="submit" name="Submit" value="�� ��"><a href=savecydh.asp?action=del&id=<%=rs("id")%>><font color="#FF0000">ɾ��</font></a></td>
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
<form name="form2" method="post" action="savecydh.asp?action=add">
  <tr align="center" valign="middle">
    <td height="20" colspan="8" align="center" class="topbg"><b>���õ绰����</b></td>
  </tr>
  <tr align="center" class="tdbg">
    <td class="tdbg"  align="center"><strong>��ϵ����</strong></td>
    <td class="tdbg"  align="center"><strong>���õ绰</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ�����</strong></td>
    <td class="tdbg"  align="center"><strong>�ͻ��绰</strong></td>
	<td class="tdbg"  align="center"><strong>ͬ������</strong></td>
    <td class="tdbg"  align="center"><strong>��˾�绰</strong></td>
    <td class="tdbg"  align="center"><strong>����</strong></td>
    <td class="tdbg"  align="center"><strong>����</strong></td>
  </tr>
  <tr align="center">
    <td class="tdbg"  align="center"><input name="cyname1" type="text" id="cyname1"  size="10"></td>
    <td class="tdbg"  align="center"><input name="cytel1" type="text" id="cytel1"  size="16"></td>
	<td class="tdbg"  align="center"><input name="khname1" type="text" id="khname1"  size="10"></td>
    <td class="tdbg"  align="center"><input name="khtel1" type="text" id="khtel1"  size="16"></td>
	<td class="tdbg"  align="center"><input name="gsname1" type="text" id="gsname1"  size="10"></td>
    <td class="tdbg"  align="center"><input name="gstel1" type="text" id="gstel1"  size="16"></td>
    <td class="tdbg"  align="center"><input name="idorder1" type="text" id="idorder1" value=<%=i+1%> size="3"></td>
    <td class="tdbg"  align="center"><input type="submit" name="Submit2" value="��ӵ绰"></td>
  </form>
	</tr> 			
</table>
</body>
</html>