<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
set conn=dbconn("conn")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title></title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<br>
<p align="center"><b><font size="+1">企业名录查询</font></b></p>
<form method="POST" action="dispinfo.asp?typenumber=3&lookstr=企业名录查询&page=1">
  <div align="center">
    <center>
    <table border="1" cellpadding="0" cellspacing="0" width="430" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
      <tr>
        <td height="25" width="59" align="center">
          <p align="center"><input type="checkbox" name="C1" value="ON"></td>
        <td height="25" width="363">&nbsp;企业名称：<input type="text" name="T1" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C2" value="ON"></td>
        <td height="25" width="363">&nbsp;联 系 人：<input type="text" name="T2" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C3" value="ON"></td>
        <td height="25" width="363">&nbsp;产品名称：<input type="text" name="T3" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C4" value="ON"></td>
        <td height="25" width="363">&nbsp;地&nbsp;&nbsp;&nbsp; 址：<input type="text" name="T4" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C5" value="ON"></td>
        <td height="25" width="363">&nbsp;电&nbsp;&nbsp;&nbsp; 话：<input type="text" name="T5" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C6" value="ON"></td>
        <td height="25" width="363">&nbsp;传&nbsp;&nbsp;&nbsp; 真：<input type="text" name="T6" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C7" value="ON"></td>
        <td height="25" width="363">&nbsp;邮&nbsp;&nbsp;&nbsp; 编：<input type="text" name="T7" size="20" style="width: 296; height: 22" class="doc_txt"></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C8" value="ON"></td>
        <td height="25" width="363">&nbsp;省&nbsp;&nbsp;&nbsp; 份：<select size="1" name="D1">
		<option value="">请选择</option>
<%
	set rs=server.createobject("adodb.recordset")
	sql="select * from diqu"
	rs.open sql,conn,1
	do while not rs.eof	
		response.write("<option value="&chr(34)&trim(rs("diqu"))&chr(34)&">"&trim(rs("diqu"))&"</option>")
		rs.movenext
	loop
%>
          </select></td>
      </tr>
      <tr>
        <td height="25" width="59" align="center"><input type="checkbox" name="C9" value="ON"></td>
        <td height="25" width="363">&nbsp;类&nbsp;&nbsp;&nbsp; 别：<select size="1" name="D2">
		<option value="">请选择</option>
<%
	set rs=nothing
	set rs=server.createobject("adodb.recordset")
	sql="select * from fenlei"
	rs.open sql,conn,1
	do while not rs.eof	
		response.write("<option value="&chr(34)&trim(rs("leibie"))&chr(34)&">"&trim(rs("leibie"))&"</option>")
		rs.movenext
	loop
	conn.close
	set conn=nothing
	set rs=nothing
%>
          </select></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value=" 查询 " name="submit"></p>
</form>

</body>

</html>
