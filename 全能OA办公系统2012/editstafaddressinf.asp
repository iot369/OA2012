<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/maillink.asp"-->
<!--#include file="asp/keepformat.asp"-->
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

'--------------------------------------
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA�칫ϵͳ.��Ե�ر��</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style6 {color: #FFFFFF}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="21"><div align="center"><span class="style2">
          <div align="center">
            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
                <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                      <td>ͨѶ����</td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
              </tr>
            </table>
          <font color="0D79B3"></font></div>
      </span></div></td>
    </tr>
</table>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><table width="1%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table align="center">
<tr>
<td>�༭<%=request("name")%>ͨѶ¼&nbsp;&nbsp;&nbsp;&nbsp;</td>
<form method="post" action="stafaddressinf.asp">
<td><input type="submit" name="submit" value="����"></td>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
</form>
</tr>
</table>



<center>

<%
'�������
if request("submit")="����" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "Insert Into stafaddressinf (username,companytel,fax,hometel,homeaddress,postcard,sex,handset,callset,remark,email) Values( "
sql = sql & SqlStr(request("username")) & ", "
sql = sql & SqlStr(request("companytel")) & ", "
sql = sql & SqlStr(request("fax")) & ", "
sql = sql & SqlStr(request("hometel")) & ", "
sql = sql & SqlStr(request("homeaddress")) & ", "
sql = sql & SqlStr(request("postcard")) & ", "
sql = sql & SqlStr(request("sex")) & ", "
sql = sql & SqlStr(request("handset")) & ", "
sql = sql & SqlStr(request("callset")) & ", "
sql = sql & SqlStr(request("remark")) & ", "
sql = sql & SqlStr(request("email")) & ")"
conn.Execute sql
%>
<br><br>
<font color="red" >����ɹ�!</font>

<%
else
if request("submit")="�޸�" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "update stafaddressinf set "
sql = sql & "companytel=" & SqlStr(request("companytel")) & ", "
sql = sql & "fax=" & SqlStr(request("fax")) & ", "
sql = sql & "hometel=" & SqlStr(request("hometel")) & ", "
sql = sql & "homeaddress=" & SqlStr(request("homeaddress")) & ", "
sql = sql & "postcard=" & SqlStr(request("postcard")) & ", "
sql = sql & "sex=" & SqlStr(request("sex")) & ", "
sql = sql & "handset=" & SqlStr(request("handset")) & ", "
sql = sql & "callset=" & SqlStr(request("callset")) & ", "
sql = sql & "remark=" & SqlStr(request("remark")) & ", "
sql = sql & "email=" & SqlStr(request("email")) & " where username=" & sqlstr(request("username"))
conn.Execute sql 
%>
<br><br>
<font color=red >�޸����!</font>
<%
else
if request("submit")="ɾ��" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete from stafaddressinf where username=" & sqlstr(request("username"))
conn.Execute sql
%>
<br><br>
<font color=red >�ɹ�ɾ��ͨѶ���ϣ�</font>

<%

else

'�����ݿ����ͨѶ��Ϣ
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from stafaddressinf where username=" & sqlstr(request("username"))
rs.open sql,conn,1
if not rs.eof and not rs.bof then
companytel=rs("companytel")
fax=rs("fax")
hometel=rs("hometel")
homeaddress=rs("homeaddress")
postcard=rs("postcard")
sex=rs("sex")
handset=rs("handset")
callset=rs("callset")
remark=rs("remark")
email=rs("email")
submit="�޸�"
else
companytel=""
fax=""
hometel=""
homeaddress=""
postcard=""
sex=""
handset=""
callset=""
remark=""
email=""
submit="����"
end if

%>

<br>
<form method="post" action="editstafaddressinf.asp">
    <table width="550"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
          </table><table width="550" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="B0C8EA">
      <tr>
        <td height=2 colspan="4" align="center" ></td>
        </tr>
      <tr>
    <td width="15%" height=25 align="center" bgcolor="D7E8F8" >Ա������</td>
    <td width="35%" bgcolor="#FFFFFF" ><%=request("name")%></td>
    <td width="15%" align="center" bgcolor="D7E8F8" >�Ա�</td>
    <td width="35%" bgcolor="#FFFFFF"  > <span class="style6">
      <select name="sex" size=1>
        <option value="��"<%=selected("��",sex)%>>��</option>
        <option value="Ů"<%=selected("Ů",sex)%>>Ů</option>
      </select>
    </span> </td>
      </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" ����������������>���ڲ���</td>
    <td bgcolor="#FFFFFF" ����������������><%=request("userdept")%></td>

    <td align="center" bgcolor="D7E8F8" ����������������>ְ��</td>
    <td bgcolor="#FFFFFF"  ><span class="style6"><%=request("userlevel")%></span></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" >�绰��ֻ�</td>
    <td bgcolor="#FFFFFF" ><input type="text" name="companytel" size=20 value="<%=companytel%>"></td>
    <td align="center" bgcolor="D7E8F8" >����</td>
    <td bgcolor="#FFFFFF"  ><input name="fax" type="text" value="<%=fax%>" size=20></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" ����������������>�ֻ�����</td>
    <td bgcolor="#FFFFFF" ����������������><input type="text" name="handset" size=20 value="<%=handset%>"></td>
    <td align="center" bgcolor="D7E8F8" ����������������>MSN</td>
    <td bgcolor="#FFFFFF"  ><input name="callset" type="text" value="<%=callset%>" size=20></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" ����������������>סլ�绰</td>
    <td bgcolor="#FFFFFF" ����������������><input type="text" name="hometel" size=20 value="<%=hometel%>"></td>
    <td align="center" bgcolor="D7E8F8" ����������������>Email</td>
    <td bgcolor="#FFFFFF"  ><input name="email" type="text" value="<%=email%>" size=20></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" ����������������>�����ַ</td>
    <td colspan="3" bgcolor="#FFFFFF"  ><input type="text" name="homeaddress" size=57 value="<%=homeaddress%>"></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" ����������������>�����ʱ�</td>
    <td colspan="3" bgcolor="#FFFFFF"  ><input type="text" name="postcard" size=57 value="<%=postcard%>"></td>
  </tr>
  <tr>
    <td height=25 align="center" bgcolor="D7E8F8" >��ע˵��</td>
    <td colspan="3" bgcolor="#FFFFFF" ><textarea rows="2" name="remark" cols="56"><%=remark%></textarea>
</td>
  </tr>
</table>
<input type="submit" name="submit" value="<%=submit%>">&nbsp;&nbsp;<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷʵҪɾ�����ͨѶ������')">
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="hidden" name="name" value="<%=request("name")%>">
<input type="hidden" name="userlevel" value="<%=request("userlevel")%>">
</form>
<%
end if
end if
end if
%>
</center>
<%

%></td>
    </tr>
  </table>
  

</body>
</html>