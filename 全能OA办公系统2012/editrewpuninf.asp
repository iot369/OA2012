<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
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
<title>oa�칫ϵͳ</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">Ա������</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<%
'�����ݿ�����û���
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(request("username"))
rs.open sql,conn,1
name=rs("name")
%>
<%
'�����ݿ�����û���
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(request("username"))
rs.open sql,conn,1
name=rs("name")
%>
<center>
  <br>
<table>
<tr>
<td>
�༭<%=name%>�Ľ�����Ϣ&nbsp;&nbsp;
</td>
<form method="post" action="rewpuninf.asp" name="form2"><td>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="submit" value="����">
</td>
</form>
</tr>
</table>
</center>

<%
if request("submit")="�޸�" then
username=request("username")
rewpunname=request("rewpunname")
rewpundate=request("rewpundate")
rewpunfile=request("rewpunfile")
rewpunsort=request("rewpunsort")
rewpuntype=request("rewpuntype")
remark=request("remark")
updatename=oabusyname
updatedate=now()
id=request("id")
set conn=opendb("oabusy","conn","accessdsn")
sql = "Update rewpuninf set "
sql = sql & "rewpunname=" & SqlStr(rewpunname) & ", "
sql = sql & "rewpundate=" & SqlStr(rewpundate) & ", "
sql = sql & "rewpunfile=" & SqlStr(rewpunfile) & ", "
sql = sql & "rewpunsort=" & SqlStr(rewpunsort) & ", "
sql = sql & "rewpuntype=" & SqlStr(rewpuntype) & ", "
sql = sql & "remark=" & SqlStr(remark) & ", "
sql = sql & "updatename=" & SqlStr(updatename) & ", "
sql = sql & "updatedate=#" & updatedate & "# where id=" & id
conn.Execute sql
%>
<br><br>
<center><font color=red >�ɹ��޸�Ա��������Ϣ��</font></center>
<%
else
if request("submit")="ɾ��" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete from rewpuninf where id=" & request("id")
conn.Execute sql
%>
<br><br>
<center><font color=red >�ɹ�ɾ��Ա��������Ϣ��</font></center>
<%
else
%>
<script Language="JavaScript">

 function form_check(){
   var l1=document.form1.rewpunname.value;
   if(l1==""){window.alert("�������Ʊ�����д��");document.form1.rewpunname.focus();return (false);}

   var l2=document.form1.rewpundate.value;
   if(l2==""){window.alert("�������ڱ�����д��");document.form1.rewpundate.focus();return (false);}
                    }

</script>

<%
'�����ݿ⣬����ְ��䶯��Ϣ
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from rewpuninf where id=" & request("id")
rs.open sql,conn,1
%>
<br>
<center>
<form method="post" action="editrewpuninf.asp" name="form1" onsubmit="return form_check();">
  <table border="0" cellpadding="0" cellspacing="0" width="540">
    <tr>
      <td height="25" width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">Ա������</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=name%>
      </td>
    </tr>
    <tr>
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">��������</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="rewpunname" size=45 value="<%=rs("rewpunname")%>"> 
      </td>
    </tr>
    <tr>
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">��������</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="rewpundate" size=10 value="<%=rs("rewpundate")%>"> 
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">�����ĺ�</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="rewpunfile" size=10 value="<%=rs("rewpunfile")%>"> 
      </td>
    </tr>
    <tr>
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">��������</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="rewpunsort" size=10 value="<%=rs("rewpunsort")%>">
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">���ͷ�ʽ</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="rewpuntype" size=10 value="<%=rs("rewpuntype")%>"> 
      </td>
    </tr>
    <tr>
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" align="center">����ԭ��<br>
        ��ע˵��</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><textarea rows="3" name="remark" cols="46"><%=rs("remark")%></textarea>
      </td>
    </tr>
  </table>
<br>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="hidden" name="id" value=<%=request("id")%>>
<input type="submit" name="submit" value="�޸�">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷ��Ҫɾ�������䶯��¼��')">
</form>
</center>
<%
end if
end if
%>


</body>
</html>










