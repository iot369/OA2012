<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
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
'�����ݿ�����༭���ʵ�Ȩ��
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select allow_edit_all_wageinf,allow_edit_dept_wageinf from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
cook_allow_edit_all_wageinf=rs("allow_edit_all_wageinf")
cook_allow_edit_dept_wageinf=rs("allow_edit_dept_wageinf")
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

<center>
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
  <br>
  <table>
<tr>
      <td> Ա�����ʵ���&nbsp;&nbsp; </td>
<%
if cook_allow_edit_all_wageinf="yes" then
%>
<form action="wageinf.asp" method=get name="form1">
<td>
<select name="userdept" size=1 onChange="document.form1.submit();">
<%
'�����ݿ��������
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
if not rs.eof and not rs.bof then firstdept=rs("userdept")
if request("userdept")<>"" then firstdept=request("userdept")
while not rs.eof and not rs.bof
%>
<option value="<%=rs("userdept")%>"<%=selected(firstdept,rs("userdept"))%>><%=rs("userdept")%></option>
<%
rs.movenext
wend
%>
</select>
</td>
</form>
<%
else
firstdept=oabusyuserdept
end if
%>
<%
if cook_allow_edit_all_wageinf="yes" or cook_allow_edit_dept_wageinf="yes" then
%>
<form action="wageinf.asp" method=get name="form2">
<td>
<input type="hidden" name="userdept" value="<%=firstdept%>">
<select name="username" size=1>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name,username from userinf where userdept=" & sqlstr(firstdept) & " and userlevel<>'�ܹ�' and forbid='no'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then username=rs("username")
if request("username")<>"" then username=request("username")
while not rs.eof and not rs.bof
%>
<option value="<%=rs("username")%>"<%=selected(username,rs("username"))%>><%=rs("name")%></option>
<%
rs.movenext
wend
%>
</select>
<td>
<input type="submit" name="submit" value="��ѯ">
</td>
</form>
<form method="post" action="addwageinf.asp">
<td>
<input type="hidden" name="userdept" value="<%=firstdept%>">
<input type="hidden" name="username" value="<%=username%>">
<input type="submit" value="���ӵ���">
</td>
</form>
<%
else
username=oabusyusername
end if
%>
</tr>
</table>
</center>

<%
'�����ݿ����Ա������
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
name=rs("name")
%>
<br>
<center>
  <%=name%>�Ĺ��ʵ���
</center>
<br>
<%
'�����ݿ⣬����Ա���������ݿ�
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from wageinf where username=" & sqlstr(username)
rs.open sql,conn,1
while not rs.eof and not rs.bof
wagelevel=rs("wagelevel")
basewage=rs("basewage")
stafjob=rs("stafjob")
jobwage=rs("jobwage")
workyear=rs("workyear")
workyearwage=rs("workyearwage")
rentwage=rs("rentwage")
carwage=rs("carwage")
prize=rs("prize")
insurance=rs("insurance")
tax=rs("tax")
affairday=rs("affairday")
affairfund=rs("affairfund")
sickday=rs("sickday")
sickfund=rs("sickfund")
mustwage=rs("mustwage")
actwage=rs("actwage")
changreason=rs("changreason")
actdate=rs("actdate")
remark=rs("remark")
recname=rs("recname")
recdate=rs("recdate")
updatename=rs("updatename")
updatedate=rs("updatedate")
id=rs("id")
%>
<center>
  ְԱ����:<%=name%> 
  <table border="0" cellpadding="4" cellspacing="0" width="95%">
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���ʼ���</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(wagelevel)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(basewage)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ա��ְ��</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(stafjob)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ְ����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(jobwage)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ա������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(workyear)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���乤��</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(workyearwage)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(prize)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���ⲹ��</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(rentwage)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���Ѳ���</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(carwage)%> 
      </td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�¼�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(affairday)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���¼ٿ�</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(affairfund)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(sickday)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�۲��ٿ�</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(sickfund)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">������˰</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(tax)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�����շ�</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(insurance)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ӧ�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(mustwage)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ʵ�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(actwage)%></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ִ��ʱ��</td>
      <td height="25" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(actdate)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�䶯ԭ��</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(changreason)%></td>
    </tr>
    <tr> 
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��ע˵��</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(remark)%></td>
    </tr>
  </table>
  <table width="95%" cellspacing="0" cellpadding="4">
    <tr> 
      <td height="25" width="15%" align="center" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��&nbsp;¼&nbsp;��</td>
      <td height="25" width="35%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(recname)%></td>
      <td height="25" width="15%" align="center" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��¼ʱ��</td>
      <td height="25" width="35%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=checked3(recdate)%></td>
    </tr>
    <tr> 
      <td height="25" width="15%" align="center" style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">��&nbsp;��&nbsp;��</td>
      <td height="25" width="35%" style="border-left: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><%=checked3(updatename)%></td>
      <td height="25" width="15%" align="center" style="border-left: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">����ʱ��</td>
      <td height="25" width="35%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><%=checked3(updatedate)%></td>
    </tr>
  </table>
<%
if cook_allow_edit_all_wageinf="yes" or cook_allow_edit_dept_wageinf="yes" then
%>
  <table border="0" cellpadding="0" cellspacing="0" width="95%">
    <tr><form method="post" action="editwageinf.asp">
      <td align=right><input type="submit" value="�༭">
<input type="hidden" name="userdept" value="<%=firstdept%>">
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="id" value=<%=id%>>
</td>
        </form>
    </tr>
  </table>
  <%
end if
%>
  <br>
</center>
<%
rs.movenext
wend 
%>


</body>
</html>










