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
<center>
  <br>
<table>
<tr>
<td>
�༭<%=name%>�Ĺ�����Ϣ&nbsp;&nbsp;
</td>
<form method="post" action="wageinf.asp" name="form2"><td>
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
wagelevel=request("wagelevel")
basewage=request("basewage")
stafjob=request("stafjob")
jobwage=request("jobwage")
workyear=request("workyear")
workyearwage=request("workyearwage")
rentwage=request("rentwage")
carwage=request("carwage")
prize=request("prize")
insurance=request("insurance")
tax=request("tax")
affairday=request("affairday")
affairfund=request("affairfund")
sickday=request("sickday")
sickfund=request("sickfund")
mustwage=request("mustwage")
actwage=request("actwage")
changreason=request("changreason")
actdate=request("actdate")
remark=request("remark")
updatename=oabusyname
updatedate=now()
id=request("id")
set conn=opendb("oabusy","conn","accessdsn")
sql = "Update wageinf set "
sql = sql & "wagelevel=" & SqlStr(wagelevel) & ", "
sql = sql & "basewage=" & SqlStr(basewage) & ", "
sql = sql & "stafjob=" & SqlStr(stafjob) & ", "
sql = sql & "jobwage=" & SqlStr(jobwage) & ", "
sql = sql & "workyear=" & SqlStr(workyear) & ", "
sql = sql & "workyearwage=" & SqlStr(workyearwage) & ", "
sql = sql & "rentwage=" & SqlStr(rentwage) & ", "
sql = sql & "carwage=" & SqlStr(carwage) & ", "
sql = sql & "prize=" & SqlStr(prize) & ", "
sql = sql & "insurance=" & SqlStr(insurance) & ", "
sql = sql & "tax=" & SqlStr(tax) & ", "
sql = sql & "affairday=" & SqlStr(affairday) & ", "
sql = sql & "affairfund=" & SqlStr(affairfund) & ", "
sql = sql & "sickday=" & SqlStr(sickday) & ", "
sql = sql & "sickfund=" & SqlStr(sickfund) & ", "
sql = sql & "mustwage=" & SqlStr(mustwage) & ", "
sql = sql & "actwage=" & SqlStr(actwage) & ", "
sql = sql & "changreason=" & SqlStr(changreason) & ", "
sql = sql & "actdate=" & SqlStr(actdate) & ", "
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
sql="delete from wageinf where id=" & request("id")
conn.Execute sql
%>
<br><br>
<center><font color=red >�ɹ�ɾ��Ա��������Ϣ��</font></center>
<%
else
%>
<script Language="JavaScript">

 function form_check(){
   var l1=document.form1.basewage.value;
   if(l1==""){window.alert("�������ʱ����������д��");document.form1.basewage.focus();return (false);}

   var l2=document.form1.mustwage.value;
   if(l2==""){window.alert("Ӧ�����ʱ�����д��");document.form1.mustwage.focus();return (false);}
                    }



</script>
<%
'�����ݿ⣬����������Ϣ
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from wageinf where id=" & request("id")
rs.open sql,conn,1
%>
<br>
<center>
<form method="post" action="editwageinf.asp" name="form1" onsubmit="return form_check();">
ְԱ����:<%=name%>
  <table border="0" cellpadding="0" cellspacing="0" width="95%">
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���ʼ���</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=wagelevel size=10 value="<%=rs("wagelevel")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=basewage size=10 value="<%=rs("basewage")%>"><font color=red>*</font></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ա��ְ��</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=stafjob size=10 value="<%=rs("stafjob")%>">
      </td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ְ����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=jobwage size=10 value="<%=rs("jobwage")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ա������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=workyear size=10 value="<%=rs("workyear")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���乤��</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=workyearwage size=10 value="<%=rs("workyearwage")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">������</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=prize size=10 value="<%=rs("prize")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���ⲹ��</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=rentwage size=10 value="<%=rs("rentwage")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���Ѳ���</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=carwage size=10 value="<%=rs("carwage")%>">
      </td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�¼�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=affairday size=10 value="<%=rs("affairday")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">���¼ٿ�</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=affairfund size=10 value="<%=rs("affairfund")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=sickday size=10 value="<%=rs("sickday")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�۲��ٿ�</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=sickfund size=10 value="<%=rs("sickfund")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">������˰</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=tax size=10 value="<%=rs("tax")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�����շ�</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=insurance size=10 value="<%=rs("insurance")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">Ӧ�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=mustwage size=10 value="<%=rs("mustwage")%>"><font color=red>*</font></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ʵ�����</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=actwage size=10 value="<%=rs("actwage")%>"></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ִ��ʱ��</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=actdate size=10 value="<%=rs("actdate")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�䶯ԭ��</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=changreason size=59 value="<%=rs("changreason")%>"></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">��ע˵��</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><input type="text" name=remark size=59 value="<%=rs("remark")%>"></td>
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










