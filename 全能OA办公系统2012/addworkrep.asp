<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->

<!--#include file="asp/sendeventemail.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
oabusyuserid=request.cookies("oabusyuserid")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<title>���ӹ�������</title>
<style type="text/css">
<!--
.style1 {color: #0d79b3}
.style4 {color: #2e4869}
.style5 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="5" leftmargin="5">
<%

username=request("username")
superior=request("superior")
recdate=request("recdate")
'�����ݿ�����û�����
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
if not rs.eof and not rs.bof then stafname=rs("name")
%>
<table width="550"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="1"><img src="images/main/l4.gif" width="1" height="21"></td>
                <td background="images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="10">&nbsp;</td>
                      <td><span class="style4">��ӹ�������</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><center>
            <br>
<table border="0" align="center" cellpadding="0"  cellspacing="0">
<tr><td>
<span class="style1">����<%=stafname%>�Ĺ�������(</span><font color=red>*</font><span class="style1">�������д)</span></td>
<form action="displayworkrec.asp" method=post name="form1">
<td><input type="submit" name="addworkrep" value="����"></td>
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="superior" value="<%=superior%>">
<input type="hidden" name="recdate" value="<%=recdate%>">
</form>
</tr>
</table>
<center>
<%
if request("addworkrep")="���������" then
title=request("title")
remark=request("remark")
set conn=opendb("oabusy","conn","accessdsn")
sql = "Insert Into workrep (username,recdate,title,remark,superior) Values( "
sql = sql & SqlStr(username) & ", "
sql = sql & "#" & recdate & "#" & ", "
sql = sql & SqlStr(title) & ", "
sql = sql & SqlStr(remark) & ", "
sql = sql & SqlStr(superior) & ")"
conn.Execute sql
if superior<>"" then
	set rs=server.createobject("adodb.recordset")
	sql="select ID from userinf where username='"&username&"'"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		emailtitle="���ã�"&oabusyname&"��������һ��������������ģ�"
		emailcontent="����������⣺["&title&"]"&chr(13)&chr(10)
		emailcontent=emailcontent&"�����������ڣ�["&recdate&"]"&chr(13)&chr(10)
		emailcontent=emailcontent&"��ϸ˵����["&remark&"]"
		errstr="�Բ���ϵͳ�Զ��������Ĺ��������ų������ֶ������ʼ�֪ͨ�Է���"
		errinfo=send_event_email(emailtitle,oabusyuserid,rs("ID"),emailcontent,errstr)
		if errinfo<>"" then
			set rs=nothing
			conn.close
			response.redirect "asp/disperrorinfo.asp?errorinfo="&errinfo
			response.end
		end if
	else
		set rs=nothing
		conn.close
		response.redirect "asp/disperrorinfo.asp?errorinfo="&errstr
		response.end
	end if
end if
%>
<br><br>
  <font color="#0033FF" >�ɹ����ӹ�������</font><br>
  <br><br>
<%
else
%>
<script Language="JavaScript">
 function maxlength(str,minl,maxl) {
    if(str.length <= maxl && str.length >= minl){return true;}else{return false;}
                                    }

 function form_check(){
   var l1=maxlength(document.form2.title.value,1,50);
   if(!l1){window.alert("����ĳ��ȴ���1λС��50λ");document.form2.title.focus();return (false);}

                    }

</script>
<br>
<form action="addworkrep.asp" method=post name="form2" onsubmit="return form_check();">
    <table border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
      <tr bgcolor="#FFFFFF">
<td width="100" height="30" bgcolor="D7E8F8"><div align="center" class="style1 style5">��Ҫ����</div></td>
<td width="380" height="30" bgcolor="#FFFFFF"><div align="center">
  <input type="text" name="title" size=50>
  <font color=red>*</font></div></td>
</tr>
<tr bgcolor="#FFFFFF">
<td height="160" bgcolor="D7E8F8"><div align="center" class="style1 style5">��ϸ˵��</div></td>
<td height="30"><div align="center">
  <textarea rows="10" name="remark" cols="50"><%=content1%></textarea>
</div></td>
</tr>
</table>
    <br>
    <input type="submit" name="addworkrep" value="���������" >
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="superior" value="<%=superior%>">
<input type="hidden" name="recdate" value="<%=recdate%>">
</form>
<%
end if
%>           
          </center></td>
        </tr>
    </table></td>
  </tr>
</table>


</center>
