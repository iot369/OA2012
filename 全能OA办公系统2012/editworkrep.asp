<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->

<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/sendeventemail.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserid=request.cookies("oabusyuserid")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
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
<title>�༭�����ƻ�</title>
<style type="text/css">
<!--
.style1 {color: #0d79b3}
.style4 {color: #2e4869}
.style5 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="5" leftmargin="5">
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
                      <td><span class="style4">�༭�����ƻ�</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><%

username=request("username")
superior=request("superior")
recdate=request("recdate")
id=request("id")
'�����ݿ�����û�����
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
if not rs.eof and not rs.bof then stafname=rs("name")
%>
<center>
  <br>
<table border="0" cellspacing="0" cellpadding="0">
<tr>
<td><span class="style1">�༭<%=stafname%>�Ĺ����ƻ���<%=recdate%>��</span></td>
<form action="displayworkrec.asp" method=post name="form1">
<td><input type="submit" name="addworkrep" value="����"></td>
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="superior" value="<%=superior%>">
<input type="hidden" name="recdate" value="<%=recdate%>">
</form>
</tr>
</table>
</center>
<%

if request("submit")="�޸�" then
title=request("title")
remark=request("remark")
finished=request("finished")
imp1=request("imp")
superior=request("superior")
set conn=opendb("oabusy","conn","accessdsn")
sql = "update workrep set finished=" & sqlstr(finished) & ",imp=" & sqlstr(imp1) & ",title=" & sqlstr(title) & ",remark=" & sqlstr(remark) & " where id=" & id
conn.Execute sql
set rs1=server.createobject("adodb.recordset")
sql1="select superior from workrep where id="&id
rs1.open sql1,conn,1
if not rs1.eof and not rs1.bof then
	superior=rs1("superior")
	set rs=server.createobject("adodb.recordset")
	sql=""
	if finished="yes" and superior<>"" and superior<>oabusyusername then
		sql="select userinf.ID from userinf,workrep where workrep.id="&id&" and workrep.superior=userinf.username"
	elseif superior=oabusyusername and superior<>"" then
		sql="select userinf.ID from userinf,workrep where workrep.id="&id&" and workrep.username=userinf.username"
	end if
	if sql<>"" then
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		if finished="yes" and superior<>"" and superior<>oabusyusername then
			emailtitle="���ã�����"&oabusyname&"����Ĺ�����������ɣ�["&title&"]"
			emailcontent="����������⣺["&title&"]"&chr(13)&chr(10)
			emailcontent=emailcontent&"��ϸ˵����["&remark&"]"
			errstr="�Բ���ϵͳ�Զ���������ɹ��������ʼ��������ֶ������ʼ�֪ͨ�Է���"
			errinfo=send_event_email(emailtitle,oabusyuserid,rs("ID"),emailcontent,errstr)
		elseif superior=oabusyusername and superior<>"" then
			emailtitle="���ã�"&oabusyname&"Ϊ���޸��˹�������["&title&"]"
			emailcontent="����������⣺["&title&"]"&chr(13)&chr(10)
			emailcontent=emailcontent&"��ϸ˵����["&remark&"]"
			errstr="�Բ���ϵͳ�Զ��������޸Ĺ��������ʼ��������ֶ������ʼ�֪ͨ�Է���"
			errinfo=send_event_email(emailtitle,oabusyuserid,rs("ID"),emailcontent,errstr)
		end if
		if errinfo<>"" then
			set rs=nothing
			set rs1=nothing
			conn.close
			response.redirect "asp/disperrorinfo.asp?errorinfo="&errinfo
			response.end
		end if
	else
		set rs=nothing
		set rs1=nothing
		conn.close
		response.redirect "asp/disperrorinfo.asp?errorinfo="&errstr
		response.end
	end if
end if
end if
%>
<center>
<br><br><br>
<font color=red >�༭�ɹ���</font>
</center>
<%
else
if request("submit")="ɾ��" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "delete from workrep where id=" & id
conn.Execute sql
%>
<center>
<br><br><br>
<font color=red >ɾ���ɹ���</font>
</center>
<%
else
%>
&nbsp;
<%
'�����ݿ����id=id�ļ�¼
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from workrep where id=" & id
rs.open sql,conn,1
%>
<center>
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
<form action="editworkrep.asp" method=post name="form2" onsubmit="return form_check();">
<table width=98% border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<tr>
<td width=80 height="30" bgcolor="D7E8F8"><div align="center" class="style1 style5">��Ҫ����</div></td>
<td width=350 bgcolor="#FFFFFF">
<%
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%>
<input type="text" name="title" size=50 value="<%=rs("title")%>"><font color=red>*</font>
<%
else
%>
<input type="hidden" name="title" value="<%=rs("title")%>"><%=rs("title")%>
<%
end if
%>
</td>
</tr>
<tr>
<td bgcolor="D7E8F8"><div align="center" class="style1 style5">��ϸ˵��</div></td>
<td bgcolor="#FFFFFF">
<%
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%>
<textarea rows="10" name="remark" cols="50"><%=rs("remark")%></textarea>
<%
else
%>
<input type="hidden" name="remark" value="<%=rs("remark")%>"><%=checked3(keepformat(rs("remark")))%>
<%
end if
%>
</td>
</tr>
<tr bgcolor="#FFFFFF">
<td height="60" colspan=2><span class="style5">��������:
    <%
if username=oabusyusername and superior="" then
%> 
    <input type="radio" name="finished" value="yes"<%=checked("yes",rs("finished"))%>>
    �����&nbsp;&nbsp;
    <input type="radio" name="finished" value="no"<%=checked("no",rs("finished"))%>>
    δ���<br>
    <%
else
%>
    <input type="hidden" name="finished" value="<%=rs("finished")%>">
    <%=checked1("yes",rs("finished"))%>�����&nbsp;&nbsp;<%=checked1("no",rs("finished"))%>δ���<br>
    <%
end if
%>
����Ҫ�̶�:
<%
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%> 
<input type="radio" name="imp" value="yes"<%=checked("yes",rs("imp"))%>>
��Ҫ&nbsp;&nbsp;
<input type="radio" name="imp" value="no"<%=checked("no",rs("imp"))%>>
һ��
<%
else
%>
<%=checked1("yes",rs("imp"))%>��Ҫ&nbsp;&nbsp;&nbsp;&nbsp;<%=checked1("no",rs("imp"))%>һ��
<input type="hidden" name="imp" value="<%=rs("imp")%>">
<%
end if
%>
</span></td>
</tr>
</table>
<br>
<%
if (oabusyusername=rs("superior")) or (oabusyusername=rs("username")) then
%>
<input type="submit" name="submit" value="�޸�" >
<%
end if
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%>
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷʵҪɾ�������ƻ���');">
<%
end if
%>
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="superior" value="<%=superior%>">
<input type="hidden" name="recdate" value="<%=recdate%>">
<input type="hidden" name="id" value="<%=id%>">
</form>
</center>
<%
end if
end if
%>
<%

%>             </td>
        </tr>
    </table></td>
  </tr>
</table>

</body>
</html>