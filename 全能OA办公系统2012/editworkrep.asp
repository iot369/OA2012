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
<title>编辑工作计划</title>
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
                      <td><span class="style4">编辑工作计划</span></td>
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
'打开数据库读出用户姓名
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
<td><span class="style1">编辑<%=stafname%>的工作计划（<%=recdate%>）</span></td>
<form action="displayworkrec.asp" method=post name="form1">
<td><input type="submit" name="addworkrep" value="返回"></td>
<input type="hidden" name="username" value="<%=username%>">
<input type="hidden" name="superior" value="<%=superior%>">
<input type="hidden" name="recdate" value="<%=recdate%>">
</form>
</tr>
</table>
</center>
<%

if request("submit")="修改" then
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
			emailtitle="您好，您给"&oabusyname&"分配的工作任务已完成！["&title&"]"
			emailcontent="工作任务标题：["&title&"]"&chr(13)&chr(10)
			emailcontent=emailcontent&"详细说明：["&remark&"]"
			errstr="对不起，系统自动发送您完成工作任务邮件出错，请手动发送邮件通知对方！"
			errinfo=send_event_email(emailtitle,oabusyuserid,rs("ID"),emailcontent,errstr)
		elseif superior=oabusyusername and superior<>"" then
			emailtitle="您好，"&oabusyname&"为您修改了工作任务！["&title&"]"
			emailcontent="工作任务标题：["&title&"]"&chr(13)&chr(10)
			emailcontent=emailcontent&"详细说明：["&remark&"]"
			errstr="对不起，系统自动发送您修改工作任务邮件出错，请手动发送邮件通知对方！"
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
<font color=red >编辑成功！</font>
</center>
<%
else
if request("submit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "delete from workrep where id=" & id
conn.Execute sql
%>
<center>
<br><br><br>
<font color=red >删除成功！</font>
</center>
<%
else
%>
&nbsp;
<%
'打开数据库读出id=id的记录
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
   if(!l1){window.alert("标题的长度大于1位小于50位");document.form2.title.focus();return (false);}

                    }

</script>
<br>
<form action="editworkrep.asp" method=post name="form2" onsubmit="return form_check();">
<table width=98% border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<tr>
<td width=80 height="30" bgcolor="D7E8F8"><div align="center" class="style1 style5">简要标题</div></td>
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
<td bgcolor="D7E8F8"><div align="center" class="style1 style5">详细说明</div></td>
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
<td height="60" colspan=2><span class="style5">　完成情况:
    <%
if username=oabusyusername and superior="" then
%> 
    <input type="radio" name="finished" value="yes"<%=checked("yes",rs("finished"))%>>
    已完成&nbsp;&nbsp;
    <input type="radio" name="finished" value="no"<%=checked("no",rs("finished"))%>>
    未完成<br>
    <%
else
%>
    <input type="hidden" name="finished" value="<%=rs("finished")%>">
    <%=checked1("yes",rs("finished"))%>已完成&nbsp;&nbsp;<%=checked1("no",rs("finished"))%>未完成<br>
    <%
end if
%>
　重要程度:
<%
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%> 
<input type="radio" name="imp" value="yes"<%=checked("yes",rs("imp"))%>>
重要&nbsp;&nbsp;
<input type="radio" name="imp" value="no"<%=checked("no",rs("imp"))%>>
一般
<%
else
%>
<%=checked1("yes",rs("imp"))%>重要&nbsp;&nbsp;&nbsp;&nbsp;<%=checked1("no",rs("imp"))%>一般
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
<input type="submit" name="submit" value="修改" >
<%
end if
if (username=oabusyusername and rs("superior")="") or (rs("superior")=oabusyusername) then
%>
&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="删除" onclick="return window.confirm('你确实要删除这条计划吗？');">
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