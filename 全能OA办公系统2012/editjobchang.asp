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
<title>OA办公系统.边缘特别版</title>
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
                  <td class="style7">员工管理</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<%
'打开数据库读出用户名
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
编辑<%=name%>的职位变动信息&nbsp;&nbsp;
</td>
<form method="post" action="jobchanginf.asp" name="form2"><td>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="submit" value="返回">
</td>
</form>
</tr>
</table>
</center>

<%
if request("submit")="修改" then
username=request("username")
oldjob=request("oldjob")
changjob=request("changjob")
changdate=request("changdate")
changfile=request("changfile")
changsort=request("changsort")
changtype=request("changtype")
changreason=request("changreason")
updateusername=oabusyname
updatedate=now()
id=request("id")
set conn=opendb("oabusy","conn","accessdsn")
sql = "Update changjob set "
sql = sql & "oldjob=" & SqlStr(oldjob) & ", "
sql = sql & "changjob=" & SqlStr(changjob) & ", "
sql = sql & "changdate=" & SqlStr(changdate) & ", "
sql = sql & "changfile=" & SqlStr(changfile) & ", "
sql = sql & "changsort=" & SqlStr(changsort) & ", "
sql = sql & "changtype=" & SqlStr(changtype) & ", "
sql = sql & "changreason=" & SqlStr(changreason) & ", "
sql = sql & "updateusername=" & SqlStr(updateusername) & ", "
sql = sql & "updatedate=#" & updatedate & "# where id=" & id
conn.Execute sql
%>
<br><br>
<center><font color=red >成功修改员工职务变动信息！</font></center>
<%
else
if request("submit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete from changjob where id=" & request("id")
conn.Execute sql
%>
<br><br>
<center><font color=red >成功删除员工职务变动信息！</font></center>
<%
else
%>
<script Language="JavaScript">

 function form_check(){
   var l1=document.form1.oldjob.value;
   if(l1==""){window.alert("原职务必须填写！");document.form1.oldjob.focus();return (false);}

   var l2=document.form1.changjob.value;
   if(l2==""){window.alert("变动职务必须填写！");document.form1.changjob.focus();return (false);}

   var l3=document.form1.changdate.value;
   if(l3==""){window.alert("变动时间必须填写！");document.form1.changdate.focus();return (false);}
                    }



</script>

<%
'打开数据库，读出职务变动信息
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from changjob where id=" & request("id")
rs.open sql,conn,1
%>
<br>
<center>
<form method="post" action="editjobchang.asp" name="form1" onsubmit="return form_check();">
  <table border="0" cellpadding="0" cellspacing="0" width="540">
    <tr>
      <td height="25" width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">员工姓名</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=name%>
      </td>
    </tr>
    <tr>
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">原&nbsp;职&nbsp;务</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="oldjob" size=10 value="<%=rs("oldjob")%>"> 
      </td>
      <td width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" align="center">变动职务</td>
      <td width="35%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="changjob" size=10 value="<%=rs("changjob")%>">  
      </td>
    </tr>
    <tr>
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">变动时间</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="changdate" size=10 value="<%=rs("changdate")%>"> 
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">变动文号</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="changfile" size=10 value="<%=rs("changfile")%>"> 
      </td>
    </tr>
    <tr>
      <td style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">变动类型</td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="changsort" size=10 value="<%=rs("changsort")%>">
      </td>
      <td style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" align="center">变动方式</td>
      <td style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name="changtype" size=10 value="<%=rs("changtype")%>"> 
      </td>
    </tr>
    <tr>
      <td width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" align="center">变动原因<br>
        备注说明</td>
      <td colspan="3" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><textarea rows="3" name="changreason" cols="46"><%=rs("changreason")%></textarea>
      </td>
    </tr>
  </table>
<br>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="hidden" name="id" value=<%=request("id")%>>
<input type="submit" name="submit" value="修改">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="删除" onclick="return window.confirm('你确定要删除这条变动记录吗？')">
</form>
</center>
<%
end if
end if
%>


</body>
</html>










