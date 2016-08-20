<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
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

id=request("id")
name=request("name")
company=request("company")
address=request("address")
project=request("project")
tel=request("tel")
callno=request("callno")
handset=request("handset")
fax=request("fax")
remark=request("remark")
email=request("email")
postcard=request("postcard")
dept=request("dept")
sex=request("sex")
position=request("position")
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
.style1 {color: #0d79b3}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                <td class="style7">客户资源</td>
              </tr>
          </table></td>
          <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
        </tr>
      </table>
    <font color="0D79B3"></font></div></td>
  </tr>
</table>
<center>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td> <br>
<table align="center">
<tr>
<td><span class="style1">编辑客户资料&nbsp;</span>&nbsp;&nbsp;&nbsp;</td>
<form method="post" action="clientinf.asp">
<td><input type="submit" value="返回"><input type="hidden" name="project" value="<%=request("project")%>">
</td>
</form>
</tr>
</table>
</center>

<%
if request("delit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "delete * from clientinf where id=" & id
conn.Execute sql
%>
<br><br>
<center>
<font color=red >
成功删除客户信息！
</font>
</center>
<%
else
if request("submit")="修改" then
set conn=opendb("oabusy","conn","accessdsn")
sql = "update clientinf set "
sql = sql & "name=" & SqlStr(name) & ", "
sql = sql & "company=" & SqlStr(company) & ", "
sql = sql & "address=" & SqlStr(address) & ", "
sql = sql & "project=" & SqlStr(project) & ", "
sql = sql & "tel=" & SqlStr(tel) & ", "
sql = sql & "callno=" & SqlStr(callno) & ", "
sql = sql & "handset=" & SqlStr(handset) & ", "
sql = sql & "fax=" & SqlStr(fax) & ", "
sql = sql & "remark=" & SqlStr(remark) & ", "
sql = sql & "email=" & SqlStr(email) & ", "
sql = sql & "postcard=" & SqlStr(postcard) & ", "
sql = sql & "dept=" & SqlStr(dept) & ", "
sql = sql & "sex=" & SqlStr(sex) & ", "
sql = sql & "position=" & SqlStr(position) & " where id=" & id
conn.Execute sql
%>
<br><br>
<center>
<font color=red >
成功编辑客户信息！
</font>
</center>
<%
else
%>
<script Language="JavaScript">

 function form_check(){
   var l1=document.form1.name.value.length;
   if(l1==0){window.alert("客户姓名必须填");document.form1.name.focus();return (false);}

   var l2=document.form1.company.value.length;
   if(l2==0){window.alert("客户所在单位必须填");document.form1.company.focus();return (false);}

   var l3=document.form1.project.value.length;
   if(l3==0){window.alert("业务项目必须填");document.form1.project.focus();return (false);}

                    }

</script>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from clientinf where id=" & id
rs.open sql,conn,1
%>
<center>
<form method="post" action="editclientinf.asp" name="form1">
<br>
    <table border="0" cellpadding="5" cellspacing="0" width="95%">
      <tr>
      <td width="15%" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">客户姓名</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="name" size=23 value="<%=rs("name")%>"><font color=red>*</font></td>
      <td width="15%" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">客户性别</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><select name="sex" size="1">
          <option value="男"<%=selected(rs("sex"),"男")%>>男</option>
          <option value="女"<%=selected(rs("sex"),"女")%>>女</option>
        </select></td>
     </tr>
     <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">职&nbsp;&nbsp;&nbsp;&nbsp;位</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="position" size=23 value="<%=rs("position")%>"></td>
      <td align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">业务项目</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="project" size=23 value="<%=rs("project")%>"><font color=red>*</font></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">所在单位</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="company" size=60 value="<%=rs("company")%>"><font color=red>*</font></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">部&nbsp;&nbsp;&nbsp;&nbsp;门</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="dept" size=23 value="<%=rs("dept")%>"></td>
      <td align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">邮政编码</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="postcard" size=23 value="<%=rs("postcard")%>"></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">单位地址</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="address" size=60value="<%=rs("address")%>"></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">传真号码</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="fax" size=60value="<%=rs("fax")%>"></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">联系电话</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="tel" size=23 value="<%=rs("tel")%>"></td>
      <td align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">手&nbsp;&nbsp;&nbsp;&nbsp;机</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="handset" size=23 value="<%=rs("handset")%>"></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">M　S　N</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="callno" size=23 value="<%=rs("callno")%>"></td>
      <td align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><span class="style8">电子邮箱</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type=text name="email" size=23 value="<%=rs("email")%>"></td>
    </tr>
    <tr>
      <td align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><span class="style8">备注说明</span></td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" colspan="3"><textarea rows="4" cols="59" name="remark"><%=rs("remark")%></textarea></td>
    </tr>
  </table>
    <br>
    <input type="hidden" name="id" value=<%=id%>>
  <font color=red>*</font><span class="style1">必须填写</span>&nbsp;&nbsp;&nbsp;&nbsp;
  <input type="submit" name="submit" value="修改" onclick="return form_check();">&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="delit" value="删除" onclick="return window.confirm('此操作后数据无法恢复，您确实要删除此员工信息吗？');">
</form>
</center>
<%
end if
end if
%></td>
    </tr>
  </table>
 

</body>
</html>










