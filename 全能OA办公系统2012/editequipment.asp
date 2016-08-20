<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/check_resource.asp"-->
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
if check_resource_setting(oabusyusername,0)<>0 then
	response.write("<script language=""javascript"">")
	response.write("alert(""对不起，您不能修改资源！"");")
	response.write("history.go(-1);")
	response.write("</script>")
	response.end
end if
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
                  <td class="style7">公共资源</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
      <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td> <table width="1%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
        <div align="center">编辑资源
          </center>

        </div>
        <center>
<%
equipment=request("equipment")
controller=request("controller")
remark=request("remark")
id=request("id")
oldequipment=request("oldequipment")

if request("submit")="修改" then
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment where equipment=" & sqlstr(equipment) & " and id<>" & id
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<br><br>新的资源名与原来资源记录冲突！<br><br>
<input type="button" value="返回" onclick="window.location.href='editequipment.asp';">
<%
else
set conn=opendb("oabusy","conn","accessdsn")
sql = "update equipment set equipment=" & sqlstr(equipment) & ",controller=" & sqlstr(controller) & ",remark=" & sqlstr(remark) & " where id=" & id
conn.Execute sql
sql="update booking set equipment=" & sqlstr(equipment) & " where equipment=" & sqlstr(oldequipment)
conn.Execute sql
%>
<br><br>资源修改成功！<br><br>
<form method="post" action="booking.asp"><input type="submit" value="返回"></form>
<%
end if


else
if request("submit")="删除" then
set conn=opendb("oabusy","conn","accessdsn")
sql="delete * from equipment where id=" & id
conn.Execute sql
sql="delete * from booking where equipment=" & sqlstr(oldequipment)
conn.Execute sql
%>
<br><br>成功删除资源！<br><br>
<form method="post" action="booking.asp"><input type="submit" value="返回"></form>
<%
else
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment where id=" & id
rs.open sql,conn,1
%>
<br>
<form method="post" name="form1" action="editequipment.asp">
  <table border="0" cellpadding="0" cellspacing="1" bgcolor="B0C8EA">
    <tr>
      <td bgcolor="D7E8F8">欲修改的资源名称：
        <input type=text size="20" name="equipment" value="<%=rs("equipment")%>"><font color=red>*</font></td>
    </tr>
    <tr>
      <td bgcolor="D7E8F8">资源维护或管理员：
        <input type=text size="20" name="controller" value="<%=rs("controller")%>"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="D7E8F8">说明</td>
    </tr>
    <tr>
      <td align="center" bgcolor="D7E8F8"><textarea rows="5" cols="35" name="remark"><%=rs("remark")%></textarea><br><font color=red>*</font>项必填&nbsp;&nbsp;<input type="button" value="返回" onclick="window.location.href='booking.asp'">
&nbsp;&nbsp;<input type="submit" name="submit" value="修改" onclick="return check_form();">&nbsp;&nbsp;<input type="submit" name="submit" value="删除" onclick="return window.confirm('此删除操作将不能恢复，您真的要删除吗？');"><input type="hidden" name="id" value=<%=id%>><input type="hidden" name="oldequipment" value="<%=rs("equipment")%>"></td>
    </tr>
  </table>
</form>
<script Language="JavaScript">

 function check_form(){
   var equipment=document.form1.equipment.value;
   if(equipment.length==0){window.alert("资源名称必须填");document.form1.equipment.focus();return (false);}
                    }

</script>

<%
end if
end if
%>
</center></td>
    </tr>
</table>
 

</body>
</html>












