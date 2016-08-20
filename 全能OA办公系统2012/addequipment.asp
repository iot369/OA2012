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
	response.write("alert(""对不起，您不能增加资源！"");")
	response.write("history.go(-1);")
	response.write("</script>")
	response.end
end if
'--------------------------------------
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<SCRIPT language=javascript>
<!--
if (window.Event) 
　document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
　if (e.which == 2 || e.which == 3)
　 return false;
 } else if (event.button == 2 || event.button == 3) {
　 event.cancelBubble = true
　 event.returnvalue = false;
　 return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;　// for IE5+
document.onmousedown = norightclick;　　 // for all others
//-->
</SCRIPT>
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
    <td><table width="1%"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td>&nbsp;</td>
  </tr>
</table>
      <div align="center">增加新的资源
        </center>

      </div>
      <center>
<%
if request("submit")="增加" then
equipment=request("equipment")
controller=request("controller")
remark=request("remark")
'打开数据库判断设备名称是否重复
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment where equipment=" & sqlstr(equipment)
rs.open sql,conn,1
if not rs.eof and not rs.bof then
%>
<br><br>输入的资源名称与原来记录重复，请重新输入！<br><br>
<input type="button" value="返回" onclick="window.location.href='addequipment.asp'">
<%
else
set conn=opendb("oabusy","conn","accessdsn")
sql = "Insert Into equipment (equipment,remark,controller) Values( "
sql = sql & SqlStr(equipment) & ", "
sql = sql & SqlStr(remark) & ", "
sql = sql & SqlStr(controller) & ")"
conn.Execute sql
%>
<br><br>成功输入资源！<br><br>
<form method="post" action="booking.asp">
<input type="submit" value="返回">
</form>
<%
end if
else
%>
<br>
<form method="post" name="form1" action="addequipment.asp">
  <table border="0" cellpadding="0" cellspacing="1" bgcolor="B0C8EA">
    <tr>
      <td bgcolor="D7E8F8">欲增加的资源名称：
        <input type=text size="20" name="equipment"><font color=red>*</font></td>
    </tr>
    <tr>
      <td bgcolor="D7E8F8">资源维护或管理员：
        <input type=text size="20" name="controller"></td>
    </tr>
    <tr>
      <td align="center" bgcolor="D7E8F8">说明</td>
    </tr>
    <tr>
      <td align="center" bgcolor="D7E8F8"><textarea rows="5" cols="35" name="remark"></textarea><br>
        <input type="button" value="返回" onclick="window.location.href='booking.asp'">
&nbsp;&nbsp;<font color=red>*</font>项必填&nbsp;&nbsp;<input type="submit" name="submit" value="增加" onclick="return check_form();"></td>
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
%>
</center>
</td>
  </tr>
</table>

</body>
</html>








