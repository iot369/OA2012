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
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>oa办公系统</title>
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
增加<%=name%>的工资档案&nbsp;&nbsp;
</td>
<form method="post" action="wageinf.asp" name="form2"><td>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<input type="submit" value="返回">
</td>
</form>
</tr>
</table>
</center>

<%
if request("submit")="增加" then
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
recname=oabusyname
updatename=oabusyname
set conn=opendb("oabusy","conn","accessdsn")
sql = "Insert Into wageinf (username,wagelevel,basewage,stafjob,jobwage,workyear,workyearwage,rentwage,carwage,prize,insurance,tax,affairday,affairfund,sickday,sickfund,mustwage,actwage,changreason,actdate,remark,recname,updatename) Values( "
sql = sql & SqlStr(username) & ", "
sql = sql & SqlStr(wagelevel) & ", "
sql = sql & SqlStr(basewage) & ", "
sql = sql & SqlStr(stafjob) & ", "
sql = sql & SqlStr(jobwage) & ", "
sql = sql & SqlStr(workyear) & ", "
sql = sql & SqlStr(workyearwage) & ", "
sql = sql & SqlStr(rentwage) & ", "
sql = sql & SqlStr(carwage) & ", "
sql = sql & SqlStr(prize) & ", "
sql = sql & SqlStr(insurance) & ", "
sql = sql & SqlStr(tax) & ", "
sql = sql & SqlStr(affairday) & ", "
sql = sql & SqlStr(affairfund) & ", "
sql = sql & SqlStr(sickday) & ", "
sql = sql & SqlStr(sickfund) & ", "
sql = sql & SqlStr(mustwage) & ", "
sql = sql & SqlStr(actwage) & ", "
sql = sql & SqlStr(changreason) & ", "
sql = sql & SqlStr(actdate) & ", "
sql = sql & SqlStr(remark) & ", "
sql = sql & SqlStr(recname) & ", "
sql = sql & SqlStr(updatename) & ")"
conn.Execute sql
%>
<br><br>
<center><font color=red >成功增加员工工资信息！</font></center>
<%
else
%>


<script Language="JavaScript">

 function form_check(){
   var l1=document.form1.basewage.value;
   if(l1==""){window.alert("基本工资必须填必须填写！");document.form1.basewage.focus();return (false);}

   var l2=document.form1.mustwage.value;
   if(l2==""){window.alert("应发工资必须填写！");document.form1.mustwage.focus();return (false);}
                    }



</script>




<br>
<center>
<form method="post" action="addwageinf.asp" name="form1" onsubmit="return form_check();">
职员姓名:<%=name%>
  <table border="0" cellpadding="0" cellspacing="0" width="95%">
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">工资级别</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=wagelevel size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">基本工资</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=basewage size=10><font color=red>*</font></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">员工职务</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=stafjob size=10>
      </td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">职务工资</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=jobwage size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">员工工龄</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=workyear size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">工龄工资</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=workyearwage size=10></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">奖金金额</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=prize size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">房租补贴</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=rentwage size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">车费补贴</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=carwage size=10>
      </td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">事假天数</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=affairday size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">扣事假款</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=affairfund size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">病假天数</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=sickday size=10></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">扣病假款</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=sickfund size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">交个人税</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=tax size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">交保险费</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=insurance size=10></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">应发金额</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=mustwage size=10><font color=red>*</font></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">实发金额</td>
      <td height="25" width="18%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=actwage size=10></td>
      <td height="25" align="center" width="15%" style="border-left: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">执行时间</td>
      <td height="25" width="19%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=actdate size=10></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">变动原因</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><input type="text" name=changreason size=59></td>
    </tr>
    <tr>
      <td height="25" align="center" width="15%" style="border-left: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">备注说明</td>
      <td colspan="5" height="25" width="85%" style="border-left: 1 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"><input type="text" name=remark size=59></td>
    </tr>
  </table>

<br>
<input type="hidden" name="userdept" value="<%=request("userdept")%>">
<input type="hidden" name="username" value="<%=request("username")%>">
<font color=red>*</font>为必填项&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="增加">
</form>
</center>
<%
end if
%>


</body>
</html>










