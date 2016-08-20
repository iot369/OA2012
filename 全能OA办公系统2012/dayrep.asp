<%@ LANGUAGE = VBScript %>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统.边缘特别版</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">个人工作计划</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/monthlycal.asp"-->

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

username=oabusyusername
'取得当前日期
mydate=date
'取得当前月
mymonth=month(now())
if request("mymonth")<>"" then mymonth=cint(request("mymonth"))
'取得当前年
myyear=year(now())
if request("myyear")<>"" then myyear=cint(request("myyear"))
%>


<center>
  <br>
<table>
<tr>
<form action="dayrep.asp" method=get name="form1">
<td>
<select name="myyear" size=1>
<option value=2001<%=selected(2001,myyear)%>>2001年</option>
<option value=2002<%=selected(2002,myyear)%>>2002年</option>
<option value=2003<%=selected(2003,myyear)%>>2003年</option>
<option value=2004<%=selected(2004,myyear)%>>2004年</option>
<option value=2005<%=selected(2005,myyear)%>>2005年</option>
<option value=2006<%=selected(2006,myyear)%>>2006年</option>
<option value=2007<%=selected(2007,myyear)%>>2007年</option>
<option value=2008<%=selected(2008,myyear)%>>2008年</option>
<option value=2009<%=selected(2009,myyear)%>>2009年</option>
</select>
<select name="mymonth" size=1>
<option value=1<%=selected(1,mymonth)%>>1月</option>
<option value=2<%=selected(2,mymonth)%>>2月</option>
<option value=3<%=selected(3,mymonth)%>>3月</option>
<option value=4<%=selected(4,mymonth)%>>4月</option>
<option value=5<%=selected(5,mymonth)%>>5月</option>
<option value=6<%=selected(6,mymonth)%>>6月</option>
<option value=7<%=selected(7,mymonth)%>>7月</option>
<option value=8<%=selected(8,mymonth)%>>8月</option>
<option value=9<%=selected(9,mymonth)%>>9月</option>
<option value=10<%=selected(10,mymonth)%>>10月</option>
<option value=11<%=selected(11,mymonth)%>>11月</option>
<option value=12<%=selected(12,mymonth)%>>12月&nbsp;&nbsp;</option>
</select>
<input type="submit" name="submit" value="查询">
</td>
</form>
</tr>
</table>
<span class="style8"><br>本人<%=myyear%>年<%=mymonth%>月工作计划</span>(<font color=red>周六处写上周工作总结，周日处写下周工作计划</font>)
</center>

<center>
<br>
<span class="style8">标题为</span>黑色<span class="style8">表示此工作性质一般且没完成，标题为</span><font color=blue>兰色</font><span class="style8">表示此工作性质一般且已完成</span><br>
<span class="style8">标题为</span><font color=red>红色</font><span class="style8">表示此工作性质重要且没完成，标题为</span><font color=#770000>棕色</font><span class="style8">表示此工作性质重要且已完成<br>
</span><br>
</center>
<%
call monthlycal(username,superior)

%>
</body>
</html>