<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/check_resource.asp"-->
<%
today1=date()
today2=date()+1
'a1=datediff("h",today1,#2001-5-6 1:00:00#)
'response.write a1
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
'删除以前的记录
resourceflag=check_resource_setting(oabusyusername,0)
set conn=opendb("oabusy","conn","accessdsn")
sql = "delete * from booking where endtime<#" & date() & "#"
conn.Execute sql
%>
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
  <table align="center">
<tr>
<%
if resourceflag=0 then
%>
<form method=post action="addequipment.asp">
<td>
<input type="submit" value="增加">
</td>
</form>
<%
set rs=nothing
end if
%>
<form method=post action="addbooking.asp">
<td>
<input type="submit" value="预约">
</td>
</form>
</tr>
</table>
  <div align="center">今天资源占用情况图（<font color=red>红色</font>表示被占用时间段）
      </center>
  
</div>
  <center>
<table width="96%" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<%
'显示图表
'打开数据库，读出设备
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr bgcolor="#FFFFFF">
<%
if resourceflag=0 then
%>
<td width=90 align=center bgcolor="D7E8F8"><a href="editequipment.asp?id=<%=rs("id")%>"><%=rs("equipment")%></a></td>
<td align=center bgcolor="#FFFFFF">
<%
else
%>
<td width=90 align=center><%=rs("equipment")%></td><td align=center bgcolor="#FFFFFF">
<%
end if
'读出预约数据
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from booking where endtime>#" & today1 & "# and starttime<#" & today2 & "# and equipment=" & sqlstr(rs("equipment")) & " order by starttime"
'response.write sql
rs1.open sql,conn,1
%>
<table border="0"  cellspacing="0" cellpadding="0" width="430" height="10">
<tr>
<td align=left>00:00</td><td align=center>12:00</td><td align=right>24:00</td>
</tr>
</table>
<%
if rs1.eof or rs1.bof then
%>

        <table border="0"  cellspacing="0" cellpadding="0" width="400" height="5">
          <tr>
            <td bgcolor="3499D0"><font color="#006699">1</font></td>
</tr>
</table>
<%
else
%>
        <table border="0"  cellspacing="0" cellpadding="0" width="400" height="5">
          <tr>
<%
'初始化前面没记录
a=0
while not rs1.eof and not rs1.bof
%>

<%
'如果是第一条记录
if a=0 then
'现在不是第一条记录了
a=1
'计算开始时间到零点的小时数
starttime=rs1("starttime")
endtime=rs1("endtime")
starttime1=rs1("starttime")
endtime1=rs1("endtime")
a1=datediff("h",today1,starttime)
'如果开始时间在今天就显示无色的长度
if a1>0 then
'计算色条长度
colorwidth=a1/24*400
'显示无色
%>
            <td bgcolor="3499D0" width="<%=colorwidth%>" align=center height="8"><%=a1%></td>
<%
end if
'计算结束时间到零点的小时数
b1=datediff("h",today1,endtime)
'如果a1<=0,b1>=24
if a1<=0 and b1>=24 then
%>
            <td bgcolor="#FF0000" width="400" align=center height="8">24</td>
<%
end if
'如果a1<=0,b1<24
if a1<=0 and b1<24 then
colorwidth=b1/24*400
%>
            <td bgcolor="#FF0000" width="<%=colorwidth%>" align=center height="8"><%=b1%></td>
<%
end if
'如果a1>0,b1>=24
if a1>0 and b1>=24 then
colorwidth=(24-a1)/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=(24-a1)%></td>
<%
end if
'如果a1>0,b1<24
if a1>0 and b1<24 then
c1=datediff("h",starttime,endtime)
colorwidth=c1/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=c1%></td>
<%
end if
else
'计算无色长度
starttime=rs1("starttime")
endtime=rs1("endtime")
a1=datediff("h",endtime1,starttime)
'如果有时间间隔就显示无色长度
if a1>0 then
colorwidth=a1/24*400
%>
            <td bgcolor="3499D0" width="<%=colorwidth%>" align=center height="8"><%=a1%></td>
<%
end if
'显示有色长度
b1=datediff("h",today1,endtime)
'如果b1>=24
if b1>=24 then
d1=datediff("h",today1,starttime)
colorwidth=(24-d1)/24*400
%>
            <td bgcolor="#FF0000" width="<%=colorwidth%>" align=center height="8">8<%=(24-d1)%></td>
<%
end if
'如果b1<24
if b1<24 then
d1=datediff("h",starttime,endtime)
colorwidth=d1/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=d1%></td>
<%
end if
starttime1=rs1("starttime")
endtime1=rs1("endtime")
end if
rs1.movenext
wend
'显示最后无色区
e1=datediff("h",today1,endtime1)
if e1<24 then
colorwidth=(24-e1)/24*400
%>
<td bgcolor="3499D0" width="<%=colorwidth%>" align=center><%=(24-e1)%></td>
<%
end if
%>
</tr>
</table>
<%
end if
%>
</td>
</tr>
<%
rs.movenext
wend
%>
</table>
  <br>
  <table width="96%" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<%
'打开数据库，读出设备
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr>
<td bgcolor="D7E8F8"><%=rs("equipment")%></td>
</tr>
<tr>
<td height=40 bgcolor="#FFFFFF">
<%
'读出预约数据
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from booking where equipment=" & sqlstr(rs("equipment")) & " order by starttime"
rs1.open sql,conn,1
while not rs1.eof and not rs1.bof
	select case rs1("auditing")
		case 0
			imgsrc="image/resource_auditing.gif"
			explainstr="<font color=""#0000ff"">未审核</font>"
		case 1
			nowtime=cdate(cstr(date())&" "&cstr(time()))
			if nowtime>=rs1("starttime") and nowtime<=rs1("endtime") then
				imgsrc="image/resource_go.gif"
				explainstr="使用中..."
			elseif nowtime<rs1("starttime") then
				imgsrc="image/stay_do.gif"
				explainstr="等待使用"
			elseif nowtime>rs1("endtime") then
				imgsrc="image/finish.gif" 
				explainstr="已完成"
			end if
		case 2
			imgsrc="image/auditing_no.gif"
			explainstr="<font color=""#ff0000"">审核未通过</font>"
	end select
%>
<img src="<%=imgsrc%>" border="0">
<a href="editbooking.asp?id=<%=rs1("id")%>"><font color="#0000ff">[<%=rs1("starttime")%>----<%=rs1("endtime")%>]</font></a>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs2=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(rs1("username"))
rs2.open sql,conn,1
if not rs2.eof and not rs2.bof then
if oabusyusername=rs1("username") then response.write "<font color=red>"
%>
&nbsp;&nbsp;部门：<%=rs2("userdept")%>&nbsp;&nbsp;预约者：<%=rs2("name")%><%
if oabusyusername=rs1("username") then response.write "</font>"
response.write("["&explainstr&"]")
end if
%>
<br>
<%
rs1.movenext
wend

rs.movenext
wend
%>
</table>
</center></td>
    </tr>
</table>
</body>
</html>










