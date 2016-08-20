<%@ LANGUAGE = VBScript %>
<!--#include file="../asp/sqlstr.asp"-->
<!--#include file="../asp/bgsub.asp"-->
<%
'session.abandon
'Server.ScriptTimeOut=500
function opendb(DBPath,sessionname,dbsort)
dim conn
'if not isobject(session(sessionname)) then
Set conn=Server.CreateObject("ADODB.Connection")
'if dbsort="accessdsn" then conn.Open "DSN=" & DBPath
'if dbsort="access" then conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
'if dbsort="sqlserver" then conn.Open "DSN=" & DBPath & ";uid=wsw;pwd=wsw"
DBPath1=server.mappath("../db/lmtof.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<!--#include file="../asp/checked.asp"-->
<!--#include file="conn.asp"-->
<!--#include file="publicsub.asp"-->
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("alert(""对不起，您已经过期，请重新登录！"");")
	response.wirte("window.close();")
	response.write("</script>")
	response.end
end if
userdept=request("dept")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>打印部门今日考勤资料</title>
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<center><br><font color="#000000" size="+1"><%=userdept%></font><br><br><%=cstr(year(date()))&"年"&cstr(month(date()))&"月"&cstr(day(date()))&"日"%>考勤资料表</center>
<%
'读取用户表
set conn=opendb("oabusy","conn","accessdsn")
set kqconn=openconn("kq")
set rs=server.createobject("adodb.recordset")
sql="select name,username from userinf where userdept=" & sqlstr(userdept)
rs.open sql,conn,1
if rs.eof or rs.bof then
	response.write("<p align=""center""><font color=""#ee0000"">该部门还没有用户！</font>")
else
%>
<br>
<center>
<table border="1" cellpadding="0" cellspacing="0" width="550" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td bgcolor="#efefef" height="30" width="66" align="center">姓名</td>
    <td bgcolor="#efefef" height="30" width="104" align="center">部门</td>
    <td bgcolor="#efefef" height="30" width="70" align="center">上班时间</td>
    <td bgcolor="#efefef" height="30" width="70" align="center">下班时间</td>
    <td bgcolor="#efefef" height="30" width="228" align="center">说明</td>
  </tr>
<%
'读取上考勤时间信息
	set kqconn=openconn("kq")
	set rs3=server.createobject("adodb.recordset")
	sql="select * from inittime"
	rs3.open sql,kqconn,1
	amcometime=rs3("amondutytime")
	amgotime=rs3("amoffdutytime")
	pmcometime=rs3("pmondutytime")
	pmgotime=rs3("pmoffdutytime")
	comedelaytime=rs3("ondutydelaytime")
	goaheadtime=rs3("offdutyaheadtime")
	kqtimephase=rs3("kqtimephase")
	amgonokq=rs3("amgonokq")
	pmcomenokq=rs3("pmcomenokq")
	pmgonokq=rs3("pmgonokq")
	set rs3=nothing
	amcometimephase1=getnewtime(amcometime,-kqtimephase)
	amcometimephase2=getnewtime(amcometime,kqtimephase)
	amgotimephase1=getnewtime(amgotime,-kqtimephase)
	amgotimephase2=getnewtime(amgotime,kqtimephase)
	pmcometimephase1=getnewtime(pmcometime,-kqtimephase)
	pmcometimephase2=getnewtime(pmcometime,kqtimephase)
	pmgotimephase1=getnewtime(pmgotime,-kqtimephase)
	pmgotimephase2=getnewtime(pmgotime,kqtimephase)
'返回上下午和上班时间还是下班时间
	nowtime=time()
	amorpmvalue=""
	goorcomevalue=""
	lookkqinfo=""
	if nowtime<amcometimephase1 then
		lookkqinfo="no"
	elseif nowtime>=amcometimephase1 and nowtime<amgotimephase1 then
		lookkqinfo="amandcome"
	elseif nowtime>=amgotimephase1 and nowtime<pmcometimephase1 then
		lookkqinfo="amall"
	elseif nowtime>=pmcometimephase1 and nowtime<pmgotimephase1 then
		lookkqinfo="amallandpmcome"
	elseif nowtime>=pmgotimephase1 then
		lookkqinfo="amandpm"
	end if
	public getamcometime,getamgotime,getamexplain,getpmcometime,getpmgotime,getpmexplain
	do while not rs.eof
		getamcometime=""
		getamgotime=""
		getamexplain=""
		getpmcometime=""
		getpmgotime=""
		getpmexplain=""
		sql1="select * from month"&cstr(month(date()))&" where day=#"&date()&"# and  username='"&rs("username")&"' and dept='"&userdept&"' and amorpm='am'"
		sql2="select * from month"&cstr(month(date()))&" where day=#"&date()&"# and username='"&rs("username")&"' and dept='"&userdept&"' and amorpm='pm'"
		set rs1=server.createobject("adodb.recordset")
		set rs2=server.createobject("adodb.recordset")
		rs1.open sql1,kqconn,1
		rs2.open sql2,kqconn,1
		select case lookkqinfo
			case "no"
				getamcometime="00:00:00"
				getamgotime="00:00:00"
				getamexplain="<font color='#ee0000'>不能取得资料</font>"
				getpmcometime="00:00:00"
				getpmgotime="00:00:00"
				getpmexplain="<font color='#ee0000'>不能取得资料</font>"
			case "amandcome"
				call disposeamcometime()'处理上午上班时间
				getpmcometime="00:00:00"
				getpmgotime="00:00:00"
				getpmexplain="<font color='#ee0000'>不能取得资料</font>"
			case "amall"
				call disposeamcometime()'处理上午上班时间
				if amgonokq=0 then
					call disposeamgotime()'处理上午下班时间
				end if
				getpmcometime="00:00:00"
				getpmgotime="00:00:00"
				getpmexplain="<font color='#ee0000'>不能取得资料</font>"
			case "amallandpmcome"
				call disposeamcometime()'处理上午上班时间
				if amgonokq=0 then
					call disposeamgotime()'处理上午下班时间
				end if
				if pmcomenokq=0 then
					call disposepmcometime()'处理下午上班时间
				end if
				getpmgotime="00:00:00"
			case "amandpm"
				call disposeamcometime()'处理上午上班时间
				if amgonokq=0 then
					call disposeamgotime()'处理上午下班时间	
				end if
				if pmcomenokq=0 then
					call disposepmcometime()'处理下午上班时间
				end if 
				if pmgonokq=0 then
					call disposepmgotime()'处理下午下班时间	
				end if
		end select 
%>
  <tr bgcolor="#ffffff" height="20">
    <td width="66" align="center" rowspan="2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><%=server.htmlencode(rs("name"))%></td>
    <td width="104" align="center" rowspan="2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><%=server.htmlencode(userdept)%></td>
    <td width="70" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><%=getamcometime%></td>
    <td width="70" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<%
if amgonokq=0 then
	response.write(getamgotime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="228" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<%
if getamexplain="" then
	getamexplain="&nbsp;"
end if
response.write(getamexplain)
%>
	</td>
  </tr>
  <tr bgcolor="<%=bgcolorvalue%>"  height="20">
    <td width="70" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<%
if pmcomenokq=0 then
	response.write(getpmcometime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="70" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<%
if pmgonokq=0 then
	response.write(getpmgotime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="228" align="center" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
<%
if getpmexplain="" then
	getpmexplain="&nbsp;"
end if
response.write(getpmexplain)
%>
	</td>
  </tr>
<%
	rs.movenext
	loop
%>
</table>
</center>
<br>
<%
end if
conn.close
set conn=nothing
set rs=nothing
%>
<script language="javascript">
if (confirm('请单击“确定”按钮开始打印，单击“取消”按钮不打印！'))
{
	window.print();
}
</script>
</body>
</html>