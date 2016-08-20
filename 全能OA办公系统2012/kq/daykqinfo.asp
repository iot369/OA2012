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
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<script language="javascript">
function printsub()
{
	window.open('printdaykqinfo.asp?dept='+document.form1.userdept.value+'&yearvalue='+document.form1.yearvalue.value+'&monthvalue='+document.form1.monthvalue.value+'&dayvalue='+document.form1.dayvalue.value,'kqprintwindow','location=no,height=450, width=600, toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no');
}
</script>
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

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
              <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">考勤系统</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <p align="center">&nbsp;</p>
  <table>
<tr>
<%
yearvalue=request("yearvalue")
monthvalue=request("monthvalue")
dayvalue=request("dayvalue")
userdept=request("userdept")
if userdept="" then
	userdept=oabusyuserdept
end if
if yearvalue="" then
	yearvalue=year(date())
end if
if monthvalue="" then
	monthvalue=month(date())
end if
if dayvalue="" then
	dayvalue=day(date())
end if
%>
<td><center><font size="+1"><%=userdept%></font><font color="#ee0000" size="+1"><%=yearvalue&"年"&monthvalue&"月"&dayvalue&"日"%></font><font size="+1">考勤统计</font></center></td></tr>
<tr>
<form method="post" action="daykqinfo.asp" name="form1">
<td align="center">
<select size="1" name="yearvalue">
<%
for i=2001 to cint(year(date()))
	if cstr(i)=cstr(yearvalue) then
		response.write("<option selected value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"年"&"</option>") 
	else
		response.write("<option value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"年"&"</option>") 
	end if
next
%>
</select>
<select size="1" name="monthvalue">
<%
for i=1 to 12
	if cstr(i)=cstr(monthvalue) then
		response.write("<option selected value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"月"&"</option>") 
	else
		response.write("<option value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"月"&"</option>") 
	end if
next
%>
</select>
<select size="1" name="dayvalue">
<%
select case cint(monthvalue)
	case 1,3,5,7,8,10,12
		monthdayvalue=31
	case 4,6,9,11
		monthdayvalue=30
	case 2
		if ((cint(yearvalue) mod 4)=0) or ((cint(yearvalue) mod 100=0) and (cint(yearvalue) mod 400)<>0) then
			monthdayvalue=29
		else
			monthdayvalue=28
		end if
end select 
for i=1 to monthdayvalue
	if cstr(i)=cstr(dayvalue) then
		response.write("<option selected value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"日</option>")
	else
		response.write("<option value="&chr(34)&cstr(i)&chr(34)&">"&cstr(i)&"日</option>")
	end if
next
%>
</select>
<%
allow_look_all_kq_info=lookkqpopedom()
if allow_look_all_kq_info="yes" then
'读取部门和人员
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select DISTINCT userdept from userinf"
	rs.open sql,conn,1
%>
<select size=1 name="userdept">
<%
	if not rs.eof and not rs.bof then
		userdept=rs("userdept")
	end if
	if request("userdept")<>"" then
		userdept=request("userdept")
	end if
	while not rs.eof and not rs.bof
%>
<option value="<%=rs("userdept")%>"<%=selected(userdept,rs("userdept"))%>><%=rs("userdept")%></option>
<%
		rs.movenext
	wend
	conn.close
	set conn=nothing
	set rs=nothing
%>
</select>
<%
end if
%>
<input type="submit" value="查看">
<%
if allow_look_all_kq_info="yes" then
%>
<input type="button" name="pintbtn" value="打印" onclick="printsub();">
<%
end if
%>
</td>
</form>
</tr>
</table>
</center>
<%

'读取用户表
set conn=opendb("oabusy","conn","accessdsn")
set kqconn=opennewdb("kq",yearvalue)
set rs=server.createobject("adodb.recordset")
sql="select name,username from userinf where userdept=" & sqlstr(userdept)
rs.open sql,conn,1
if rs.eof or rs.bof then
	response.write("<p align=""center""><font color=""#ee0000"">该部门还没有用户！</font>")
else
%>
<br>
<center>
<table border="1" cellpadding="0" cellspacing="0" width="540" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
  <tr bgcolor="D7E8F8">
    <td width="66" height="30" align="center">姓名</td>
    <td width="104" height="30" align="center">部门</td>
    <td width="70" height="30" align="center">上班时间</td>
    <td width="70" height="30" align="center">下班时间</td>
    <td width="228" height="30" align="center">说明</td>
  </tr>
<%
'读取上考勤时间信息
	set kqconn=opennewdb("kq",yearvalue)
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
	i=0
	newdate=dateserial(yearvalue,monthvalue,dayvalue)
	do while not rs.eof
		getamcometime=""
		getamgotime=""
		getamexplain=""
		getpmcometime=""
		getpmgotime=""
		getpmexplain=""
		sql1="select * from month"&monthvalue&" where day=#"&newdate&"# and  username='"&rs("username")&"' and dept='"&userdept&"' and amorpm='am'"
		sql2="select * from month"&monthvalue&" where day=#"&newdate&"# and username='"&rs("username")&"' and dept='"&userdept&"' and amorpm='pm'"
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
		if i mod 2=0 then
			bgcolorvalue="#EBF3FC"
		else
			bgcolorvalue="#ffffff"
		end if
		i=i+1
%>
  <tr bgcolor="<%=bgcolorvalue%>" height="20">
    <td width="66" align="center" rowspan="2" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF"><%=server.htmlencode(rs("name"))%></td>
    <td width="104" align="center" rowspan="2" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF"><%=server.htmlencode(userdept)%></td>
    <td width="70" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF"><%=getamcometime%></td>
    <td width="70" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
<%
if amgonokq=0 then
	response.write(getamgotime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="228" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
<%
if getamexplain="" then
	getamexplain="&nbsp;"
end if
response.write(getamexplain)
%>
	</td>
  </tr>
  <tr bgcolor="<%=bgcolorvalue%>"  height="20">
    <td width="70" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
<%
if pmcomenokq=0 then
	response.write(getpmcometime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="70" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
<%
if pmgonokq=0 then
	response.write(getpmgotime)
else
	response.write("&nbsp;")
end if
%>
	</td>
    <td width="228" align="center" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF">
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
</body>
</html>