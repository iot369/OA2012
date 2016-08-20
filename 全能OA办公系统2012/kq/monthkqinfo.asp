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
<style type="text/css">
<!--
-->
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
</style>
<script language="javascript1.2" src="js/openwin.js"></script>
<script language="javascript">
function printsub()
{
	window.open('printmonthkqinfo.asp?dept='+document.form1.userdept.value+'&name='+document.form1.username.value+'&yearvalue='+document.form1.yearvalue.value+'&monthvalue='+document.form1.monthvalue.value,'kqprintwindow','location=no,height=450, width=600, toolbar=no, menubar=no, scrollbars=yes, resizable=no, location=no, status=no');
}
</script>
<title>oa办公系统</title>
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
  <br>
  <table>
<tr><td align="center">
<font color="#dd0000" size="+1">月考勤统计</font>
</td></tr>
<tr><td>
<form method="post" action="monthkqinfo.asp" name="form1">
<%
yearvalue=request("yearvalue")
monthvalue=request("monthvalue")
if yearvalue="" then
	yearvalue=year(date())
end if
if monthvalue="" then
	monthvalue=month(date())
end if
%>
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
<%
allow_look_all_kq_info=lookkqpopedom()
if allow_look_all_kq_info="no" then
	userdept=oabusyuserdept
	username=oabusyusername
%>
<%
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select DISTINCT username,name from userinf where userdept='"&userdept&"'"
	rs1.open sql,conn,1
%>
<select size=1 name="username">
<%
	if not rs1.eof and not rs1.bof then
		username=rs1("username")
	end if
	if request("username")<>"" then
		username=request("username")
	end if
	while not rs1.eof and not rs1.bof
		if rs1("username")=username then
			namevalue=rs1("name")
		end if
%>
<option value="<%=rs1("username")%>"<%=selected(username,rs1("username"))%>><%=rs1("name")%></option>
<%
		rs1.movenext
	wend
	conn.close
	set conn=nothing
	set rs=nothing
	set rs1=nothing
%>
</select>
<%
else
'读取部门和人员
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select DISTINCT userdept from userinf"
	rs.open sql,conn,1
%>
<select size=1 name="userdept" onChange="document.form1.submit();">
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
%>
</select>
<%
	set rs1=server.createobject("adodb.recordset")
	sql="select DISTINCT username,name from userinf where userdept='"&userdept&"'"
	rs1.open sql,conn,1
%>
<select size=1 name="username">
<%
	if not rs1.eof and not rs1.bof then
		username=rs1("username")
	end if
	if session("userdept")=userdept and request("username")<>"" then
		username=request("username")
	end if
	while not rs1.eof and not rs1.bof
		if rs1("username")=username then
			namevalue=rs1("name")
		end if
%>
<option value="<%=rs1("username")%>"<%=selected(username,rs1("username"))%>><%=rs1("name")%></option>
<%
		rs1.movenext
	wend
	conn.close
	set conn=nothing
	set rs=nothing
	set rs1=nothing
%>
</select>
<%
session("userdept")=userdept
end if
%>
<input type="submit" value="查看">
<%
if allow_look_all_kq_info="yes" then
%>
<input type="button" name="printbtn" value="打印" onclick="printsub();">
<%
end if
%>
</td>
</form>
</tr>
</table>
</center>
<br>
<%

if username<>"" then
%>
<center>
<font color="#dd0000" size="+1"><%=namevalue%></font><font size="+1"><%=cstr(yearvalue)&"年"&cstr(monthvalue)&"月考勤统计表"%></font></center>
<br>
<center>
<table border="1" cellpadding="0" cellspacing="0" width="540" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
  <tr bgcolor="D7E8F8">
    <td width="104" height="30" align="center">日期</td>
    <td width="66" height="30" align="center">星期</td>
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
	public getamcometime,getamgotime,getamexplain,getpmcometime,getpmgotime,getpmexplain
	public amlatesums,amleaveearlysums,amnocomesums,pmlatesums,pmleaveearlysums,pmnocomesums
	amlatesums=0
	amleaveearlysums=0
	amnocomesums=0
	pmlatesums=0
	pmleaveearlysums=0
	pmnocomesums=0
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
		datevalueis=dateserial(cint(yearvalue),cint(monthvalue),i)
		weekvalue=weekday(datevalueis)
		select case weekvalue
			case 1
				weekvalue="星期日"
			case 2
				weekvalue="星期一"
			case 3
				weekvalue="星期二"
			case 4
				weekvalue="星期三"
			case 5
				weekvalue="星期四"
			case 6
				weekvalue="星期五"
			case 7
				weekvalue="星期六"
		end select
		if weekvalue<>"星期六" and weekvalue<>"星期日" then
		getamcometime=""
		getamgotime=""
		getamexplain=""
		getpmcometime=""
		getpmgotime=""
		getpmexplain=""
		sql1="select * from month"&monthvalue&" where day=#"&cstr(datevalueis)&"# and  username='"&username&"' and dept='"&userdept&"' and amorpm='am'"
		sql2="select * from month"&monthvalue&" where day=#"&cstr(datevalueis)&"# and username='"&username&"' and dept='"&userdept&"' and amorpm='pm'"
		set rs1=server.createobject("adodb.recordset")
		set rs2=server.createobject("adodb.recordset")
		rs1.open sql1,kqconn,1
		rs2.open sql2,kqconn,1
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
		if i mod 2=0 then
			bgcolorvalue="#ffffff"
		else
			bgcolorvalue="#ffffff"
		end if
	else
		getamcometime=""
		getamgotime=""
		getamexplain=""
		getpmcometime=""
		getpmgotime=""
		getpmexplain=""
		bgcolorvalue="#EBF3FC"
	end if
%>
  <tr bgcolor="<%=bgcolorvalue%>" height="20">
    <td width="104" align="center" rowspan="2" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF"><%=cstr(datevalueis)%></td>
    <td width="66" align="center" rowspan="2" bordercolorlight="#6FECFF" bordercolordark="#FFFFFF"><%=weekvalue%></td>
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
	next
%>
</table>
</center>
<br>
<table width="540" border="1" align="center" cellpadding="0" cellspacing="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
  <tr bgcolor="D7E8F8">
    <td height="25" colspan="6"><div align="center"><strong><%=namevalue%>本月考勤数据统计</strong></div></td>
  </tr>
  <tr>
    <td width="110" height="25" bgcolor="EBF3FC"><div align="center">上午未到次数：</div></td>
    <td width="66" height="25"><div align="center"><font color="#0000ee"><%=cstr(amnocomesums)%></font>次&nbsp;</div></td>
    <td width="101" height="25" bgcolor="EBF3FC"><div align="center">上午迟到次数：</div></td>
    <td width="75" height="25"><div align="center"><font color="#0000ee"><%=cstr(amlatesums)%></font>次&nbsp;</div></td>
    <td width="101" height="25" bgcolor="EBF3FC"><div align="center">上午早退次数：</div></td>
    <td width="83" height="25"><div align="center"><font color="#0000ee"><%=cstr(amleaveearlysums)%></font>次&nbsp;&nbsp;</div></td>
  </tr>
  <tr>
    <td height="25" bgcolor="EBF3FC"><div align="center">下午未到次数：</div></td>
    <td height="25"><div align="center"><font color="#0000ee"><%=cstr(pmnocomesums)%></font>次&nbsp;</div></td>
    <td height="25" bgcolor="EBF3FC"><div align="center">下午迟到次数：</div></td>
    <td height="25"><div align="center"><font color="#0000ee"><%=cstr(pmlatesums)%></font>次&nbsp;</div></td>
    <td height="25" bgcolor="EBF3FC"><div align="center">下午早退次数：</div></td>
    <td height="25"><div align="center"><font color="#0000ee"><%=cstr(pmleaveearlysums)%></font>次&nbsp;</div></td>
  </tr>
</table>
<table width="70%"  border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td><div align="center"><font color="#dd0000" size="+1"><br>
  </font><br>
  <%
kqconn.close
set kqconn=nothing
end if

%>
    </div></td>
  </tr>
</table>
<br>

</body>
</html>