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
yearvalue=request("yearvalue")
monthvalue=request("monthvalue")
username=request("name")
userdept=request("dept")
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>��ӡ�¿�������</title>
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<center>
<br>
<font color="#000000" size="+1"><%=userdept%></font>
<%
	set conn=opendb("oabusy","conn","accessdsn")
	set rs1=server.createobject("adodb.recordset")
	sql="select DISTINCT username,name from userinf where userdept='"&userdept&"' and username='"&username&"'"
	rs1.open sql,conn,1
	while not rs1.eof and not rs1.bof
		if rs1("username")=username then
			namevalue=rs1("name")
		end if
		rs1.movenext
	wend
	conn.close
	set conn=nothing
	set rs1=nothing
if username<>"" then
%>
<br><br><center><font color="#dd0000"><%=namevalue%></font><%=cstr(yearvalue)&"��"&cstr(monthvalue)&"�¿���ͳ�Ʊ�"%></center>
<br>
<center>
<table border="1" cellpadding="0" cellspacing="0" width="550" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
  <tr>
    <td bgcolor="#efefef" height="30" width="104" align="center">����</td>
    <td bgcolor="#efefef" height="30" width="66" align="center">����</td>
    <td bgcolor="#efefef" height="30" width="70" align="center">�ϰ�ʱ��</td>
    <td bgcolor="#efefef" height="30" width="70" align="center">�°�ʱ��</td>
    <td bgcolor="#efefef" height="30" width="228" align="center">˵��</td>
  </tr>
<%
'��ȡ�Ͽ���ʱ����Ϣ
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
'������������ϰ�ʱ�仹���°�ʱ��
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
				weekvalue="������"
			case 2
				weekvalue="����һ"
			case 3
				weekvalue="���ڶ�"
			case 4
				weekvalue="������"
			case 5
				weekvalue="������"
			case 6
				weekvalue="������"
			case 7
				weekvalue="������"
		end select
		if weekvalue<>"������" and weekvalue<>"������" then
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
		call disposeamcometime()'���������ϰ�ʱ��
		if amgonokq=0 then
			call disposeamgotime()'���������°�ʱ��
		end if
		if pmcomenokq=0 then
			call disposepmcometime()'���������ϰ�ʱ��
		end if
		if pmgonokq=0 then
			call disposepmgotime()'���������°�ʱ��
		end if
		bgcolorvalue="#ffffff"
	else
		getamcometime="&nbsp;"
		getamgotime="&nbsp;"
		getamexplain="&nbsp;"
		getpmcometime="&nbsp;"
		getpmgotime="&nbsp;"
		getpmexplain="&nbsp;"
		bgcolorvalue="#efefef"
	end if
%>
  <tr bgcolor="<%=bgcolorvalue%>" height="20">
    <td width="104" align="center" rowspan="2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><%=cstr(datevalueis)%></td>
    <td width="66" align="center" rowspan="2" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF"><%=weekvalue%></td>
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
	next
%>
</table>
</center>
<br><br>
<font color="#000000" size="+1">����ͳ��</font><br><br>
����δ��������<font color="#0000ee"><%=cstr(amnocomesums)%></font>��&nbsp;&nbsp;
����ٵ�������<font color="#0000ee"><%=cstr(amlatesums)%></font>��&nbsp;&nbsp;
�������˴�����<font color="#0000ee"><%=cstr(amleaveearlysums)%></font>��&nbsp;&nbsp;
<br><br>
����δ��������<font color="#0000ee"><%=cstr(pmnocomesums)%></font>��&nbsp;&nbsp;
����ٵ�������<font color="#0000ee"><%=cstr(pmlatesums)%></font>��&nbsp;&nbsp;
�������˴�����<font color="#0000ee"><%=cstr(pmleaveearlysums)%></font>��&nbsp;&nbsp;
<br>
<%
kqconn.close
set kqconn=nothing
end if
%>
<script language="javascript">
if (confirm('�뵥����ȷ������ť��ʼ��ӡ��������ȡ������ť����ӡ��'))
{
	window.print();
}
</script>
</body>
</html>