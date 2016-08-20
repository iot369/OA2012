<%@ LANGUAGE = VBScript %>
<%Response.Expires=0%>
<!--#include file="conn.asp"-->
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
<%
Function DispErrInfo(ErrInfo)
	Response.Write("<script language=""javascript"">")
	Response.Write("alert("&chr(34)&ErrInfo&chr(34)&");")
	Response.Write("parent(""banner1"").location.href=""kqcheck.asp"";")
	response.write("parent(""banner"").location.href=""kqmain.asp"";")
	Response.Write("</script>")
End Function
'返回新时间
function getnewtime(oldtime,addtime)
	dim hourvalue,minutevalue,newminute,newtime
	hourvalue=hour(oldtime)
	minutevalue=minute(oldtime)+addtime
	hourvalue=hourvalue+fix(minutevalue/60)
	newminute=minutevalue mod 60
	newtime=timeserial(hourvalue,newminute,0)
	getnewtime=newtime
end function
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>用户登录</title>
<link rel="stylesheet" type="text/css" href="../style.css">
</head>
<body>
<%
randnumber=request("randnumber")
If Ucase(Session("key"))=Ucase(Request("key")) Then
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select name,username,userdept from userinf where epassserialnumber='"&randnumber&"'"
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		call DispErrInfo("对不起，没有这个用户！")
	else
		username=rs("username")
		name=rs("name")
		dept=rs("userdept")
		set rs=nothing
		conn.close
		set conn=nothing
		set kqconn=openconn("kq")
		set rs=server.createobject("adodb.recordset")
		sql="select * from inittime"
		rs.open sql,kqconn,1
		amcometime=rs("amondutytime")
		amgotime=rs("amoffdutytime")
		pmcometime=rs("pmondutytime")
		pmgotime=rs("pmoffdutytime")
		comedelaytime=rs("ondutydelaytime")
		goaheadtime=rs("offdutyaheadtime")
		kqtimephase=rs("kqtimephase")
		set rs=nothing
		amcometimephase1=getnewtime(amcometime,-kqtimephase)
		amcometimephase2=getnewtime(amcometime,kqtimephase)
		amgotimephase1=getnewtime(amgotime,-kqtimephase)
		amgotimephase2=getnewtime(amgotime,kqtimephase)
		pmcometimephase1=getnewtime(pmcometime,-kqtimephase)
		pmcometimephase2=getnewtime(pmcometime,kqtimephase)
		pmgotimephase1=getnewtime(pmgotime,-kqtimephase)
		pmgotimephase2=getnewtime(pmgotime,kqtimephase)
		nowtime=time()
		if amcometimephase1<=nowtime and amcometimephase2>=nowtime then
			amorpmvalue="am"
			goorcomevalue="come"
		elseif amgotimephase1<=nowtime and amgotimephase2>=nowtime then
			amorpmvalue="am"
			goorcomevalue="go"
		elseif pmcometimephase1<=nowtime and pmcometimephase2>=nowtime then
			amorpmvalue="pm"
			goorcomevalue="come"
		elseif pmgotimephase1<=nowtime and pmgotimephase2>=nowtime then
			amorpmvalue="pm"
			goorcomevalue="go"
		else
			amorpmvalue=""
		end if
		if amorpmvalue<>"" then
			set rs=server.createobject("adodb.recordset")
			sql="select * from month"&cstr(month(date()))&" where day=#"&date()&"# and username='"&username&"' and amorpm='"&amorpmvalue&"'"
			rs.open sql,kqconn,3,2
			if rs.eof or rs.bof then
				if goorcomevalue="go" then
					comedatevalue="00:00:00"
					godatevalue=time()
				else
					comedatevalue=time()
					godatevalue="00:00:00"
				end if
				sql="insert into month"&cstr(month(date()))&" (username,name,dept,day,comedate,leavedate,amorpm) values('"&username&"','"&name&"','"&dept&"',#"&date()&"#,#"&comedatevalue&"#,#"&godatevalue&"#,'"&amorpmvalue&"')"
				kqconn.execute(sql)
			else
				if goorcomevalue="go" then
					if cstr(rs("leavedate"))<>"" and rs("leavedate")<>#0:00:00# then
						kqconn.close
						set rs=nothing
						set kqconn=nothing
						call disperrinfo("对不起，您不能重复考勤！")
						response.end
					else
						rs("leavedate")=time()
					end if
				else
					if cstr(rs("comedate"))<>"" and rs("comedate")<>#0:00:00# then
						kqconn.close
						set rs=nothing
						set kqconn=nothing
						call disperrinfo("对不起，您不能重复考勤！")
						response.end
					else
						rs("comedate")=time()
					end if
				end if
				rs.update
			end if
			kqconn.close
			set rs=nothing
			set kqconn=nothing
			response.write("<script language=""javascript"">")
			response.write("parent(""banner"").location.href=""kqmain.asp"";")
			response.write("alert(""考勤成功，请确定并拨出设备！"");")
			response.write("location.href=""kqcheck.asp"";")
			response.write("</script>")
		else
			call disperrinfo("对不起，现在不能考勤，请拨出，单击“手工考勤”补考勤！")
		end if
	end if
Else
	call disperrinfo("对不起，出现错误，没有该用户！")
End If
%>
</body>
</html>