<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
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
	response.write("history.go(-1);")
	Response.Write("</script>")
	response.end
End Function
'������ʱ��
function getnewtime(oldtime,addtime)
	dim hourvalue,minutevalue,newminute,newtime
	hourvalue=hour(oldtime)
	minutevalue=minute(oldtime)+addtime
	hourvalue=hourvalue+fix(minutevalue/60)
	newminute=minutevalue mod 60
	newtime=timeserial(hourvalue,newminute,0)
	getnewtime=newtime
end function
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("alert(""���Ѿ�����,�����µ�¼ϵͳ!"");")
	response.write("window.close()")
	response.write("</script>")
	response.end
end if
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
amgonokq=rs("amgonokq")
pmcomenokq=rs("pmcomenokq")
pmgonokq=rs("pmgonokq")
set rs=nothing
amcometimephase1=getnewtime(amcometime,-kqtimephase)
amcometimephase2=getnewtime(amcometime,kqtimephase)
amgotimephase1=getnewtime(amgotime,-kqtimephase)
amgotimephase2=getnewtime(amgotime,kqtimephase)
pmcometimephase1=getnewtime(pmcometime,-kqtimephase)
pmcometimephase2=getnewtime(pmcometime,kqtimephase)
pmgotimephase1=getnewtime(pmgotime,-kqtimephase)
pmgotimephase2=getnewtime(pmgotime,kqtimephase)
'�жϴ����ı���ֵ
nowtime=time()
lookkqinfo=0
if nowtime<amcometimephase1 then
	lookkqinfo=0
elseif nowtime>=amcometimephase1 and nowtime<amgotimephase1 then
	lookkqinfo=1
elseif nowtime>=amgotimephase1 and nowtime<pmcometimephase1 then
	lookkqinfo=2
elseif nowtime>=pmcometimephase1 and nowtime<pmgotimephase1 then
	lookkqinfo=3
elseif nowtime>=pmgotimephase1 then
	lookkqinfo=4
end if
kqtimephase=request.form("kqtimephase")
if kqtimephase="amcome" then
	amorpmvalue="am"
	goorcomevalue="come"
elseif kqtimephase="amgo" then
	amorpmvalue="am"
	goorcomevalue="go"
elseif kqtimephase="pmcome" then
	amorpmvalue="pm"
	goorcomevalue="come"
elseif kqtimephase="pmgo" then
	amorpmvalue="pm"
	goorcomevalue="go"
end if
username=request.form("username")
if username<>"" then
	yystr=trim(request.form("yy"))
	if yystr="" then
		call disperrinfo("������ԭ��")
		response.end
	end if
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select name,username,userdept from userinf where username='"&username&"'"
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		call DispErrInfo("�Բ���û������û���")
	else
		name=rs("name")
		conn.close
		set rs=nothing
		set conn=nothing
	end if
	set rs=server.createobject("adodb.recordset")
	sql="select * from month"&cstr(month(date()))&" where day=#"&date()&"# and username='"&username&"' and amorpm='"&amorpmvalue&"'"
	rs.open sql,kqconn,3,2
	if rs.eof or rs.bof then
		if goorcomevalue="go" then
			comedatevalue="00:00:00"
			godatevalue=time()
			explain1=""
			explain2=yystr
		else
			comedatevalue=time()
			godatevalue="00:00:00"
			explain1=yystr
			explain2=""
		end if
		sql="insert into month"&cstr(month(date()))&" (username,name,dept,day,comedate,leavedate,amorpm,explain1,explain2) values('"&username&"','"&name&"','"&oabusyuserdept&"',#"&date()&"#,#"&comedatevalue&"#,#"&godatevalue&"#,'"&amorpmvalue&"','"&explain1&"','"&explain2&"')"
		kqconn.execute(sql)
	else
		if goorcomevalue="go" then
			if cstr(rs("leavedate"))<>"" and rs("leavedate")<>#0:00:00# then
				kqconn.close
				set rs=nothing
				set kqconn=nothing
				call disperrinfo("�Բ����������ظ����ڣ�")
			else
				rs("leavedate")=time()
				rs("explain2")=yystr
			end if
		else
			if cstr(rs("comedate"))<>"" and rs("comedate")<>#0:00:00# then
				kqconn.close
				set rs=nothing
				set kqconn=nothing
				call disperrinfo("�Բ����������ظ����ڣ�")
			else
				rs("comedate")=time()
				rs("explain1")=yystr
			end if
		end if
		rs.update
	end if
	kqconn.close
	set rs=nothing
	set kqconn=nothing
	response.write("<script language=""javascript"">")
	response.write("opener.location.href=""kqmain.asp"";")
	response.write("window.close();")
	response.write("</script>")
	response.end
end if
%> 
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="pragma" content="no-cache">
<title>������</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<style type="text/css">
<!--
.style4 {color: #2e4869}
-->
</style>
</head>
<body bgcolor="#F9F9FF">
<table width="420"  border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="567B98">
  <tr>
    <td bgcolor="#FFFFFF"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="1"><img src="../images/main/l4.gif" width="1" height="21"></td>
                <td background="../images/main/m4.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="10">&nbsp;</td>
                      <td><span class="style4">������</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="../images/main/r4.gif" width="1" height="21"></td>
              </tr>
          </table></td>
        </tr>
        <tr>
          <td><%
	if lookkqinfo=0 then
		response.write("<p align=""center"">�Բ������ڲ��ǿ���ʱ�䣡</p>")
	else
%>
<br><div align="center">
  <center>
<form method="POST" action="handkq.asp" name="form1">
  <div align="center">
    <center>
   <table border="1" cellpadding="0" cellspacing="0" width="400" height="198" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
      <tr bgcolor="D7E8F8">
        <td width="396" height="25" colspan="2"><font color="#ee0000">ע�⣺</font>����д��������ԭ��</td>
      </tr>
      <tr>
        <td height="25" width="396" colspan="2">�û�����<%=oabusyname%>
		<input type="hidden" name="username" value="<%=oabusyusername%>">
</td>
      </tr>
      <tr bgcolor="D7E8F8">
        <td width="396" height="25" colspan="2">ѡ�񲹿���ʱ���:
          <input type="radio" name="kqtimephase" value="amcome" checked>�����ϰ�
<%
if lookkqinfo>=2 and amgonokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""amgo"">�����°�")
end if
if lookkqinfo>=3 and pmcomenokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""pmcome"">�����ϰ�")
end if
if lookkqinfo>=4 and pmgonokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""pmgo"">�����°�")
end if
%>	</td>
      </tr>
      <tr>
        <td width="49" height="90" bgcolor="D7E8F8">ԭ��<br>        </td>
        <td height="90" width="345"><textarea rows="8" name="yy" cols="46"></textarea></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center">
  <input type="submit" value="ȷ��" name="okbutton">&nbsp;&nbsp;&nbsp;
  <input type="button" value="�ر�" onclick="window.close();">
  </p>
</form>
  </center>
</div>
<%
end if
%></td>
        </tr>
    </table></td>
  </tr>
</table>

</body>

</html>
