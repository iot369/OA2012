<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
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
kqconn.close
set kqconn=nothing
set rs=nothing
amcometimephase1=getnewtime(amcometime,-kqtimephase)
amcometimephase2=getnewtime(amcometime,kqtimephase)
amgotimephase1=getnewtime(amgotime,-kqtimephase)
amgotimephase2=getnewtime(amgotime,kqtimephase)
pmcometimephase1=getnewtime(pmcometime,-kqtimephase)
pmcometimephase2=getnewtime(pmcometime,kqtimephase)
pmgotimephase1=getnewtime(pmgotime,-kqtimephase)
pmgotimephase2=getnewtime(pmgotime,kqtimephase)
%> 
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="pragma" content="no-cache">
<title>����ϵͳ</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script language="javascript">
var winhandle;
function bekq()
{
	winhandle=window.open('bekq.asp','handkqwin','width=450,height=380,toolbar=no,scrollbars=no,resizable=0,menubar=no');
}

function handkq()
{
	winhandle=window.open('handkq.asp','handkqwin','width=450,height=380,toolbar=no,scrollbars=no,resizable=0,menubar=no');
}

function closewin()
{
	if (winhandle!=null && !winhandle.closed)
		winhandle.close();
}
</script>
<style type="text/css">
<!--
.style1 {
	font-size: 12px;
	font-weight: bold;
}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body onunload="closewin();" topmargin="0" leftmargin="0" style="overflow-x:hidden;overflow-y:hidden">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">����ϵͳ</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<p align="center">&nbsp;</p>
<div align="center">
  <center>
  <table border="1" width="346" cellspacing="0" cellpadding="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF" bgcolor="#B0C8EA">
	<tr bgcolor="D7E8F8">
	<td width="448" height="21" colspan="2" align="center" valign="middle">
	<center>
<img src="../images/time/space.gif" name="one"><img src="../images/time/space.gif" name="two"><img src="../images/time/dgon.gif" name="three"><img src="../images/time/space.gif" name="four"><img src="../images/time/space.gif" name="five"><img src="../images/time/space.gif" name="six">
	</center>
<script language="javascript">
<!--//
img = new Array()
for(var i=0; i <= 14; i++) {
img[i] = new Image()
}
img[1].src = "../images/time/dg0.gif"
img[2].src = "../images/time/dg1.gif"
img[3].src = "../images/time/dg2.gif"
img[4].src = "../images/time/dg3.gif"
img[5].src = "../images/time/dg4.gif"
img[6].src = "../images/time/dg5.gif"
img[7].src = "../images/time/dg6.gif"
img[8].src = "../images/time/dg7.gif"
img[9].src = "../images/time/dg8.gif"
img[10].src = "../images/time/dg9.gif"
img[11].src = "../images/time/dgon.gif"
img[12].src = "../images/time/dgoff.gif"
img[13].src = "../images/time/dgam.gif"
img[14].src = "../images/time/dgpm.gif"
var base = "../images/time/dg"
var space = "../images/time/space.gif" 
var per = false

function go() {
per = true
start()
}

function start() {
if(per == true) {
var now = new Date()
var hours = now.getHours();
var ampm = (hours < 12) ? "am" : "pm"
hours = (hours > 12) ? (hours - 12) + "" : hours + ""
hours = (hours == "0") ? "12" : hours
hours = (hours < 10) ? "0" + hours : hours + ""
var minutes = now.getMinutes();
minutes = (minutes < 10) ? "0" + minutes : minutes + ""
var seconds = now.getSeconds();
seconds = (seconds < 10) ? "0" + seconds : seconds + ""
document.one.src = (hours.charAt(0)=="0") ? space : add(hours.charAt(0))
document.two.src = add(hours.charAt(1))
//document.three.src = (now.getSeconds() % 2) ? add("on") : add("off")
document.four.src = add(minutes.charAt(0))
document.five.src = add(minutes.charAt(1))
document.six.src = add(ampm)
setTimeout("start()",1000)
}
}

secflag=1;
function secondgo()
{
document.three.src = (secflag % 2) ? add("on") : add("off")
if (secflag==1)
{	secflag=2;}
else
{	secflag=1;}
setTimeout("secondgo()",500);
}

function add(it) {
return base + it + ".gif"
}
go();
secondgo();
//-->
</script>	</td>
	</tr>
    <tr bgcolor="D7E8F8">
      <td width="224" height="25">
	  �����ϰ�ʱ�䣺<font color="#0000ee"><%=cstr(amcometime)%></font><br>
	  ����ʱ��Σ�<font color="#0000ee"><%=cstr(amcometimephase1)&"-"&cstr(amcometimephase2)%></font>	  </td>
      <td width="224" height="25">
	  �����°�ʱ�䣺<font color="#0000ee"><%=cstr(amgotime)%></font><br>
	  ����ʱ��Σ�<font color="#0000ee"><%=cstr(amgotimephase1)&"-"&cstr(amgotimephase2)%></font>	  </td>
    </tr>
    <tr bgcolor="D7E8F8">
      <td width="224" height="25">
	  �����ϰ�ʱ�䣺<font color="#0000ee"><%=cstr(pmcometime)%></font><br>
	  ����ʱ��Σ�<font color="#0000ee"><%=cstr(pmcometimephase1)&"-"&cstr(pmcometimephase2)%></font>	  </td>
      <td width="224" height="25">
	  �����°�ʱ�䣺<font color="#0000ee"><%=cstr(pmgotime)%></font><br>
	  ����ʱ��Σ�<font color="#0000ee"><%=cstr(pmgotimephase1)&"-"&cstr(pmgotimephase2)%></font>	  </td>
    </tr>
    <tr bgcolor="D7E8F8">
      <td width="224" height="25">�ϰ��ӳ�ʱ�䣺<font color="#0000ee"><%=cstr(comedelaytime)%>����</font></td>
      <td width="224" height="25">�°���ǰʱ�䣺<font color="#0000ee"><%=cstr(goaheadtime)%>����</font></td>
    </tr>
    <tr bgcolor="D7E8F8">
      <td width="384" height="25" colspan="2">
        <p align="center">
        <marquee align="middle" width="446" height="12" behavior="alternate">
		<%
		nowtime=time()
		if amcometimephase1<=nowtime and amcometimephase2>=nowtime then
			strvalue="�����ǡ������ϰ࿼��ʱ�䡱������������п��ڣ�"
			strvalue1="�����ϰ�ʱ��"
			amorpmvalue="am"
			goorcomevalue="come"
		elseif amgotimephase1<=nowtime and amgotimephase2>=nowtime then
			strvalue="�����ǡ������°࿼��ʱ�䡱������������п��ڣ�"
			strvalue1="�����°�ʱ��"
			amorpmvalue="am"
			goorcomevalue="go"
		elseif pmcometimephase1<=nowtime and pmcometimephase2>=nowtime then
			strvalue="�����ǡ������ϰ࿼��ʱ�䡱������������п��ڣ�"
			strvalue1="�����ϰ�ʱ��"
			amorpmvalue="pm"
			goorcomevalue="come"
		elseif pmgotimephase1<=nowtime and pmgotimephase2>=nowtime then
			strvalue="�����ǡ������°࿼��ʱ�䡱������������п��ڣ�"
			amorpmvalue="pm"
			strvalue1="�����°�ʱ��"
			goorcomevalue="go"
		else
			strvalue="���ڲ��ǿ���ʱ�䣬��Ҫ�������뵥���������ڡ���ť��"
			amorpmvalue=""
		end if
		response.write(strvalue)
		%>
		</marquee>
      </td>
    </tr>
  </table>
  </center>
</div>
<%
if nowtime<amcometimephase1 then
	amorpmvalue=""
elseif nowtime>=amcometimephase1 and nowtime<amgotimephase1 then
	strvalue1="�����ϰ�ʱ��"
	amorpmvalue="am"
	goorcomevalue="come"
elseif nowtime>=amgotimephase1 and nowtime<pmcometimephase1 then
	strvalue1="�����°�ʱ��"
	amorpmvalue="am"
	goorcomevalue="go"
elseif nowtime>=pmcometimephase1 and nowtime<pmgotimephase1 then
	strvalue1="�����ϰ�ʱ��"
	amorpmvalue="pm"
	goorcomevalue="come"
elseif nowtime>=pmgotimephase1 then
	strvalue1="�����°�ʱ��"
	amorpmvalue="pm"
	goorcomevalue="go"
end if
if amorpmvalue<>"" then
	set kqconn=openconn("kq")
	set rs=server.createobject("adodb.recordset")
	sql="select * from month"&cstr(month(date()))&" where day=#"&date()&"# and dept='"&oabusyuserdept&"' and amorpm='"&amorpmvalue&"'"
	rs.open sql,kqconn,1
	if not rs.eof and not rs.bof then
%>
<br><div align="center">
  <center>
  <table border="1" width="450" bordercolorlight="#B0C8EA" cellspacing="0" cellpadding="0" bordercolordark="#FFFFFF">
    <tr bgcolor="D7E8F8">
      <td width="55" height="20">�û���</td>
      <td width="90" height="20">����</td>
      <td width="133" height="20"><%=strvalue1%></td>
      <td width="172" height="20">״̬</td>
    </tr>
<%
	do while not rs.eof
	  if goorcomevalue="come" then
	  	if rs("comedate")<>#0:00:00# then
%>
	<tr>
      <td width="55" height="20"><%=server.htmlencode(rs("name"))%></td>
      <td width="90" height="20"><%=server.htmlencode(rs("dept"))%></td>
      <td width="129" height="20"><%=cstr(rs("comedate"))%></td>
	  <td width="168" height="20"><font color="#ee0000">��</font><%=rs("explain1")%></td>
	</tr>
<%
		end if
	  else
	  	if rs("leavedate")<>#0:00:00# then
%>
	  <tr>
      <td width="55" height="20"><%=server.htmlencode(rs("name"))%></td>
      <td width="90" height="20"><%=server.htmlencode(rs("dept"))%></td>
      <td width="129" height="20"><%=cstr(rs("leavedate"))%></td>
	  <td width="168" height="20"><font color="#ee0000">��</font><%=rs("explain2")%></td>
	  </tr>
<%
	  	end if
 	end if
	rs.movenext
	loop
%>
</table>
</center>
</div>
<%
	end if
	kqconn.close
	set rs=nothing
	set kqconn=nothing
end if
%>
<p align="center">
<input type="button" value="��ʼ����" onclick="bekq()">
&nbsp;&nbsp;&nbsp;
<input name="button" type="button" onclick="handkq()" value="�� �� ��">

</p>
</body>
</html>
