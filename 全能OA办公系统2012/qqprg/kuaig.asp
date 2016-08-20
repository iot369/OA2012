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
DBPath1=server.ma<%@ LANGUAGE = VBScript %>
<%Response.Expires=0%>
<!--#include file="error.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")

if oabusyusername="" then
	session.abandon
	response.write("<script language=""javascript"">")
	response.write("window.close();")
	response.write("</script>")
	Response.End
end if
%>
<!--#include file="function.asp"-->
<html>
<head>
<title>OA系统即时通信工具</title>
<meta http-equiv="pragma" content="no-cache">
<meta http-equiv="expires" content="web,26 Feb 1960 08:21:57 GMT">
<meta http-equiv="Content-Type" content="text/html;charset=gb2312">
<link rel="stylesheet" type="text/css" href="style.css">
</head>
<body bgcolor="#ffffff"  scroll=no leftmargin="0" topmargin="0">
<script language="javascript">
<!--hide
var y;//页面位置
var randtime;
var ok=null;
var listtimeid=null;
var winhandle=null;
randtime=0;
y=0;

function sx()
{
	if (opener.closed)
		{
			window.open('lostuser.asp','lostwin','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
			window.close();
		}
	if (refflag.value=="1" && dispinfoflag.value=="0")
		{
			if (pageflag.value=="1")
			{
				listflag.value="0";
				parent("dispuserwin").location.href="disp_online_user.asp";
			}
			else if (pageflag.value=="0")
			{	
				listflag.value="0";
				parent("dispuserwin").location.href="disp_online_manager.asp?sitenumber="+sitenumber.value;	
			}
			else
			{
				listflag.value="0";
				parent("hiddenwin").location.href="disp_online_user.asp?sitenumber="+sitenumber.value;
			}
			listtime();
		}
	ok=setTimeout("sx()",15000+randtime*1000);
	randtime=parseInt(Math.random() * 10);
}

function window_onload()
{
	sx();
}

function listtime()
{
if (listflag.value=="1")
	{
		if (pageflag.value=="1" && dispinfoflag.value=="0")
			listonlineuser();
		else if (dispinfoflag.value=="0")
			disp_get_info();
		clearTimeout(listtimeid);
	}
else
	listtimeid=setTimeout("listtime()",1);
}

function listonlinehead(title)
{
if (dispinfoflag.value=="0")
	disp_get_info();
window("dispuserwin").document.write("<html>");
window("dispuserwin").document.write("<head>");
window("dispuserwin").document.write("<title>"+title+"</title>");
window("dispuserwin").document.write("<meta http-equiv=\"pragma\" content=\"no-cache\">");
window("dispuserwin").document.write("<meta http-equiv=\"expires\" content=\"web,26 Feb 1960 08:21:57 GMT\">");
window("dispuserwin").document.write("<meta http-equiv=\"Content-Type\" content=\"text/html; charset=gb2312\">");
window("dispuserwin").document.write("<style type=\"text/css\">");
window("dispuserwin").document.write("<!--");
window("dispuserwin").document.write("body {  font-family: \"宋体\"; font-size: 9pt;}");
window("dispuserwin").document.write("td {font-family:\"宋体\";font-size:9pt;}");
window("dispuserwin").document.write("-->");
window("dispuserwin").document.write("</style>");
window("dispuserwin").document.write("</head>");
window("dispuserwin").document.write("<body bgcolor=\"#0099FF\" text=\"#000000\" leftmargin=\"0\" topmargin=\"2\">");
}

function listonlineuser()
{
window("dispuserwin").document.open();
listonlinehead("本站在线用户");
var userinfostr,userinfodim,i,dimsums,infstr,onlineflag,onlinedate,onlineflag,okcode;
onlineflag=0;
userinfostr=onlineuser.value;
if (userinfostr.length==0)

{
	refflag.value=0;
}
else
{
	window("dispuserwin").document.write("<table border='0' bgcolor='#0099FF' width='100%'>");
	userinfodim=userinfostr.split("|");
	for (i in userinfodim)
	{	
		if(userinfodim[i]!="")
		{
		infstr=userinfodim[i].split("$");
		window("dispuserwin").document.write("<tr>");
		if (infstr[0]==usersessionid.value && infstr[2]=="0")
		{
			window("dispuserwin").document.write("<td width='100%' height='20' ONMOUSEOVER=\"this.bgColor='#0066FF';this.style.cursor='hand';\" ONMOUSEOUT=\"this.bgColor='#0099FF';\" ONCLICK=\"parent.clickdo('"+infstr[0]+"','"+infstr[1]+"','"+infstr[5]+"')\" title=\""+infstr[3]+"上线\"><img src='../qqpic/face/face"+infstr[5]+".gif' height=16 width=16>&nbsp;"+infstr[1]+"</td>");
			onlineflag=1;
		}
		else if (infstr[0]==usersessionid.value && infstr[2]=="2")
		{
			window("dispuserwin").document.write("<td width='100%' height='20' ONMOUSEOVER=\"this.bgColor='#0066FF';this.style.cursor='hand';\" ONMOUSEOUT=\"this.bgColor='#0099FF';\" ONCLICK=\"parent.clickdo('"+infstr[0]+"','"+infstr[1]+"','"+infstr[5]+"')\" title=\""+infstr[3]+"上线(注册用户)\"><img src='../qqpic/face/face"+infstr[5]+".gif' height=16 width=16>&nbsp;<font color=#e8e8e8>"+infstr[1]+"</font></td>");
			onlineflag=1;
		}
		else if (infstr[2]=="2")
		{
			window("dispuserwin").document.write("<td width='100%' height='20' ONMOUSEOVER=\"this.bgColor='#0066FF' this.style.cursor='hand';\" ONMOUSEOUT=\"this.bgColor='#0099FF';\" ONCLICK=\"parent.clickdo('"+infstr[0]+"','"+infstr[1]+"','"+infstr[5]+"')\" title=\""+infstr[3]+"上线(注册用户)\"><img src='../qqpic/face/face"+infstr[5]+".gif' height=16 width=16>&nbsp;<font color=#e8e8e8>"+infstr[1]+"</td>");
		}
		else if (infstr[2]=="0")
			window("dispuserwin").document.write("<td width='100%' height='20' ONMOUSEOVER=\"this.bgColor='#0066FF';this.style.cursor='hand';\" ONMOUSEOUT=\"this.bgColor='#0099FF';\" ONCLICK=\"parent.clickdo('"+infstr[0]+"','"+infstr[1]+"','"+infstr[5]+"')\" title=\""+infstr[3]+"上线\"><img src='../qqpic/face/face"+infstr[5]+".gif' height=16 width=16>&nbsp;"+infstr[1]+"</td>");
		else
		{
			window("dispuserwin").document.write("<td width='100%' height='20' ONMOUSEOVER=\"this.bgColor='#0066FF';this.style.cursor='hand';\" ONMOUSEOUT=\"this.bgColor='#0099FF';\" ONCLICK=\"parent.clickdo('"+infstr[0]+"','"+infstr[1]+"','"+infstr[5]+"')\" title=\""+infstr[3]+"上线(站长)\"><img src='../qqpic/face/face"+infstr[5]+".gif' height=16 width=16>&nbsp;<font color=#ffff00>"+infstr[1]+"</font></td>");
			if (infstr[0]==usersessionid.value)
				onlineflag=1;
		}
		window("dispuserwin").document.write("</tr>");
		}
	}
window("dispuserwin").document.write("</table>");
refflag.value=onlineflag;
}
window("dispuserwin").document.write("</body></html>");
window("dispuserwin").document.close();
reset_win_site();
}

function reset_win_site()
{
	screeny=window("dispuserwin").document.body.scrollHeight;
	y=removewin.value;
	if (screeny<y)
	{
	y=0;
	removewin.value=y;
	}
	window("dispuserwin").scroll(1,y);
}

//收信息
function disp_get_info()
{
var i,getinfostr,getinfodim,infstr;
getinfostr=getinfo.value;
if (getinfostr.length>0)
{
	getinfodim=getinfostr.split("|");
	for (i in getinfodim)
	{
		if (getinfodim[i]!="")
		{
			infstr=getinfodim[i].split("$");
			history.value=getinfodim[i]+"|"+history.value;
			urlstr="get_info_win.asp?id="+infstr[1]+"&yhm="+infstr[4]+"&sitemc="+infstr[2]+"&url="+infstr[3]+"&infostr="+infstr[5];
			infowin=window.open(urlstr,'','height=245,width=355,toolbar=no,scrollbars=no,resizable=0,menubar=no');
			
		}
	}
}
getinfo.value="";
}

//关闭窗口
function closewindow()
{
	if (winhandle!=null)
		if (!winhandle.closed)
			winhandle.close();
	window.open('lostuser.asp','lostwin','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
	window.close();
}

function clickup()
{
if (pageflag.value!="1" && dispinfoflag.value=="0")
{
	listflag.value="0";
	listtime();
	window("dispuserwin").location.href="disp_online_user.asp?sitenumber="+sitenumber.value;
	pageflag.value="1";
	refflag.value="1";
}
}

function changeface()
{
window.open('changeface.asp','facewindow','toolbar=no,scrollbars=no,resizable=0,menubar=no,width=302,height=330');
}

function sendsms()
{
	window.open('../sendsms.asp','smswin','toolbar=no,scrollbars=no,resizable=0,menubar=no,width=620,height=350');
}

function talk_face()
{
	winhandle=window.open('../newface/talkface.htm','talkface','toolbar=no,scrollbars=no,resizable=0,menubar=no,width=360,height=279');
	dispinfoflag.value="1";
}

function gohomepage()
{
var urlmc;
urlmc="../default.htm";
window.open(urlmc,'','toolbar=yes,scrollbars=yes,status=yes,resizable=1,menubar=yes,location=yes,width=600,height=400');
}

function clickdo(number,name,facenumber)
{
if (dispinfoflag.value=="0")
	sendwin=window.open('send_info_win.htm',number,'height=245,width=355,toolbar=no,scrollbars=no,resizable=0,menubar=no');
else
{
	winhandle.userid.value=number;
	winhandle.userlist.document.close();
	winhandle.talkinfo.value="对"+name+"说：";
	winhandle.info.focus();
}
}

//-->
</script>
<SCRIPT language=javascript>
<!--
if (window.Event) 
　document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
　if (e.which == 2 || e.which == 3)
　 return false;
 } else if (event.button == 2 || event.button == 3) {
　 event.cancelBubble = true
　 event.returnvalue = false;
　 return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;　// for IE5+
document.onmousedown = norightclick;　　 // for all others
//-->
</SCRIPT>
</head>
<body bgcolor="#ffffff" scroll=no leftmargin="0" topmargin="0" onload="window_onload()">
<script language="javascript">
//关闭鼠标右键
if (document.all)
document.body.onmousedown=  new Function("if(event.button==2||event.button==3)alert('禁止鼠标右键！')");
</script>
<%
	url="http://200.0.0.90"'取得注册网址
	if instr(url,"http://")<=0 and instr(url,"HTTP://")<=0 then
		url="http://"&url
	end if
	mc="伴江行"'取得注册网站名称
	lx="政府机关"
	jj="伴江行办公自动化系统"
	id=1'取得注册网址的唯一id号
	longmc=mc&"：通信工具"
	shortmc=GetNewStr(longmc)
	randomize
	faceid=Int((25 - 0 ) * Rnd + 0)
	do while faceid=0
		faceid=Int((25 - 0 ) * Rnd + 0)
	loop
	session("username")=oabusyname
	session("siteid")=id
	if isempty(application("reftime"&session("siteid"))) then
		application("reftime"&session("siteid"))=now()
	end if
	if isempty(application("onlinesite")) then'如果在线队列不存在
		call create_online_site'建立在线队列
		sitenumber=write_online_site(id,mc,lx,url,jj)'写入当前站点信息
	else
		sitenumber=find_online_site(id)
		if sitenumber=-1 then'如果在线队列中不存在该队列
			sitenumber=write_online_site(id,mc,lx,url,jj)'写入当前站点信息
		end if
	end if
	if isempty(application("onlineuser"&id)) then'如果该站点的在线用户队列不存在，则建立队列
		call create_online_user(id)'建立在线用户队列
		call write_online_user(id,faceid)'写入当前用户信息
	elseif find_online_user(id)=0 then
		call write_online_user(id,faceid)'写入当前用户信息
	end if
%>	
<input id="refflag" name="refflag" type="hidden" value="1">
<input id="faceid" name="faceid" type="hidden" value="1">
<input id="sitenumber" name="sitenumber" type="hidden" value="-1">
<input id="sitename" name="sitename" type="hidden" >
<input id="siteidnumber" name="siteidnumber" type="hidden" >
<input id="siteurl" name="siteurl" type="hidden" >
<input id="username" name="username" type="hidden" >
<input id="usernoid" name="usernoid" type="hidden" >
<input id="usersessionid" name="usersessionid" type="hidden" >
<input id="removewin" name="removewin" type="hidden" value="0">
<input id="onlineuser" name="onlineuser" type="hidden" >
<input id="getinfo" name="getinfo" type="hidden" value="">
<input id="sendover" name="sendover" type="hidden" >
<input id="pageflag" name="pageflag" type="hidden" value="1">
<input id="listflag" name="listflag" type="hidden" value="0">
<input id="history" name="history" type="hidden" value="">
<input id="userdlflag" name="userdlflag" type="hidden" value="0">
<input id="userregisterflag" name="userregisterflag" type="hidden" value="0">
<input id="dispinfoflag" name="dispinfoflag" type="hidden" value="0">
<IFRAME name="hiddenwin" frameBorder=0 scrolling="no" height="0" src="" width="0" bgcolor="#8482C6"></IFRAME> 
<%
Response.Write("<script language=""javascript"">")
Response.Write("faceid.value="&chr(34)&faceid&chr(34)&";")
Response.Write("sitenumber.value="&chr(34)&sitenumber&chr(34)&";")
Response.Write("sitename.value="&chr(34)&mc&chr(34)&";")
Response.Write("siteidnumber.value="&chr(34)&id&chr(34)&";")
Response.Write("siteurl.value="&chr(34)&url&chr(34)&";")
Response.Write("username.value="&chr(34)&session("username")&chr(34)&";")
Response.Write("usersessionid.value="&chr(34)&session.SessionID&chr(34)&";")
response.write("history.value="&chr(34)&" "&chr(34)&";")
Response.Write("</script>")
%>
<script language="JavaScript" src="move.js"></script>
<div id="Layer1" style="HEIGHT: 25px; LEFT: 0px; POSITION: absolute; TOP: 0px; WIDTH: 112px; Z-INDEX: 1"  title="<%=longmc%>">
    <table border="0" cellpadding="0" cellspacing="0" width="100%" height="100%">
	  <tr class="handle" handlefor="Layer1"> 
        <td width="70%"></td>
        <td width="16%">
</td>
        <td width="14%">
          <p align="center"><A href="#" onclick=closewindow() title="关闭寻呼"><IMG border=0  src="images/kuaig_r2_c5.gif"></A></p></td>
      </tr>
    </table>
</div>
<table border="0" cellpadding="0" cellspacing="0" width="112">
  <tr>
   <td><img src="images/spacer.gif" width="3" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="52" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="2" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="39" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="13" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="2" height="1" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
   <td rowspan="3" colspan="4"><img name="kuaig_r1_c1" src="images/kuaig_r1_c1.gif" width="96" height="40" border="0"></td>
   <td colspan="3"><img name="kuaig_r1_c5" src="images/kuaig_r1_c5.gif" width="16" height="3" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="3" border="0"></td>
  </tr>
  <tr>
   <td colspan="3"><img name="kuaig_r2_c5" src="images/kuaig_r2_c5.gif" width="16" height="14" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr>
   <td colspan="3"><img name="kuaig_r3_c5" src="images/kuaig_r3_c5.gif" width="16" height="23" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="23" border="0"></td>
  </tr>
  <tr>
   <td rowspan="6"><img name="kuaig_r4_c1" src="images/kuaig_r4_c1.gif" width="3" height="260" border="0"></td>
    <td colspan="4" bgcolor="#0099FF"><IFRAME name="dispuserwin" frameBorder=0 scrolling="auto" marginwidth=1 marginheight=1 height="100%" noresize src="" width="100%" bgcolor="#0099FF"></IFRAME> 
    </td>
   <td rowspan="6" colspan="2"><img name="kuaig_r4_c6" src="images/kuaig_r4_c6.gif" width="3" height="260" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="224" border="0"></td>
  </tr>
  <tr>
   <td colspan="4"><img name="kuaig_r5_c2" src="images/kuaig_r5_c2.gif" width="106" height="3" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="3" border="0"></td>
  </tr>
  <tr>
   <td ONMOUSEOVER="this.bgColor='#c7c7c7';this.style.cursor='hand';" ONMOUSEOUT="this.bgColor='#e7e7e7';" ONCLICK="changeface()"><img name="kuaig_r6_c2" src="images/kuaig_r6_c2.gif" width="52" height="14" border="0"></td>
   <td rowspan="3"><img name="kuaig_r6_c3" src="images/kuaig_r6_c3.gif" width="2" height="30" border="0"></td>
   <td ONMOUSEOVER="this.bgColor='#c7c7c7';this.style.cursor='hand';" ONMOUSEOUT="this.bgColor='#e7e7e7';" ONCLICK="talk_face()" colspan="2"><img name="kuaig_r6_c4" src="images/kuaig_r6_c4.gif" width="52" height="14" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="14" border="0"></td>
  </tr>
  <tr>
   <td><img name="kuaig_r7_c2" src="images/kuaig_r7_c2.gif" width="52" height="3" border="0"></td>
   <td colspan="2"><img name="kuaig_r7_c4" src="images/kuaig_r7_c4.gif" width="52" height="3" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="3" border="0"></td>
  </tr>
  <tr>
   <td ONMOUSEOVER="this.bgColor='#c7c7c7';this.style.cursor='hand';" ONMOUSEOUT="this.bgColor='#e7e7e7';" ONCLICK="sendsms()"><img name="kuaig_r8_c2" src="images/kuaig_r8_c2.gif" width="52" height="13" border="0"></td>
   <td ONMOUSEOVER="this.bgColor='#c7c7c7';this.style.cursor='hand';" ONMOUSEOUT="this.bgColor='#e7e7e7';" ONCLICK="gohomepage()" colspan="2"><img name="kuaig_r8_c4" src="images/kuaig_r8_c4.gif" width="52" height="13" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="13" border="0"></td>
  </tr>
  <tr>
   <td colspan="4"><img name="kuaig_r9_c2" src="images/kuaig_r9_c2.gif" width="106" height="3" border="0"></td>
   <td><img src="images/spacer.gif" width="1" height="3" border="0"></td>
  </tr>
</table>
</body>
</html>
ppath("../db/lmtof.mdb")
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
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("alert(""您已经过期,请重新登录系统!"");")
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
'判断传来的表单域值
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
		call disperrinfo("请输入原因！")
		response.end
	end if
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select name,username,userdept from userinf where username='"&username&"'"
	rs.open sql,conn,1
	if rs.eof or rs.bof then
		call DispErrInfo("对不起，没有这个用户！")
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
				call disperrinfo("对不起，您不能重复考勤！")
			else
				rs("leavedate")=time()
				rs("explain2")=yystr
			end if
		else
			if cstr(rs("comedate"))<>"" and rs("comedate")<>#0:00:00# then
				kqconn.close
				set rs=nothing
				set kqconn=nothing
				call disperrinfo("对不起，您不能重复考勤！")
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
<title>手工考勤</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body>
<p align="center"><font size="3">补考勤</font></p>
<%
	if lookkqinfo=0 then
		response.write("<p align=""center"">对不起，现在不是考勤时间！</p>")
	else
%>
<div align="center">
  <center>
<form method="POST" action="handkq.asp" name="form1">
  <div align="center">
    <center>
   <table border="1" cellpadding="0" cellspacing="0" width="400" height="198" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
      <tr>
        <td height="25" bgcolor="#EFEFEF" width="396" colspan="2"><font color="#ee0000">注意：</font>请您写明补考勤原因。</td>
      </tr>
      <tr>
        <td height="25" width="396" colspan="2">用户名：<%=oabusyname%>
		<input type="hidden" name="username" value="<%=oabusyusername%>">
</td>
      </tr>
      <tr>
        <td height="25" bgcolor="#EFEFEF" width="396" colspan="2">选择考勤时间段:
<input type="radio" name="kqtimephase" value="amcome" checked>上午上班
<%
if lookkqinfo>=2 and amgonokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""amgo"">上午下班")
end if
if lookkqinfo>=3 and pmcomenokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""pmcome"">下午上班")
end if
if lookkqinfo>=4 and pmgonokq=0 then
	response.write("<input type=""radio"" name=""kqtimephase"" value=""pmgo"">下午下班")
end if
%>
	</td>
      </tr>
      <tr>
        <td height="90" width="49">原因：<br>
        </td>
        <td height="90" width="345"><textarea rows="8" name="yy" cols="46"></textarea></td>
      </tr>
    </table>
    </center>
  </div>
  <p align="center">
  <input type="submit" value="确定" name="okbutton">&nbsp;&nbsp;&nbsp;
  <input type="button" value="关闭" onclick="window.close();">
  </p>
</form>
  </center>
</div>
<%
end if
%>
</body>

</html>
