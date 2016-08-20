<%response.expires=0%>
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/sqlstr.asp"-->
<HTML><HEAD><TITLE>伴江行Office办公系统</TITLE>
<META http-equiv=Content-Type content="text/html; charset=gb2312">
<LINK href="css/css.css" rel=stylesheet>
<SCRIPT language=JavaScript1.2>
<!--
window.parent("main").location.href="bbsnew/index.asp";
scores = new Array(20);
var numTotal=0;
NS4 = (document.layers) ? 1 : 0;
IE4 = (document.all) ? 1 : 0;
ver4 = (NS4 || IE4) ? 1 : 0;

if (ver4) {    with (document) {        write("<STYLE TYPE='text/css'>");        if (NS4) {            write(".parent {position:absolute; visibility:visible}");            write(".child {position:absolute; visibility:visible}");            write(".regular {position:absolute; visibility:visible}")        }        else {            write(".child {display:none}")        }        write("</STYLE>");    }}



function getIndex(el) {    ind = null;    for (i=0; i<document.layers.length; i++) {        whichEl = document.layers[i];        if (whichEl.id == el) {            ind = i;            break;        }    }    return ind;}


function arrange() {    nextY = document.layers[firstInd].pageY +document.layers[firstInd].document.height;    for (i=firstInd+1; i<document.layers.length; i++) {        whichEl = document.layers[i];        if (whichEl.visibility != "hide") {            whichEl.pageY = nextY;            nextY += whichEl.document.height;        }    }}function initIt(){    if (!ver4) return;    if (NS4) {        for (i=0; i<document.layers.length; i++) {            whichEl = document.layers[i];            if (whichEl.id.indexOf("Child") != -1) whichEl.visibility = "hide";       }        arrange();    }    else {        divColl = document.all.tags("DIV");        for (i=0; i<divColl.length; i++) {            whichEl = divColl(i);            if (whichEl.className == "child") whichEl.style.display = "none";        }    }}


function expandIt(el) {	if (!ver4) return;    if (IE4) {        whichEl1 = eval(el + "Child");		for(i=1;i<=numTotal;i++){			whichEl = eval(scores[i] + "Child");			if(whichEl!=whichEl1) {				whichEl.style.display = "none";			}		}        whichEl1 = eval(el + "Child");        if (whichEl1.style.display == "none") {            whichEl1.style.display = "block";        }        else {            whichEl1.style.display = "none";        }    }    else {        whichEl = eval("document." + el + "Child");		for(i=1;i<=numTotal;i++){			whichEl = eval("document." + scores[i] + "Child");			if(whichEl!=whichEl1) {				whichEl.visibility = "hide";			}		}        if (whichEl.visibility == "hide") {            whichEl.visibility = "show";        }        else {            whichEl.visibility = "hide";        }        arrange();    }}



onload = initIt;

function MM_swapImgRestore() { //v3.0
  var i,x,a=document.MM_sr; for(i=0;a&&i<a.length&&(x=a[i])&&x.oSrc;i++) x.src=x.oSrc;
}

function MM_preloadImages() { //v3.0
  var d=document; if(d.images){ if(!d.MM_p) d.MM_p=new Array();
    var i,j=d.MM_p.length,a=MM_preloadImages.arguments; for(i=0; i<a.length; i++)
    if (a[i].indexOf("#")!=0){ d.MM_p[j]=new Image; d.MM_p[j++].src=a[i];}}
}

function MM_findObj(n, d) { //v4.0
  var p,i,x;  if(!d) d=document; if((p=n.indexOf("?"))>0&&parent.frames.length) {
    d=parent.frames[n.substring(p+1)].document; n=n.substring(0,p);}
  if(!(x=d[n])&&d.all) x=d.all[n]; for (i=0;!x&&i<d.forms.length;i++) x=d.forms[i][n];
  for(i=0;!x&&d.layers&&i<d.layers.length;i++) x=MM_findObj(n,d.layers[i].document);
  if(!x && document.getElementById) x=document.getElementById(n); return x;
}

function MM_swapImage() { //v3.0
  var i,j=0,x,a=MM_swapImage.arguments; document.MM_sr=new Array; for(i=0;i<(a.length-2);i+=3)
   if ((x=MM_findObj(a[i]))!=null){document.MM_sr[j++]=x; if(!x.oSrc) x.oSrc=x.src; x.src=a[i+2];}
}

function openkg()
{	
	var openwinflag=0;
	openwinflag=window.open('','hrz8315','fullscreen=1,toolbar=no,scrollbars=no,resizable=0,menubar=no');
	openwinflag.resizeTo(112,301);
	openwinflag.moveTo(600,163);
	openwinflag.focus();
	openwinflag.location.href="qqprg/kuaig.asp";
}

function opennetmeeting()
{
	var openwinflag=0;
	openwinflag=window.open('','netmeetingwin','toolbar=no,scrollbars=no,resizable=0,menubar=no');
	openwinflag.resizeTo(465,455);
	openwinflag.moveTo(100,100);
	openwinflag.focus();
	openwinflag.location.href="netmeeting.asp";
}
//-->
</SCRIPT>
<link rel="stylesheet" href="9pp.css" type="text/css">
</HEAD>
<BODY background="images/head_r1_c3.jpg" leftMargin=0 topMargin=0 marginwidth="0" marginheight="0" onLoad="MM_preloadImages('image/menu1-1-b.gif','image/menu1-2-b.gif','image/menu2-1-b.gif','image/menu2-2-b.gif','image/menu2-3-b.gif','image/menu3-1-b.gif','image/menu3-2-b.gif','image/menu3-3-b.gif','image/menu3-4-b.gif','image/menu5-1-b.gif','image/menu5-2-b.gif','image/menu5-3-b.gif','image/menu6-1-b.gif','image/menu6-2-b.gif')" link="#336699" vlink="#336699">
<table width="26" border="0" cellspacing="0" cellpadding="0">
  <tr> 
    
    <td style="height: 100%" id="outLookBarShow" name="outLookBarShow" valign="top" align="center" width="133"> 
      <table border="0" cellspacing="0" cellpadding="0" style="height:100%;width:100%;border-bottom:0pt solid #ebf5d6;" valign="middle" align="center" width="90%">
        <tr valign="top"> 
          <td style="position:relative"> 
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
'打开数据库，读出权限
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if
on error resume next
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
rs.open "select count(*) as countss from userinf",conn,1,1
usercount=rs("countss")
if usercount >500 then
   rs.close
   set rs=nothing
   %>
   <script language=javascript>
       window.alert("对不起，超过了最大使用用户数");
       parent("main").location.href="/usercontrol.asp";
   </script>
<%
end if
rs.close
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1

cook_allow_manage_workthings=rs("allow_manage_workthings")
cook_allow_auditing_workthings=rs("allow_auditing_workthings")
%>
          </td>
        </tr>
        <tr> 

          </td>
        </tr>
        <tr name="outlookdiv5" id="outlookdiv5" style="width:100%;display:none;height:0%;"> 
          <td valign="top" align="left" height="120" bgcolor="#FFFFFF"> 
              <div name="outlookdivin5" id="outlookdivin5" style="overflow:auto;width:100%;height:100%"> 
        </tr>
        <tr> 
          <td><a href="jishuang.asp" target="main"><img src="image/new4.gif" height="20" border="0"></a></td>
        </tr>		
        <tr> 
          <td><a href="bbs/list.asp?boardid=6" target="main"><img src="image/new9.gif" height="20" border="0"></a></td>
        </tr>		
        <tr> 
          <td><a href="web/index.asp?typeid=25" target="main"><img src="image/new7.gif" height="20" border="0"></a></td>
        </tr>		
        <tr> 
          <td><a href="url/index.html" target="main"><img src="image/new10.gif" height="20" border="0"></a></td>
        </tr>		

        <tr> 
          <td><a href="shouji/index.asp" target="main"><img src="image/new22.gif" height="20" border="0"></a></td>
        </tr>		
        <tr> 
          <td><a href="yzqh/default.asp" target="main"><img src="image/yzqh.gif" height="20" border="0"></a></td>
        </tr>		
        <tr> 
          <td><a href="rl/cal.htm" target="main"><img src="image/new11.gif" height="20" border="0"></a></td>
        </tr>		

        <tr> 
          <td><a href="qyml/default.asp" target="main"><img src="image/qyml.gif" height="20" border="0"></a></td>
        </tr>		

		<div name="blankdiv" id="blankdiv" style="overflow:auto;width:100%;height:100%"> 
	    </div>
      </table>
    </td>
  </tr>
  </table>
<script language="javascript">

</script>
</BODY>
</HTML>
