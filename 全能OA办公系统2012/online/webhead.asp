<%
sub webhead()
	oabusyname=request.cookies("oabusyname")
	oabusyusername=request.cookies("oabusyusername")
	oabusyuserdept=request.cookies("oabusyuserdept")
	oabusyuserlevel=request.cookies("oabusyuserlevel")
%>
<body leftmargin="0" topmargin="0" marginwidth="0" marginheight="0">
<table border="0" cellpadding="0" cellspacing="0" width="779">
  <tr>
   <td><img src="../images/spacer.gif" width="241" height="1" border="0"></td>
   <td><img src="../images/spacer.gif" width="463" height="1" border="0"></td>
   <td><img src="../images/spacer.gif" width="75" height="1" border="0"></td>
   <td><img src="../images/spacer.gif" width="1" height="1" border="0"></td>
  </tr>

  <tr>
    <td rowspan="4" height="56"><object classid="clsid:D27CDB6E-AE6D-11cf-96B8-444553540000" codebase="http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=5,0,0,0" width="241" height="56" vspace="0" hspace="0">
        <param name="_cx" value="6376">
        <param name="_cy" value="1482">
        <param name="Movie" value="../images/logo.swf">
        <param name="Src" value="../images/logo.swf">
        <param name="WMode" value="Window">
        <param name="Play" value="-1">
        <param name="Loop" value="-1">
        <param name="Quality" value="High">
        <param name="SAlign" value>
        <param name="Menu" value="-1">
        <param name="Base" value>
        <param name="Scale" value="ExactFit">
        <param name="DeviceFont" value="0">
        <param name="EmbedMovie" value="0">
        <param name="BGColor" value>
        <param name="SWRemote" value>
        <param name="Stacking" value="below"><embed src="../images/logo.swf" quality="high" pluginspage="http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash" type="application/x-shockwave-flash" width="241" height="59" scale="exactfit" vspace="0" hspace="0"> 
      </object></td>
   <td><img name="top_r1_c2" src="../images/top_r1_c2.gif" width="463" height="10" border="0"></td>
   <td rowspan="2"><img name="top_r1_c3" src="../images/top_r1_c3.gif" width="75" height="26" border="0"></td>
   <td><img src="../images/spacer.gif" width="1" height="10" border="0"></td>
  </tr>
  <tr>
   <td rowspan="2" background="../images/top_r2_c2.gif">
    <table border="0" width="100%">
      <tr>
        <td width="22%">部门：<%=oabusyuserdept%></td>
        <td width="22%">姓名：<%=oabusyname%></td>
        <td width="22%">职位：<%=oabusyuserlevel%></td>
        <td width="34%">
<script language="JavaScript">
<!---
//取得日期和星期
   today=new Date();
   function initArray(){
	 this.length=initArray.arguments.length
	 for(var i=0;i<this.length;i++)
	 this[i+1]=initArray.arguments[i]  }
	 
  var d=new initArray("<font color=RED>星期日","<font color=black>星期一","<font color=black>星期二","<font color=black>星期三","<font color=black>星期四","<font color=black>星期五","<font color=GREEN>星期六"); 
document.write(today.getYear(),"年",today.getMonth()+1,"月",today.getDate(),"日",d[today.getDay()+1]);  
//-->
</script>
</td>
      </tr>
    </table>
   </td>
   <td><img src="../images/spacer.gif" width="1" height="16" border="0"></td>
  </tr>
  <tr>
   <td><img name="top_r3_c3" src="../images/top_r3_c3.gif" width="75" height="20" border="0"></td>
   <td><img src="../images/spacer.gif" width="1" height="20" border="0"></td>
  </tr>
  <tr>
    <td rowspan="2" background="../images/top_r4_c2.gif"><img border="0" src="../images/tetle.gif" usemap="#Map"></td>
   <td rowspan="2" background="../images/top_r4_c3.gif" valign="middle">
   <center>
<img src="../images/time/space.gif" name="one"><img src="../images/time/space.gif" name="two"><img src="../images/time/dgon.gif" name="three"><img src="../images/time/space.gif" name="four"><img src="../images/time/space.gif" name="five"></center>
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
</script>

</td>
   <td><img src="../images/spacer.gif" width="1" height="10" border="0"></td>
  </tr>
  <tr>
   <td width="241" height="15" background="../images/top_r5_c1.gif">
   <IFRAME name="msgwin" frameBorder=0 scrolling="no" height="15" src="top1.asp" width="241"  height="15"></IFRAME>
   </td>
   <td><img src="../images/spacer.gif" width="1" height="16" border="0"></td>
  </tr>
</table>
  <%
end sub
%>
<script language="javascript">
//单击"注销"连接时，弹出对话框是否要求退出系统
function closesystem()
{
	window.open('logout.asp?closeflag=0','closesystem','location=no,height=10, width=10, top=600, left=10,toolbar=no, menubar=no, scrollbars=no, resizable=no, location=no, status=no');
	window.location.href="default.asp";
}
</script>

<map name="Map">
<area href="../left.asp" target="contents" shape="rect" coords="0, 0, 86, 18">
<area href="../doc_storeroom/left.asp" target="contents" shape="rect" coords="92, 0, 184, 18">
<area href="public_serve_left.asp" target="contents" shape="rect" coords="188, 0, 268, 18">
<area href="../setting.asp" target="contents" shape="rect" coords="274, 2, 326, 18">
<area href="../helpleft.asp" target="contents" shape="rect" coords="334, 2, 386, 16">
<area href="#" onclick="closesystem();" shape="rect" coords="394, 0, 444, 18">
</map>
</body>
</html>
