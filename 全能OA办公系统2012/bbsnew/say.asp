<!--#include file="conn.asp"--><%
name=Request.Cookies(cn)("lgname")
pwd=Request.Cookies(cn)("lgpwd")
if name="" then
Response.Redirect"login.asp"
response.end
end if
application(name)=1
myconn.close
set myconn=nothing%>
<!--#include file="up.asp"-->
<%
function ubbs(str)
dim re
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
re.Pattern="(\[showtoname=(.[^\[]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")	
re.Pattern="(\[showtoreply\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")	
re.Pattern="(\[showtograde=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")
re.Pattern="(\[smoney=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")
re.Pattern="(\[smeili=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")
re.Pattern="(\[sjingyan=*([0-9]*)\])(.[^\[]*)(\[\/s\])"
str=re.Replace(str,"[enter][color="&c1&"]…… 加密内容不能引用 ……[/color][enter]")
	
str = replace(str, ">", "&gt;")
str = replace(str, "<", "&lt;")
	
set re=Nothing
ubbs=str
end function
%>
<SCRIPT>
function emoticon(theSmilie){
document.kbbs.body.value +=theSmilie + '';
document.kbbs.body.focus();
}
</SCRIPT><SCRIPT src="ybbcode.js"></SCRIPT>
<SCRIPT>var i=0;
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('帖子正在发出，请耐心等待！');return false;}this.document.kbbs.submit();}}
</SCRIPT><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%bdlogin(2)
%>
<%
session("upsize")=upsize
session("upnum")=upnum
pagenum=request.querystring("pagenum")
quoteid=request.querystring("quoteid")
re=request.querystring("re")
if re="no" then
action="save.asp?bd="&bd&"&re=no&pagenum="&pagenum&""
zhuti=""%>
<SCRIPT>function showvote(){
if (document.kbbs.voteyn.checked == true) {
vote.style.display = "";
}else{
vote.style.display = "none";
}
}
</SCRIPT>
<%
elseif re="yes" then
action="save.asp?bd="&bd&"&id="&id&"&re=yes&pagenum="&pagenum&""
zhuti="回复ID:"&id&""
set lock=myconn.execute("select type from min where id="&id&"")
if lock("type")<>4 or canre="yes" then
else
response.end
end if
set lock=nothing
end if
%>
<form method=POST name=kbbs action=<%=action%>><div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/<%=sp%>3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b> 发 表 帖 子</b></font></td><td background='pic/<%=sp%>5.gif'><img border='0' src='pic/<%=sp%>4.gif'></td></tr></table></center></div>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
    <tr>
  <td width="24%" height="30">	<b>&nbsp;用户名：</b></td>
  <td width="76%" height="30"> &nbsp;<input type=text name=name size=20 value="<%=name%>">
  <font color="<%=c1%>">*</font> <a href="zhuce.asp">没有注册？</a></td>    </tr>
    <tr>
      <td width="24%" height="30">&nbsp;<b>密&nbsp; 码：</b></td>
      <td width="76%">&nbsp;<input type=password name=password size=20 value="<%=pwd%>">
      <font color="<%=c1%>">*</font> <a href="getpwd.asp">忘记密码？</a></td>
    </tr>
    <tr>
      <td width="24%" height="30">&nbsp;<b>帖子主题：</b></td>
      <td width="76%">&nbsp;<input type=text name=zhuti size=80 maxlength=59 value="<%=zhuti%>">
  <font color="<%=c1%>">*</font></td>
    </tr>
    <tr>
      <td width="24%">
      <p style="line-height: 150%; margin-left: 5; margin-top: 5"><b>你的表情：</b> <br>·放在帖子前面。</td>
      <td width="76%">
  <p style="margin: 4"><input type=radio value=face1 name=face checked> 
  <img border=0 src=face/face1.gif width="16" height="16"> <input type=radio value=face2 name=face> 
  <img border=0 src=face/face2.gif width="16" height="16"> <input type=radio value=face3 name=face> 
  <img border=0 src=face/face3.gif width="16" height="16"> <input type=radio value=face4 name=face> 
  <img border=0 src=face/face4.gif width="16" height="16"> <input type=radio value=face5 name=face> 
  <img border=0 src=face/face5.gif width="16" height="16"> <input type=radio value=face6 name=face> 
  <img border=0 src=face/face6.gif width="16" height="16"> <input type=radio value=face7 name=face> 
  <img border=0 src=face/face7.gif width="16" height="16"> <input type=radio value=face8 name=face> 
  <img border=0 src=face/face8.gif width="16" height="16"> <input type=radio value=face9 name=face> 
  <img border=0 src=face/face9.gif width="16" height="16"><br>
    <input type=radio value=face10 name=face> 
  <img border=0 src=face/face10.gif width="16" height="16"> <input type=radio value=face11 name=face> 
  <img border=0 src=face/face11.gif width="16" height="16"> <input type=radio value=face12 name=face> 
  <img border=0 src=face/face12.gif width="16" height="16"> <input type=radio value=face13 name=face> 
  <img border=0 src=face/face13.gif width="16" height="16"> <input type=radio value=face14 name=face> 
  <img border=0 src=face/face14.gif width="16" height="16"> <input type=radio value=face15 name=face> 
  <img border=0 src=face/face15.gif width="16" height="16"> <input type=radio value=face16 name=face> 
  <img border=0 src=face/face16.gif width="16" height="16"> <input type=radio value=face17 name=face> 
  <img border=0 src=face/face17.gif width="16" height="16"> <input type=radio value=face18 name=face> 
  <img border=0 src=face/face18.gif width="16" height="16"></td>
    </tr>
    <%if re="no" then%><tr>
      <td width="24%" height="28">
      <p style="margin: 5"><b>帖子类型：</b></td>
      <td width="76%">
<input type="radio" name="gonggao" value="0" checked>普通帖子<a title="（将在公告栏中显示主题，使用该类型帖子需要扣除：金钱1000，魅力200，经验200）"> <input type="radio" name="gonggao" value="4"></a>公告帖子 <%if admin="yes" then%><input type="radio" name="gonggao" value="1">系统公告<%end if%></td>
    </tr><%end if%>
<tr>
      <td width="24%" valign="top">
<p style="line-height: 150%; margin-left: 5; margin-top: 5">
<b>帖子内容：</b><br>
      ·HTML标签： 不可用<br>
·UBB标签： 可用<br>
·贴图标签： 可用<br>
·Flash标签：不可用<br>
·表情字符转换：不可用<br>
·上传图片：可用<br>
·最多15KB<br><b>特殊内容：</b><br>
·<a href="javascript:grade()" title='格式：[showtograde=1]内容[/s]			表示只有等级在“1”以上才能浏览“内容”			“1”和“内容”可以自定，				使用该项请点击！'>等级可见内容</a><br>·<a href="javascript:reply()" title='格式：[showtoreply]内容[/s]			表示只有回复者才能浏览“内容”，		“内容”可以自定，				使用该项请点击！'>回复可见内容</a><br>·<a href="javascript:name()" title='格式：[showtoname=zym]内容[/s]			表示只有“zym”才能浏览“内容”			“zym”和“内容”可以自定，				使用该项请点击！'>指定读者内容</a><br>·<a href="javascript:smoney()" title='格式：[smoney=1000]内容[/s]			表示只有金钱不小于“1000”才能浏览“内容”“1000”和“内容”可以自定，使用该项请点击！'>金钱</a>·<a href="javascript:smeili()" title='格式：[smeili=1]内容[/s]			表示只有魅力在“1000”以上才能浏览“内容”		“1000”和“内容”可以自定，使用该项请点击！'>魅力</a>·<a href="javascript:sjingyan()" title='格式：[sjingyan=1000]内容[/s]			表示只有经验不小于“1000”才能浏览“内容”		“1000”和“内容”可以自定，使用该项请点击！'>经验</a>·
<p style="line-height: 150%; margin-left: 5; margin-top: 5">
<%if re="no" then%><input type="checkbox" onclick=showvote() name="voteyn" value="1"> 显示投票选项<br><%end if%><br>
</td>      <td width="76%" valign="top">
      <p style="margin-left: 4; margin-top: 4">
        <p>
<IFRAME name=ad src="upload.asp" frameBorder=0 
            width="100%" scrolling=no height=25></IFRAME><br>&nbsp;功能按钮：<IMG onclick=fly() alt=飞行字 src="pic/fly.gif" border=0> 
        <IMG onclick=move() alt=移动字 
      src="pic/move.gif" border=0> 
        <IMG 
      onclick=light() alt=发光字 src="pic/glow.gif" border=0> 
        <IMG onclick=ying() alt=阴影字 
      src="pic/shadow.gif" border=0> 
        <IMG 
      onclick=image() alt=图片 src="pic/image.gif" border=0> 
        <IMG onclick=Cswf() alt=Flash图片 
      src="pic/swf.gif" border=0> 
        <IMG onclick=Cdir() alt=Shockwave文件 src="pic/Shockwave.gif" border=0> 
        <IMG onclick=Crm() alt=realplay视频文件 
      src="pic/rm.gif" border=0> 
        <IMG onclick=Cwmv() alt="Media Player视频文件" src="pic/mp.gif" border=0>
        <img onclick=center() alt="居中" border="0" src="pic/center.gif">
        <img onclick=javascript:emoticon('[url=超连接]连接文本[/url]') alt="超连接" border="0" src="pic/url1.gif"><br><br>&nbsp;文字大小：<select onchange=ybbsize(this.options[this.selectedIndex].value) name=a style="font-size: 9pt"><OPTION value=1>1</OPTION><OPTION value=2>2</OPTION><OPTION value=3>3</OPTION><OPTION value=4>4</OPTION></SELECT> <span lang=zh-cn>颜色：<select onchange=COLOR(this.options[this.selectedIndex].value) name="111" style="font-size: 9pt"><option style='COLOR:000000;BACKGROUND-COLOR:000000' value=000000>黑色</option><option style='COLOR:FFFFFF;BACKGROUND-COLOR:FFFFFF' value=FFFFFF>白色</option><option style='COLOR:008000;BACKGROUND-COLOR:008000' value=008000>绿色</option><option style='COLOR:800000;BACKGROUND-COLOR:800000' value=800000>褐色</option><option style='COLOR:808000;BACKGROUND-COLOR:808000' value=808000>橄榄色</option><option style='COLOR:000080;BACKGROUND-COLOR:000080' value=000080>深蓝色</option><option style='COLOR:800080;BACKGROUND-COLOR:800080' value=800080>紫色</option><option style='COLOR:808080;BACKGROUND-COLOR:808080' value=808080>灰色</option><option style='COLOR:FFFF00;BACKGROUND-COLOR:FFFF00' value=FFFF00>黄色</option><option style='COLOR:00FF00;BACKGROUND-COLOR:00FF00' value=00FF00>浅绿色</option><option style='COLOR:00FFFF;BACKGROUND-COLOR:00FFFF' value=00FFFF>浅蓝色</option><option style='COLOR:FF00FF;BACKGROUND-COLOR:FF00FF' value=FF00FF>粉红色</option><option style='COLOR:C0C0C0;BACKGROUND-COLOR:C0C0C0' value=C0C0C0>银白色</option><option style='COLOR:FF0000;BACKGROUND-COLOR:FF0000' value=FF0000>红色</option><option style='COLOR:0000FF;BACKGROUND-COLOR:0000FF' value=0000FF>蓝色</option><option style='COLOR:008080;BACKGROUND-COLOR:008080' value=008080>蓝绿色</option></select><br><br>
  &nbsp;
<%if quoteid<>"" then
set quote=myconn.execute("select top 1 name,riqi,body,bd,gonggao from min where id="&quoteid&"")
qqbody=quote("body")
%>
<textarea onkeydown=presskey(); rows=8 name=body cols=79 Title='按 Ctrl+Enter 直接发送'>
[quote][color=<%=c1%>]该内容引用 <%=quote("name")%> 在 <%=quote("riqi")%> 发表的内容：[/color]
<%=ubbs(qqbody)%>
[/quote]
</textarea>
<%
set quote=nothing
else%>
<textarea onkeydown=presskey(); rows=8 name=body cols=79 Title='按 Ctrl+Enter 直接发送'></textarea><%end if%><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="30">
          <tr>
            <td>
            &nbsp;点击图标插入表情：</td>
            <td>
            <A href=javascript:emoticon('{f1)')><IMG alt=笑脸 src=face/xl.gif border=0></A> <A href=javascript:emoticon('{f2)')><IMG alt=开口笑脸 src=face/kk.gif border=0></A> <A href=javascript:emoticon('{f3)')><IMG alt=惊讶的笑脸 src=face/jy.gif border=0></A> <A href=javascript:emoticon('{f4)')><IMG alt=吐舌笑脸 src=face/ts.gif border=0></A> <A href=javascript:emoticon('{f5)')><IMG alt=眨眼微笑 src=face/zy.gif border=0></A> <A href=javascript:emoticon('{f6)')><IMG alt=难过的脸 src=face/ng.gif border=0></A> <A href=javascript:emoticon('{f7)')><IMG alt=困惑的笑脸 src=face/kh.gif border=0></A> <A href=javascript:emoticon('{f8)')><IMG alt=失望的脸 src=face/sw.gif border=0></A> <a href=javascript:emoticon('{f9)')><IMG alt=尴尬的笑脸 src=face/gg.gif border=0></a>

            </td>
          </tr>
        </table>
</td>
    </tr>  
    <tr id=vote style="DISPLAY: none">
      <td valign="top">
<p style="margin: 5; ; line-height:150%"><b>投票项目</b><br>·各个项目用回车隔开<br>
<input type="radio" name="votetype" value="1" checked>单选&nbsp;
<input type="radio" name="votetype" value="2">多选</p>
<p style="margin: 5; ; line-height:150%"><b>过期时间：</b><select size="1" name="outtime" style="font-size: 9pt">
<option value="1">一天</option>
<option value="3">三天</option>
<option value="7">一周</option>
<option value="15">半个月</option>
<option value="31">一个月</option>
<option value="93">三个月</option>
<option value="365">一年</option>
<option value="10000" selected>不过期</option>
</select><br></p>
</td>      <td valign="top">
      <p style="margin: 5"> &nbsp;<textarea onkeydown=presskey(); rows=6 name=vote cols=79 Title='按 Ctrl+Enter 直接发送'></textarea></td>
    </tr>
    <tr>
      <td width="100%" height="40" colspan="2" align="center" bgcolor="<%=c2%>">
      &nbsp;<input type=submit value=OK_好了！发表帖子 name=B1> <input type=reset value=NO_不行！我要重写 name=B2> [按  Ctrl+Enter 直接发送]</td>      
    </tr>
  </table>
  </center>
</div></form><!--#include file="down.asp"-->