<!--#include file="up.asp"--><style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%
function hell(str)
dim re
Set re=new RegExp
re.IgnoreCase=true
re.Global=True
re.Pattern=""&chr(10)&""&chr(10)&""&chr(10)&"(\[right\])(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])(\[\/right\])"
str=re.Replace(str,"")	
str = replace(str, ">", "&gt;")
str = replace(str, "<", "&lt;")
set re=Nothing
hell=str
end function
t1="<div align=center><center><table border='0' cellpadding='0' cellspacing='0' style='border-collapse: collapse' width='94%'><tr><td width='180' background='pic/"&sp&"3.gif'>&nbsp;<img border='0' src='pic/fl.gif'> <font color='#FFFFFF'><b>"
t2="</b></font></td><td background='pic/"&sp&"5.gif'><img border='0' src='pic/"&sp&"4.gif'></td></tr></table></center></div><div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% >"
d1="<tr><td width=100% ><P style='MARGIN: 15px'>"
d2="</td></tr></table></center></div>"
id=request.querystring("id")
set edit=myconn.execute("select*from min where id="&id&"")
nnn=edit("name")
if lgname<>nnn then
abc="no"
end if
set edit=nothing
set che=myconn.execute("select name from user where name='"&lgname&"' and password='"&lgpwd&"'")
if che.eof then
abc="no"
end if
set che=nothing
set adok=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='70767766'")
if not adok.eof then
abc="yes"
end if
set adok=nothing
set adok=myconn.execute("select name from admin where name='"&lgname&"' and password='"&lgpwd&"' and bd='"&bd&"'")
if not adok.eof then
abc="yes"
end if
set adok=nothing
if abc="no" then
%><br><br>
<%=t1%>错 误 信 息<%=t2&d1%>·你不是该帖的作者或该版面的斑竹，不能编辑该帖子·<%=d2%>
<%
response.end
end if
set edit=nothing
%>
<%
id=request.querystring("id")
ed=request.querystring("ed")
reid=request.querystring("reid")
bd=request.querystring("bd")
set edit=myconn.execute("select*from min where id="&id&"")
%>
<%if ed=1 then
add=edit("zhuti")
else
add="re"
end if
zhuti=Replace(Request.Form("zhuti"),"'","''")
body=Replace(Request.Form("body"),"'","''")
body=""&body&chr(10)&chr(10)&chr(10)&"[right][color="&c1&"]「该帖子被 "&lgname&" 在 "&now&" 编辑过」[/color][/right]"
face=Replace(Request.Form("face"),"'","''")
if zhuti="" or body="" or face="" then
%><br>
<SCRIPT>
function emoticon(theSmilie){
document.kbbs.body.value +=theSmilie + '';
document.kbbs.body.focus();
}
</SCRIPT><SCRIPT src="ybbcode.js"></SCRIPT>
<SCRIPT>var i=0;
function presskey(eventobject){if(event.ctrlKey && window.event.keyCode==13){i++;if (i>1) {alert('帖子正在发出，请耐心等待！');return false;}this.document.kbbs.submit();}}
</SCRIPT><form method=POST name=kbbs>

<%=t1%>编 辑 帖 子<%=t2%>

<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="<%=c1%>" width="94%">
        <tr>
      <td width="24%" height="30">&nbsp;<b>帖子主题：</b></td>
      <td width="76%">&nbsp;<input type=text name=zhuti size=80 maxlength=59 value="<%=add%>">
  <font color="<%=c1%>">*</font></td>
    </tr>
    <tr>
      <td width="24%" height="50">
      <p style="line-height: 150%; margin-left: 5; margin-top: 5"><b>你的表情：</b> <br>·放在帖子前面。</td>
      <td width="76%"><input type=radio value=face1 name=face checked> 
  <img border=0 src=face/face1.gif width="16" height="16"> <input type=radio value=face2 name=face> 
  <img border=0 src=face/face2.gif width="16" height="16"> <input type=radio value=face3 name=face> 
  <img border=0 src=face/face3.gif width="16" height="16"> <input type=radio value=face4 name=face> 
  <img border=0 src=face/face4.gif width="16" height="16"> <input type=radio value=face5 name=face> 
  <img border=0 src=face/face5.gif width="16" height="16"> <input type=radio value=face6 name=face> 
  <img border=0 src=face/face6.gif width="16" height="16"> <input type=radio value=face7 name=face> 
  <img border=0 src=face/face7.gif width="16" height="16"> <input type=radio value=face8 name=face> 
  <img border=0 src=face/face8.gif width="16" height="16"> <input type=radio value=face9 name=face> 
  <img border=0 src=face/face9.gif width="16" height="16"><p style='margin-top: 2; margin-bottom: 2'>
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
<tr>
      <td width="24%" height="296" valign="top">
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
<br></td>      <td width="76%" valign="top">
      <p style="margin-left: 4; margin-top: 4">
        <p>
        <IFRAME name=ad 
            src="upload.asp" frameBorder=0 
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
        <IMG onclick=Cwmv() alt="Media Player视频文件" src="pic/mp.gif" border=0><br><br>&nbsp;文字大小：<select onchange=ybbsize(this.options[this.selectedIndex].value) name=a style="font-size: 9pt"><OPTION value=1>1</OPTION><OPTION value=2>2</OPTION><OPTION value=3>3</OPTION><OPTION value=4>4</OPTION></SELECT> <span lang=zh-cn>颜色：<select onchange=COLOR(this.options[this.selectedIndex].value) name="111" style="font-size: 9pt"><option style='COLOR:000000;BACKGROUND-COLOR:000000' value=000000>黑色</option><option style='COLOR:FFFFFF;BACKGROUND-COLOR:FFFFFF' value=FFFFFF>白色</option><option style='COLOR:008000;BACKGROUND-COLOR:008000' value=008000>绿色</option><option style='COLOR:800000;BACKGROUND-COLOR:800000' value=800000>褐色</option><option style='COLOR:808000;BACKGROUND-COLOR:808000' value=808000>橄榄色</option><option style='COLOR:000080;BACKGROUND-COLOR:000080' value=000080>深蓝色</option><option style='COLOR:800080;BACKGROUND-COLOR:800080' value=800080>紫色</option><option style='COLOR:808080;BACKGROUND-COLOR:808080' value=808080>灰色</option><option style='COLOR:FFFF00;BACKGROUND-COLOR:FFFF00' value=FFFF00>黄色</option><option style='COLOR:00FF00;BACKGROUND-COLOR:00FF00' value=00FF00>浅绿色</option><option style='COLOR:00FFFF;BACKGROUND-COLOR:00FFFF' value=00FFFF>浅蓝色</option><option style='COLOR:FF00FF;BACKGROUND-COLOR:FF00FF' value=FF00FF>粉红色</option><option style='COLOR:C0C0C0;BACKGROUND-COLOR:C0C0C0' value=C0C0C0>银白色</option><option style='COLOR:FF0000;BACKGROUND-COLOR:FF0000' value=FF0000>红色</option><option style='COLOR:0000FF;BACKGROUND-COLOR:0000FF' value=0000FF>蓝色</option><option style='COLOR:008080;BACKGROUND-COLOR:008080' value=008080>蓝绿色</option></select><br><br>
  &nbsp;<textarea onkeydown=presskey(); rows=10 name=body cols=79 Title='按 Ctrl+Enter 直接发送'><%=hell(edit("body"))%></textarea><table border="0" cellpadding="0" cellspacing="0" style="border-collapse: collapse" bordercolor="#111111" height="30">
          <tr>
            <td>
            &nbsp;点击图标插入表情：</td>
            <td>
            <A href=javascript:emoticon('{f1)')><IMG alt=笑脸 src=face/xl.gif border=0></A> <A href=javascript:emoticon('{f2)')><IMG alt=开口笑脸 src=face/kk.gif border=0></A> <A href=javascript:emoticon('{f3)')><IMG alt=惊讶的笑脸 src=face/jy.gif border=0></A> <A href=javascript:emoticon('{f4)')><IMG alt=吐舌笑脸 src=face/ts.gif border=0></A> <A href=javascript:emoticon('{f5)')><IMG alt=眨眼微笑 src=face/zy.gif border=0></A> <A href=javascript:emoticon('{f6)')><IMG alt=难过的脸 src=face/ng.gif border=0></A> <A href=javascript:emoticon('{f7)')><IMG alt=困惑的笑脸 src=face/kh.gif border=0></A> <A href=javascript:emoticon('{f8)')><IMG alt=失望的脸 src=face/sw.gif border=0></A> <a href=javascript:emoticon('{f9)')><IMG alt=尴尬的笑脸 src=face/gg.gif border=0></a>

            </td>
          </tr>
        </table>
<br>
&nbsp;<input type=submit value=OK_好了！修改帖子 name=B1>&nbsp; <input type=reset value=NO_不行！我要重写 name=B2> [按  Ctrl+Enter 直接发送]<br><br>
        
      </td>
    </tr>  </table>
  </center>
</div></form><%else
riqi=now+timeset/24
myconn.execute("update [min] set zhuti='"&zhuti&"',body='"&body&"',face='"&face&"',orders='"&riqi&"' where id="&id&"")
%><br><br><%=t1%>修 改 成 功<%=t2&d1%>·修改成功 <a href="show.asp?id=<%=reid%>&bd=<%=bd%>">
<font color="<%=c1%>">回到帖子页面·</font></a><%=d2%><%
end if%><br><!--#include file="down.asp"-->