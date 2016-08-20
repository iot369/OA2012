<%
tt1="<font style='font-size:9pt;line-height:15pt'><hr color=#CDCFCE size=1><font color="&c1&">"
tt2="<hr color=#CDCFCE size=1>"
function ubb(str)
str = kb6k(str)
dim re
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
	re.Pattern="(javascript)"
str=re.Replace(str,"&#106avascript")
re.Pattern="(jscript:)"
str=re.Replace(str,"&#106script:")
re.Pattern="(js:)"
str=re.Replace(str,"&#106s:")
re.Pattern="(value)"
str=re.Replace(str,"&#118alue")
re.Pattern="(about:)"
str=re.Replace(str,"about&#58")
re.Pattern="(file:)"
str=re.Replace(str,"file&#58")
re.Pattern="(document.cookie)"
str=re.Replace(str,"documents&#46cookie")
re.Pattern="(vbscript:)"
str=re.Replace(str,"&#118bscript:")
re.Pattern="(vbs:)"
str=re.Replace(str,"&#118bs:")
re.Pattern="(on(mouse|exit|error|click|key))"
str=re.Replace(str,"&#111n$2")
re.Pattern="(script)"
str=re.Replace(str,"&#115cript")

	re.Pattern="\[IMG\](http|https|ftp):\/\/(.[^\[]*)\[\/IMG\]"
	str=re.Replace(str,"<a onfocus=this.blur() href=""$1://$2"" target=_blank><IMG SRC=""$1://$2"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></a>")
	re.Pattern="\[UPLOAD=(gif|jpg|jpeg|bmp|png)\](.[^\[]*)(gif|jpg|jpeg|bmp|png)\[\/UPLOAD\]"
	str= re.Replace(str,"<br><IMG SRC=""pic/$1.gif"" border=0> 此主题相关图片如下：<br><A HREF=""upload$2$1"" TARGET=_blank><IMG SRC=""upload$2$1"" border=0 alt=按此在新窗口浏览图片 onload=""javascript:if(this.width>screen.width-333)this.width=screen.width-333""></A>")
	re.Pattern="(\[UPLOAD=(.[^\[]*)\])(.[^\[]*)(\[\/UPLOAD\])"
	str= re.Replace(str,"<IMG SRC=""pic/$2.gif"" border=0> <a href=upload$3>点击浏览该文件</a>")
	re.Pattern="(\[FLASH\])(http://.[^\[]*(.swf))(\[\/FLASH\])"
	str= re.Replace(str,"<a href=""$2"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=500 height=400><PARAM NAME=movie VALUE=""$2""><PARAM NAME=quality VALUE=high><embed src=""$2"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=500 height=400>$2</embed></OBJECT>")
	re.Pattern="(\[FLASH=*([0-9]*),*([0-9]*)\])(http://.[^\[]*(.swf))(\[\/FLASH\])"
	str= re.Replace(str,"<a href=""$4"" TARGET=_blank><IMG SRC=pic/swf.gif border=0 alt=点击开新窗口欣赏该FLASH动画! height=16 width=16>[全屏欣赏]</a><br><OBJECT codeBase=http://download.macromedia.com/pub/shockwave/cabs/flash/swflash.cab#version=4,0,2,0 classid=clsid:D27CDB6E-AE6D-11cf-96B8-444553540000 width=$2 height=$3><PARAM NAME=movie VALUE=""$4""><PARAM NAME=quality VALUE=high><embed src=""$4"" quality=high pluginspage='http://www.macromedia.com/shockwave/download/index.cgi?P1_Prod_Version=ShockwaveFlash' type='application/x-shockwave-flash' width=$2 height=$3>$4</embed></OBJECT>")
	re.Pattern="\[DIR=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/DIR]"
	str=re.Replace(str,"<object classid=clsid:166B1BCA-3F9C-11CF-8075-444553540000 codebase=http://download.macromedia.com/pub/shockwave/cabs/director/sw.cab#version=7,0,2,0 width=$1 height=$2><param name=src value=$3><embed src=$3 pluginspage=http://www.macromedia.com/shockwave/download/ width=$1 height=$2></embed></object>")
	re.Pattern="\[MP=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/MP]"
	str=re.Replace(str,"<object align=middle classid=CLSID:22d6f312-b0f6-11d0-94ab-0080c74c7e95 class=OBJECT id=MediaPlayer width=$1 height=$2 ><param name=ShowStatusBar value=-1><param name=Filename value=$3><embed type=application/x-oleobject codebase=http://activex.microsoft.com/activex/controls/mplayer/en/nsmp2inf.cab#Version=5,1,52,701 flename=mp src=$3  width=$1 height=$2></embed></object>")
	re.Pattern="\[RM=*([0-9]*),*([0-9]*)\](.[^\[]*)\[\/RM]"
	str=re.Replace(str,"<OBJECT classid=clsid:CFCDAA03-8BE4-11cf-B84B-0020AFBBCCFA class=OBJECT id=RAOCX width=$1 height=$2><PARAM NAME=SRC VALUE=$3><PARAM NAME=CONSOLE VALUE=Clip1><PARAM NAME=CONTROLS VALUE=imagewindow><PARAM NAME=AUTOSTART VALUE=true></OBJECT><br><OBJECT classid=CLSID:CFCDAA03-8BE4-11CF-B84B-0020AFBBCCFA height=32 id=video2 width=$1><PARAM NAME=SRC VALUE=$3><PARAM NAME=AUTOSTART VALUE=-1><PARAM NAME=CONTROLS VALUE=controlpanel><PARAM NAME=CONSOLE VALUE=Clip1></OBJECT>")
	
	
re.Pattern="\[fly\](.*)\[\/fly\]"
str=re.Replace(str,"<marquee width=90% behavior=alternate scrollamount=3>$1</marquee>")
re.Pattern="\[move\](.*)\[\/move\]"
str=re.Replace(str,"<MARQUEE scrollamount=3>$1</marquee>")	
re.Pattern="\[SHADOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/SHADOW]"
str=re.Replace(str,"<table width=$1 ><tr><td style=""filter:shadow(color=$2, strength=$3)"">$4</td></tr></table>")
re.Pattern="\[GLOW=*([0-9]*),*(#*[a-z0-9]*),*([0-9]*)\](.[^\[]*)\[\/GLOW]"
str=re.Replace(str,"<table width=$1 ><tr><td style=""filter:glow(color=$2, strength=$3)"">$4</td></tr></table>")
	
	
	
	re.Pattern = "^((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	str = re.Replace(str,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
	re.Pattern = "((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)$"
	str = re.Replace(str,"<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$1>$1</a>")
	re.Pattern = "([^>=""])((http|https|ftp|rtsp|mms):(\/\/|\\\\)[A-Za-z0-9\./=\?%\-&_~`@':+!]+)"
	str = re.Replace(str,"$1<img align=absmiddle src=pic/url.gif border=0><a target=_blank href=$2>$2</a>")
	re.Pattern="(\[size=1\])(.[^\[]*)(\[\/size\])"
	str=re.Replace(str,"<font size=1 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=2\])(.[^\[]*)(\[\/size\])"
	str=re.Replace(str,"<font size=2 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=3\])(.[^\[]*)(\[\/size\])"
	str=re.Replace(str,"<font size=3 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[size=4\])(.[^\[]*)(\[\/size\])"
	str=re.Replace(str,"<font size=4 style=""line-height:"&FontHeight&"pt"">$2</font>")
	re.Pattern="(\[color=(.[^\[]*)\])(.[^\[]*)(\[\/color\])"
	str=re.Replace(str,"<font color=$2 style=""font-size:"&FontSize&"pt;line-height:"&FontHeight&"pt"">$3</font>")

re.pattern="(\{f1)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/xl.gif align=absmiddle>")
re.pattern="(\{f2)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/kk.gif align=absmiddle>")
re.pattern="(\{f3)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/jy.gif align=absmiddle>")
re.pattern="(\{f4)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/ts.gif align=absmiddle>")
re.pattern="(\{f5)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/zy.gif align=absmiddle>")
re.pattern="(\{f6)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/ng.gif align=absmiddle>")
re.pattern="(\{f7)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/kh.gif align=absmiddle>")
re.pattern="(\{f8)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/sw.gif align=absmiddle>")
re.pattern="(\{f9)\)"
	str=re.replace(str,"<IMG border=0 SRC=face/gg.gif align=absmiddle>")
	re.Pattern="(\[right\])(.*)(\[\/right\])"
	str=re.Replace(str,"<div align=right>$2</div>")

re.Pattern="(^.*)(\[smoney=*([0-9]*)\])(.[^\[]*)(\[\/s\])(.*)"
po=re.Replace(str,"$3")
if IsNumeric(po) then ii=int(po) else ii=0
if lgname="" then
qq11=0
else
set ain=myconn.execute("select qian from user where name='"&lgname&"'")
qq11=ain("qian")
set ain=nothing
end if
if lgname=myname or qq11>=ii then
str=re.Replace(str,"$1"&tt1&"此内容需要拥有金钱 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>$4"&tt2&"$6")
else
str=re.Replace(str,"$1"&tt1&"此内容需要拥有金钱 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>"&tt2&"$6")
end if

re.Pattern="(^.*)(\[smeili=*([0-9]*)\])(.[^\[]*)(\[\/s\])(.*)"
po=re.Replace(str,"$3")
if IsNumeric(po) then ii=int(po) else ii=0
if lgname="" then
qq11=0
else
set ain=myconn.execute("select meili from user where name='"&lgname&"'")
qq11=ain("meili")
set ain=nothing
end if
if lgname=myname or qq11>=ii then
str=re.Replace(str,"$1"&tt1&"此内容需要拥有魅力值 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>$4"&tt2&"$6")
else
str=re.Replace(str,"$1"&tt1&"此内容需要拥有魅力值 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>"&tt2&"$6")
end if
re.Pattern="(^.*)(\[sjingyan=*([0-9]*)\])(.[^\[]*)(\[\/s\])(.*)"
po=re.Replace(str,"$3")
if IsNumeric(po) then ii=int(po) else ii=0
if lgname="" then
qq11=0
else
set ain=myconn.execute("select jingyan from user where name='"&lgname&"'")
qq11=ain("jingyan")
set ain=nothing
end if
if lgname=myname or qq11>=ii then
str=re.Replace(str,"$1"&tt1&"此内容需要拥有经验值 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>$4"&tt2&"$6")
else
str=re.Replace(str,"$1"&tt1&"此内容需要拥有经验值 <b>"&ii&"</b> 以上的用户以及作者才能浏览：</font><BR>"&tt2&"$6")
end if

re.Pattern="(^.*)(\[showtograde=*([0-9]*)\])(.[^\[]*)(\[\/s\])(.*)"
po=re.Replace(str,"$3")
if IsNumeric(po) then
ii=int(po) 
else
	ii=0
end if
if lgname="" then
dj=0
else
set ain=myconn.execute("select qian,meili,jingyan from user where name='"&lgname&"'")
q1=ain("qian")
j1=ain("jingyan")
m1=ain("meili")
set ain=nothing	
sqltype="lg"
%><!--#include file="upji.asp"--><%end if%>
<%if lgname=myname or dj>=ii then%><%
str=re.Replace(str,"$1"&tt1&"此内容需要等级为 <b>"&ii&"</b> 或更高的等级以及作者才能浏览：</font><BR>$4"&tt2&"$6")
else
str=re.Replace(str,"$1"&tt1&"此内容需要等级为 <b>"&ii&"</b> 或更高的等级以及作者才能浏览：</font><BR>"&tt2&"$6")
end if

re.Pattern="(^.*)(\[showtoname=(.[^\[]*)\])(.[^\[]*)(\[\/s\])(.*)"
usna=re.replace(str,"$3")
if lgname=usna or lgname=myname then
str=re.Replace(str,"$1"&tt1&"此内容只有作者和 <b>$3</b> 能浏览：</font><BR>$4"&tt2&"$6")
else
str=re.Replace(str,"$1"&tt1&"此内容只有作者和 <b>$3</b> 能浏览：</font><BR>"&tt2&"$6")
end if

re.Pattern="(^.*)(\[showtoreply\])(.[^\[]*)(\[\/s\])(.*)"
set see=myconn.execute("select riqi from min where bid="&id&" and name='"&lgname&"'")
if not see.eof or lgname=myname then
str=re.Replace(str,"$1"&tt1&"此内容只有作者和已经回复此帖的浏览者能浏览：</font><BR>$3"&tt2&"$5")
else
str=re.Replace(str,"$1"&tt1&"此内容只有作者和已经回复此帖的浏览者能浏览：</font><BR>"&tt2&"$5")
end if
re.Pattern="(^.*)(\[quote\])(.[^\[]*)(\[\/quote\])(.*)"
str=re.Replace(str,"$1"&tt1&"</font>$3"&tt2&"$5")
re.Pattern="\[align=(center|left|right)\](.*)\[\/align\]"
str=re.Replace(str,"<div align=$1>$2</div>")
re.Pattern="(\[URL=(.[^\[]*)\])(.[^\[]*)(\[\/URL\])"
str= re.Replace(str,"<A HREF=""$2"" TARGET=_blank>$3</A>")
set re=Nothing
ubb=str
end function

Rem 过滤HTML代码
function kb6k(fweing)
if not isnull(fweing) then
	fweing = replace(fweing, ">", "&gt;")
	fweing = replace(fweing, "<", "&lt;")
	fweing = Replace(fweing, CHR(32), " ")
	fweing = Replace(fweing, CHR(9), "&nbsp;")
	fweing = Replace(fweing, CHR(34), "&quot;")
	fweing = Replace(fweing, CHR(39), "&#39;")
fweing = Replace(fweing, CHR(13), "")
fweing = Replace(fweing, CHR(10), "<BR> ")
fweing = Replace(fweing, "[enter]", "<BR> ")

kb6k = fweing
end if
end function



%>