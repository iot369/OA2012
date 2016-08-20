<%width=40
height=40
%>
<!--#include file="up.asp"--><!--#include file="fun.asp"--><!--#include file="md5.asp"-->
<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<br><br><%
function strLength(str)
       ON ERROR RESUME NEXT
       dim WINNT_CHINESE
       WINNT_CHINESE    = (len("论坛")=2)
       if WINNT_CHINESE then
          dim l,t,c
          dim i
          l=len(str)
          t=l
          for i=1 to l
             c=asc(mid(str,i,1))
             if c<0 then c=c+65536
             if c>255 then
                t=t+1
             end if
          next
          strLength=t
       else 
          strLength=len(str)
       end if
       if err.number<>0 then err.clear
end function
function ubbg(str)
dim re
	Set re=new RegExp
	re.IgnoreCase=true
	re.Global=True
re.Pattern="(height|javascript|jscript:|js:|value|about:|file:|document.cookie|vbscript:|vbs:|script|width|)"
str=re.Replace(str,"")

re.Pattern="(on(mouse|exit|error|click|key))"
str=re.Replace(str,"")

set re=Nothing
ubbg=str
end function
sty="<P style='MARGIN: 8px'>"
t1="<div align=center><center><table border=1 cellpadding=0 cellspacing=0 style='border-collapse: collapse' bordercolor="&c1&" width=94% ><tr><td width=100% background=pic/"&sp&"3.gif height=25 bgcolor="&c1&">&nbsp;<img border=0 src=pic/fl.gif> <font color=#FFFFFF><b>"
t2="</b></font></td></tr>"
d1="<tr><td width=100% ><P style='MARGIN: 14px'>"
d2="</td></tr></table></center></div>"
newpwd=Replace(Request.Form("newpwd"),"'","''")
repwd=Replace(Request.Form("repwd"),"'","''")
email=Replace(Request.Form("email"),"'","''")
anhao=Replace(Request.Form("anhao"),"'","''")
sex=Replace(Request.Form("sex"),"'","''")
burn=Replace(Request.Form("burn"),"'","''")
home=Replace(Request.Form("home"),"'","''")
qq=Replace(Request.Form("qq"),"'","''")
toupic=Replace(Request.Form("toupic"),"'","''")
mypic=Replace(Request.Form("mypic"),"'","''")
mypic=ubbg(mypic)
ch=Replace(Request.Form("ch"),"'","''")
ku=Replace(Request.Form("ku"),"'","''")
gxqm=Replace(Request.Form("gxqm"),"'","''")
set canl=myconn.execute("select*from user where name='"&lgname&"' and password='"&lgpwd&"'")
if canl.eof or canl.bof then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·你还没有登陆或者你登陆的用户名或密码错误！·<%=d2%><%
response.end
end if
anan=canl("anhao")
if strlength(newpwd)>16 or Instr(newpwd,"=")>0 or Instr(newpwd,"%")>0 or Instr(newpwd,chr(32))>0 or Instr(newpwd,"?")>0 or Instr(newpwd,"&")>0 or Instr(newpwd,";")>0 or Instr(newpwd,",")>0 or Instr(newpwd,"'")>0 or Instr(newpwd,",")>0 or Instr(newpwd,chr(34))>0 or Instr(newpwd,chr(9))>0 or Instr(newpwd,"")>0 or Instr(newpwd,"$")>0 then
can="no"
end if
if strlength(anhao)>16 or Instr(anhao,"=")>0 or Instr(anhao,"%")>0 or Instr(anhao,chr(32))>0 or Instr(anhao,"?")>0 or Instr(anhao,"&")>0 or Instr(anhao,";")>0 or Instr(anhao,",")>0 or Instr(anhao,"'")>0 or Instr(anhao,",")>0 or Instr(anhao,chr(34))>0 or Instr(anhao,chr(9))>0 or Instr(anhao,"")>0 or Instr(anhao,"$")>0 then
can="no"
end if
if newpwd="" or repwd="" or anhao="" or email="" then
%>
<%=t1%>错 误 信 息<%=t2&d1%>·对不起！请填写完整必填的项目· <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><%=d2%>
<%response.end
end if%>
<%if can="no" then%>
<%=t1%>错 误 信 息<%=t2&d1%>·你的密码 或 备用密码 含有非法字符或者字符过多· <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><%=d2%>
<%response.end
end if%>

<%if repwd<>newpwd then%>
<%=t1%>错 误 信 息<%=t2&d1%>·你的重复密码与你的新密码不匹配· <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><%=d2%>
<%response.end
end if%>
<%
function IsValidEmail(email)

 dim names, name, i, c


 IsValidEmail = true
 names = Split(email, "@")
 if UBound(names) <> 1 then
   IsValidEmail = false
   exit function
 end if
 for each name in names
   if Len(name) <= 0 then
     IsValidEmail = false
     exit function
   end if
   for i = 1 to Len(name)
     c = Lcase(Mid(name, i, 1))
     if InStr("abcdefghijklmnopqrstuvwxyz_-.", c) <= 0 and not IsNumeric(c) then
       IsValidEmail = false
       exit function
     end if
   next
   if Left(name, 1) = "." or Right(name, 1) = "." then
      IsValidEmail = false
      exit function
   end if
 next
 if InStr(names(1), ".") <= 0 then
   IsValidEmail = false
   exit function
 end if
 i = Len(names(1)) - InStrRev(names(1), ".")
 if i <> 2 and i <> 3 then
   IsValidEmail = false
   exit function
 end if
 if InStr(email, "..") > 0 then
   IsValidEmail = false
 end if

end function
email=request.form("email")
email=server.HTMLEncode(email)
if not IsValidEmail(email) then
%><%=t1%>错 误 信 息<%=t2&d1%>·请检查你的 <b>E-mail</b> 是否填写准确· <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><%=d2%>
<%response.end
end if%>
<%mytp=mypic
if mypic="" then
mytp="headpic/"&toupic&".gif"
ch=height
ku=width
end if
if not isnumeric(ch) or not isnumeric(ku) then
%><%=t1%>错 误 信 息<%=t2&d1%>·你的图像大小设置错误· <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><%=d2%>
<%response.end
end if%>
<%
if ch>120 or ku>120 then
ch=height
ku=width
end if%>
<%if newpwd<>lgpwd then
newpwd=md5(newpwd)
end if
if anhao<>anan then
anhao=md5(anhao)
end if
myconn.execute("update [user] set password='"&newpwd&"',burn='"&burn&"' ,anhao='"&anhao&"',ch='"&ch&"',ku='"&ku&"',email='"&email&"',qq='"&qq&"',sex='"&sex&"',toupic='"&mytp&"',home='"&home&"',gxqm='"&gxqm&"' WHERE name='"&lgname&"'")
myconn.execute("update admin set password='"&newpwd&"' where name='"&lgname&"'")
%>
<%=t1%>修 改 成 功<%=t2&d1%>·你已经成功的修改了你的用户信息！如果你有修改密码，请 <a href="login.asp">重新登陆</a>·<%=d2%>
<br><!--#include file="down.asp"-->