<!--#include file="md5.asp"--><%width=40
height=40
function laiyuan()
laiyuan=false
come=Request.ServerVariables("HTTP_REFERER")
here=Request.ServerVariables("SERVER_NAME")
if mid(come,8,len(here))<>here then
laiyuan=false
else
laiyuan=true
end if
end function
laiyuan()
if laiyuan=false then
response.redirect"index.asp"
end if
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
%>
<!--#include file="up.asp"-->
<style>TABLE {BORDER-TOP: 0px; BORDER-LEFT: 0px; BORDER-BOTTOM: 1px; }TD {BORDER-RIGHT: 0px; BORDER-TOP: 0px;}</style>
<%
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
noyes="注 册 失 败 ！"
name=Replace(Request.Form("name"),"'","''")
if strlength(name)>16 or Instr(name,"=")>0 or Instr(name,"%")>0 or Instr(name,chr(32))>0 or Instr(name,"?")>0 or Instr(name,"&")>0 or Instr(name,";")>0 or Instr(name,",")>0 or Instr(name,"'")>0 or Instr(name,",")>0 or Instr(name,chr(34))>0 or Instr(name,chr(9))>0 or Instr(name,"")>0 or Instr(name,"$")>0 then
can="no"
end if
password=Replace(Request.Form("password"),"'","''")
if strlength(password)>16 or Instr(password,"=")>0 or Instr(password,"%")>0 or Instr(password,chr(32))>0 or Instr(password,"?")>0 or Instr(password,"&")>0 or Instr(password,";")>0 or Instr(password,",")>0 or Instr(password,"'")>0 or Instr(password,",")>0 or Instr(password,chr(34))>0 or Instr(password,chr(9))>0 or Instr(password,"")>0 or Instr(password,"$")>0 then
can="no"
end if
repassword=Replace(Request.Form("repassword"),"'","''")
anhao=Replace(Request.Form("anhao"),"'","''")
if strlength(anhao)>16 or Instr(anhao,"=")>0 or Instr(anhao,"%")>0 or Instr(anhao,chr(32))>0 or Instr(anhao,"?")>0 or Instr(anhao,"&")>0 or Instr(anhao,";")>0 or Instr(anhao,",")>0 or Instr(anhao,"'")>0 or Instr(anhao,",")>0 or Instr(anhao,chr(34))>0 or Instr(anhao,chr(9))>0 or Instr(anhao,"")>0 or Instr(anhao,"$")>0 then
can="no"
end if
nameok=Replace(Request.Form("name")," ","")
passwordok=Replace(Request.Form("password")," ","")
repasswordok=Replace(Request.Form("repassword")," ","")
anhaook=Replace(Request.Form("anhao")," ","")
email=Replace(Request.Form("email"),"'","''")
set rs=myconn.execute("SELECT*FROM user where name='"&name&"'")
if not rs.eof and not rs.bof then
mes="<br>对不起！"&kbbs(name)&" 已被人注册了！！！ <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
%><%else
if nameok="" or passwordok="" or repasswordok="" or anhaook="" or email="" then
mes="<br>对不起！你不能成功地注册用户！！！请填写完整必填的项目 <a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
%><%
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
elseif not IsValidEmail(email) then
mes="<br>对不起！你不能成功地注册用户！！！请检查你的E-mail是否出错！！<a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
%><%
else
if can="no" then
mes="<br>你的用户名、密码 或 备用密码 含有非法字符或者字符过多！！<a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
else
if repassword<>password then
mes="<br>你的重复密码与原密码不相同！！<a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
%><%
else
%>
<%
name=Replace(Request.Form("name"),"'","''")
password=Replace(Request.Form("password"),"'","''")
repassword=Replace(Request.Form("repassword"),"'","''")
anhao=Replace(Request.Form("anhao"),"'","''")
mypic=Replace(Request.Form("mypic"),"'","''")
mypic=ubbg(mypic)
toupic=Replace(Request.Form("headpic"),"'","''")
email=Replace(Request.Form("email"),"'","''")
home=Replace(Request.Form("home"),"'","''")
sex=Replace(Request.Form("sex"),"'","''")
burn=Replace(Request.Form("burn"),"'","''")
qq=Replace(Request.Form("qq"),"'","''")
gxqm=Replace(Request.Form("gxqm"),"'","''")
ch=Replace(Request.Form("ch"),"'","''")
ku=Replace(Request.Form("ku"),"'","''")
mytp=mypic
if mypic="" then
mytp="headpic/"&toupic&".gif"
ch=height
ku=width
end if
if not isnumeric(ch) or not isnumeric(ku) then
mes="<br>你的图像大小设置错误！！<a href='javascript:history.go(-1)'><img border=0 src=pic/re.gif align=absmiddle> 返 回</a><br><br>"
%>
<%else%>
<%if ch>120 or ku>120 then
ch=height
ku=width
end if%>

<%
passworda=md5(password)
anhao=md5(anhao)
abc="insert into user(name,password,anhao,email,home,sex,burn,qq,toupic,ch,ku,gxqm,qian,meili,jingyan)VALUES('"&name&"','"&passworda&"','"&anhao&"','"&email&"','"&home&"','"&sex&"','"&burn&"','"&qq&"','"&mytp&"','"&ch&"','"&ku&"','"&gxqm&"',1000,200,200)"
myconn.Execute(abc)
noyes="注 册 成 功！"
mes="<br><form method=POST action=bbselse.asp name=login>恭喜你！ <b>"&kbbs(name)&"</b> 成功注册！<input type=hidden name=lgname size=20 value="&name&"><input type=hidden name=lgpwd size=20 value="&password&"><a href='javascript:document.login.submit()'><img border=0 src=pic/go.gif align=absmiddle> 进入论坛</a></form>"
%>    
<%end if
end if
end if
end if
end if
%><!--#include file="mes.asp"--><br><!--#include file="down.asp"-->