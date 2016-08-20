<%
'If1
If len(Request("UserName"))=0 or len(Request("CRNAME"))=0 then Response.Redirect("index.asp")
'EIf1

Session("CRNAME")=Server.HtmlEncode(Request("CRNAME"))
Session("username")=Server.HtmlEncode(Request("UserName"))

'IF1
If IsEmpty(Application(SESSION("CRNAME")&"sentences")) then
dim sentences(60)
dim whotowho(60,2)
dim OUN(60) 'Online User Name
dim OULT(60) 'Online User Last Timer
OUN(1)=Session("username")
OULT(1)=Now
cur=1
sentences(cur)="<font color=#FF0000>网络会议中心 Running at "&Now&"</font>"
whotowho(cur,1)="System"
whotowho(cur,2)="大家"

cur=cur+1
If cur>60 then cur=1
sentences(cur)="<font color=#FF0000>[公告]</font>"&Session("username")&"刚刚进入<u>"&Session("CRNAME")&"</u>……<font color=#B0B0B0>("&Now&")</font>"
whotowho(cur,1)="System"
whotowho(cur,2)="大家"
Application.Lock
Application(SESSION("CRNAME")&"sentences")=sentences
Application(SESSION("CRNAME")&"whotowho")=whotowho
Application(SESSION("CRNAME")&"OUN")=OUN
Application(SESSION("CRNAME")&"OULT")=OULT
Application(SESSION("CRNAME")&"cur")=cur
Application(SESSION("CRNAME")&"usernum")=1
Application.UnLock

else

OUN=Application(SESSION("CRNAME")&"OUN")
OULT=Application(SESSION("CRNAME")&"OULT")
for i=1 to 60
'IF2
If Session("username")=OUN(i) and abs(DateDiff("s",OULT(i),Now))<130 then
Response.Write("<Font color=#FF0000>ERROR!<BR>您希望使用的这个名字正在被其它人使用，请换用其它名字！</FONT>")
Response.End
End If
'EIF2
next
'检测人数是否超出最大限量（60）
'IF2
If Application(SESSION("CRNAME")&"usernum")>=59 then
Response.Write("<Font color=#FF0000>ERROR!<R>对不起，本会议室已达到最大同时在线人数，请进入其它会议室</font>")
Response.End
End If
'EIF2
End If
'EIF1

Response.Redirect("chat.asp")
%>
