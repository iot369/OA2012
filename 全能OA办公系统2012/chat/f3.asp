<%Response.Expires=0%>

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='20;url=f3.asp'>
<meta http-equiv="pragma" content="no-cache">

<style type="TEXT/CSS"> 
<!--
body,table {color:#FFFFFF;font-family: 宋体_GB2312; font-size: 9pt; line-height: 12pt}
A:link {text-decoration: none; color:#E0E0FF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:visited {text-decoration: none; color: #E0E0FF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:active {text-decoration: underline; color: #E0E0FF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:hover {text-decoration: none; color: E0FFFF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
-->
</style>


<script language="JavaScript">
<!--
function selectwho(list){
parent.f2.document.forms[0].towho.options[0].value=list;
parent.f2.document.forms[0].towho.options[0].text=list;
parent.f2.document.forms[0].saystemp.focus();
}
//-->
</script>
</head>
<body bgcolor=#0000FF text="#FFFFFF" background="rbg.gif">
<%
usernum=0 '在线人数

OUN=Application(SESSION("CRNAME")&"OUN")
OULT=Application(SESSION("CRNAME")&"OULT")
sentences=Application(SESSION("CRNAME")&"sentences")
cur=Application(SESSION("CRNAME")&"cur")
whotowho=Application(SESSION("CRNAME")&"whotowho")

for i=1 to 60
If len(OUN(i))>0 then usernum=usernum+1 '统计在线人数
next


'删除过期会议用户
for i=usernum to 1 step -1
If abs(DateDiff("s",OULT(i),Now))>120 then
cur=cur+1
if cur>60 then cur=1

sentences(cur)="<font color=#FF0000>[公告]</font>"&OUN(i)&"刚刚离开<u>"&Session("CRNAME")&"</u>……<font color=#B0B0B0>("&OULT(i)&")</font>"
whotowho(cur,1)="System"
whotowho(cur,2)="大家"

for os=i to usernum-1
OUN(os)=OUN(os+1)
OULT(os)=OULT(os+1)
next
OUN(usernum)=EMPTY
OULT(usernum)=EMPTY
usernum=usernum-1
End If
next

Application.Lock
Application(SESSION("CRNAME")&"sentences")=sentences
Application(SESSION("CRNAME")&"whotowho")=whotowho
Application(SESSION("CRNAME")&"usernum")=usernum
Application(SESSION("CRNAME")&"cur")=cur
Application(SESSION("CRNAME")&"OUN")=OUN
Application(SESSION("CRNAME")&"OULT")=OULT
Application.unLock


Response.Write("在线人数："&usernum)
%>
<BR>
<HR>
<a href="javascript:selectwho('大家');" onMouseOver="window.status='选择对话或动作对象为：大家'" onMouseOut="window.status=''">大家</a><br>
<%for x=1 to usernum%>
<a href="javascript:selectwho('<%=OUN(x)%>');" onMouseOver="window.status='选择对话或动作对象为：<%=OUN(x)%>'" onMouseOut="window.status=''"><%=OUN(x)%></a><br>
<%next%>



</body></html>
