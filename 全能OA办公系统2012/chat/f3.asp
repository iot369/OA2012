<%Response.Expires=0%>

<html>
<head>
<meta http-equiv='content-type' content='text/html; charset=gb2312'>
<meta http-equiv=refresh content='20;url=f3.asp'>
<meta http-equiv="pragma" content="no-cache">

<style type="TEXT/CSS"> 
<!--
body,table {color:#FFFFFF;font-family: ����_GB2312; font-size: 9pt; line-height: 12pt}
A:link {text-decoration: none; color:#E0E0FF; font-family: "����"; font-size: 9pt; line-height: 12pt}
A:visited {text-decoration: none; color: #E0E0FF; font-family: "����"; font-size: 9pt; line-height: 12pt}
A:active {text-decoration: underline; color: #E0E0FF; font-family: "����"; font-size: 9pt; line-height: 12pt}
A:hover {text-decoration: none; color: E0FFFF; font-family: "����"; font-size: 9pt; line-height: 12pt}
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
usernum=0 '��������

OUN=Application(SESSION("CRNAME")&"OUN")
OULT=Application(SESSION("CRNAME")&"OULT")
sentences=Application(SESSION("CRNAME")&"sentences")
cur=Application(SESSION("CRNAME")&"cur")
whotowho=Application(SESSION("CRNAME")&"whotowho")

for i=1 to 60
If len(OUN(i))>0 then usernum=usernum+1 'ͳ����������
next


'ɾ�����ڻ����û�
for i=usernum to 1 step -1
If abs(DateDiff("s",OULT(i),Now))>120 then
cur=cur+1
if cur>60 then cur=1

sentences(cur)="<font color=#FF0000>[����]</font>"&OUN(i)&"�ո��뿪<u>"&Session("CRNAME")&"</u>����<font color=#B0B0B0>("&OULT(i)&")</font>"
whotowho(cur,1)="System"
whotowho(cur,2)="���"

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


Response.Write("����������"&usernum)
%>
<BR>
<HR>
<a href="javascript:selectwho('���');" onMouseOver="window.status='ѡ��Ի���������Ϊ�����'" onMouseOut="window.status=''">���</a><br>
<%for x=1 to usernum%>
<a href="javascript:selectwho('<%=OUN(x)%>');" onMouseOver="window.status='ѡ��Ի���������Ϊ��<%=OUN(x)%>'" onMouseOut="window.status=''"><%=OUN(x)%></a><br>
<%next%>



</body></html>
