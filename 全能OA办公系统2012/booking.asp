<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/check_resource.asp"-->
<%
today1=date()
today2=date()+1
'a1=datediff("h",today1,#2001-5-6 1:00:00#)
'response.write a1
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

'--------------------------------------
'ɾ����ǰ�ļ�¼
resourceflag=check_resource_setting(oabusyusername,0)
set conn=opendb("oabusy","conn","accessdsn")
sql = "delete * from booking where endtime<#" & date() & "#"
conn.Execute sql
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA�칫ϵͳ.��Ե�ر��</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">������Դ</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
      <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td><table width="1%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table align="center">
<tr>
<%
if resourceflag=0 then
%>
<form method=post action="addequipment.asp">
<td>
<input type="submit" value="����">
</td>
</form>
<%
set rs=nothing
end if
%>
<form method=post action="addbooking.asp">
<td>
<input type="submit" value="ԤԼ">
</td>
</form>
</tr>
</table>
  <div align="center">������Դռ�����ͼ��<font color=red>��ɫ</font>��ʾ��ռ��ʱ��Σ�
      </center>
  
</div>
  <center>
<table width="96%" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<%
'��ʾͼ��
'�����ݿ⣬�����豸
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr bgcolor="#FFFFFF">
<%
if resourceflag=0 then
%>
<td width=90 align=center bgcolor="D7E8F8"><a href="editequipment.asp?id=<%=rs("id")%>"><%=rs("equipment")%></a></td>
<td align=center bgcolor="#FFFFFF">
<%
else
%>
<td width=90 align=center><%=rs("equipment")%></td><td align=center bgcolor="#FFFFFF">
<%
end if
'����ԤԼ����
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from booking where endtime>#" & today1 & "# and starttime<#" & today2 & "# and equipment=" & sqlstr(rs("equipment")) & " order by starttime"
'response.write sql
rs1.open sql,conn,1
%>
<table border="0"  cellspacing="0" cellpadding="0" width="430" height="10">
<tr>
<td align=left>00:00</td><td align=center>12:00</td><td align=right>24:00</td>
</tr>
</table>
<%
if rs1.eof or rs1.bof then
%>

        <table border="0"  cellspacing="0" cellpadding="0" width="400" height="5">
          <tr>
            <td bgcolor="3499D0"><font color="#006699">1</font></td>
</tr>
</table>
<%
else
%>
        <table border="0"  cellspacing="0" cellpadding="0" width="400" height="5">
          <tr>
<%
'��ʼ��ǰ��û��¼
a=0
while not rs1.eof and not rs1.bof
%>

<%
'����ǵ�һ����¼
if a=0 then
'���ڲ��ǵ�һ����¼��
a=1
'���㿪ʼʱ�䵽����Сʱ��
starttime=rs1("starttime")
endtime=rs1("endtime")
starttime1=rs1("starttime")
endtime1=rs1("endtime")
a1=datediff("h",today1,starttime)
'�����ʼʱ���ڽ������ʾ��ɫ�ĳ���
if a1>0 then
'����ɫ������
colorwidth=a1/24*400
'��ʾ��ɫ
%>
            <td bgcolor="3499D0" width="<%=colorwidth%>" align=center height="8"><%=a1%></td>
<%
end if
'�������ʱ�䵽����Сʱ��
b1=datediff("h",today1,endtime)
'���a1<=0,b1>=24
if a1<=0 and b1>=24 then
%>
            <td bgcolor="#FF0000" width="400" align=center height="8">24</td>
<%
end if
'���a1<=0,b1<24
if a1<=0 and b1<24 then
colorwidth=b1/24*400
%>
            <td bgcolor="#FF0000" width="<%=colorwidth%>" align=center height="8"><%=b1%></td>
<%
end if
'���a1>0,b1>=24
if a1>0 and b1>=24 then
colorwidth=(24-a1)/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=(24-a1)%></td>
<%
end if
'���a1>0,b1<24
if a1>0 and b1<24 then
c1=datediff("h",starttime,endtime)
colorwidth=c1/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=c1%></td>
<%
end if
else
'������ɫ����
starttime=rs1("starttime")
endtime=rs1("endtime")
a1=datediff("h",endtime1,starttime)
'�����ʱ��������ʾ��ɫ����
if a1>0 then
colorwidth=a1/24*400
%>
            <td bgcolor="3499D0" width="<%=colorwidth%>" align=center height="8"><%=a1%></td>
<%
end if
'��ʾ��ɫ����
b1=datediff("h",today1,endtime)
'���b1>=24
if b1>=24 then
d1=datediff("h",today1,starttime)
colorwidth=(24-d1)/24*400
%>
            <td bgcolor="#FF0000" width="<%=colorwidth%>" align=center height="8">8<%=(24-d1)%></td>
<%
end if
'���b1<24
if b1<24 then
d1=datediff("h",starttime,endtime)
colorwidth=d1/24*400
%>
<td bgcolor="#FF0000" width="<%=colorwidth%>" align=center><%=d1%></td>
<%
end if
starttime1=rs1("starttime")
endtime1=rs1("endtime")
end if
rs1.movenext
wend
'��ʾ�����ɫ��
e1=datediff("h",today1,endtime1)
if e1<24 then
colorwidth=(24-e1)/24*400
%>
<td bgcolor="3499D0" width="<%=colorwidth%>" align=center><%=(24-e1)%></td>
<%
end if
%>
</tr>
</table>
<%
end if
%>
</td>
</tr>
<%
rs.movenext
wend
%>
</table>
  <br>
  <table width="96%" border="0" cellpadding="0"  cellspacing="1" bgcolor="B0C8EA">
<%
'�����ݿ⣬�����豸
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from equipment"
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
<tr>
<td bgcolor="D7E8F8"><%=rs("equipment")%></td>
</tr>
<tr>
<td height=40 bgcolor="#FFFFFF">
<%
'����ԤԼ����
set conn=opendb("oabusy","conn","accessdsn")
set rs1=server.createobject("adodb.recordset")
sql="select * from booking where equipment=" & sqlstr(rs("equipment")) & " order by starttime"
rs1.open sql,conn,1
while not rs1.eof and not rs1.bof
	select case rs1("auditing")
		case 0
			imgsrc="image/resource_auditing.gif"
			explainstr="<font color=""#0000ff"">δ���</font>"
		case 1
			nowtime=cdate(cstr(date())&" "&cstr(time()))
			if nowtime>=rs1("starttime") and nowtime<=rs1("endtime") then
				imgsrc="image/resource_go.gif"
				explainstr="ʹ����..."
			elseif nowtime<rs1("starttime") then
				imgsrc="image/stay_do.gif"
				explainstr="�ȴ�ʹ��"
			elseif nowtime>rs1("endtime") then
				imgsrc="image/finish.gif" 
				explainstr="�����"
			end if
		case 2
			imgsrc="image/auditing_no.gif"
			explainstr="<font color=""#ff0000"">���δͨ��</font>"
	end select
%>
<img src="<%=imgsrc%>" border="0">
<a href="editbooking.asp?id=<%=rs1("id")%>"><font color="#0000ff">[<%=rs1("starttime")%>----<%=rs1("endtime")%>]</font></a>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs2=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(rs1("username"))
rs2.open sql,conn,1
if not rs2.eof and not rs2.bof then
if oabusyusername=rs1("username") then response.write "<font color=red>"
%>
&nbsp;&nbsp;���ţ�<%=rs2("userdept")%>&nbsp;&nbsp;ԤԼ�ߣ�<%=rs2("name")%><%
if oabusyusername=rs1("username") then response.write "</font>"
response.write("["&explainstr&"]")
end if
%>
<br>
<%
rs1.movenext
wend

rs.movenext
wend
%>
</table>
</center></td>
    </tr>
</table>
</body>
</html>










