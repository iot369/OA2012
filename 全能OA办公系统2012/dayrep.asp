<%@ LANGUAGE = VBScript %>
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
.style8 {color: #2b486a}
-->
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">���˹����ƻ�</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/monthlycal.asp"-->

<!--#include file="asp/checked.asp"-->
<%
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

username=oabusyusername
'ȡ�õ�ǰ����
mydate=date
'ȡ�õ�ǰ��
mymonth=month(now())
if request("mymonth")<>"" then mymonth=cint(request("mymonth"))
'ȡ�õ�ǰ��
myyear=year(now())
if request("myyear")<>"" then myyear=cint(request("myyear"))
%>


<center>
  <br>
<table>
<tr>
<form action="dayrep.asp" method=get name="form1">
<td>
<select name="myyear" size=1>
<option value=2001<%=selected(2001,myyear)%>>2001��</option>
<option value=2002<%=selected(2002,myyear)%>>2002��</option>
<option value=2003<%=selected(2003,myyear)%>>2003��</option>
<option value=2004<%=selected(2004,myyear)%>>2004��</option>
<option value=2005<%=selected(2005,myyear)%>>2005��</option>
<option value=2006<%=selected(2006,myyear)%>>2006��</option>
<option value=2007<%=selected(2007,myyear)%>>2007��</option>
<option value=2008<%=selected(2008,myyear)%>>2008��</option>
<option value=2009<%=selected(2009,myyear)%>>2009��</option>
</select>
<select name="mymonth" size=1>
<option value=1<%=selected(1,mymonth)%>>1��</option>
<option value=2<%=selected(2,mymonth)%>>2��</option>
<option value=3<%=selected(3,mymonth)%>>3��</option>
<option value=4<%=selected(4,mymonth)%>>4��</option>
<option value=5<%=selected(5,mymonth)%>>5��</option>
<option value=6<%=selected(6,mymonth)%>>6��</option>
<option value=7<%=selected(7,mymonth)%>>7��</option>
<option value=8<%=selected(8,mymonth)%>>8��</option>
<option value=9<%=selected(9,mymonth)%>>9��</option>
<option value=10<%=selected(10,mymonth)%>>10��</option>
<option value=11<%=selected(11,mymonth)%>>11��</option>
<option value=12<%=selected(12,mymonth)%>>12��&nbsp;&nbsp;</option>
</select>
<input type="submit" name="submit" value="��ѯ">
</td>
</form>
</tr>
</table>
<span class="style8"><br>����<%=myyear%>��<%=mymonth%>�¹����ƻ�</span>(<font color=red>������д���ܹ����ܽᣬ���մ�д���ܹ����ƻ�</font>)
</center>

<center>
<br>
<span class="style8">����Ϊ</span>��ɫ<span class="style8">��ʾ�˹�������һ����û��ɣ�����Ϊ</span><font color=blue>��ɫ</font><span class="style8">��ʾ�˹�������һ���������</span><br>
<span class="style8">����Ϊ</span><font color=red>��ɫ</font><span class="style8">��ʾ�˹���������Ҫ��û��ɣ�����Ϊ</span><font color=#770000>��ɫ</font><span class="style8">��ʾ�˹���������Ҫ�������<br>
</span><br>
</center>
<%
call monthlycal(username,superior)

%>
</body>
</html>