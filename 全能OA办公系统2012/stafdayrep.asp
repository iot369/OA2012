<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/monthlycal.asp"-->

<!--#include file="asp/opendb.asp"-->
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

cook_allow_see_all_workrep=request.cookies("cook_allow_see_all_workrep")
cook_allow_see_dept_workrep=request.cookies("cook_allow_see_dept_workrep")

if request("username")<>"" then username=request("username")
'ȡ�õ�ǰ����
mydate=date
'ȡ�õ�ǰ��
mymonth=month(now())
if request("mymonth")<>"" then mymonth=cint(request("mymonth"))
'ȡ�õ�ǰ��
myyear=year(now())
if request("myyear")<>"" then myyear=cint(request("myyear"))
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>�ƶ��Զ���OA�칫ϵͳ</title>
<style type="text/css">
<!--
.style3 {color: #098abb}
.h14 {
	line-height: 20px;
}
.style5 {color: #2d4865}
.style6 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                <td><span class="style5">���������ƻ�</span></td>
              </tr>
          </table></td>
          <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
        </tr>
      </table>
    <font color="0D79B3"></font></div></td>
  </tr>
  <tr>
    <td><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
        <center>
          <center>
            <%
if cook_allow_see_all_workrep="yes" then
%>
            <table>
              <tr>
                <form action="stafdayrep.asp" method=get name="form1">
                  <td>
                    <select name="userdept" size=1 onChange="document.form1.submit();">
                      <%
'�����ݿ��������
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
if not rs.eof and not rs.bof then firstdept=rs("userdept")
if request("userdept")<>"" then firstdept=request("userdept")
while not rs.eof and not rs.bof
%>
                      <option value="<%=rs("userdept")%>"<%=selected(firstdept,rs("userdept"))%>><%=rs("userdept")%></option>
                      <%
rs.movenext
wend
%>
                    </select>
                  </td>
                </form>
                <form action="stafdayrep.asp" method=get name="form2">
                  <td>
                    <input type="hidden" name="userdept2" value="<%=firstdept%>">
                    <select name="username" size=1>
                      <%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select name,username from userinf where userdept=" & sqlstr(firstdept) & " and userlevel<>'�ܹ�' and forbid='no'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then username=rs("username")
if request("username")<>"" then username=request("username")
while not rs.eof and not rs.bof
%>
                      <option value="<%=rs("username")%>"<%=selected(username,rs("username"))%>><%=rs("name")%></option>
                      <%
rs.movenext
wend
%>
                    </select>
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
            <%
else
%>
            <table>
              <tr>
                <form action="stafdayrep.asp" method=get name="form1">
                  <td>
                    <select name="username" size=1>
                      <%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where userdept=" & sqlstr(oabusyuserdept) & " and userlevel<>'�ܹ�' and forbid='no'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then username=rs("username")
if request("username")<>"" then username=request("username")
while not rs.eof and not rs.bof

%>
                      <option value="<%=rs("username")%>"<%=selected(username,rs("username"))%>><%=rs("name")%></option>
                      <%
rs.movenext
wend
%>
                    </select>
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
            <br>
            <%
end if
%>
            <%
'��������
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(username)
rs.open sql,conn,1
if not rs.eof and not rs.bof then
name=rs("name")
else
%>
          </center>
          <br>
          <br>
          <center>
            <font color=red > ����������Ա����������<br> 
            </font>
          </center>
          <span class="style3">
          <br>
          </span>          <span class="style3">
          <%
response.end
end if
%>
          <br>
          <font color="#FF0048"><%=name%></font><span class="style6">��</span><font color="#FF0048"><%=myyear%>��<%=mymonth%>��</font><span class="style6">��������</span></span> <br>
          <center>
            <p class="h14"><span class="style6">����Ϊ</span>��ɫ<span class="style6">��ʾ�˹�������һ����û��ɣ�����Ϊ</span><font color=blue>��ɫ</font><span class="style6">��ʾ�˹�������һ���������</span><br>
            <span class="style6">����Ϊ</span><font color=red>��ɫ</font><span class="style6">��ʾ�˹���������Ҫ��û��ɣ�����Ϊ</span><font color=#770000>��ɫ</font><span class="style6">��ʾ�˹���������Ҫ�������</span> </p>
          </center>
          
          <% call monthlycal(username,oabusyusername) %>
<br>
          <br>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>