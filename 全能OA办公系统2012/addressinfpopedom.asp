<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
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
if request("submit")="应用" then
allow_edit_person_addressinf=request("allow_edit_person_addressinf")
if allow_edit_person_addressinf="" then allow_edit_person_addressinf="no"

allow_edit_dept_addressinf=request("allow_edit_dept_addressinf")
if allow_edit_dept_addressinf="" then allow_edit_dept_addressinf="no"

allow_edit_all_addressinf=request("allow_edit_all_addressinf")
if allow_edit_all_addressinf="" then allow_edit_all_addressinf="no"

allow_edit_client_addressinf=request("allow_edit_client_addressinf")
if allow_edit_client_addressinf="" then allow_edit_client_addressinf="no"

set conn=opendb("oabusy","conn","accessdsn")
sql="update userinf set "
sql=sql & "allow_edit_person_addressinf=" & sqlstr(allow_edit_person_addressinf) & ","
sql=sql & "allow_edit_all_addressinf=" & sqlstr(allow_edit_all_addressinf) & ","
sql=sql & "allow_edit_client_addressinf=" & sqlstr(allow_edit_client_addressinf) & ","
sql=sql & "allow_edit_dept_addressinf=" & sqlstr(allow_edit_dept_addressinf) & " where id=" & request("id")
conn.Execute sql

end if
%>
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='" & oabusyusername&"'"
rs.open sql,conn,1
cook_allow_control_all_user=rs("allow_control_all_user")     
conn.close
set conn=nothing
set rs=nothing
if cook_allow_control_all_user="no" then
response.write("<font color=red size=""+1"">对不起，您没有这个权限！</font>")
	response.end
	end if
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统.边缘特别版</title>
<style type="text/css">
<!--
.z14 {font-size: 14px;
	font-weight: bold;
	color: #098abb;
}
.style8 {color: #0d79b3}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style9 {color: #000000}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">
<SCRIPT language=javascript>
<!--
if (window.Event) 
　document.captureEvents(Event.MOUSEUP); 
 
function nocontextmenu() {
 event.cancelBubble = true
 event.returnvalue = false;
 return false;
}
 
function norightclick(e) {
 if (window.Event) {
　if (e.which == 2 || e.which == 3)
　 return false;
 } else if (event.button == 2 || event.button == 3) {
　 event.cancelBubble = true
　 event.returnvalue = false;
　 return false;
 } 
}
 
document.oncontextmenu = nocontextmenu;　// for IE5+
document.onmousedown = norightclick;　　 // for all others
//-->
</SCRIPT>
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="25"><div align="center">
      <table width="583"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="21"><div align="center">
              <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
                  <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                      <tr>
                        <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                        <td class="style7">用户设置</td>
                      </tr>
                  </table></td>
                  <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
                </tr>
              </table>
              <font color="0D79B3"></font></div></td>
        </tr>
      </table>
    <font color="0D79B3"></font></div></td>
  </tr>
  <tr>
    <td><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><center>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="30"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                        <tr>
                          <td height="20"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                              <tr>
                                <td>&nbsp;</td>
                              </tr>
                            </table>
                              <center>
                                <table>
                                  <tr>
                                    <td><span class="style8">编辑通讯录权限设置&nbsp;&nbsp;&nbsp;&nbsp;</span></td>
                                    <%
'打开数据库，读出部门
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select DISTINCT userdept from userinf"
rs.open sql,conn,1
%>
                                    <form method="post" action="addressinfpopedom.asp">
                                      <td>
                                        <select size=1 name="userdept">
                                          <%
if not rs.eof and not rs.bof then userdept=rs("userdept")
if request("userdept")<>"" then userdept=request("userdept")
while not rs.eof and not rs.bof
%>
                                          <option value="<%=rs("userdept")%>"<%=selected(userdept,rs("userdept"))%>><%=rs("userdept")%></option>
                                          <%
rs.movenext
wend
%>
                                        </select>
                                        <input name="submit" type="submit" value="查看">
                                      </td>
                                    </form>
                                  </tr>
                                </table>
                              </center>
                              <br>
                              <center>
                                <table border="1"  cellspacing="0" cellpadding="0" width="95%"  bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
                                  <tr bgcolor="D7E8F8">
                                    <td height=25 align=center><span class="style9">姓名</span></td>
                                    <td align=center><span class="style9">部门</span></td>
                                    <td align=center><span class="style9">职位</span></td>
                                    <td align=center><span class="style9">可编辑本人</span></td>
                                    <td align=center><span class="style9">可编辑本部门</span></td>
                                    <td align=center><span class="style9">可编辑全员</span></td>
                                    <td align=center><span class="style9">可编辑客户</span></td>
                                    <td>&nbsp;</td>
                                  </tr>
                                  <%
'显示用户表
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where userdept=" & sqlstr(userdept)
rs.open sql,conn,1
while not rs.eof and not rs.bof
%>
                                  <form method="post" action="addressinfpopedom.asp">
                                    <tr>
                                      <td align=center><%=rs("name")%></td>
                                      <td align=center><%=rs("userdept")%></td>
                                      <td align=center><%=rs("userlevel")%></td>
                                      <td align=center><input type="checkbox" name="allow_edit_person_addressinf" value="yes"<%=checked(rs("allow_edit_person_addressinf"),"yes")%>></td>
                                      <td align=center><input type="checkbox" name="allow_edit_dept_addressinf" value="yes"<%=checked(rs("allow_edit_dept_addressinf"),"yes")%>></td>
                                      <td align=center><input type="checkbox" name="allow_edit_all_addressinf" value="yes"<%=checked(rs("allow_edit_all_addressinf"),"yes")%>></td>
                                      <td align=center><input type="checkbox" name="allow_edit_client_addressinf" value="yes"<%=checked(rs("allow_edit_client_addressinf"),"yes")%>></td>
                                      <td align=center><input type="submit" name="submit" value="应用"></td>
                                    </tr>
                                    <input type="hidden" name="id" value=<%=rs("id")%>>
                                    <input type="hidden" name="userdept2" value=<%=userdept%>>
                                  </form>
                                  <%
rs.movenext
wend
%>
                                </table>
                              </center>
                              <br>
                              <%

%>
                              <p><br>
                                <br>
                            </p>
                          </td>
                        </tr>
                    </table></td>
                  </tr>
                </table>
            </center></td>
          </tr>
        </table>
        <center>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>
