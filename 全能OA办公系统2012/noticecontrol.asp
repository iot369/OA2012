<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->

<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
cook_allow_send_note=request.cookies("cook_allow_send_note")
cook_allow_control_note=request.cookies("cook_allow_control_note")

if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1

cook_allow_send_note=rs("allow_send_note")
cook_allow_control_note=rs("allow_control_note")

'取得当前月
mymonth=month(now())
if request("mymonth")<>"" then mymonth=cint(request("mymonth"))
'取得当前年
myyear=year(now())
if request("myyear")<>"" then myyear=cint(request("myyear"))

if request("detel")="删 除 通 告" and request("delid")<>"" then
set conn=opendb("oabusy","conn","accessdsn")
count=0
condition=""
for each idno in request("delid")
count=count+1
condition=condition+"id=" & idno
if count<request("delid").count then
condition=condition+" or "
end if
next
'删除数据库中的记录
sql = "delete * from newnotice where " & condition
conn.Execute sql
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<title>菲尔自动化办公系统 管理通告</title>
<style type="text/css">
<!--
.style6 {color: #098abb}
.z14 {	font-size: 14px;
	font-weight: bold;
	color: #098abb;
}
.style7 {color: #FF0000}
.style5 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                <td><span class="style5">公司通告</span></td>
              </tr>
          </table></td>
          <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
        </tr>
      </table>
    </div></td>
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
                        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                          <tr>
                            <td><table width="550"  border="0" align="center" cellpadding="0" cellspacing="0">
                              <tr>
                                <td width="350" height="50"><table width="300" align="center">
                                  <tr>
                                    <form method=post action="noticecontrol.asp">
                                      <td width="45"><input type="submit" name="submit" value="全部"></td>
                                    </form>
                                    <form method=post action="noticecontrol.asp">
                                      <td width="219">
                                        <select name="select" size=1>
                                          <%
for i=2001 to 2010
%>
                                          <option value=<%=i%><%=selected(i,myyear)%>><%=i%>年</option>
                                          <%
next
%>
                                        </select>
                                        <select name="select" size=1>
                                          <%
for i=1 to 12
%>
                                          <option value=<%=i%><%=selected(i,mymonth)%>><%=i%>月</option>
                                          <%
next
%>
                                        </select>
                                        <input type="hidden" name="lookfor2" value="yes">
                                        <input name="submit" type="submit" value="查询">
                                      </td>
                                    </form>
                                  </tr>
                                </table></td>
                                <td width="70"><div align="right"><a href="noticelook.asp"><img src="images/bt_tonggaoliebiao.gif" width="58" height="18" border="0"></a></div></td>
                                <td width="70"><div align="right"><a href="newnotice.asp"><img src="images/bt_fabutonggao.gif" width="58" height="18" border="0"></a></div></td>
                                <td width="70"><div align="right"><a href="noticecontrol.asp"><img src="images/bt_guanlitonggao.gif" width="58" height="18" border="0"></a></div></td>
                              </tr>
                            </table></td>
                          </tr>
                        </table>                        
                          <%     
if cook_allow_send_note="yes" or cook_allow_control_note="yes" then     
%>
                                                 <%
if myyear<>"" then
mydate=myyear & "-" & mymonth & "-" & 1
mydate1=dateadd("m",1,mydate)
else
mydate=""
end if


set conn=opendb("oabusy","conn","accessdsn")
Set rs=Server.CreateObject("ADODB.recordset")
sql="select * from newnotice order by id desc"
if request("lookfor")="yes" then sql="select * from newnotice where noticedate between " & "#" & mydate & "# and #" & mydate1 & "# order by id desc"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
rs.pagesize=20
page=request("page")
if not isnumeric(page) then
	page=1
end if
page=clng(page)
if err.number<>0 then
	page=1
end if
if page<1 then page=1
if page>rs.pagecount then page=rs.pagecount
href="noticecontrol.asp"
rs.absolutepage=page
%>
                                                 <form method="post" action="<%=href%>?page=<%=page%>">
      <table width="550"  border="0" cellspacing="0" cellpadding="0" align="center">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
          </table><table width="550" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="B0C8EA">
      <tr>
        <td height="2" colspan="3" align="center"></td>
        </tr>
      <tr bgcolor="D7E8F8">
        <td align="center"><span class="style8">选择</span></td>
        <td height="24"><span class="style8">　通告标题</span></td>
        <td width="136"><span class="style8">　发布时间</span></td>
      </tr><%
for ipage=1 to rs.pagesize
%>
      <tr>
        <td width="45" align="center" bgcolor="#FFFFFF">
          <input type="checkbox" name="delid" value=<%=rs("id")%>>
        </td>
        <td height="24" bgcolor="#FFFFFF"><a href="shownotice.asp?id=<%=rs("id")%>"><font color="#2b486a">　<%=rs("title")%></font></a><font color="FF0048">&nbsp; </font></td>
        <td bgcolor="#FFFFFF"><font color="#2b486a">　<%=rs("noticedate")%></font></td>
      </tr>
      <%
rs.movenext
if rs.eof then exit for
next
%>
    </table>
    <div align="center"><span class="style6"><br>
        <span class="style8">请选择已过期的通告进行删除</span></span>      
      <input type="submit" value="删 除 通 告" name="detel" onClick="return window.confirm('你确定要删除通告吗？');">
      <input type="hidden" name="myyear" value="<%=myyear%>">
      <input type="hidden" name="mymonth" value="<%=mymonth%>">
      <input type="hidden" name="lookfor" value="<%=request("lookfor")%>">
      <input type="hidden" name="page" value="<%=page%>">
    </div>
</form>
<table width=450 border=0 align="center">
<tr><form action=<%=href%> method=get>
<td align="center">
<%
response.write "<a href=" & href & "?page=1&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2b486a>第一页</font></a>"
%>
</td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & (page-1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2b486a>上一页</font></a>"
%>
</td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & (page+1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2b486a>下一页</font></a>"
%>
</td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & rs.pagecount & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2b486a>最后一页</font></a>"
%>
</td>
<input type="hidden" name="myyear" value="<%=myyear%>">
<input type="hidden" name="mymonth" value="<%=mymonth%>">
<input type="hidden" name="lookfor" value="<%=request("lookfor")%>">
<td align="center"><font color="#2b486a">第<%=page%>/<%=rs.pagecount%>页</font></td></form></tr></table>
<br>



<%
else
%>
<br><br><br>
<table width="400" border="0" align="center" cellpadding="0" cellspacing="0">
<tr>
      <td width="400" align="center"><span class="style8">对不起,没有相关纪录</span></td>
    </tr></table>
<p>
    <%
end if
%>
</p>
<p>
  <%
else
%>
</p>
<p align="center" class="style7"><br>
  对不起,您没有管理公司通告的权限</p>
<p>
  <%
end if
%>
</p></td>
                      </tr>
                    </table>                      </td>
                  </tr>
                </table>
                </center>
            </td>
          </tr>
        </table>
        <center>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>