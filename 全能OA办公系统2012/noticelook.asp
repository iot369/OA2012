<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->
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

if request("myday")<>"" then myday=cint(request("myday"))
'取得当前月
mymonth=month(now())
if request("mymonth")<>"" then mymonth=cint(request("mymonth"))
'取得当前年
myyear=year(now())
if request("myyear")<>"" then myyear=cint(request("myyear"))
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="inc/style.css" type="text/css">
<title>OA办公系统.边缘特别版 公司通知</title>
<style type="text/css">
<!--
.style2 {
	color: #0d79b3;
	font-weight: bold;
}
.style3 {color: #098abb}
.td10 {
	height: 12px;
	border: 1px solid 098ABB;
	width: 14px;
}
.style4 {color: #0d79b3}
.style5 {color: #2d4865}
.style7 {color: #2b486a}
-->
</style>
</head>
<body topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
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
    </table></td>
  </tr>
  <tr>
    <td height="21"><table width="550"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td height="50"><table width="350" align="center">
          <tr>
            <td><span class="style3 style7">历史通告信息查询</span><span class="style7">&nbsp;&nbsp;</span>&nbsp;&nbsp; </td>
            <form method=post action="noticelook.asp">
              <td><input type="submit" name="submit" value="全部"></td>
            </form>
            <form method=post action="noticelook.asp">
              <td>
                <select name="myyear" size=1>
                  <%
for i=2001 to 2010
%>
                  <option value=<%=i%><%=selected(i,myyear)%>><%=i%>年</option>
                  <%
next
%>
                </select>
                <select name="mymonth" size=1>
                  <%
for i=1 to 12
%>
                  <option value=<%=i%><%=selected(i,mymonth)%>><%=i%>月</option>
                  <%
next
%>
                </select>
                <input type="hidden" name="lookfor" value="yes">
                <input name="submit2" type="submit" value="查询">
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
  <tr>
    <td><div align="center">
        <center>
          
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
if page<1 then page=1
if page>rs.pagecount then page=rs.pagecount
href="noticelook.asp"
rs.absolutepage=page
%>
          <table width="550"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
          </table>
          <table  border="0"  cellspacing="1" cellpadding="0" width="550" bgcolor="B0C8EA"  align="center">
            <tr>
              <td height="2" colspan="2"></td>
            </tr>
            <tr bgcolor="D7E8F8">
              <td height="24" bgcolor="D7E8F8">　<span class="style7">通告标题</span></td>
              <td height="20" bgcolor="D7E8F8">　<span class="style7">发布时间</span></td>
            </tr>
            <%
for ipage=1 to rs.pagesize
%>
            <tr bgcolor="#FFFFFF">
              <td width="411" height="24" bgcolor="#FFFFFF">　 <a href="shownotice.asp?id=<%=rs("id")%>" target="main_body"><font color="#2B486A"><%=rs("title")%></font></a> </td>
              <td width="136"><div align="left"></div>
              <div align="left"><font color="#2B486A">　<%=rs("noticedate")%></font></div></td>
            </tr>
            <%
rs.movenext
if rs.eof then exit for
next
%>
          </table>
          <br>
          <table border=0 width=450>
            <tr valign="middle">
              <form action=<%=href%> method=get>
                <td width="116" align="center">
                <%
response.write "<a href=" & href & "?page=1&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2B486A>第一页</font></a>"
%>                </td>
                <td width="93" align="center">
                <%
response.write "<a href=" & href & "?page=" & (page-1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2B486A>上一页</font></a>"
%>                </td>
                <td width="90" align="center">
                <%
response.write "<a href=" & href & "?page=" & (page+1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&lookfor=" & server.urlencode(request("lookfor")) & "><font color=#2B486A>下一页</font></a>"
%>                </td>
                <td width="133" align="center"><font color="#2B486A">第<%=page%>/<%=rs.pagecount%>页</font></td>
              </form>
            </tr>
          </table>
          <br>
          <%
else
%>
          <br>
          <br>
          <br>
          <table border="0" cellpadding="0" cellspacing="0" width="400">
            <tr>
              <td width="400" align="center"><span class="style7">对不起,没有相关纪录</span></td>
            </tr>
          </table>
          <%
end if
%>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>