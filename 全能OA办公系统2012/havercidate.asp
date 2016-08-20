<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/keepformat.asp"-->
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
'取得当前日
myday=day(now())
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
<meta http-equiv="expires" content="no-cache">
<link rel="stylesheet" href="css/css.css">
<title>已收公文</title>
<script language="javascript">
function openwin(href)
{
	window.open(href,'answerwin','location=no,height=450, width=640, toolbar=no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no');
}
</script>
<style type="text/css">
<!--
.style1 {color: #2b486a}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body bgcolor="#ffffff" topmargin="0" leftmargin="0">

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
                    <td class="style7">公文传阅</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <table width="1%"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td>&nbsp;</td>
    </tr>
  </table>
  <table>
<tr>
<form method=post action="havercidate.asp">
<td><input type="submit" name="submit" value="全部"></td>
</form>
<form method=post action="havercidate.asp">
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
<select name="myday" size=1>
<%
for i=1 to 31
%>
<option value=<%=i%><%=selected(i,myday)%>><%=i%>日</option>
<%
next
%>
</select>
<input type="hidden" name="lookfor" value="yes">
<input type="submit" value="查询">
</td>
</form>
</tr>
</table>
</center>

<center>
<br>
<%
myday1=myday+1
if myyear<>"" then
	mydate=myyear & "-" & mymonth & "-" & myday
	mydate1=dateadd("d",cdate(mydate),1)
else
	mydate=""
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from senddate,texttype where (senddate.recipientusername=" & sqlstr(oabusyusername) & " or (senddate.recipientusername='所有人' and senddate.recipientuserdept=" & sqlstr(oabusyuserdept) & ")) and senddate.sender<>" & sqlstr(oabusyusername) & " and senddate.reid=0 and senddate.documenttype=texttype.number order by id desc"
if request("lookfor")="yes" then sql="select * from senddate,texttype where (senddate.recipientusername=" & sqlstr(oabusyusername) & " or (senddate.recipientusername='所有人' and senddate.recipientuserdept=" & sqlstr(oabusyuserdept) & ")) and senddate.sender<>" & sqlstr(oabusyusername) & " and senddate.reid=0 and senddate.documenttype=texttype.number and inputdate between " & "#" & mydate & "# and #" & mydate1 & "# order by id desc"
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
		href="havercidate.asp"
		rs.absolutepage=page
%>
 <table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table width="540" border="0" cellpadding="0" cellspacing="1" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="B0C8EA">
<tr>
  <td height="2" colspan="5" align=center ></td>
  </tr>
<tr bgcolor="D7E8F8" height="25" >
<td width="252" height="24" align=center><span class="style1">公文标题</span></td>
<td width=58 align=center><div align="center" class="style1">附件</div></td>
<td width=99 align=center><span class="style1">发送者所在部门</span></td>
<td width=50 align=center><span class="style1">发送者</span></td>
<td width=75 align=center><span class="style1">发送日期</span></td>
</tr>
<%
		for ipage=1 to rs.pagesize
%>
<tr height="25" bgcolor="#ffffff">
<td align="center">
<a href="javascript:openwin('showdate.asp?id=<%=rs("id")%>');">
<%
			response.write(keepformat(rs("title")))
			if year(rs("inputdate"))=year(date()) and month(rs("inputdate"))=month(date()) and day(rs("inputdate"))=day(date()) then
				response.write("<img src=""image/new.gif"" border=""0"">")
			end if
%>
</a>
<br>
（<font color="#ee0000"><%=rs("typename")%></font>）
</td>
<td align="center">
        <!--  <%
if rs("filename")<>"" then
%>
        <a href="../listsendfile.asp?id=<%=rs("id")%>" target="_blank"><img src="../images/attach.gif" width=30 height=30 border=0 alt="文件名：<%=rs("filename")%>"></a> 
        <%
else
%>
        &nbsp; 
        <%
end if
%>-->
        <!--#include file="showfile.asp"-->
      </td><td align="center">
<%
			set rs1=Server.CreateObject("ADODB.recordset")
			sql="select name,userdept from userinf where username=" & sqlstr(rs("sender"))
			rs1.open sql,conn,1
			if not rs1.eof and not rs1.bof then
				response.write(rs1("userdept"))
			end if
%>
</td>
<td align="center">
<%
			if not rs1.eof and not rs1.bof then
				response.write(rs1("name"))
			end if
%>
</td>
<td align="center"><%=rs("inputdate")%></td>
</tr>
<%
			rs.movenext
			if rs.eof then exit for
next
%>
</table>
<table border=0 width=540>
<tr><form action=<%=href%> method="get">
<td height="24" align="center">
<%
response.write "<a href=" & href & "?page=1&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">第一页</a>"
%></td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & (page-1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">上一页</a>"
%>
</td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & (page+1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">下一页</a>"
%>
</td>
<td align="center">
<%
response.write "<a href=" & href & "?page=" & rs.pagecount & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">最后一页</a>"
%>
</td>
<td align="center">&nbsp;
</td>
<input type="hidden" name="myyear" value="<%=myyear%>">
<input type="hidden" name="mymonth" value="<%=mymonth%>">
<input type="hidden" name="myday" value="<%=myday%>">
<input type="hidden" name="lookfor" value="<%=request("lookfor")%>">
<td align="center">第<%=page%>/<%=rs.pagecount%>页</td></form></tr></table><br>
<%
else
%>
<br><br><br>
<table border="0" cellpadding="0" cellspacing="0" width="400">
<tr>
      <td width="400" align="center"><font size="3" color="red">对不起,没有相关纪录</font></td>
    </tr></table>
<%
end if
%>
</center>
<%
%>
         </td>
            
          
          </tr>
        </table>
    </tr>
    <tr>
      
  <td height=19>&nbsp; </td>
    </tr>
  </table>
<%
conn.close
set conn=nothing
%>
</body>
</html>