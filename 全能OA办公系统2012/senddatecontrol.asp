<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/keepformat.asp"-->
<%
On Error Resume Next
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

set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from texttype where delflag=false"
rs.open sql,conn,1
if rs.eof or rs.bof then
	conn.close
	set rs= nothing
	response.redirect "documentaddtype.asp"
	response.end
end if
if request("detel")="删除" and request("delid")<>empty then
 set fso = Server.Createobject("Scripting.FileSystemObject")
 showid=split(trim(request("delid")),",")
 for a=0 to Ubound(showid)
   set rsd=server.createobject("adodb.recordset")
   sql="select * from senddate where id="&showid(a)
   rsd.open sql,conn,1,2
   if not rsd.eof then
       if trim(rsd("filename"))=null then
	   response.write ""
	   else
       delfile=split(rsd("filename"),",")
       patha=server.mappath("file/")
	   for i=0 to Ubound(delfile)
    	  fso.DeleteFile(patha&"/"&delfile(i))
	   next
	   end if
   end if
   rsd.close
   set rsd=nothing
 next
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
if request.form("detel")="删除" and request.form("delid")<>"" then
	count=0
	condition=""
	condition1=""
	condition2=""
	for each idno in request.form("delid")
		count=count+1
		condition=condition+"id=" & idno
		condition1=condition1+"reid=" & idno
		if count<request.form("delid").count then
			condition=condition+" or "
			condition1=condition1+" or "
		end if
	next
	set rs=server.createobject("adodb.recordset")
	sql="select * from senddate where " & condition1
	rs.open sql,conn,1
	while not rs.bof and not rs.eof
		condition2=condition2+"senddateid=" & rs("id")
		rs.movenext
		if not rs.bof and not rs.eof then condition2=condition2+" or "
	wend
	'删除数据库中的记录
	sql = "delete * from senddate where " & condition
	conn.Execute sql
	sql = "delete * from senddate where " & condition1
	conn.Execute sql
	sql = "delete * from seesenddate where " & condition2
	conn.Execute sql
end if
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta http-equiv="expires" content="no-cache">
<link rel="stylesheet" href="css/css.css">
<title>OA办公系统.边缘特别版</title>
<script language="javascript">
function openwin(href)
{
	window.open(href,'answerwin','location=no,height=450, width=640, toolbar=no, menubar=no, scrollbars=yes, resizable=yes, location=no, status=no');
}
</script>
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
<form method=post action="senddatecontrol.asp">
<td>公文管理&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="全部"></td>
</form>
<form method=post action="senddatecontrol.asp">
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
<input type="button" value="公文类型管理" onclick="location.href='documentaddtype.asp'">
</td>
</form>
</tr>
</table>
</center>

<br>
<center>
<%
myday1=myday+1
if myyear<>"" then
	mydate=myyear & "-" & mymonth & "-" & myday
	mydate1=dateadd("d",cdate(mydate),1)  'myyear & "-" & mymonth & "-" & myday1
else
	mydate=""
end if
set rs=Server.CreateObject("ADODB.recordset")
sql="select * from senddate,texttype where reid=0 and senddate.documenttype=texttype.number order by id desc"
if request("lookfor")="yes" then sql="select * from senddate,texttype where reid=0 and  senddate.documenttype=texttype.number and inputdate between " & "#" & mydate & "# and #" & mydate1 & "# order by id desc"
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
	href="senddatecontrol.asp"
	rs.absolutepage=page
%>
<form method="post" action="<%=href%>?page=<%=page%>">
<table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table width="540" border="0" cellpadding="0" cellspacing="1"  bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" bgcolor="B0C8EA">
<tr>
  <td height="2" colspan="5" align=center ></td>
  </tr>
<tr bgcolor="D7E8F8" >
<td width="42" align="center"><input name="detel" type="submit" onclick="return confirm('您真的要删除选定的公文吗？');" value="删除"></td>
        <td width="129" height="24" align=center ><span class="style8">公文标题</span></td>
        <td width=53 align=center><span class="style8">附件</span></td>
        <td width=59 align=center><span class="style8">发送者</span></td>
        <td width=92 align=center><span class="style8">接收者所在部门</span></td>
        <td width=59 align=center><span class="style8">接收者</span></td>
        <td width=98 align=center><span class="style8">发布日期</span></td>
</tr>
<%
	for ipage=1 to rs.pagesize
%>
<tr bgcolor="#ffffff">
<td align="center"><input type="checkbox" name="delid" value=<%=rs("id")%>></td><td bgcolor="#ffffff"><div align="center"><a href="javascript:openwin('showsenddate.asp?id=<%=rs("id")%>')"><%=keepformat(rs("title"))%></a><br>
    （<font color="#dd0000"><%=server.htmlencode(rs("typename"))%></font>）
</div></td>
<td width=53 bgcolor="#ffffff">
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
          <div align="center">
            <!--#include file="showfile.asp"-->
          </div></td>
<td align="center">
<%
	set rs2=Server.CreateObject("ADODB.recordset")
	sql="select name from userinf where username=" & sqlstr(rs("sender"))
	rs2.open sql,conn,1
	if not rs12.eof and not rs2.bof then
		response.write(rs2("name"))
	end if
%>
</td>
<td align="center"><%=rs("recipientuserdept")%></td><td align="center">
<%
	if rs("recipientusername")<>"所有人" then
		set rs1=Server.CreateObject("ADODB.recordset")
		sql="select name from userinf where username=" & sqlstr(rs("recipientusername"))
		rs1.open sql,conn,1
		if not rs1.eof and not rs1.bof then
			response.write(rs1("name"))
		end if
	else
		response.write(rs("recipientusername"))
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
<input type="hidden" name="myyear" value="<%=myyear%>">
<input type="hidden" name="mymonth" value="<%=mymonth%>">
<input type="hidden" name="myday" value="<%=myday%>">
<input type="hidden" name="lookfor" value="<%=request("lookfor")%>">
<input type="hidden" name="page" value="<%=page%>">
</form>
<table border=0 width=550>
<tr><form action=<%=href%> method=get>
<td align="center">
<%
response.write "<a href=" & href & "?page=1&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">第一页</a>"
%>
</td>
<td align="center">
  <div align="center">
      <%
response.write "<a href=" & href & "?page=" & (page-1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">上一页</a>"
%>
  </div></td>
<td align="center">
  <div align="center">
      <%
response.write "<a href=" & href & "?page=" & (page+1) & "&myyear=" & server.urlencode(myyear) & "&mymonth=" &  server.urlencode(mymonth) & "&myday=" & server.urlencode(myday) & "&lookfor=" & server.urlencode(request("lookfor")) & ">下一页</a>"
%>
  </div></td>
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
<td align="center">第<%=page%>/<%=rs.pagecount%>页</td></form></tr></table>
<br>
<%
else
%>
<br><br><br>
<table border="0" cellpadding="0" cellspacing="0" width="400">
<tr>
<td width="400" align="center"><font size="4" color="red">对不起,没有相关纪录</font></td></tr></table>
<%
end if
%>
</center>
<%
%>
          <tr>
      
    </tr>
<%
conn.close
set conn=nothing
%>
</body>
</html>