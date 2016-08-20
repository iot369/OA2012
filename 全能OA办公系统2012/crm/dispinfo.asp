<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
typenumber=request("typenumber")
lookstr=request("lookstr")
if typenumber<>"" and lookstr<>"" then
	sql=""
	select case typenumber
		case "1"
			sql="select * from qiye where diqu='"&lookstr&"' order by id desc"
		case "2"
			sql="select * from qiye where companystyle='"&lookstr&"' order by id desc"
		case "3"
			if request.form("submit")=" 查询 " then
				dim conditionflag,sums
				dim sqlstr(9)
				sums=0
				conditionflag=false
				for i=1 to 9
					sqlstr(i)=""
				next
				for i=1 to 7
					findstr=request.form("C"&cstr(i))
					if findstr="ON" then
						select case i
							case 1
								fieldname="企业名称"
							case 2
								fieldname="contact"
							case 3
								fieldname="production"
							case 4
								fieldname="address"
							case 5
								fieldname="phone"
							case 6
								fieldname="fax"
							case 7
								fieldname="postcode"
						end select
						sqlstr(i)=fieldname&" like '%"&request.form("T"&cstr(i))&"%'"
						sums=sums+1
						conditionflag=true
					end if
				next
				if request.form("C8")="ON" and request.form("D1")<>"" then
					conditionflag=true
					sqlstr(8)="diqu='"&request.form("D1")&"'"
					sums=sums+1
				end if
				if request.form("C9")="ON" and request.form("D2")<>"" then
					conditionflag=true
					sqlstr(9)="companystyle='"&request.form("D2")&"'"
					sums=sums+1
				end if
				if conditionflag then
					sql="select * from qiye where "
					for i=1 to 9
						if sums=1 and sqlstr(i)<>"" then
							sql=sql&sqlstr(i)
						elseif sqlstr(i)<>"" then
							if i=1 then
								sql=sql&sqlstr(i)
							elseif sql<>"select * from qiye where " then
								sql=sql&" and "&sqlstr(i)
							else
								sql=sql&sqlstr(i)
							end if
						end if
					next
					response.cookies("findcondiction")=sql
				end if
			else
				sql=request.cookies("findcondiction")
			end if
	end select
	if sql="" then
		response.write("<script language=""javascript"">")
		response.write("alert(""请至少选择一个条件！"");")
		response.write("history.go(-1);")
		response.write("</script>")
		response.end
	elseif sql<>"" then
		set conn=dbconn("conn")
		set rs=server.createobject("adodb.recordset")
		rs.open sql,conn,1
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><%=server.htmlencode(lookstr)%>销售管理系统</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
<script src="openwin.js"></script>
<script language="vbscript">
sub checkkey()
    if window.event.keyCode  >57 or window.event.keyCode <48  then 
		window.event.keyCode=0
	end if    
end sub
</script>
</head>
<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<br>
<p align="center"><b><font size="4"><%=server.htmlencode(lookstr)%>企业名录<br>
</font></b>
<%
		if rs.eof and rs.bof then
			response.write("<br><br><center><font color=""red"" size=""+1"">对不起，该类别中没有企业！</font></center>")
			conn.close
			set rs=nothing
			set conn=nothing
		else
			rs.pagesize     = 10
			maxpages        = rs.pagecount
			if maxpages=0 then
				maxpages=1
			end if
			pagenumber=request("page")
			if pagenumber="" then
				page=1
			elseif not isnumeric(pagenumber) then
				page=1
			else
				page=clng(pagenumber)
			end if
			if page<=0 then
				page=1
			elseif page>maxpages then
				page=maxpages
			end if
			rs.absolutepage = page
			total           = rs.recordcount
%>
<div align="center">
  <center>
  <table border="0" cellpadding="0" cellspacing="0" width="450" height="40">
    <tr>
      <td>
<%
	if maxpages>1  then
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page=1"">首&nbsp;&nbsp页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&page-1&chr(34)&">上一页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&page+1&chr(34)&">下一页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&maxpages&chr(34)&">尾&nbsp;&nbsp页</a>&nbsp;&nbsp;&nbsp;")
	end if
%>共<font color="blue"><%=total%></font>条记录&nbsp;&nbsp;页码：<font color="blue"><%=page%></font>/<font color="blue"><%=maxpages%></font>&nbsp;&nbsp;第<input type="text" name="T1" size="20" style="width: 26; height: 22" class="doc_txt" onkeypress="vbscript:checkkey()">页<input type="button" value="Go" name="B3" onclick="javascript:location.href='dispinfo.asp?typenumber=<%=typenumber%>&lookstr=<%=lookstr%>&page='+T1.value">
</td>
    </tr>
  </table>
  </center>
</div>
<%
	if total>0 then
		for ipage=1 to rs.pagesize
%>
<div align="center">
  <center>
  <table border="1" cellpadding="0" cellspacing="0" width="450" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr>
      <td colspan="3" height="25" bgcolor="#EFEFEF" width="446"><a href="#" onclick ="javascript:openwinfun('dispcontent.asp?id=<%=cstr(rs("id"))%>&printflag=0','qymlwin',400,550);"><font color="red"><%=server.htmlencode(trim(rs("企业名称")))%></font></a></td>
    </tr>
    <tr>
      <td colspan="3" height="25" width="446">地址：
<%
		if not isnull(rs("address")) then
			response.write(server.htmlencode(trim(rs("address"))))
		end if
%>
</td>
    </tr>
    <tr>
      <td height="25" width="149">联系人：
<%
		if not isnull(rs("contact")) then
		  	response.write(server.htmlencode(trim(rs("contact"))))
		end if
%>
	  </td>
      <td height="25" width="149">邮政编码：
<%
		if not isnull(rs("postcode")) then
			response.write(server.htmlencode(trim(rs("postcode"))))
		end if
%>
	  </td>
      <td height="25" width="150">电话：
<%
		if not isnull(rs("phone")) then
			response.write(server.htmlencode(trim(rs("phone"))))
		end if
%>
	  </td>
    </tr>
  </table>
      <td colspan="3" height="25" bgcolor="#EFEFEF" width="446"><a href="#" onclick ="javascript:openwinfun('delcontent.asp?id=<%=cstr(rs("id"))%>&printflag=0','qymlwin',400,550);"><font color="red"><%=server.htmlencode("删除此企业记录")%></font></a></td>
  </center>
</div>
<br>
<%
			rs.movenext
			if rs.eof then exit for
		next
	end if
%>
<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="450" height="40">
    <tr>
      <td>
<%
	if maxpages>1  then
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page=1"">首&nbsp;&nbsp页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&page-1&chr(34)&">上一页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&page+1&chr(34)&">下一页</a>&nbsp;&nbsp;&nbsp;")
		response.write("<a href=""dispinfo.asp?typenumber="&typenumber&"&lookstr="&lookstr&"&page="&maxpages&chr(34)&">尾&nbsp;&nbsp页</a>&nbsp;&nbsp;&nbsp;")
	end if
%>共<font color="blue"><%=total%></font>条记录&nbsp;&nbsp;页码：<font color="blue"><%=page%></font>/<font color="blue"><%=maxpages%></font>&nbsp;&nbsp;第<input type="text" name="T2" size="20" style="width: 26; height: 22" class="doc_txt" onkeypress="vbscript:checkkey()">页<input type="button" value="Go" name="B3" onclick="javascript:location.href='dispinfo.asp?typenumber=<%=typenumber%>&lookstr=<%=lookstr%>&page='+T2.value">
</td>
    </tr>
  </table>
</div>
</body>
</html>
<%
			conn.close
			set rs=nothing
			set conn=nothing
		end if
	end if
else
	
end if
%>
