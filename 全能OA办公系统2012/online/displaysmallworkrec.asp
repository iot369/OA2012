<!--#include file="opendb.asp"-->
<!--#include file="sqlstr.asp"-->
<%
function displaysmallworkrec(myday,username,superior)
	displaysmallworkrec="<a href='displayworkrec.asp?username=" & username & "&superior=" & superior & "&recdate=" & myday & "' onclick='return js_openpage(this.href);'>" & Day(myday) & "</a>"
	if cdate(myday)=cdate(mydate) then
		displaysmallworkrec="<a href='displayworkrec.asp?username=" & username & "&superior=" & superior & "&recdate=" & myday & "' onclick='return js_openpage(this.href);'><font color=red><b>" & Day(myday) & "</b></font></a>"
	end if
	'打开数据库，读出日期为myday，用户名为username的记录
	set conn=opendb("oabusy","conn","accessdsn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from workrep where username=" & sqlstr(username) & " and recdate=#" & myday & "#"
	rs.open sql,conn,1
	if not rs.eof and not rs.bof then
		displaysmallworkrec=displaysmallworkrec & "<hr size=1 color=red width='80%'><table align=center>"
		while not rs.eof and not rs.bof
			if rs("finished")="yes" then
				if rs("imp")="yes" then
					titlecolorfront="<b><font color='#770000'>"
				else
					titlecolorfront="<font color='#0000ff'>"
				end if
                                if rs("imp")="yes" then
  				       titlecolorback="</font></b>"
                                else
                                       titlecolorback="</font>"
                                end if
			else
				if rs("imp")="yes" then
					titlecolorfront="<b><font color='#ff0000'>"
				else
					titlecolorfront="<font color='#cccccc'>"
				end if
                                if rs("imp")="yes" then
  				       titlecolorback="</font></b>"
                                else
                                       titlecolorback="</font>"
                                end if
				'titlecolorback="</font>"
			end if
			displaysmallworkrec=displaysmallworkrec & "<tr><td><font color=red>△</font>" & titlecolorfront & server.htmlencode(rs("title")) & titlecolorback & "</td></tr>"
rs.movenext
		wend
		displaysmallworkrec=displaysmallworkrec & "</table>"
	end if
end function
%>
