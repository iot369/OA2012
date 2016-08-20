<%
'session.abandon
'Server.ScriptTimeOut=500
function opendb(DBPath,sessionname,dbsort)
dim conn
'if not isobject(session(sessionname)) then
Set conn=Server.CreateObject("ADODB.Connection")
'if dbsort="accessdsn" then conn.Open "DSN=" & DBPath
'if dbsort="access" then conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
'if dbsort="sqlserver" then conn.Open "DSN=" & DBPath & ";uid=wsw;pwd=wsw"
DBPath1=server.mappath("../db/lmtof.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<!--#include file="config.asp" -->
<%
Q_ID = RequestText(Request("Q_ID"))
strF=" order by Q_Active"
if Q_ID<>"" and IsNumeric(Q_ID) then strF=" and Q_ID=" & Q_ID
Call UpdatePoll()
%>

<table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
              <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">公共服务</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
<!--#include file="top.asp" -->
<%
Call OpenDB()
Set rs = Dbconn.Execute("SELECT Q_ID,Q_Title,Q_Vote FROM Question WHERE Q_StartDate <= #" & Date() & "# AND Q_EndDate >= #" & Date() & "# "& strF)

If Not rs.EOF Then 
	N_Quest = rs("Q_Title")
	Q_ID = Cint(rs("Q_ID"))
	PollNum = Cint(Request.Cookies("poll")("s"&Q_ID))

	sql = "SELECT A_ID,A_Answer,Q_Vote FROM Answer WHERE Q_ID=" & Q_ID
	Set ars = Server.CreateObject("ADODB.Recordset")
	ars.Open sql, DBConn, 1, 1

	C=ars.recordcount
	ReDim A_Count(c)
		i=65
		p_total=0
		For N=1 to C
			A_Count(N)=ars("Q_Vote")
			aID = ars("A_ID")
			answer = ars("A_Answer")
			p_total=p_total+A_Count(N)
        	Content1=Content1 & "[" & chr(i) & "]<input type=""radio"" name=""poll"" value=""" & aID & """>"
        	Content1=Content1 & answer & "<br>"
			ars.MoveNext
			i=i+1
		Next
	ars.Close
	Set ars = Nothing
	rs.Close
	Set rs = Nothing
	i=65
	For N=1 to C
		if A_Count(N) = 0 then
			p_percent = 0
		Else
			p_percent = (A_Count(N)/p_total) * 100
		End If
		
		Content20=Content20 & "<tr><td align=""right"" valign=""bottom"">[" & chr(i) & "]</td></tr>"
	 	Content21=Content21 & "<tr><td valign=""bottom""> <img src=""images/p1.gif"" width=""" & p_percent & """ height=""8"">&nbsp; <font color=""#7C96B8"">" & FormatNumber(p_percent,2) & "%&nbsp;[" & A_Count(N) & "人]</font></td></tr>"

		Content30=Content30 & "<td width=""40"" valign=""bottom"" align=""center""><font color=""#7C96B8"">" & FormatNumber(p_percent,2) & "%</font><br><img src=""images/p2.gif"" width=""8"" height=""" & p_percent & """ vspace=""1""></td>"
		Content31=Content31 & "<td width=""40"" align=""center"">[" & chr(i) & "]</td>"
		i=i+1
	Next
%>
	<form action="default.asp" method="post">
	<input type="hidden" name="Q_ID" value="<%=Q_ID%>">
	<input type="hidden" name="N_Question" value="<%=N_Quest%>">

  <table width="98%" cellspacing="1" cellpadding="3" border="0" bgcolor="#7C96B8" align="center">
    <tr> 
      <td colspan="2"><font color="#000000">当前调查项目</font></td>
    </tr>
	</table>
	 
  <table width="98%" cellspacing="1" cellpadding="3" border="0" bgcolor="#7C96B8" align="center">
    <tr bgcolor="#F2F2F2"> 
      <td width="49%"> 
        <font color="#000000">主题：<%=N_Quest%> </font>
      </td>
      <td width="51%"><font color="#000000">[调查结果显示]</font> 
        <input type="radio" name="list" onclick="L1.style.display='inline';L2.style.display='none';" checked>
        横向显示
<input type="radio" name="list" onclick="L2.style.display='inline';L1.style.display='none';">纵向显示
</td>
    </tr>
    <tr> 
      <td width="49%" bgcolor="#FFFFFF" valign="top"> 
        <%response.write Content1 %>
      </td>
      <td width="51%" bgcolor="#FFFFFF"> 
        <table width="100%" >
          <tr> 
            <td width="48%" class="small" id="L1" height="120"> 
              <table width="100%" border="0" cellspacing="0" cellpadding="1">
                <tr> 
                  <td valign="bottom" width="22"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="120">
                      <tr>
                        <td align="right" valign="bottom" height="18">&nbsp;</td>
                      </tr>
                      <%response.write Content20 %>
                      <tr> 
                        <td height="18">&nbsp; </td>
                      </tr>
                    </table>
                  </td>
                  <td class="L1" valign="bottom"> 
                    <table width="100%" border="0" cellspacing="0" cellpadding="0" height="120">
                      <tr>
                        <td valign="bottom" height="18">&nbsp;</td>
                      </tr>
                      <%response.write Content21%>
                      <tr> 
                        <td height="18">&nbsp; </td>
                      </tr>
                    </table>
                  </td>
                </tr>
              </table>
            </td>
            <td width="52%" class="small" valign="bottom" id="L2" style="display:none;">
              <table border="0" cellspacing="0" cellpadding="0" class="L2">
                <tr> 
				  <td valign="bottom" align="center" height="120" width="19"> 
                  </td>
                  <%response.write Content30 %>
                  <td valign="bottom" align="center" height="120" width="20"> 
                  </td>
                </tr>
			  </table>
				 
              <table border="0" cellspacing="0" cellpadding="0" align="left">
                <tr>
				  <td valign="bottom" align="center"  width="19" height="18"> 
                  </td>
				   <%response.write Content31%>
                </tr>
              </table>
            </td>
          </tr>
        </table>
      </td>
    </tr>
	<tr bgcolor="#F2F2F2"> 
      <td width="49%" bgcolor="#F2F2F2" valign="bottom" height="2"> 
        <input type="submit" value="提交" <%If PollNum = Q_ID Then response.write "disabled"%>>
      </td>
      <td width="51%" height="2">总投票数: <%=p_total%> </td>
    </tr>
  </table>
</form>
  

<%
Dim PageNo	
PageNo = 10
Dim PageNum		
Dim RCount
Dim mpage		
Dim i
			
Dim sql	
sql = "Select Q_ID, Q_Title, Q_StartDate,Q_Vote,Q_Active FROM Question WHERE Q_StartDate <= #" & Date() & "# AND Q_EndDate >= #" & Date() & "# and Q_ID <> " & Q_ID
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open sql, DBConn, 1, 1

%>
<br>
<table width="98%" cellspacing="1" cellpadding="3" border="0" bgcolor="#7C96B8" align="center">
  <tr> 
      
    <td colspan="2"><font color="#000000">其他调查项目</font></td>
  </tr>
</table>
      
<table width="98%" border="0" cellspacing="1" cellpadding="1" align="center">
  <%if rs.eof then %>
  <tr> 
    <td colspan="4" class="small" align="center" bgcolor="#FFFFFF">没有其他调查项目了</td>
  </tr>
  <%
  else
  call CheckPage()
  %>
  <tr> 
    <td width="3%">&nbsp;</td>
    <td width="57%" height="20" >调查主题</td>
    <td width="8%" align="center">参加人数</td>
    <td width="32%">开始日期</td>
  </tr>
  <%
rs.AbsolutePage = PageNum
i=1
For RCount = 1 To PageNo
%>
  <tr> 
    <td width="3%" bgcolor="#FFFFFF" align="center"><img src="images/<%if rs("Q_Active") then response.write "poll" else response.write "poll1"%>.gif" width="13" height="15"> 
    </td>
    <td width="57%" bgcolor="#FFFFFF"><a href="default.asp?Q_ID=<%=rs("Q_ID")%>"> 
      <%=rs("Q_Title")%></a> </td>
    <td width="8%" bgcolor="#FFFFFF" align="center"><%=rs("Q_Vote")%></td>
    <td width="32%" bgcolor="#FFFFFF">[<%=rs("Q_StartDate")%>]</td>
  </tr>
  <%
		i=i+1
rs.MoveNext
If rs.EOF Then Exit For
Next
end if
%>
</table>
<%
if mpage>1 then
%>
<table width="98%" border="0" cellspacing="1" cellpadding="2" align="center" >
  <tr> 
    <td> 
      <%call DisplayPage()%>
    </td>
  </tr>
</table>
<%
end if
rs.Close
Set rs = Nothing
Call CloseDB()
%>
<%
Else
%>
<table width="98%" cellspacing="0" cellpadding="3" border="0" align="center">
  <tr>
    <td bgcolor="#C9CFCD"> 调查项目 </td>
</tr>
<tr>
    <td align="center"> 
      <%Response.Write ("对不起，此调查项目不存在。")%>
      <br>
</td>
</tr>
</table>
<%
rs.Close
Set rs=Nothing
End IF
%>
<!--#include file="foot.asp" -->
<%
Sub UpdatePoll()
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	aID = Request.Form("poll")
	N_Quest = Request.Form("N_Question")
	PollNum = Request.Cookies("poll")("s" & Q_ID)
	If Q_ID = "" or aID = "" Then out "请选择一个选项!"
	if PollNum=Q_ID then
		out "对不起，此项目你已经投过票了。"
	else
		Call OpenDB()
		DbConn.Execute("Update Answer SET Q_Vote = Q_Vote+1 Where A_ID = " & aID)
		DbConn.Execute("Update Question SET Q_Vote = Q_Vote+1 Where Q_ID = " & Q_ID)
		Call CloseDB()
		Response.Cookies("poll")("s" & Q_ID) = Q_ID
		Response.Cookies("poll").Expires = Date + 365
		response.redirect ("default.asp?Q_ID=" & Q_ID)
	end if
end if
End Sub
%>