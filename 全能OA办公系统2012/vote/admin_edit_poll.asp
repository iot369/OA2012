
<!--#include file="chkadmin.asp"-->
<!--#include file="config.asp"-->
<%
call UpdatePoll()
%>

<!--#include file="top.asp"-->
<!--#include file="menu.asp"-->
<br><br>
<%
Dim pID		
Dim Qrs	
Dim Ars	
Dim i		
pID = CINT(Request("id"))
call OpenDB()
Set Qrs = DbConn.Execute("SELECT * FROM Question WHERE Q_ID = " & pID )
If Qrs.EOF Then
	Response.Redirect ("admin_poll.asp")
End IF
Set Ars = DbConn.Execute("SELECT * FROM Answer WHERE Q_ID = " & pID )
If Ars.EOF Then
	Response.Redirect ("admin_poll.asp")
End IF
%>
<form action="" method="post" name="form1">
        
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#7C96B8">
    <tr > 
      <td colspan="2"><font color="#FF0000">修改调查项目</font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="right" width="22%">调查主题：</td>
      <td> 
        <input name="title" type="text" value="<%=Qrs("Q_Title")%>" size="30" maxlength="100">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="22%" align="right">开始时间：</td>
      <td> 
        <input name="startt" type="text" value="<%=Qrs("Q_StartDate")%>" size="8">
        (如：2003-1-1)</td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="22%" align="right">结束时间：</td>
      <td> 
        <input name="endt" type="text" value="<%=Qrs("Q_EndDate")%>" size="8">
        (如：2003-1-30)</td>
    </tr>
    <%i = 0%>
    <%Do until Ars.EOF
    i = i + 1%>
    <tr bgcolor="#FFFFFF"> 
      <td width="22%" align="right">调查选项<%=i%>：</td>
      <td width="61%"><%=Ars("A_Answer")%> <a href="admin_edit_option.asp?id=<%=Qrs("Q_ID")%>&aid=<%=Ars("A_ID")%>&option=<%=i%>"><img src="images/edit.gif" width="15" height="13" border="0" vspace="2" hspace="2" align="absmiddle"></a><a href="admin_edit_poll.asp?T=DELA&aid=<%=Ars("A_ID")%>&pid=<%=pID%>"><img src="images/delete.gif" width="15" height="13" border="0" vspace="2" hspace="2" align="absmiddle"></a></td>
    </tr>
    <%
	Ars.MoveNext
	Loop
	%>
    <tr bgcolor="#FFFFFF"> 
      <td align="right">当前调查项目？</td>
      <td> 
        <input type="radio" name="active" value="1" <%if Qrs("Q_Active") Then response.write "checked"%>>
        是 
        <input name="active" type="radio" value="0"  <%if not Qrs("Q_Active") Then response.write "checked"%>>
        否 </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center" colspan="2"> 
        <input type="hidden" name="pid" value="<%=pID%>">
        <input type="submit" name="Submit" value="更 新">
        <input type="button" name="add" value="添加新选项" onClick="window.location.href='admin_add_option.asp?id=<%=pID%>';">
      </td>
    </tr>
  </table>
</form>
<%
Qrs.Close
Ars.Close
Set Qrs = Nothing
Set Ars = Nothing
call CloseDB()
%>

<!--#include file="foot.asp"-->

<%
Sub UpdatePoll()
	T = RequestText(Request.Querystring("T"))
	aID = RequestText(Request("aid"))
	pID = RequestText(Request("pid"))
	if T="DELA" and aID<>"" then
		call OpenDB()
		set rs=DbConn.Execute("Select Q_ID,Q_Vote From Answer WHERE A_ID = " & aID )
		DbConn.Execute("UPDATE Question SET Q_Vote=Q_Vote-"&rs("Q_Vote")&" Where Q_ID="&rs("Q_ID"))
		rs.close
		set rs=nothing
		DbConn.Execute("Delete * From Answer WHERE A_ID = " & aID )
		call CloseDB()
		Response.Redirect ("admin_edit_poll.asp?id=" & pID)
	end if

If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	title = RequestText(Request.Form("title"))
	if title="" then out "调查主题不能为空"
	if not ISDATE(RequestText(Request.Form("startt"))) or not isdate(RequestText(Request.Form("endt"))) then
		out "日期格式不对。"
	end if 
	startt = CDATE(RequestText(Request.Form("startt")))
	endt = CDATE(RequestText(Request.Form("endt")))
	active = RequestText(Request.Form("active"))
	call OpenDB()
	If active = 1 Then
	DbConn.Execute("UPDATE Question SET Q_Active=0")
	End If
	DbConn.Execute("UPDATE Question SET Q_Title = '" & title & "', Q_StartDate = '" & startt & "',Q_EndDate = '" & endt & "', Q_Active = '" & active & "' WHERE Q_ID = " & pID )
call CloseDB()

Response.Redirect "admin_poll.asp"
end if
End Sub
%>
