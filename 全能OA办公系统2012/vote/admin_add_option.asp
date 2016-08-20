
<!--#include file="chkadmin.asp"-->
<!--#include file="config.asp"-->
<%
Call AddOption()
%>
<!--#include file="top.asp"-->
<!--#include file="menu.asp"-->
<%
Dim pID	
Dim rs		
Dim choice	
pID = CINT(Request.QueryString("id"))
call OpenDB()
Set rs = DbConn.Execute("SELECT p.Q_Title, p.Q_StartDate, p.Q_EndDate, a.A_Answer FROM Question p, Answer a WHERE p.Q_ID = a.Q_ID AND p.Q_ID = " & pID)
If rs.EOF Then
	Response.Redirect ("admin_poll.asp")
End If
%>
    
<form action="" method="post" >
        
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#7C96B8">
    <tr > 
      <td colspan="2"><font color="#FFFFFF"><img src="images/poll.gif" width="13" height="15" align="absmiddle" vspace="2" hspace="2"><font color="#FF0000">调查主题：</font><%=rs("Q_Title")%></font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td width="22%" align="center">开始时间：</td>
      <td width="78%"><%=rs("Q_StartDate")%></td>
    </tr>
    <tr  bgcolor="#FFFFFF"> 
      <td align="center" width="22%">结束时间：</td>
      <td align="left" width="78%"><%=rs("Q_EndDate")%></td>
    </tr>
    <%choice = 0%>
    <%Do until rs.EOF
    choice = choice + 1%>
    <tr  bgcolor="#FFFFFF"> 
      <td align="center" width="22%">调查选项<%=choice%>：</td>
      <td align="left" width="78%"><%=rs("A_Answer")%></td>
    </tr>
    <%
	rs.MoveNext
	Loop
	
	rs.Close
	Set rs = Nothing
	%>
    <tr bgcolor="#FFFFFF"> 
      <td align="center" width="22%">调查选项<%=choice + 1%>：</td>
      <td width="78%"> 
        <input name="content" type="text" size="20" maxlength="50">
      </td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td align="center" width="22%"><font color="#EFEFEF">设 置：</font></td>
      <td width="78%"> <font color="#EFEFEF">
        <input type="radio" name="G" value="T" checked>
        添加下一个 
        <input type="radio" name="G" value="F">
        提交后返回 </font></td>
    </tr>
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" align="center"> 
        <input type="hidden" name="pID" value="<%=pID%>">
        <input name="action" type="submit" id="action" value="提 交">
      </td>
    </tr>
  </table>
</form>
   
<!--#include file="foot.asp"-->
<%
Sub AddOption()
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	content = RequestText(Request.Form("content"))
	G = RequestText(Request.Form("G"))
	pID = RequestText(Request.Form("pID"))
	if content="" then out "选项内容不能为空"
	call OpenDB()
	DbConn.Execute("Insert into Answer (Q_ID,A_Answer) values('" & pID & "','" & content & "')")
	call CloseDB()
	if G="T" then
		Response.Redirect "admin_add_option.asp?id="&pID
	else
		Response.Redirect "admin_edit_poll.asp?id="&pID
	end if
end if
End Sub
%>
