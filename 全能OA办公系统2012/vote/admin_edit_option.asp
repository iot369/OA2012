
<!--#include file="chkadmin.asp"-->
<!--#include file="config.asp"-->
<%
Call UpdateOption()
%>
<!--#include file="top.asp"-->
<!--#include file="menu.asp"-->
<%
Dim pID
Dim aID	
Dim rs		
aID = CINT(Request.QueryString("aid"))
pID = CINT(Request.QueryString("id"))

call OpenDB()
Set rs = DbConn.Execute("SELECT p.Q_Title, p.Q_StartDate, p.Q_EndDate, a.A_Answer FROM Question p, Answer a WHERE p.Q_ID = a.Q_ID AND p.Q_ID = " & pID & " AND a.A_ID = " & aID )
If rs.EOF Then
	Response.Redirect ("admin_poll.asp")
End If
%>
<form action="" name="form1" method="post">
        
  <table width="98%" border="0" align="center" cellpadding="5" cellspacing="1" bgcolor="#7C96B8">
    <tr > 
      <td colspan="2" ><img src="images/poll.gif" width="13" height="15" align="absmiddle" vspace="2" hspace="2"><font color="#FF0000">调查主题：</font><%=rs("Q_Title")%></td>
          </tr>
          
    <tr bgcolor="#FFFFFF"> 
      <td width="21%" align="center">开始时间：</td>
            
      <td width="79%"><%=rs("Q_StartDate")%></td>
          </tr>
          
    <tr bgcolor="#FFFFFF"> 
      <td width="21%" align="center">结束时间：</td>
            
      <td width="79%"><%=rs("Q_EndDate")%></td>
          </tr>
          
    <tr bgcolor="#FFFFFF"> 
      <td align="center">调查选项<%=Request.QueryString("option")%>： </td>
            
      <td> 
        <input name="content" type="text" value="<%=rs("A_Answer")%>" size="20" maxlength="50">
            </td>
          </tr>
          
    <tr bgcolor="#FFFFFF"> 
      <td colspan="2" align="center"> 
        <input type="hidden" name="aID" value="<%=aID%>">
		 <input type="hidden" name="pID" value="<%=pID%>">
        <input name="action" type="submit" id="action" value="更新选项">
            </td>
          </tr>
        </table>
</form>
<%call CloseDB()%>

<!--#include file="foot.asp"-->
<%
Sub UpdateOption()
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	aID = RequestText(Request.Form("aID"))
	pID = RequestText(Request.Form("pID"))
	content = RequestText(Request.Form("content"))
	if content="" then out "选项内容不能为空"
	call OpenDB()
	DbConn.Execute("UPDATE Answer SET A_Answer = '" & content & "' WHERE A_ID = " & aID )
call CloseDB()
Response.Redirect ("admin_edit_poll.asp?id=" & pID)
end if
End Sub
%>
