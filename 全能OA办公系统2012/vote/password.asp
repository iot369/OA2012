
<!--#include file="chkadmin.asp"-->
<!--#include file="config.asp"-->
<%
Call ChangePassword()
%>
<!--#include file="top.asp"-->
<!--#include file="menu.asp"-->
<% 
dim rs		
call OpenDB()
set rs = DbConn.Execute("select top 1 Adminname, Adminpassword from admin")
%>
<form action="" method="post" name="form1">
 <TABLE align="center" cellSpacing=0 cellPadding=0 width="98%" border=0>
     <TR> 
        <TD vAlign=top width="100%" height="184"> 
        <TABLE cellSpacing=1 cellPadding=5 width="100%" border=0 bgcolor="#7C96B8">
          <TR > 
            <TD colspan="2"><font color="#FF0000">修改密码 </font></TD>
           </TR>
          <TR bgcolor="#FFFFFF"> 
            <TD width="26%" align="right">用户名：</TD>
            <TD width="74%"> 
              <input name="username" type="text" value="<%=rs("Adminname")%>" size="12" maxlength="20">
                  </TD>
                </TR>
          <TR bgcolor="#FFFFFF"> 
            <TD width="26%" align="right">密码：</TD>
            <TD width="74%"> 
              <input name="password" type="password" value="<%=rs("Adminpassword")%>" size="12" maxlength="20">
                  </TD>
                </TR>
          <TR bgcolor="#FFFFFF"> 
            <TD width="26%" align="right">较验密码：</TD>
                  
            <TD width="74%"> 
              <input name="password2" type="password" value="<%=rs("Adminpassword")%>" size="12" maxlength="20">
                  </TD>
                </TR>
          <TR align="center" bgcolor="#FFFFFF"> 
            <TD colspan="2"> 
              <input type="submit" name="Submit" value="确定">
                  </TD>
                </TR>
              </TABLE>
              </TD>
            </TR>
      </TABLE>
</form>
<%
call CloseDB()
%>

<!--#include file="foot.asp"-->
<%
Sub ChangePassword()
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	adminname = RequestText(Request.Form("username"))
	adminpassword = RequestText(Request.Form("password"))
	adminpassword2 = RequestText(Request.Form("password2"))
	if adminname="" or adminpassword="" then 
		out "用户名和密码不能为空"
	elseif adminpassword<>adminpassword2 then
		out "密码验证不一样"
	end if
	call OpenDB()
	DbConn.Execute("UPDATE Admin SET adminname = '" & adminname & "',adminpassword = '" & adminpassword & "'")
	call CloseDB()
	Response.Redirect "admin_poll.asp"
end if
End Sub
%>
