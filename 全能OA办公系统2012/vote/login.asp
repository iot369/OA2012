
<%If session("MoonDowner_Poll") = "MoonDowner_Poll" Then Response.Redirect "admin_poll.asp"%>
<!--#include file="config.asp"-->
 <% 
  Call Chklogin() 		
 %>
<!--#include file="top.asp" -->
  <br>
  <br>
  <br>
  <form method="post" action="login.asp">
  <TABLE cellSpacing=1 cellPadding=5 width="60%" border=0 align="center" bgcolor="#7C96B8">
    <TR  align="center"> 
      <TD colspan="2"><font color="#FF0000">-- 在线调查系统管理登录 
        --</font></TD>
    </TR>
          <TR bgcolor="#FFFFFF"> 
            <TD width="50%" align="right">用户名：</TD>
            <TD width="50%"> 
              <input name="username" type="text" size="12" maxlength="20">
            </TD>
          </TR>
          <TR bgcolor="#FFFFFF"> 
            <TD width="50%" align="right">密　码：</TD>
            <TD width="50%"> 
              <input name="password" type="password" size="12" maxlength="20">
            </TD>
          </TR>
          <TR align="center" > 
            <TD colspan="2"> 
              <input type="submit" name="Submit" value=" 登录 ">
            </TD>
          </TR>
        </TABLE>
                  
  <br>
  <br>
  <br>
  <br>
  <br>
</form>
<!--#include file="foot.asp" -->
<%
Sub Chklogin()
If Request.ServerVariables("REQUEST_METHOD")="POST" Then
	dim username	
   	dim password	
   	dim rs
   	username = RequestText(Request.Form("username"))
   	password = RequestText(Request.Form("password"))
	if username = "" then
  		out "请输入用户名。"
    elseif password = "" then
		out "请输入登录密码。"
    else
  		call OpenDB()
         set rs = DbConn.Execute("select * from Admin where Adminname='" & username & "'")
         if rs.eof then
             out "用户不存在。"
         elseif  password <> rs("adminpassword") then
             out "密码输入错误。"
	   	 else
			session("MoonDowner_Poll") = "MoonDowner_Poll"
			Response.Redirect "admin_poll.asp"
         end if
		 Set rs = Nothing
		 Call CloseDB()
	end if
end if
End Sub
%>
