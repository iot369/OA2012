<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Option Explicit
Response.Buffer = True
Response.Expires = 0
Response.Expiresabsolute = Now() - 1 
Response.AddHeader "pragma","no-cache" 
Response.AddHeader "cache-control","private" 
Response.CacheControl = "no-cache"
%>
<!--#include file="Connections/conn.asp" -->

<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
<script language="JavaScript">
<!--
if (this.location.href == top.location.href){
    top.location.href = "";
}

function checkInput()
{
    var strItem1 = document.all.item1.value;
	var strItem2 = document.all.item2.value;
	if (strItem1 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">请输入账号。</font>"
		document.all.item1.focus();
		return false;
	}
	if (strItem2 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">请输入密码。</font>"
		document.all.item2.focus();
		return false;
	}
}

function checkChr(chr)
{
    var str0 = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
	var intLen = chr.length;
	for (var i=0;i<intLen;i++){
	    
	}
}
-->
</script>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<style type="text/css">
body {SCROLLBAR-FACE-COLOR:#3499D0;
SCROLLBAR-HIGHLIGHT-COLOR:#CCFFFF;
SCROLLBAR-SHADOW-COLOR:#2587C3;
SCROLLBAR-ARROW-COLOR:#CCFFFF;
SCROLLBAR-BASE-COLOR:#1068A4; 
SCROLLBAR-DARK-SHADOW-COLOR:#3499D0} 
</style>
</head>

<body style="background-color: menu;" topmargin="0" leftmargin="0">
<table width="702"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21" background="../images/r_bg.gif"><div align="center"><font color="0D79B3"><b>— 销 售 系 统 —</b></font></div></td>
  </tr>
</table>
<br>
<td align="center">
</td>
<%
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And Session("CRM_level") > 0 Then Response.Redirect("listAll.asp")
Dim action
action = Trim(Request("action"))
Select Case action
Case "login"
    Call login()
Case Else
    Call loginForm()
End Select

Sub login()
    Dim account,password
	account = Trim(Request("item1"))
	password = Trim(Request("item2"))
	If account = "" Or password = "" Then
	    Response.Redirect("login.asp?errMsg=2")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uAccount = '" & account & "'",conn,3,1
	If rs.RecordCount <> 1 Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If password <> rs("uPassword") Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If rs("uBlock") = True Then
	    Response.Redirect("login.asp?errMsg=3")
		Response.End()
	End If
	Session("CRM_account") = account
	Session("CRM_name") = rs("uName")
	Session("CRM_level") = rs("uLevel")
	Session("CRM_group") = rs("uGroup")
	rs.Close
	Set rs = Nothing
	Response.Redirect("listAll.asp")
End Sub

Sub loginForm()
    Dim errMsg
	errMsg = CInt(ABS(Request("errMsg")))
	Select Case errMsg
	Case 2
	    errMsg = "<font color=""#FF0000"">请输入账号和密码。</font>"
	Case 1
	    errMsg = "<font color=""#FF0000"">账号密码错误。</font>"
	Case 3
	    errMsg = "<font color=""#FF0000"">账号被冻结。</font>"
	Case Else
	    errMsg = ""
	End Select
%>
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="middle"><table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
      <tr>
        <td align="center" valign="middle"><table width="364" border="0" cellspacing="0" cellpadding="2" style="border-left: 1px solid #FFFFFF; border-right: 1px solid #888888; border-top: 1px solid #FFFFFF; border-bottom: 1px solid #888888;">
            <form name="loginForm" action="?action=login" method="post" onSubmit="return checkInput();">
              <tr> </tr>
              <tr>
                <td height="24" id="errMsg"><% = errMsg %></td>
              </tr>
              <tr>
                <td height="24" align="center"><span style="font-size:11px"></span>账号：
                    <input name="item1" type="text" id="item1" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="24" align="center">密码：
                    <input name="item2" type="password" id="item2" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" value="123" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="24" align="center">
                  <input type="submit" name="Submit" value="提交">
&nbsp;&nbsp;
              <input name="Reset" type="reset" id="Reset" value="重置"></td>
              </tr>
              <tr>
                <td height="24" align="center">&nbsp;</td>
              </tr>
            </form>
        </table></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
End Sub
%>
</body>
</html>
