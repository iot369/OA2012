<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<% Option Explicit %>
<!--#include file="conn/conn.asp" -->
<!--#include file="setup.asp" -->
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title><% = siteTitle %>,</title>
<link href="cn001.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function chkLogin()
{
    var strUser = document.elogin.item1.value;
	var strPws = document.elogin.item2.value;
	if(strUser == ""){
	    alert("�������˺š�");
		document.elogin.item1.focus();
		return false;
	}
	if(strPws == ""){
        alert("���������롣");
		document.elogin.item2.focus();
		return false;
	}
	window.setTimeout("cleanPws()",1000);
}

function cleanPws()
{
    document.all.item2.value = "";
}
-->
</script>
</head>

<body>
<!--#include file="head.asp" -->
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td valign="top"><hr align="center" size="1" noshade color="#E9F3FE">
      <table width="360" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="#E9F3FE">
        <form name="elogin" action="http://www.china-net.com/crm02/login.asp?action=login" method="post" onSubmit="return chkLogin();" target="_blank">
		<tr align="center"> 
          <td height="22" colspan="2" class="titletd">Ա �� �� ¼</td>
        </tr>
        <tr> 
          <td width="80" align="right">�˺ţ�</td>
          <td height="24" bgcolor="#FFFFFF">
<input name="item1" type="text" id="item1" size="24" maxlength="12"></td>
        </tr>
        <tr> 
          <td width="80" align="right">���룺</td>
          <td height="24" bgcolor="#FFFFFF">
<input name="item2" type="password" id="item2" size="24" maxlength="12"></td>
        </tr>
        <tr> 
          <td width="80" align="right">&nbsp;</td>
          <td height="24" bgcolor="#FFFFFF">
<input name="imageField53" type="image" src="images/login.gif" width="36" height="20" border="0"></td>
        </tr>
		</form>
      </table>
      <br>
      <br>
      <table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
        <tr>
          <td><p>ע�����<br>
              1.�˺š������ɹ���Աͨ���ʼ��ṩ��<br>
              2.�������룬��ֱ�Ӻ͹���Ա��ϵ��</p>
            </td>
        </tr>
      </table> </td>
    <td width="17" align="center" valign="top" background="images/dot17x4.gif">&nbsp;</td>
    <td width="161" align="center" valign="top" bgcolor="#E9F3Fe"> 
      <%
If Session("userName") = "" Or Session("userPass") <> True Then
    Session("thisPage") = Request.ServerVariables("HTTP_REFERER")
%>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E9F3FE">
        <form name="login" action="login.asp?action=login" method="post">
          <tr> 
            <td height="20" align="center" class="titletd">�� �� �� ¼</td>
          </tr>
          <tr> 
            <td height="24" align="center">�˺ţ� 
              <input name="userName" type="text" id="userName" size="12" maxlength="16"> 
            </td>
          </tr>
          <tr> 
            <td height="24" align="center">���룺 
              <input name="userPasswords" type="password" id="pws4" size="12" maxlength="16"></td>
          </tr>
          <tr> 
            <td height="24" align="center"> <input name="imageField5" type="image" src="images/login.gif" width="36" height="20" border="0"></td>
          </tr>
          <tr> 
            <td height="24" align="center"><a href="reg.asp" class="redem">ע��</a> 
              �һ�����</td>
          </tr>
        </form>
      </table>
      <%
Else
%>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E9F3FE">
        <form name="isLogin" action="login.asp" method="post">
          <tr> 
            <td height="20" align="center" class="titletd">�� �� �� Ϣ</td>
          </tr>
          <tr> 
            <td height="24" align="center" class="redem">
              <% = Session("userName") %>
            </td>
          </tr>
          <tr> 
            <td height="24" align="center">���Ѿ���¼</td>
          </tr>
          <tr> 
            <td height="24" align="center"> <input name="imageField52" type="image" src="images/logout.gif" width="36" height="20" border="0" onClick="location.href='login.asp?action=logout';"></td>
          </tr>
          <tr> 
            <td height="24" align="center"><a href="reg.asp?action=edit&userId=<% = Session("userId") %>">�޸�����</a></td>
          </tr>
        </form>
      </table>
      <%
End If
%>
      <table width="100%" height="10" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFFFFF"><img src="images/null5.gif" width="5" height="5"></td>
        </tr>
      </table> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E9F3FE">
        <tr> 
          <td height="20" align="center" class="titletd">�� �� �� ��</td>
        </tr>
        <tr> 
          <td valign="top" class="mar5td"> 
            <%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select Top 10 * From ser Order By serId Desc",conn,3,1
Do While Not rs.BOF And Not rs.EOF
    Response.Write("<span class='raquo'>&raquo;</span>&nbsp;<a href='services.asp?serId=" & rs("serId") & "' target='_blank'>" & rs("serTitle") & "</a><br>")
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
          </td>
        </tr>
        <tr> 
          <td align="right"><a href="services.asp"><span class="raquo">&raquo;</span>&nbsp;�������...</a></td>
        </tr>
      </table>
      <table width="100%" height="10" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td bgcolor="#FFFFFF"><img src="images/null5.gif" width="5" height="5"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="#E9F3FE">
        <tr> 
          <td height="20" align="center" class="titletd">�� �� �� ��</td>
        </tr>
        <tr> 
          <td valign="top" class="mar5td"> 
            <%
Call showLinks("�Ƽ��ͻ�",5)
%>
          </td>
        </tr>
        <tr> 
          <td align="right"><a href="clients.asp" target="_blank"><span class="raquo">&raquo;</span>&nbsp;����ͻ�...</a></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<table width="778" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr> 
    <td align="center" valign="top"><hr align="center" size="2" noshade color="#E9F3FE"> 
      <a href="about_us.asp"> 
      <%
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From menu Where menuLocal = '1' Order By menuList Asc",conn,3,1
Do While Not rs.BOF And Not rs.EOF
    Response.Write("| <a href='menu.asp?menuId=" & rs("menuId") & "' target='_blank'>" & rs("menuName") & "</a> ")
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
Response.Write("|")
%>
      <br>
      </a></td>
  </tr>
</table>
</body>
</html>
