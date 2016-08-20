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
<%
If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Session("CRM_planWin") = True
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>拜访计划提醒</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function view(i)
{
    opener.location.href = "view.asp?cId=" + i;
}
-->
</script>
</head>

<body style="border: 0px;" >
<br>
<%
Function getCompany(cId)
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select cCompany From baidu_client Where cId = " & cId,conn,3,1
	If rs.RecordCount = 1 Then
	    getCompany = rs("cCompany")
	Else
	    getCompany = ""
	End If
    rs.Close
	Set rs = Nothing
End Function
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_recordsPlan Where rUser = '" & Session("CRM_name") & "'",conn,3,1
Do While Not rs.BOF And Not rs.EOF
    If DateAdd("d",rs("rDay"),Date()) >= rs("rDate") Then
%>
<table width="360" border="0" align="center" cellpadding="2" cellspacing="0">
  <tr>
    <td><strong>客户名称：</strong><a href="JavaScript: view('<% = rs("cId") %>');"><% = getCompany(rs("cId")) %></a></td>
  </tr>
  <tr>
    <td><strong>拜访时间：</strong><% = rs("rDate") %></td>
  </tr>
  <tr>
    <td><strong>拜访类型：</strong><% = rs("rType") %></td>
  </tr>
  <tr>
    <td><% = rs("rContent") %></td>
  </tr>
</table>
<br>
<%
    End If
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<table width="360" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center">
      <input type="button" name="Submit" value=" 关闭窗口 " onClick="window.close();">
    </td>
  </tr>
</table>
<br>
</body>
</html>
