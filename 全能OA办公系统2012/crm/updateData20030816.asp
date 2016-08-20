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
Function getGroupId(s)
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uName = '" & s & "'",conn,3,1
	If rs.RecordCount = 1 Then
	    getGroupId = rs("uGroup")
	Else
	    getGroupId = 0
	End If
	rs.Close
	Set rs = Nothing
End Function

Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_client",conn,3,2
Do While Not rs.BOF And Not rs.EOF
    rs("cGroup") = getGroupId(rs("cUser"))
	rs.Update
	Response.Write(rs("cUser") & rs("cGroup") & "<br>")
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="http://www.web87.9126.com">
<meta name="Date" content="2003-08">
<title>客户关系管理系统</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
</head>

<body>

</body>
</html>
