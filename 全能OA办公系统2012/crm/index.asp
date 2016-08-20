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
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="Author" content="http://www.web87.9126.com">
<meta name="Date" content="2003-08">
<title>客户关系管理系统联系邮箱：  ,客户服务qq:客户关系管理系统---咨询电话:(短信)</title>
</head>

<%
Dim url
url = "login.asp"
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And IsNumeric(Session("CRM_level")) Then url = "listAll.asp"
    
%>
<frameset rows="0,*" frameborder="0" framespacing="0">
  <frame name="top" scrolling="no" src="about:blank">
  <frame name="main" scrolling="yes" src="<% = url %>">
</frameset>
<noframes>

<body>

</body>
</noframes>
</html>
