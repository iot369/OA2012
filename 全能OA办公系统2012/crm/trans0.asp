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
<title>����ת��,��ѯ����  ,��ϵqq:�ͻ���ϵ����ϵͳ��ϵ���䣺  ,�ͻ�����qq:�ͻ���ϵ����ϵͳ---��ѯ�绰:(����)</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
</head>

<body style="overflow: hidden; border: 0px;" bgcolor="menu">
<%
Dim errMsg,flag
errMsg = CInt(Abs(Request("errMsg")))
flag = 0

Response.Write(errMsg)
Select Case errMsg
Case 1
    errMsg = "<center><br><br><font color=""#FF0000"">�ύ�����ݲ�������</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""window.close();""><br><br>"
	Response.Write(errMsg)
	flag = 1
Case 2
    errMsg = "<center><br><br><font color=""#FF0000"">��ת���û���Ŀ���û���ͬ��</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""window.close();""><br><br>"
	Response.Write(errMsg)
	flag = 1
Case 3
    errMsg = "<center><br><br><font color=""#FF0000"">��ת���û���Ŀ���û�<br>������һ�������ڡ�</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""window.close();""><br><br>"
	Response.Write(errMsg)
	flag = 1
Case 4
    errMsg = "<center><br><br><font color=""#FF0000"">����ת�����</font><br><br>"
	errMsg = errMsg & "<input type=""button"" value="" �� �� "" onClick=""opener.loaction.refresh();""><br><br>"
	Response.Write(errMsg)
	flag = 1
Case Else
    
End Select

%>
</body>
</html>
<%
If flag = 1 Then Response.End()
Dim selNum,arrayId,transFrom,transTo
selNum = Trim(Request("selNum"))
arrayId = Trim(Request("arrayId"))
transFrom = Trim(Request("transFrom"))
transTo = Trim(Request("transTo"))

If selNum = "" Or arrayId = "" Or transTo = "" Or transFrom = "" Then Response.Redirect("?errMsg=1")

If transFrom = transTo Then Response.Redirect("?errMsg=2")
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_user Where uName In ('" & transFrom & "','" & transTo & "')",conn,3,1
If rs.RecordCount <> 2 Then Response.Redirect("?errMsg=3")
rs.Close

Response.Write(selNum)
Response.End()
If selNum = "seled" Then
End If
%>