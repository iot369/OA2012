<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<%
Function dl(f,n)
    On Error Resume Next
    Set S = CreateObject("Adodb.Stream") 
    S.Mode = 3 
    S.Type = 1 
    S.Open 
    S.LoadFromFile(f)
    If Err.Number > 0 Then 
        Response.Status = "404"
    Else
        Response.ContentType = "application/octet-stream"
        Response.AddHeader "Content-Disposition:","attachment; filename=" & n
        Range = Mid(Request.ServerVariables("HTTP_RANGE"),7)
        If Range = "" Then
           Response.BinaryWrite(S.Read)
        Else
            S.position = Clng(Split(Range,"-")(0))
            Response.BinaryWrite(S.Read)
        End If
    End If
    Response.End
End Function

Dim f
f = Trim(Request("file"))
If f <> "" Then Call dl(Server.MapPath(f),f)
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>文件下载</title>
</head>

<body>
<script language="JavaScript">
<!--
function clsWin()
{
    window.close();
}
setTimeout("clsWin()",1000);
-->
</script>
</body>
</html>
