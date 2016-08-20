<!--#include file="conn.asp"-->
<%
Response.Buffer = true
Server.ScriptTimeOut = 999
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<style type="text/css">
td,p,body {
	font-size: 12px;
	font-family: "宋体", "Times New Roman", "sans-serif","Arial";
}
</style>

<title>CuteCounterV1.5 IP表修复程序_本程序由IT学习者提供</title></head>
<body leftmargin="0">
<h3 align=center>CuteCounterV1.5 IP表修复程序</h3>
<table cellpadding="0" cellspacing="0" border="0" align=center>
<tr><td>
<%
dim rs,C1
set rs = hx.execute("select count(id) from CC_I where len(ip)<15")
    C1 = rs(0)
    if C1>0 then
        response.write " <p>共有"&rs(0)&"条记录需要更新</p>"
        response.write " <li><a href=?module=1>清空所有不符合条件的IP记录(推荐)</a></li>" 
        response.write " <li><a href=?module=2>清空所有IP记录</a></li>" 
        response.write " <li><a href=?module=3>更新这些IP记录(1分钟约可以更新3000条记录)</a></li>"     
        response.write " <p>建议：更新前请先备份数据库，更新后可以压缩一下数据库</p>"
    else
        response.write "IP表数据正确，无需更新:)"
    end if
set rs = nothing
%>
</td></tr></table>
<%
dim module
module = request("module")
Select Case module
case 1
    Call UpdateIp1()
case 2
    Call UpdateIp2()
case 3
    Call UpdateIp3()
End select

Sub UpdateIp1()
    hx.execute("delete from CC_I where len(ip)<15")
    Response.write "<p align=center>操作成功</p>"
End Sub

Sub UpdateIp2()
    hx.execute("delete from CC_I")
    Response.write "<p align=center>操作成功</p>"
End Sub

Sub UpdateIp3()
%><br>
<table cellpadding="0" cellspacing="0" border="0" align=center>
<tr>
    <td colspan=2> 正在更新，预计本次共有<%=C1%>个需要更新 
      <table width="400" border="0" cellspacing="1" cellpadding="1">
<tr> 
<td bgcolor=000000>
<table width="400" border="0" cellspacing="0" cellpadding="1">
<tr><td bgcolor=ffffff height=9><img src="bar3.gif" width=10 height=10 id=img name=img align=absmiddle>
</td></tr></table>
</td></tr></table>
<span id=txt name=txt style="font-size:9pt">0</span><span style="font-size:9pt">%</span>
</td></tr>
</table>
<%
    dim rs,rs1
    dim ip,i
    dim C2
    if C1 > 5000 then
        C2 = 50
    elseif C1 > 1000 then
        C2 = 10
    else
        C2 = 5
    end if
        
set rs = hx.execute("select id,ip,cip,vtime from CC_I where len(ip)<15")
    do while not rs.eof
    ip = rs("ip")
        
    ip = Getip(ip)

    set rs1 = hx.execute("select ip,id from CC_I where ip = '"& ip &"'")

    if rs1.eof then
        hx.execute("Update CC_I set ip='"& ip &"' where id = " & rs("id"))
    else
        hx.execute("Update CC_I set cip = cip + " & rs("cip") & ",vtime = '"& rs("vtime")&"' where id = " & rs1("id"))
        hx.execute("Delete from CC_I where id = " & rs("id"))
    end if
    set rs1 = nothing
    i = i +1
   
    if i mod C2 = 0 then
      		Response.Write "<script>img.width=" & Fix((i/C1) * 400) & ";" & VbCrLf
       		Response.Write "txt.innerHTML=""已更新到第"&i&"条记录，" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
       		Response.Write "</script>" & VbCrLf
       		Response.Flush
	    response.flush 
	end if   
    rs.movenext
    loop
    Response.Write "<script>img.width=400;txt.innerHTML=""100"";</script>"
    Response.write "<p align=center>更新完成！</p>"
set rs = nothing
End Sub


	Function Getip(ip)
		Dim a,i
		a = Split(ip,".")
		if ubound(a)<>3 then Getip=0:Exit Function
		For i=0 to 3
 			Getip=Getip & String(3-Len(a(i)),"0") & a(i) & "."
		Next
		Getip=left(Getip,15)
	End Function	

%>