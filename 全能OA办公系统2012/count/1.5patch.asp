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
	font-family: "����", "Times New Roman", "sans-serif","Arial";
}
</style>

<title>CuteCounterV1.5 IP���޸�����_��������ITѧϰ���ṩ</title></head>
<body leftmargin="0">
<h3 align=center>CuteCounterV1.5 IP���޸�����</h3>
<table cellpadding="0" cellspacing="0" border="0" align=center>
<tr><td>
<%
dim rs,C1
set rs = hx.execute("select count(id) from CC_I where len(ip)<15")
    C1 = rs(0)
    if C1>0 then
        response.write " <p>����"&rs(0)&"����¼��Ҫ����</p>"
        response.write " <li><a href=?module=1>������в�����������IP��¼(�Ƽ�)</a></li>" 
        response.write " <li><a href=?module=2>�������IP��¼</a></li>" 
        response.write " <li><a href=?module=3>������ЩIP��¼(1����Լ���Ը���3000����¼)</a></li>"     
        response.write " <p>���飺����ǰ���ȱ������ݿ⣬���º����ѹ��һ�����ݿ�</p>"
    else
        response.write "IP��������ȷ���������:)"
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
    Response.write "<p align=center>�����ɹ�</p>"
End Sub

Sub UpdateIp2()
    hx.execute("delete from CC_I")
    Response.write "<p align=center>�����ɹ�</p>"
End Sub

Sub UpdateIp3()
%><br>
<table cellpadding="0" cellspacing="0" border="0" align=center>
<tr>
    <td colspan=2> ���ڸ��£�Ԥ�Ʊ��ι���<%=C1%>����Ҫ���� 
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
       		Response.Write "txt.innerHTML=""�Ѹ��µ���"&i&"����¼��" & FormatNumber(i/C1*100,4,-1) & """;" & VbCrLf
       		Response.Write "</script>" & VbCrLf
       		Response.Flush
	    response.flush 
	end if   
    rs.movenext
    loop
    Response.Write "<script>img.width=400;txt.innerHTML=""100"";</script>"
    Response.write "<p align=center>������ɣ�</p>"
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