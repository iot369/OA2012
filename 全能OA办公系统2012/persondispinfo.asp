<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
id=request("id")
if id="" then
	response.redirect "personlist.asp"
	response.end
end if
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

%>
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<script language="javascript">
function deleteask()
{
	if (confirm("���Ҫɾ���ü�¼��"))
		location.href="persondelete.asp?id=<%=id%>";
}
</script>
<title>OA�칫ϵͳ</title>
<style type="text/css">
<!--
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<center>
  <table width="583"  border="0" cellspacing="0" cellpadding="0">
    <tr>
      <td height="21"><div align="center">
          <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
            <tr>
              <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
              <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">����ͨѶ¼</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td> <br>
  <form method="post" action="personlist.asp">
<table align="center">
<tr>
<td>
<input type="button" value="ɾ��" onclick="deleteask()">&nbsp;&nbsp;�鿴����&nbsp;&nbsp;&nbsp;&nbsp;</td>
<td><input type="button" value="�༭" onclick="location.href='personeditrecord.asp?id=<%=cstr(id)%>'">&nbsp;<input type="submit" name="submit" value="����">
</td>
</tr>
</table></form>
</center>
<%

set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personrecord,persontype where personrecord.id="&id&" and persontype.id=personrecord.recordtype"
rs.open sql,conn,1
if rs.eof or rs.bof then
	conn.close
	set rs=nothing
	response.write("<font color=""#ff0000"" size=""+3"">")
	response.write("<center>���󣬸�����Ϣ�����Ѿ���ɾ����</center>")
	response.write("</font>")
	response.end
end if
if rs("thisinfousername")<>oabusyusername then
	conn.close
	set rs=nothing
	response.write("<font color=""#ff0000"" size=""+1"">")
	response.write("<center>�Բ��������ܲ鿴�������ϣ�</center>")
	response.write("</font>")
	response.end
end if
xm=server.htmlencode(rs("xm"))
xb=rs("sex")
zw=server.htmlencode(rs("userzw"))
if zw="" then zw="-"
lb=rs("typename")
dw=server.htmlencode(rs("company"))
zzdh=server.htmlencode(rs("hometel"))
if zzdh="" then zzdh="-"
yzbm=server.htmlencode(rs("postcard"))
if yzbm="" then yzbm="-"
dhfj=server.htmlencode(rs("companytel"))
if dhfj="" then dhfj="-"
sj=server.htmlencode(rs("handset"))
if sj="" then sj="-"
hj=server.htmlencode(rs("callset"))
if hj="" then hj="-"
dzyj=server.htmlencode(rs("email"))
if dzyj="" then dzyj="-"
zzdz=server.htmlencode(rs("homeaddress"))
if zzdz="" then zzdz="-"
cz=server.htmlencode(rs("fax"))
if cz="" then cz="-"
bz=server.htmlencode(rs("remark"))
if bz="" then bz="-"
%>
<center>     
<br>     
   <table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table border="1" cellpadding="0" cellspacing="0" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF" width="540">
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">����</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="227" height="25">��<%=xm%></td>
      <td width="96" height="25" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�Ա�</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="219" height="25">��<%=xb%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ְ��</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="227" height="25">��<%=zw%></td>
      <td width="96" height="25" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ͨѶ¼���</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="219" height="25">��<%=lb%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��λ</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3" width="544" height="25">��<%=dw%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">סլ�绰</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="227" height="25">��<%=zzdh%></td>
      <td width="96" height="25" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="219" height="25">��<%=yzbm%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�绰��ֻ�</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="228" height="25">��<%=dhfj%></td>
      <td width="97" height="25" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <p align="center">��&nbsp;&nbsp;&nbsp;&nbsp;��</p>      </td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="220" height="25">��<%=sj%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">M S N</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="228" height="25">��<%=hj%></td>
      <td width="97" height="25" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <p align="center">Email
      </td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="220" height="25">��<%=dzyj%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">סլ��ַ</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="542" colspan="3" height="25">��<%=zzdz%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">����</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="542" colspan="3" height="25">��<%=cz%></td>
    </tr>
    <tr> 
      <td width="94" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">��ע˵��</td>
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" colspan="3" width="544" height="25">��<%=bz%></td>
    </tr>
  </table>    
    
  <input type="button" value="����" onclick="window.location.href=history.go(-1);">    
</center>     
<%     
conn.close 
set rs=nothing 
     
%>     </td>
    </tr>
  </table>
 
</body>     
</html>