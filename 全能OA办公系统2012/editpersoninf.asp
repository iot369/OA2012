<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='default.asp';")
	response.write("</script>")
	response.end
end if

'--------------------------------------
%>
<html>

<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA�칫ϵͳ.��Ե�ر��</title>
<style type="text/css">
<!--
.style1 {color: #098abb}
-->
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
</style>
</head>
<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
            <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">���˻�������</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<center>
  <br>
  <table bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr> 
      <td> �༭<%=oabusyname%>�ĸ��˻�������&nbsp;&nbsp; </td>
      
      <form method="post" action="personinf.asp">
      <td>  
        <input type="submit" value="����">
      </td>
      <form method="post" action="personinf.asp">
        <td>  
          <input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷ��Ҫɾ����ĸ��˻���������');">
        </td>
      </form>
      </form>
    </tr>
  </table>
</center>

<%
dim a(33)
'�����ݿ⣬�������˵���
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
if not rs.eof and not rs.bof then
for i=1 to 33
a(i)=rs("a" & i)
next
inputdate=rs("inputdate")
updatedate=rs("updatedate")
havephoto=rs("havephoto")
else
for i=1 to 33
a(i)=""
next
a(26)="���������λ����ַ���绰"
a(30)="�밴������ϵ���������Ա𡢹�����λ��ְ�񡢵�ַ˳����д"
a(32)="�밴��ϵ����������ַ���绰˳����д"
inputdate=""
updatedate=""
havephoto="no"
end if
%>
<center>
<br>
<form method="post" ENCTYPE="multipart/form-data" action="editpersoninfindb.asp">
  <table border="0" cellpadding="0" cellspacing="0" width="95%">
    <tr> 
      <td width="33%">Ա����ţ� 
        <input type=text size=10 name=a1 value="<%=a(1)%>">
      </td>
      <td width="33%"></td>
      <td align="right"></td>
    </tr>
  </table>

  <table border="0" cellpadding="0" cellspacing="0" width="540" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%" height="20">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="35%"><%=oabusyname%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="15%">�� 
        �� ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="35%"> 
        <input type=text size=10 name=a2 value="<%=a(2)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <select name=a3 size=1>
          <option value="��"<%=selected(a(3),"��")%>>��&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="Ů"<%=selected(a(3),"Ů")%>>Ů</option>
        </select>
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a4 value="<%=a(4)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��������</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=oabusyuserdept%></td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ְ&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"><%=oabusyuserlevel%></td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">ְ&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a5 value="<%=a(5)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��������</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a6 value="<%=a(6)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">������ò</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a7 value="<%=a(7)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">����״��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a8 value="<%=a(8)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a9 value="<%=a(9)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a10 value="<%=a(10)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">���֤����</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a11 value="<%=a(11)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a12 value="<%=a(12)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">����״��</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <select name=a13 size=1>
          <option value="δ��"<%=selected(a(13),"δ��")%>>δ��&nbsp;&nbsp;&nbsp;&nbsp;</option>
          <option value="�ѻ�"<%=selected(a(13),"�ѻ�")%>>�ѻ�</option>
        </select>
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">��ҵԺУ</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a14 value="<%=a(14)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">���˳ɷ�</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a15 value="<%=a(15)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ר&nbsp;&nbsp;&nbsp; 
        ҵ</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a16 value="<%=a(16)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a17 value="<%=a(17)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">ѧ&nbsp;&nbsp;&nbsp; 
        ��</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a18 value="<%=a(18)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��������</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a19 value="<%=a(19)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">����ˮƽ</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a20 value="<%=a(20)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��ͨ���̶�</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a21 value="<%=a(21)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">����̶�</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a22 value="<%=a(22)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">���������</td>
      <td style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a23 value="<%=a(23)%>">
      </td>
      <td align="center" style="border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">�������ڵ�</td>
      <td style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=10 name=a24 value="<%=a(24)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">�� 
        ס ַ</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="85%"> 
        <input type=text size=50 name=a25 value="<%=a(25)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">������ŵ�</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <input type=text size=50 name=a26  value="<%=a(26)%>">
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">����ר��<br>
        �Լ�����</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a27 rows="3" cols="49"><%=a(27)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��������<br>
        �����ֽ�<br>
        ���ʹ���</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a28 rows="3" cols="49"><%=a(28)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��������</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a29 rows="3" cols="49"><%=a(29)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��ͥ���</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a30 rows="3" cols="49"><%=a(30)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        ��<br>
        ��ϵ��ʽ</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a31 rows="3" cols="49"><%=a(31)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" height="20">��������<br>
        �������<br>
        ֪ͨ����</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA"> 
        <textarea name=a32 rows="3" cols="49"><%=a(32)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        ע</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"> 
        <textarea name=a33 rows="3" cols="49"><%=a(33)%></textarea>
      </td>
    </tr>
    <tr> 
      <td align="center" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" height="20">��&nbsp;&nbsp;&nbsp; 
        Ƭ</td>
      <td colspan="3" style="border-right: 2 solid #B0C8EA; border-bottom: 2 solid #B0C8EA"> 
        <input type="file" name="file1" size=40>
      </td>
    </tr>
  </table>
  <br>
  <table>
<tr>
<td>
<%
if inputdate="" then
%>
<input type="submit" name="submit" value="����">
<%
else
%>
<input type="submit" name="submit" value="�޸�"> 
<%
end if
%>
</td>
</form>
<form method="post" action="personinf.asp"><td>
<input type="submit" name="submit" value="ɾ��" onclick="return window.confirm('��ȷ��Ҫɾ����ĸ��˻���������');">
</td>
</form>
</table>
</center>

</body>
</html>










