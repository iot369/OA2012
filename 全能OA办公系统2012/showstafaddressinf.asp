<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<!--#include file="asp/maillink.asp"-->
<!--#include file="asp/keepformat.asp"-->
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
'�����ݿ⣬�����༭ͨѶ¼Ȩ��
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username=" & sqlstr(oabusyusername)
rs.open sql,conn,1
allow_edit_person_addressinf=rs("allow_edit_person_addressinf")
allow_edit_dept_addressinf=rs("allow_edit_dept_addressinf")
allow_edit_all_addressinf=rs("allow_edit_all_addressinf")
allow_edit_client_addressinf=rs("allow_edit_client_addressinf")
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
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style3 {color: #0d79b3}
.style5 {color: #2d4865}
.style6 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">


  <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
    <tr>
      <td height="21"><div align="center"><span class="style2">
          <div align="center">
            <table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
                <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                      <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
                      <td><span class="style5">ͨѶ����</span></td>
                    </tr>
                </table></td>
                <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
              </tr>
            </table>
          <font color="0D79B3"></font></div>
      </span></div></td>
    </tr>
    <tr>
      <td><div align="center">
        <center>
            <br>
            <center>
              <table>
                <tr>
                  <td><span class="style3"><%=request("name")%>��ͨѶ����&nbsp;&nbsp;</span>&nbsp;&nbsp;</td>
                  <%
if (allow_edit_person_addressinf="yes" and request("username")=oabusyusername) or (allow_edit_dept_addressinf="yes" and request("userdept")=oabusyuserdept) or allow_edit_all_addressinf="yes" then
%>
                  <form method="post" action="editstafaddressinf.asp">
                    <td><input type="submit" name="submit" value="�༭"></td>
                    <input type="hidden" name="userdept" value="<%=request("userdept")%>">
                    <input type="hidden" name="username" value="<%=request("username")%>">
                    <input type="hidden" name="name" value="<%=request("name")%>">
                    <input type="hidden" name="userlevel" value="<%=request("userlevel")%>">
                  </form>
                  <%
end if
%>
                  <form method="post" action="stafaddressinf.asp">
                    <td><input type="submit" name="submit2" value="����"></td>
                    <input type="hidden" name="userdept" value="<%=request("userdept")%>">
                  </form>
                </tr>
              </table>
              <center>
                <table height=20 border="0" cellpadding="0" cellspacing="0" width="95%">
                  <tr>
                    <td></td>
                  </tr>
                </table>
                <%
'�����ݿ����ͨѶ��Ϣ
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from stafaddressinf where username=" & sqlstr(request("username"))
rs.open sql,conn,1
if not rs.eof and not rs.bof then
companytel=rs("companytel")
fax=rs("fax")
hometel=rs("hometel")
homeaddress=rs("homeaddress")
postcard=rs("postcard")
sex=rs("sex")
handset=rs("handset")
callset=rs("callset")
remark=rs("remark")
email=rs("email")
else
companytel=""
fax=""
hometel=""
homeaddress=""
postcard=""
sex=""
handset=""
callset=""
remark=""
email=""
end if

%>
</table>                       
              <table width="583"  border="0" align="center" cellpadding="0" cellspacing="0">
                <tr>
                  <td><table width="550"  border="0" cellspacing="0" cellpadding="0" align="center">
                    <tr>
                      <td height="1" bgcolor="4B789F" align="center"></td>
                    </tr>
                  </table>
                    <table width="550" border="0" align="center" cellpadding="0" cellspacing="1" bgcolor="B0C8EA" >
                      <tr>
                        <td height=2 colspan="4" align="center" ></td>
                      </tr>
                      <tr>
                        <td width="15%" height=25 align="center" bgcolor="D7E8F8" ><span class="style6">Ա������</span></td>
                        <td width="35%" bgcolor="#FFFFFF"  >��<%=request("name")%></td>
                        <td width="15%" align="center" bgcolor="D7E8F8"><span class="style6">�Ա�</span></td>
                        <td width="35%" bgcolor="#FFFFFF" >��<%=checked3(sex)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8" ><span class="style6">���ڲ���</span></td>
                        <td bgcolor="#FFFFFF" >��<%=request("userdept")%></td>
                        <td align="center" bgcolor="D7E8F8" ><span class="style6">ְ��</span></td>
                        <td bgcolor="#FFFFFF" >��<%=request("userlevel")%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8" ><span class="style6">�绰��ֻ�</span></td>
                        <td bgcolor="#FFFFFF" >��<%=checked3(companytel)%></td>
                        <td align="center" bgcolor="D7E8F8"  ><span class="style6">����</span></td>
                        <td bgcolor="#FFFFFF" >��<%=checked3(fax)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8"   ><span class="style6">�ֻ�����</span></td>
                        <td bgcolor="#FFFFFF"  >��<%=checked3(handset)%></td>
                        <td align="center" bgcolor="D7E8F8"  ><span class="style6">MSN</span></td>
                        <td bgcolor="#FFFFFF" >��<%=checked3(callset)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8"   ><span class="style6">סլ�绰</span></td>
                        <td bgcolor="#FFFFFF"  >��<%=checked3(hometel)%></td>
                        <td align="center" bgcolor="D7E8F8"  ><span class="style6">Email</span></td>
                        <td bgcolor="#FFFFFF" >��<%=maillink(email)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8"   ><span class="style6">סլ��ַ</span></td>
                        <td colspan="3" bgcolor="#FFFFFF" >��<%=checked3(homeaddress)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8"   ><span class="style6">סլ�ʱ�</span></td>
                        <td colspan="3" bgcolor="#FFFFFF" >��<%=checked3(postcard)%></td>
                      </tr>
                      <tr>
                        <td height=25 align="center" bgcolor="D7E8F8" ><span class="style6">��ע˵��</span></td>
                        <td colspan="3" bgcolor="#FFFFFF" >��<%=checked3(keepformat(remark))%></td>
                      </tr>
                    </table></td>
                </tr>
              </table>
              <br>
                <%

%>
              </center>
            </center>
          <center>
              <table border="0"  cellspacing="0" cellpadding="0" width="95%" height=10>
              </table>
            <br>
              <br>
          </center>
        </center>
      </div></td>
    </tr>
</table>
</body>
</html>