<%@ LANGUAGE = VBScript %>
<!--#include file="asp/sqlstr.asp"-->

<!--#include file="asp/opendb.asp"-->
<!--#include file="asp/checked.asp"-->
<%
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
.style4 {color: #0d79b3}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
<SCRIPT language=javascript>function chkinput(f){var tmp=f.name.value;if(!tmp){alert("����д��Ҫ��ѯ������!");return false;}var tmp2=f.tiaojian.value;if(!tmp2){alert("��ѡ����Ҫ��ѯ������!");return false;}return true;}function chkinput2(f){var tmp=f.user.value;if(!tmp){alert("�ʺŲ���Ϊ��!");return false;}var tmp2=f.pass.value;if(!tmp2){alert("���벻��Ϊ��!");return false;}var tmp3=f.site.value;if(!tmp3){alert("��û��ѡ������!");return false;}return true;}function MM_openBrWindow(theURL,winName,features){window.open(theURL,winName,features);}</SCRIPT>

<SCRIPT language=javascript>
<!--
function clearpass(){document.loginmail.pass.value="";}//--></SCRIPT>
</head>
<body  topmargin="0" leftmargin="0" bgcolor="#F9F9FF">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
      <table width="100%"  border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
          <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
              <tr>
                <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                <td class="style7">�����ʼ�</td>
              </tr>
          </table></td>
          <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
        </tr>
      </table>
    <font color="0D79B3"></font></div></td>
  </tr>

  <tr>
    <td><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td>&nbsp;</td>
          </tr>
        </table>
        <center>
        </center>
        <br>
        <table width="300"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td height="30"><div align="center">���������½</div></td>
          </tr>
        </table>

        <TABLE cellSpacing=0 cellPadding=0 width=550 border=0 
valign="top">
          <TBODY>
            <TR>
              <FORM name=loginmail onsubmit="return chkinput2(this);" 
    action=http://www.hao123.com/sendmail.php method=post target=_blank>
                <TD align=middle><B></B>�ʺţ�
                    <INPUT 
      style="FONT-SIZE: 12px" tabIndex=1 size=14 name=user>
                    ���䣺
                    <SELECT tabIndex=2 
      size=1 name=site>
                      <OPTION selected>��ѡ������</OPTION>
                      <OPTION 
        value=163.com>@163.com ����</OPTION>
                      <OPTION value=sina.com>@sina.com ����</OPTION>
                      <OPTION value=126.com>@126.com ����</OPTION>
                      <OPTION 
        value=cn.yahoo.com>@yahoo.com.cn �Ż�</OPTION>
                      <OPTION 
        value=163.net>@163.net</OPTION>
                      <OPTION 
        value=21cn.com>@21cn.com</OPTION>
                      <OPTION value=sohu.com>@sohu.com �Ѻ�</OPTION>
                      <OPTION value=tom.com>@tom.com</OPTION>
                      <OPTION 
        value=vip.163.com>@vip.163.com</OPTION>
                      <OPTION 
        value=263.net>@263.net</OPTION>
                      <OPTION 
        value=vip.sina.com>@vip.sina.com����VIP</OPTION>
                      <OPTION 
        value=mail.china.com>@mail.china.com</OPTION>
                      <OPTION 
        value=china.com>@china.com</OPTION>
                      <OPTION 
        value=netease.com>@netease.com</OPTION>
                      <OPTION 
        value=yeah.net>@yeah.net</OPTION>
                      <OPTION value=etang.com>@etang.com ����</OPTION>
                      <OPTION value=xinhuanet.com>@xinhuanet.com�»���</OPTION>
                      <OPTION 
        value=eyou.com>@eyou.com ����</OPTION>
                      <OPTION value=citiz.net>@citiz.net �Ϻ�����</OPTION>
                      <OPTION value=56.com>@56.com</OPTION>
                      <OPTION 
        value=188.com>@188.com</OPTION>
                      <OPTION 
        value=gmail.com>@gmail.com</OPTION>
                    </SELECT>
                    ���룺
                    <INPUT 
      style="FONT-SIZE: 12px" tabIndex=3 type=password size=13 name=pass>
                    <INPUT style="FONT-SIZE: 12px" onclick="setTimeout('clearpass()',1000)" tabIndex=4 type=submit value=��¼ name=Submit2>          </TD>
              </FORM>
            </TR>
          </TBODY>
        </TABLE>
        <center>
   
        </center>
        <br>
        <br>
        <br>
        <center>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>