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
<title>���۹���ϵͳ</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function showHideHead(strSrc)
{
	var strFile = strSrc.substring(strSrc.lastIndexOf("/"),strSrc.length);
    if (strFile == "/arrow_up.gif"){
	    oHead.style.display = "none";
		oHeadCtrl.src = "images/arrow_down.gif";
		oHeadCtrl.alt = "��ʾͷ��";
		oHeadBar.title = "��ʾͷ��";		
	}
	else {
	    oHead.style.display = "block";
		oHeadCtrl.src = "images/arrow_up.gif";
		oHeadCtrl.alt = "����ͷ��";
		oHeadBar.title = "����ͷ��";
	}
}
-->
</script>
</head>

<body>
<table width="100%" border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="5"><img src="images/null.gif" width="1" height="1"></td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0" bgcolor="menu">
  <tr> 
	<td align="left" background='images/tab_top_background_runner.gif'> <table width="5" border="0" align="left" cellpadding="0" cellspacing="0">
        <tr>
          <td><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr > 
          <td><img src="images/tab_on_left.gif" width="16" height="24"></td>
          <td background="images/tab_on_middle.gif">�鿴��������</td>
          <td><img src="images/tab_on_right.gif" width="16" height="24"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr>   
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">�������</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>	  
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr> 
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">�߼�����</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr> 
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">���ݱ���</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr> 
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">���ݵ���</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr> 
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">�û�����</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>
      <table onclick="window.location.replace('')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
        <tr>    
          <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
          <td background="images/tab_off_middle.gif">ע��</td>
          <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
        </tr>
      </table>      
    </td>
  </tr>
</table>
<table width="100%" border="0" cellpadding="0" cellspacing="0">
  <tr> 
    <td height="5"><img src="images/null.gif" width="1" height="1"></td>
  </tr>
  <tr>
    <td bgcolor="#999999">&nbsp;</td>
  </tr>
</table>
  <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
    <tr> 
      <td width="40" align="right">&nbsp;</td>
	  <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <form name="searchCom" method="post" action="?action=com">
		  <tr>
            <td>����˾���������� 
              <input name="searchCom" type="text" id="searchCom" size="24" maxlength="36">
              <input name="Search" type="submit" id="Search" value="����"></td>
          </tr></form>
        </table></td>
    </tr>
    <tr> 
      <td width="40" align="right">&nbsp;</td>
      <td><table width="100%" border="0" cellspacing="0" cellpadding="0">
          <form name="searchUrl" method="post" action="?action=url">
		  <tr>
            <td>����˾��ַ������ 
              <input name="searchUrl" type="text" id="searchUrl" size="24" maxlength="36">
              <input name="Submit" type="submit" id="Submit" value="����"></td>
          </tr></form>
        </table></td>
    </tr>
  </table>


</body>
</html>
