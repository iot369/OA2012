<%
info=request("info")
title=request("title")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>即时消息窗口</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body topmargin="1" leftmargin="1" scroll=no>
<bgsound src="../msg.wav" loop="1">
<table border="1" cellpadding="0" cellspacing="0" width="150">
  <tr>
    <td height="150" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" height="93">
        <tr>
          <td height="35">
            <p align="center"><img border="0" src="../image/msg.gif" width="34" height="32"><b><font color="#FF0000"><%=title%></font></b></td>
        </tr>
        <tr>
          <td height="58">&nbsp;&nbsp;&nbsp;&nbsp;<%=info%></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<script language="javascript">
window.focus();
setTimeout("window.close()",5000);
</script>
</body>
</html>
