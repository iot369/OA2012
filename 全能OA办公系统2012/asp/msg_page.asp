<%
info=request("info")
title=request("title")
%>
<html>
<head>
<meta http-equiv="Content-Language" content="zh-cn">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>消息通知</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>
<body topmargin="1" leftmargin="1" scroll=no>
<bgsound src="../msg.wav" loop="1">
<table border="0" cellpadding="0" cellspacing="0" width="150" height="83"  >
  <tr>
    <td height="150" bordercolorlight="#C0C0C0" bordercolordark="#FFFFFF">
      <table border="0" cellpadding="0" cellspacing="0" width="100%" height="121" >
        <tr>
          <td height="44"> 
            <p align="center"><img border="0" src="../image/msg.gif" width="34" height="32"><font color="#FF0000"><%=title%></font></td>
        </tr>
        <tr>
          <td height="100">&nbsp;&nbsp;&nbsp;&nbsp;<%=info%><p>&nbsp;</p></td>
        </tr>
      </table>
    </td>
  </tr>
</table>
<script language="javascript">
window.focus();
setTimeout("window.close()",8000);
</script>
</body>
</html>
