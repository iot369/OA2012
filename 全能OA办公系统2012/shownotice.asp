<%@ LANGUAGE = VBScript %>
<!--#include file="asp/keepformat.asp"-->
<!--#include file="asp/opendb.asp"-->

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
<title>查看公司通告</title>
<style type="text/css">
<!--
.z14 {	font-size: 14px;
	font-weight: bold;
	color: #098abb;
}
.style5 {color: #2d4865}
.style8 {color: #2b486a}
-->
</style>
</head>
<body  topmargin="0" leftmargin="0">

<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
      <tr>
        <td width="2" height="25"><img src="images/main/l3.gif" width="2" height="25"></td>
        <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td width="21"><div align="center"><img src="images/main/icon.gif" width="15" height="12"></div></td>
              <td><span class="style5">公司通告</span></td>
            </tr>
        </table></td>
        <td width="1"><img src="images/main/r3.gif" width="1" height="25"></td>
      </tr>
    </table></td>
  </tr>

  <tr>
    <td><div align="center">
        <table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td><center>
                <table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td height="30"><%
set conn=opendb("oabusy","conn","accessdsn")
Set rs=Server.CreateObject("ADODB.recordset")
sql="select * from newnotice where id=" & request("id")
rs.open sql,conn,1
%>
                      <table width="550"  border="0" align="center" cellpadding="0" cellspacing="0">
                        <tr>
                          <td width="350" height="50">&nbsp;</td>
                          <td width="70"><div align="right"><a href="noticelook.asp"><img src="images/bt_tonggaoliebiao.gif" width="58" height="18" border="0"></a></div></td>
                          <td width="70"><div align="right"><a href="newnotice.asp"><img src="images/bt_fabutonggao.gif" width="58" height="18" border="0"></a></div></td>
                          <td width="70"><div align="right"><a href="noticecontrol.asp"><img src="images/bt_guanlitonggao.gif" width="58" height="18" border="0"></a></div></td>
                        </tr>
                      </table></td>
                  </tr>
                </table>
                <table width="550"  border="0" cellspacing="0" cellpadding="0">
                    <tr>
                    <td height="1" bgcolor="B0C8EA"></td>
                  </tr>
				  <tr>
                    <td height="30"><div align="center" class="z14 style8">通告标题：<%=rs("title")%></div></td>
                  </tr>
                  <tr>
                    <td height="1" bgcolor="B0C8EA"></td>
                  </tr>
                <tr>
                    <td height="15" bgcolor="D7E8F8"><table width="96%"  border="0" align="center" cellpadding="0" cellspacing="0">
                      <tr>
                        <td>&nbsp;</td>
                      </tr>
                      <tr>
                        <td><span class="style8"><%=keepformat(rs("content"))%>
                            <%

%>
                        </span></td>
                      </tr>
                    </table></td>
                  </tr>
                <tr>
                  <td height="30" bgcolor="D7E8F8"><div align="right"><span class="style8">[发布日期：<%=rs("noticedate")%>] </span>　</div></td>
                </tr>
			      <tr>
                    <td height="1" bgcolor="B0C8EA"></td>
                  </tr>
				</table>
                <br>
            </center>
                <center>              
                </center></td>
          </tr>
        </table>
        <center>
        </center>
    </div></td>
  </tr>
</table>
</body>
</html>