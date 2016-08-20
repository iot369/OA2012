<%
'session.abandon
'Server.ScriptTimeOut=500
function opendb(DBPath,sessionname,dbsort)
dim conn
'if not isobject(session(sessionname)) then
Set conn=Server.CreateObject("ADODB.Connection")
'if dbsort="accessdsn" then conn.Open "DSN=" & DBPath
'if dbsort="access" then conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath 
'if dbsort="sqlserver" then conn.Open "DSN=" & DBPath & ";uid=wsw;pwd=wsw"
DBPath1=server.mappath("../db/lmtof.mdb")
conn.Open "driver={Microsoft Access Driver (*.mdb)};dbq=" & DBPath1
set session(sessionname)=conn
'end if
set opendb=session(sessionname)
end function
%>
<%
'-----------------------------------------
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
%>
<html>
<head>
<title>
on-line meeting
</title>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">

<style type="TEXT/CSS"> 
<!--
body,table {color:#202030;font-family: 宋体_GB2312; font-size: 9pt; line-height: 12pt}
A:link {text-decoration: none; color:#6060FF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:visited {text-decoration: none; color: #6060FF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:active {text-decoration: underline; color: #FF6000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A:hover {text-decoration: underline; color:#FF6000; font-family: "宋体"; font-size: 9pt; line-height: 12pt}

A.B:link {text-decoration: none; color:#EFEFEF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A.B:visited {text-decoration: none; color: #EFEFEF; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A.B:active {text-decoration: underline; color: #EFEFFE; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
A.B:hover {text-decoration: underline; color:#EFEFFE; font-family: "宋体"; font-size: 9pt; line-height: 12pt}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}

-->
</style>

</head>

<%
LoginID=Session("LoginID")
if LoginID="" then LoginID=0
%>

<body style='margin:0;background:#F9F9FF' >
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">公共服务</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<table width="583" height="500">
<tr><td>
<form action="enter.asp" >
        <table border="0" align="center" bgcolor="D8F7FF">
          <tr align="center" bgcolor="D8F7FF"> 
            <td colspan="2"><font color="#308080"><Big><a href="index.asp">一般会议</a></Big></font></td>
          </tr>
          <tr bgcolor="#6FECFF"> 
            <td> 用户名:</td>
            <td> 
            <input type="text" name="UserName" value="<%=oabusyname%>">            </td>
          </tr>
          <tr bgcolor="#6FECFF"> 
            <td> 房间:</td>
            <td> 
              <select name="CRNAME">
                <option value="Room1">房间一</option>
                <option value="Room2">房间二</option>
                <option value="Room3">房间三</option>
                <option value="Room4">房间四</option>
            </select>            </td>
          </tr>
          <tr align="center" bgcolor="#6FECFF"> 
            <td colspan="2"> 
            <input type="submit" value=" 进 入 ">            </td>
          </tr>
        </table>        
        <br>
        <table border="0" cellspacing="0" cellpadding="0" align="center" width="180">
          <tr> 
            <td> 
              <div align="center"><font color="#308080">……</font><font color="#308080">……<big><a href="Meeting.asp">进入加密会议</a></big>…………</font></div>
            </td>
          </tr>
        </table>
      </form>
</td></tr></table>      
</body>
</html>
<%Response.End%>

