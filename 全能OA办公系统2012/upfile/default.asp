<!--#include file='connect.asp'-->
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
<title>�ϴ��ļ�ҳ��</title>
<style>
body		{font-size:9pt}
td			{font-size:9pt}
input		{font-size:9pt}
textarea	{font-size:9pt}
.style2 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
</style>
<script language='Jscript'>
function checkSub(src){
	if(src.filename.value==''){
		alert("�ļ������������!");
		src.filename.focus();
		return (false)
	}
}
</script>
</head>
<body style='margin:0;background:#F9F9FF'>
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">��������</td>
                </tr>
            </table></td>
            <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<table border='0' width='540' align='center' cellspacing="0" cellpadding="0">
  <tr height='100' bgcolor='#003399'> 
    <td align='cenater' bgcolor="#F9F9FF" height="45" > <div align="center"><font color="#0066FF" size="3"><b><font color="#FF0000">�� 
      �� �� ��</font></b></font></div></td>
  </tr>
  <tr> 
    <td valign='top' bgcolor='#F9F9FF'><br> 
      &nbsp;&nbsp;�ļ��б�(�����ļ�ֻ����30��):<br> 
      <hr size=0 width='96%' align='center'> <table border='0' width='90%' align='center' cellspacing="2" cellpadding="0">
        <%
	Dim Sql
	Dim Rs
	Sql="select top 100 * from upfile_table order by id desc"
	set Rs=Server.CreateObject("Adodb.Recordset")
	Rs.Open Sql,conn,1,1
	if not Rs.EOF then
		while not Rs.EOF
			Response.Write "<tr height='20'><td width='10%' nowrap>����:</td><td>"& Rs("id") &"��"& Rs("Subject") &" ["& Rs("filesize") &" Bytes]</td></tr>"
			Response.Write "<tr height='20'><td width='10%' nowrap>�ļ�:</td><td><a href='"& Rs("filePath") &"/"& Rs("Filename") &"' target=main_wanglongdai>"& Rs("Filename") &"</a></td></tr>"
			Response.Write "<tr height='20'><td width='10%' nowrap>���:</td><td>"& Rs("Expit") &"</td></tr>"
			Response.Write "<tr height='1' bgcolor='#003399'><td colspan='2'></td></tr>"
		Rs.MoveNext
		wend
	else
		Response.Write "<tr><td>�޼�¼</td></tr>"
	end if
	Rs.Close
	set Rs=nothing
	Conn.close
	set conn=nothing
	%>
    </table>    </td>
  </tr>
  <tr bgcolor='#003399' height='1'> 
    <td bgcolor="#F9F9FF"></td>
  </tr>
  <tr> 
    <td bgcolor='#F9F9FF'>
      <div align="center"><br>
        &nbsp;��ҳ���趨���ϴ��ļ���СΪ<font color='red'> 5M </font>����,������Ѿ�����ͬ���ļ�������(����ɰ����趨) 
      </div>
      <table border='0' align="center" cellpadding="5">
        <form method="POST" action="upfile.asp" enctype="multipart/form-data" id="form1" name="form1" onsubmit='return checkSub(this)'>
          <tr>
            <td>����</td>
            <td><input type='radio' name='errnumber' value='0'>
              �Զ�����&nbsp;
              <input type='radio' name='errnumber' value='1' checked>
              ����ʽ&nbsp;
              <input type='radio' name='errnumber' value='2'>
              ֱ�Ӹ���</td>
          </tr>
          <tr>
            <td>���⣺</td>
            <td><input type='text' name='filename' size='30'></td>
          </tr>
          <tr>
            <td>�ļ���</td>
            <td><input type="file" name="fruit" size="30"></td>
          </tr>
          <tr>
            <td valign='top'>��飺</td>
            <td><textarea name='fileExt' cols='40' rows='5'></textarea></td>
          </tr>
          <tr>
            <td colspan='2'><input type="submit" value="�ϴ��ļ�" name="subbutt"></td>
          </tr>
        </form>
      </table>
    <div align="center"></div>    </td>
  </tr>
</table>
</body>
</html>