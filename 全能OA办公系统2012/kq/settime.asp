<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
<!--#include file="conn.asp"-->
<!--#include file="../asp/bgsub.asp"-->
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
oabusyname=request.cookies("oabusyname")
oabusyuserid=request.cookies("oabusyuserid")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" or oabusyuserid="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='../default.asp';")
	response.write("</script>")
	response.end
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from userinf where username='"&oabusyusername&"'"
rs.open sql,conn,1
if not rs.eof and not rs.bof then
	allow_edit_work_time=rs("allow_edit_work_time")
	if allow_edit_work_time="no" then
		response.write("<script language=""javascript"">")
		response.write("alert(""�Բ������������ÿ���ʱ�䣡"");")
		response.write("</script>")
		response.end
	end if
else
	if allow_edit_work_time="no" then
		response.write("<script language=""javascript"">")
		response.write("alert(""�Բ���û�е��ҵ���Ӧ���û���"");")
		response.write("</script>")
		response.end
	end if
end if
conn.close
set conn=nothing
set rs=nothing
%>
<html>

<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<link rel="stylesheet" href="../css/css.css">
<script language="javascript">
function checkform()
{
	var time1,time2,value1,value2;
	time1=document.form1.amgohour.value;
	time2=document.form1.pmcomehour.value;
	value1=document.form1.amgominute.value;
	value2=document.form1.pmcomeminute.value;
	if (time1==time2)
		{
			if ((value1==value2) || (value1=="30" && value2=="0"))
				{
					alert("�����°�ʱ���������ϰ�ʱ���ͻ��");
					return (false);
				}
		}
	return (true);
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
              <td width="2" height="25"><span class="style2"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
              <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">����ϵͳ</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
<tr>
<td>���ÿ���ʱ��</td>
</tr>
</table>
</center>

<br> 
<center>
<%
if request("submit")="ȷ��" then
	amcomehour=request.form("amcomehour")
	amcomeminute=request.form("amcomeminute")
	amgohour=request.form("amgohour")
	amgominute=request.form("amgominute")
	pmcomehour=request.form("pmcomehour")
	pmcomeminute=request.form("pmcomeminute")
	pmgohour=request.form("pmgohour")
	pmgominute=request.form("pmgominute")
	comedelaytime=request.form("comedelaytime")
	goaheadtime=request.form("goaheadtime")
	kqtimephase=request.form("kqtimephase")
	'checkvalue=request.form("check")
	amgonokq=request.form("amgonokq")
	if amgonokq="" then amgonokq="0"
	pmcomenokq=request.form("pmcomenokq")
	if pmcomenokq="" then pmcomenokq="0"
	pmgonokq=request.form("pmgonokq")
	if pmgonokq="" then pmgonokq="0"
	set conn=openconn("kqconn")
	amcometime=timeserial(cint(amcomehour),cint(amcomeminute),0)
	amgotime=timeserial(cint(amgohour),cint(amgominute),0)
	pmcometime=timeserial(cint(pmcomehour),cint(pmcomeminute),0)
	pmgotime=timeserial(cint(pmgohour),cint(pmgominute),0)
	sql="update inittime set amondutytime=#"&amcometime&"#,amoffdutytime=#"&amgotime&"#,pmondutytime=#"&pmcometime&"#,pmoffdutytime=#"&pmgotime&"#,ondutydelaytime="&comedelaytime&",offdutyaheadtime="&goaheadtime&",kqtimephase="&kqtimephase&",amgonokq="&amgonokq&",pmcomenokq="&pmcomenokq&",pmgonokq="&pmgonokq
	conn.execute(sql)
	response.write("<p align=""center""><font color=""#dd0000"">�ɹ����ÿ���ʱ�䣡</font>")
	conn.close
		
	response.end
end if
set conn=openconn("kqconn")
set rs=server.createobject("adodb.recordset")
sql="select * from inittime"
rs.open sql,conn,1
%>
<form method="POST" action="settime.asp" onsubmit="return checkform();" name="form1">
  <div align="center">
    <center>
    <table border="1" width="540" cellspacing="0" cellpadding="0" bordercolorlight="#B0C8EA" bordercolordark="#FFFFFF">
      <tr bgcolor="D7E8F8">
        <td width="496" height="40" colspan="2"><font color="#DD0000">ע�⣺</font>�ϰ��ӳ�ʱ����°���ǰʱ�䲻��Ϊ0���ӣ�����������ֵ����������Ϊ���ڵĻ���ʱ�䡣</td>
      </tr>
      <tr>
<%
amcomehour=hour(rs("amondutytime"))
amcomeminute=minute(rs("amondutytime"))
%>
        <td width="246" height="30" bgcolor="#FFFFFF">�����׼�ϰ�ʱ�䣺<select size="1" name="amcomehour">
            <option value="6">6��</option>
            <option value="7">7��</option>
            <option value="8">8��</option>
            <option selected value="9">9��</option>
          </select><select size="1" name="amcomeminute">
            <option selected value="0">00��</option>
            <option value="30">30��</option>
          </select>
<script language="javascript">
document.form1.amcomehour.value=<%=amcomehour%>;
document.form1.amcomeminute.value=<%=amcomeminute%>;
</script>
		  </td>
<%
amgohour=hour(rs("amoffdutytime"))
amgominute=minute(rs("amoffdutytime"))
%>
        <td width="248" height="30" bgcolor="#FFFFFF">�����׼�°�ʱ�䣺<select size="1" name="amgohour">
            <option value="11">11��</option>
            <option value="12" selected>12��</option>
            <option value="13">13��</option>
          </select><select size="1" name="amgominute">
            <option selected value="0">00��</option>
            <option value="30">30��</option>
          </select>
<script language="javascript">
document.form1.amgohour.value=<%=amgohour%>;
document.form1.amgominute.value=<%=amgominute%>;
</script>
		  </td>
      </tr>
      <tr>
<%
pmcomehour=hour(rs("pmondutytime"))
pmcomeminute=minute(rs("pmondutytime"))
%>
        <td width="246" height="30" bgcolor="#FFFFFF">�����׼�ϰ�ʱ�䣺<select size="1" name="pmcomehour">
            <option value="13">13��</option>
            <option value="14" selected>14��</option>
            <option value="15">15��</option>
          </select><select size="1" name="pmcomeminute">
            <option selected value="0">00��</option>
            <option value="30">30��</option>
          </select>
<script language="javascript">
document.form1.pmcomehour.value=<%=pmcomehour%>;
document.form1.pmcomeminute.value=<%=pmcomeminute%>;
</script>
		  </td>
<%
pmgohour=hour(rs("pmoffdutytime"))
pmgominute=minute(rs("pmoffdutytime"))
%>
        <td width="248" height="30" bgcolor="#FFFFFF">�����׼�°�ʱ�䣺<select size="1" name="pmgohour">
            <option value="16">16��</option>
            <option value="17" selected>17��</option>
            <option value="18">18��</option>
            <option value="19">19��</option>
          </select><select size="1" name="pmgominute">
            <option selected value="0">00��</option>
            <option value="30">30��</option>
          </select>
<script language="javascript">
document.form1.pmgohour.value=<%=pmgohour%>;
document.form1.pmgominute.value=<%=pmgominute%>;
</script>
		  </td>
      </tr>
      <tr>
        <td width="246" height="30" bgcolor="#FFFFFF">�ϰ࿼���ӳ�ʱ�䣺<select size="1" name="comedelaytime">
			<option value="0">0����</option>
            <option value="10">10����</option>
            <option value="15">15����</option>
            <option value="20">20����</option>
            <option value="25">25����</option>
            <option value="30">30����</option>
            <option value="35">35����</option>
            <option value="40">40����</option>
            <option value="45">45����</option>
            <option value="50">50����</option>
            <option value="55">55����</option>
          </select>
<script language="javascript">
document.form1.comedelaytime.value=<%=rs("ondutydelaytime")%>;
</script>
		  </td>
        <td width="248" height="30" bgcolor="#FFFFFF">�°࿼����ǰʱ�䣺<select size="1" name="goaheadtime">
			<option value="0">0����</option>
            <option value="10">10����</option>
            <option value="15">15����</option>
            <option value="20">20����</option>
            <option value="25">25����</option>
            <option value="30">30����</option>
            <option value="35">35����</option>
            <option value="40">40����</option>
            <option value="45">45����</option>
            <option value="50">50����</option>
            <option value="55">55����</option>
          </select>
<script language="javascript">
document.form1.goaheadtime.value=<%=rs("offdutyaheadtime")%>;
</script>
		  </td>
      </tr>
	  <tr>
        <td width="100%" height="30" bgcolor="#FFFFFF" colspan="2">����ʱ��Σ�<select size="1" name="kqtimephase">
            <option value="10">10����</option>
            <option value="15">15����</option>
            <option value="20" selected>20����</option>
            <option value="25">25����</option>
            <option value="30">30����</option>
            <option value="35">35����</option>
            <option value="40">40����</option>
            <option value="45">45����</option>
            <option value="50">50����</option>
            <option value="55">55����</option>
          </select>
<script language="javascript">
document.form1.kqtimephase.value=<%=rs("kqtimephase")%>;
</script>
		  </td>
	  </tr>
	  <tr>
        <td width="100%" height="30" bgcolor="#FFFFFF" colspan="2">
		<input type="checkbox" name="amgonokq" value=1>�����°಻����<input type="checkbox" name="pmcomenokq" value=1>�����ϰ಻����<input type="checkbox" name="pmgonokq" value=1>�����°಻����
<script language="javascript">
<%
if rs("amgonokq")=1 then
	response.write("document.form1.amgonokq.checked=true;")
end if
if rs("pmcomenokq")=1 then
	response.write("document.form1.pmcomenokq.checked=true;")
end if
if rs("pmgonokq")=1 then
	response.write("document.form1.pmgonokq.checked=true;")
end if
%>
</script>
		</td>
	</tr>
    </table>
    </center>
  </div>
  <p align="center"><input type="submit" value="ȷ��" name="submit">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;  
  <input type="reset" value="����" name="resetbutton"></p> 
</form> </center>
<%
set rs=nothing
conn.close
set conn=nothing

%>
</body>
</html>
