<%@ LANGUAGE = VBScript %>
<%response.expires=0%>
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

xm=trim(request("xm"))
xb=request("xb")
zw=trim(request("zw"))
lb=trim(request("lb"))
dw=trim(request("dw"))
zzdh=trim(request("zzdh"))
yzbm=trim(request("yzbm"))
dhfj=trim(request("dhfj"))
sj=trim(request("sj"))
dzyj=trim(request("dzyj"))
hj=trim(request("hj"))
zzdz=trim(request("zzdz"))
cz=trim(request("cz"))
bz=trim(request("bz"))
%>
<html>
<head>
<meta http-equiv="expires" content="no-cache">
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<meta name="GENERATOR" content="Microsoft FrontPage 4.0">
<meta name="ProgId" content="FrontPage.Editor.Document">
<link rel="stylesheet" href="css/css.css">
<script language="javascript1.2" src="js/openwin.js"></script>
<title>OA办公系统</title>
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
              <td width="2" height="25"><span class="style2"><img src="images/main/l3.gif" width="2" height="25"></span></td>
              <td background="images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                  <tr>
                    <td width="21"><div align="center"><span class="style2"><img src="images/main/icon.gif" width="15" height="12"></span></div></td>
                    <td class="style7">个人通讯录</td>
                  </tr>
              </table></td>
              <td width="1"><span class="style2"><img src="images/main/r3.gif" width="1" height="25"></span></td>
            </tr>
          </table>
          <font color="0D79B3"></font></div></td>
    </tr>
  </table>
  <br>
  <table>
<tr>
<td>修改个人通讯录资料&nbsp;&nbsp;&nbsp;&nbsp;</td>
<form method="post" action="personlist.asp">
<td><input type="submit" value="返回">
</td>
</form>
</tr>
</table>
</center>
<%

id=request("id")
if id="" then
	response.redirect "personlist.asp"
	response.end
end if
if request("submit")="修改" then
if xm="" or dw="" then
	response.write("<script language=""javascript"">")
	response.write("alert(""姓名和单位不能为空！"");")
	response.write("history.go(-1);")
	response.write("</script>")
	response.end
end if
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personrecord where xm='"&xm&"' and thisinfousername='"&oabusyusername&"' and id<>"&id
rs.open sql,conn,1
if not rs.eof or not rs.bof then
	conn.close
	set rs=nothing
	response.write("<script language=""javascript"">")
	response.write("alert(""已有该用户的资料，请重新输入姓名！"");")
	response.write("history.go(-1);")
	response.write("</script>")
	response.end
end if
sql = "update personrecord set "
sql = sql & "thisinfousername="&SqlStr(oabusyusername) & ", "
sql = sql & "recordtype="& cstr(lb) & ", "
sql = sql & "xm="& SqlStr(xm) & ", "
sql = sql & "company="&SqlStr(dw) & ", "
sql = sql & "userzw="&SqlStr(zw) & ", "
sql = sql & "companytel="&SqlStr(dhfj) & ", "
sql = sql & "fax="&SqlStr(cz) & ", "
sql = sql & "hometel="&SqlStr(zzdh) & ", "
sql = sql & "email="&SqlStr(dzyj) & ", "
sql = sql & "homeaddress="&SqlStr(zzdz) & ", "
sql = sql & "postcard="&SqlStr(yzbm) & ", "
sql = sql & "sex="&SqlStr(xb) & ", "
sql = sql & "handset="&SqlStr(sj) & ", "
sql = sql & "callset="&SqlStr(hj) & ","
sql=sql & "remark="&sqlstr(bz)& "where id="&id
conn.Execute sql
conn.close
set rs=nothing
%>
<br><br>
<center>
<font color=red >
成功修改通讯录资料！    
</font>    
</center>    
<%    
else    
%>    
<script Language="JavaScript">    
    
 function form_check(){    
   var l1=document.form1.xm.value.length;    
   if(l1==0){window.alert("姓名为必填项！");document.form1.xm.focus();return (false);}    
    
   var l2=document.form1.dw.value.length;    
   if(l2==0){window.alert("单位为必填项！");document.form1.dw.focus();return (false);}    
                    }    
</script>    
<%
set conn=opendb("oabusy","conn","accessdsn")
set rs=server.createobject("adodb.recordset")
sql="select * from personrecord,persontype where personrecord.id="&id&" and persontype.id=personrecord.recordtype"
rs.open sql,conn,1
if rs.eof or rs.bof then
	conn.close
	set rs=nothing
	response.write("<font color=""#ff0000"" size=""+3"">")
	response.write("<center>错误，该条信息可能已经被删除！</center>")
	response.write("</font>")
	response.end
end if
if rs("thisinfousername")<>oabusyusername then
	conn.close
	set rs=nothing
	response.write("<font color=""#ff0000"" size=""+1"">")
	response.write("<center>对不起，您不能修改该条资料！</center>")
	response.write("</font>")
	response.end
end if
xm=server.htmlencode(rs("xm"))
xb=rs("sex")
zw=server.htmlencode(rs("userzw"))
lb=rs("recordtype")
dw=server.htmlencode(rs("company"))
zzdh=server.htmlencode(rs("hometel"))
yzbm=server.htmlencode(rs("postcard"))
dhfj=server.htmlencode(rs("companytel"))
sj=server.htmlencode(rs("handset"))
hj=server.htmlencode(rs("callset"))
dzyj=server.htmlencode(rs("email"))
zzdz=server.htmlencode(rs("homeaddress"))
cz=server.htmlencode(rs("fax"))
bz=server.htmlencode(rs("remark"))
%>
<center>    
<br>    
<form method="post" action="personeditrecord.asp?id=<%=id%>" name="form1" onsubmit="return form_check();">    
  <table width="540"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td height="1" bgcolor="4B789F"></td>
            </tr>
  </table><table border="0" cellpadding="0" cellspacing="0" width="540">    
    <tr>    
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">姓名</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="169"><input type=text name="xm" size=23 value="<%=xm%>"><font color=red>*</font></td>   
      <td width="84" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">性别</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 2 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="190"><select name="xb" size="1">   
          <option value="男">男</option>   
          <option value="女">女</option>   
        </select>
		<%
		response.write("<script language=""javascript"">")
		response.write("form1.xb.value="&chr(34)&xb&chr(34)&";")
		response.write("</script>")
		%>
		</td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">职务</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="169"><input type=text name="zw" size=23 value="<%=zw%>"></td>   
      <td width="84" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">通讯录类别</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="190"> 
		<select size="1" name="lb">   
        <% 
		set conn=opendb("oabusy","conn","accessdsn") 
		set rs=server.createobject("adodb.recordset") 
		sql="select * from persontype where username='"&oabusyusername&"'" 
		rs.open sql,conn,1 
		if rs.eof or rs.bof then 
			conn.close 
			set rs=nothing 
			response.write("<script language=""javascript"">") 
			response.write("location.href=""personaddtype.asp"";") 
			response.write("</script>") 
			response.end 
		end if 
		do while not rs.eof 
			if rs("id")=clng(lb) then
		%> 
		<option value="<%=rs("id")%>" selected><%=rs("typename")%></option> 
		<%
			else
		%>
		<option value="<%=rs("id")%>"><%=rs("typename")%></option> 
		<% 
			end if
			rs.movenext 
		loop 
		conn.close 
		set rs=nothing 
		%> 
		</select> 
		</td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">单位</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="dw" size=60 value="<%=dw%>"><font color=red>*</font></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">住宅电话</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="169"><input type=text name="zzdh" size=23 value="<%=zzdh%>"></td>   
      <td width="84" align=center bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">邮政编码</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="190"><input type=text name="yzbm" size=23  maxlength="6" value="<%=yzbm%>"></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">电话或分机</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="169"><input type=text name="dhfj" size=23 value="<%=dhfj%>"></td>   
      <td width="84" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">   
        <p align="center">手&nbsp;&nbsp;&nbsp;&nbsp;机</p>      </td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="190"><input type=text name="sj" size=23 value="<%=sj%>"></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">M S N</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="169"><input type=text name="hj" size=23 value="<%=hj%>"></td>   
      <td width="84" bgcolor="D7E8F8" style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">   
        <p align="center">Email</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" width="190"><input type=text name="dzyj" size=23 value="<%=dzyj%>"></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">住宅地址</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="zzdz" size=50 value="<%=zzdz%>"></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA">传真</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 1 solid #B0C8EA" colspan="3"><input type=text name="cz" size=50 value="<%=cz%>"></td>   
    </tr>   
    <tr>   
      <td width="97" height="25" align=center bgcolor="D7E8F8" style="border-left: 2 solid #B0C8EA; border-right: 1 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA">备注说明</td>   
      <td style="border-left: 0 solid #B0C8EA; border-right: 2 solid #B0C8EA; border-top: 0 solid #B0C8EA; border-bottom: 2 solid #B0C8EA" colspan="3"><textarea rows="4" cols="59" name="bz"><%=bz%></textarea></td>   
    </tr>   
  </table>   
   
  <font color=red>*</font>必须填写&nbsp;&nbsp;&nbsp;&nbsp;<input type="submit" name="submit" value="修改">&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value="返回" onclick="window.location.href='personlist.asp';">   
</form>   
</center>   
<%   
end if   
   
%>   
</body>   
</html>