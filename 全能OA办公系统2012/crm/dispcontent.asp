<%response.expires=0%>
<!--#include file="conn.asp"-->
<%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
if oabusyusername="" then 
	response.write("<script language=""javascript"">")
	response.write("window.top.location.href='index.asp';")
	response.write("</script>")
	response.end
end if
id=request("id")
printflag=request("printflag")
if id="" then
	response.write("<script language=""javascript"">")
	response.write("alert(""对不起，数据传输出错！"");")
	response.write("window.close();")
	response.write("</script>")
else
	set conn=dbconn("conn")
	set rs=server.createobject("adodb.recordset")
	sql="select * from qiye where id="&id
	rs.open sql,conn,1
	if rs.eof and rs.bof then
		conn.close
		set rs=nothing
		set conn=nothing
		response.write("<script language=""javascript"">")
		response.write("alert(""对不起，没有找到对应的记录！"");")
		response.write("window.close();")
		response.write("</script>")
	else
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
<link rel="stylesheet" type="text/css" href="../css/css.css">
</head>

<body bgcolor="#ffffff" topmargin="5" leftmargin="5">
<div align="center">
<br><b><font color="red" style="font-family:黑体;font-size:24px">企业名录</font></b><br><br>
  <table width="365" border="0" cellspacing="0" cellpadding="0">
    <tr> 
      <td width="360" height="200" valign="top">
        <table width="460" border="1" cellspacing="0" cellpadding="0" height="200">
          <tr bordercolor="#000000">
            <td valign="top">
              <div align="right">
                <table border="0" cellpadding="0" cellspacing="0" width="100%" height="75">
                  <tr>
                    <td width="50%" height="34" align="center"><br><font color="red" style="font-family:黑体;font-size:16px"><%=server.htmlencode(trim(rs("企业名称")))%></font></td>
                    <td width="50%" rowspan="2" height="75" align="center" valign="top">
<div align="center">
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="25%" valign="top">
        <p align="right"><br>
<%
		if not isnull(rs("production")) then
%><b>产品：</b>
<%    
		end if
%>
        </p>
      </td>
      <td width="75%" valign="top"><br>
<%
		if not isnull(rs("production")) then
			response.write(server.htmlencode(rs("production")))
		end if
%>
      </td>
    </tr>
  </table>
</div>
					</td>
                  </tr>
                  <tr>
                    <td width="50%" height="41" align="center">
<%
		if not isnull(rs("contact")) then
			response.write("<b>联系人：</b>"&server.htmlencode(rs("contact")))
		end if
%>
					</td>
                  </tr>
                </table>
              </div>
              <div align="right">
                <table border="0" cellpadding="0" cellspacing="0" width="100%">
                  <tr>
                    <td width="100%" height="25" colspan="2">
                      <hr color="#000000">
                    </td>
                  </tr>
                  <tr>
                    <td width="100%" height="25" colspan="2">
                      <div align="right">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr>
                            <td width="11%" valign="top">&nbsp;<b>地址：</b></td>
                            <td width="89%" valign="top">
<%       
		if not isnull(rs("address")) then
			response.write(server.htmlencode(rs("address")))
		end if
%>                            </td>
                          </tr>
                        </table>
                      </div>
					</td>
                  </tr>
                  <tr>
                    <td width="50%" height="25">
                      <div align="right">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr>
                            <td width="22%" valign="top">&nbsp;<b>电话：</b></td>
                            <td width="78%" valign="top">
<%     
		if not isnull(rs("phone")) then
			response.write(server.htmlencode(rs("phone")))
		end if
%>                            
                            </td>
                          </tr>
                        </table>
                      </div>
					</td>
                    <td width="50%" height="25">&nbsp;<b>传真：</b>
<%         
		if not isnull(rs("fax")) then
			response.write(server.htmlencode(rs("fax")))
		end if
%>					
					</td>
                  </tr>
                  <tr>
                    <td width="50%" height="25">
                      <div align="right">
                        <table border="0" cellpadding="0" cellspacing="0" width="100%">
                          <tr>
                            <td width="23%" valign="top">&nbsp;<b>Email：</b></td>
                            <td width="77%" valign="top">
<%       
		if not isnull(rs("email")) then
			response.write(server.htmlencode(rs("email")))
		end if
%>                            
                            </td>
                          </tr>
                        </table>
                      </div>
					</td>
                    <td width="50%" height="25">&nbsp;<b>邮编：</b>
<%         
		if not isnull(rs("postcode")) then
			response.write(server.htmlencode(rs("postcode")))
		end if
%>					
					</td>
					</td>
</tr>
                    <td width="50%" height="25">&nbsp;<b>网站：</b>
<%         
		if not isnull(rs("web")) then
%><A HREF="<%=rs("web")%>" target=_blank><%
			response.write(server.htmlencode(rs("web")))
		end if
%>					
					</td>

                  </tr>
                  <tr>
                    <td width="100%" height="25" colspan="2">
<div align="right">
  <table border="0" cellpadding="0" cellspacing="0" width="100%">
    <tr>
      <td width="11%" valign="top">&nbsp;<b>备注：</b></td>
      <td width="89%" valign="top">
<%
		if not isnull(rs("other")) then
			response.write(server.htmlencode(rs("other")))
		end if
%>					
      </td>

    </tr>
    </tr>
      <td width="11%" valign="top">&nbsp;<b>附件：</b></td>
      <td width="89%" valign="top">
<%
		if not isnull(rs("iaddfile")) then
if rs("iaddfile")<>"无" then 
  response.write "<a target=_blank href=" & rs("iaddfile") & ">" & rs("iaddfile") & "</a>&nbsp;"
end if
if rs("iaddfile")="无" then 
response.write(server.htmlencode(rs("iaddfile")))
end if		
		end if
%>					
      </td>
  </table>
</div>
					</td>
                  </tr>
                </table>
              </div>
            </td>
          </tr>
        </table>
      </td>
      <td width="5" valign="top" bgcolor="#E0E0E0"><img src="images/dot.gif" width="6" height="10"></td>
    </tr>
    <tr> 
      <td colspan="2" height="5" bgcolor="#E0E0E0"><img src="images/dot.gif" width="10" height="6"></td>
    </tr>
  </table>
<%
	if printflag="0" or printflag="" then
%>
<br>
<input type="button" value=" 打 印 " onclick="javascript:location.href='dispcontent.asp?printflag=1&id=<%=id%>'">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<input type="button" value=" 关 闭 " onclick="javascript:window.close();">
<%
	else
%>
<script language="javascript">
if (confirm('请单击“确定”按钮开始打印！'))
	window.print();
else
	history.go(-1);
</script>
<%
	end if
%>
</div>
</body>
</html>
<%
		conn.close
		set rs=nothing
		set conn=nothing
	end if
end if
%>
