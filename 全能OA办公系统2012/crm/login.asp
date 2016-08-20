<%@LANGUAGE="VBSCRIPT" CODEPAGE="936"%>
<!--#include file="../asp/opendb.asp"-->
<!--#include file="../asp/sqlstr.asp"-->
<!--#include file="Connections/conn.asp" -->

    <%
oabusyname=request.cookies("oabusyname")
oabusyusername=request.cookies("oabusyusername")
oabusyuserdept=request.cookies("oabusyuserdept")
oabusyuserlevel=request.cookies("oabusyuserlevel")
%>
<html>
<style type="text/css">
<!--
.box {
	color: #FFFFFF;
	background-color: #FFFFFF;
	height: 1px;
	width: 1px;
	border: #FFFFFF;
}
-->
</style>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
<script language="JavaScript">
<!--
if (this.location.href == top.location.href){
    top.location.href = "";
}

function checkInput()
{
    var strItem1 = document.all.item1.value;
	var strItem2 = document.all.item2.value;
	if (strItem1 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">您没有权限进入</font>"
		document.all.item1.focus();
		return false;
	}
	if (strItem2 == ""){
	    document.all.errMsg.innerHTML = "<font color=\"#FF0000\">您没有权限进入</font>"
		document.all.item2.focus();
		return false;
	}
}

function checkChr(chr)
{
    var str0 = "1234567890abcdefghijklmnopqrstuvwxyzABCDEFGHIJKLMNOPQRSTUVWXYZ";
	var intLen = chr.length;
	for (var i=0;i<intLen;i++){
	    
	}
}
-->
</script>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<style type="text/css">

.style1 {color: #0d79b3}
.style2 {color: #FF0000}
.style8 {color: #0d79b3;
	font-weight: bold;
}
.style7 {color: #2d4865}
</style>
</head>

<body  topmargin="0" leftmargin="0">
<table width="583"  border="0" cellspacing="0" cellpadding="0">
  <tr>
    <td height="21"><div align="center">
        <table width="100%"  border="0" align="center" cellpadding="0" cellspacing="0">
          <tr>
            <td width="2" height="25"><span class="style8"><img src="../images/main/l3.gif" width="2" height="25"></span></td>
            <td background="../images/main/m3.gif"><table width="100%"  border="0" cellspacing="0" cellpadding="0">
                <tr>
                  <td width="21"><div align="center"><span class="style8"><img src="../images/main/icon.gif" width="15" height="12"></span></div></td>
                  <td class="style7">销售系统</td>
                </tr>
            </table></td>
            <td width="1"><span class="style8"><img src="../images/main/r3.gif" width="1" height="25"></span></td>
          </tr>
        </table>
        <font color="0D79B3"></font></div></td>
  </tr>
</table>
<br>
<td align="center">
</td>
<%
If Session("CRM_account") <> "" And Session("CRM_name") <> "" And Session("CRM_level") > 0 Then Response.Redirect("listAll.asp")
Dim action
action = Trim(Request("action"))
Select Case action
Case "login"
    Call login()
Case Else
    Call loginForm()
End Select

Sub login()
    Dim account,password
	account = Trim(Request("item1"))
	password = Trim(Request("item2"))
	If account = "" Or password = "" Then
	    Response.Redirect("login.asp?errMsg=2")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uAccount = '" & account & "'",conn,3,1
	If rs.RecordCount <> 1 Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If password <> rs("uPassword") Then
	    Response.Redirect("login.asp?errMsg=1")
		Response.End()
	End If
	If rs("uBlock") = True Then
	    Response.Redirect("login.asp?errMsg=3")
		Response.End()
	End If
	Session("CRM_account") = account
	Session("CRM_name") = rs("uName")
	Session("CRM_level") = rs("uLevel")
	Session("CRM_group") = rs("uGroup")
	rs.Close
	Set rs = Nothing
	Response.Redirect("listAll.asp")
End Sub

Sub loginForm()
    Dim errMsg
	errMsg = CInt(ABS(Request("errMsg")))
	Select Case errMsg
	Case 2
	    errMsg = "<font color=""#FF0000"">对不起,您不是业务部员工,您没有使用销售系统的权限</font>"
	Case 1
	    errMsg = "<font color=""#FF0000"">对不起,您不是业务部员工,您没有使用销售系统的权限</font>"
	Case 3
	    errMsg = "<font color=""#FF0000"">您的账号暂时被冻结</font>"
	Case Else
	    errMsg = ""
	End Select
%>
<table width="570" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr>
    <td align="center" valign="top"><table width="92%"  border="0" align="center" cellpadding="0" cellspacing="0">
      <tr>
        <td><table width="100%"  border="0" cellspacing="0" cellpadding="0">
          <tr>
            <td bgcolor="#88ADDF">&nbsp;</td>
          </tr>
        </table>          
          <br>          <span class="style1">销售系统工作流:</span><br>
          <span class="style1">从客户资源中寻找并选定目标客户→点击按钮进入销售系统→进入增加数据功能,添加您选择的客户→进入您的个人客户→添加拜访计划→按照计划进行客户拜访→当天拜访结束后再次进入此个人客户填写拜访记录(总结)→如成功签单则请部门经理或财务人员在合同管理中添加记录<br>
          </span><br>
          <table width="100%"  border="0" cellspacing="0" cellpadding="0">
            <tr>
              <td bgcolor="#88ADDF">&nbsp;</td>
            </tr>
          </table>
          <br>          <span class="style2">注意事项:</span><br>
          1.如果客户数据为您添加则此客户成为您的个人客户,其他业务人员将不能添加或使用此客户信息。<br> 
          2.如果您想把此客户转交其他业务人员拜访,请使用转移数据功能。<br>
          3.您可以不填写拜访计划,但您需填写拜访记录(总结)。<br>
          4.客户较多时,请使用查询功能。<br>
          5.其他业务人员无法看到您的拜访记录及个人客户资源。<br>
          6.非业务部员工无法使用此系统。</td>
      </tr>
      <tr>
        <td>&nbsp;</td>
      </tr>
      <tr>
        <td><div align="center">
          <table width="364" border="0" cellspacing="0">
            <form name="loginForm" action="?action=login" method="post" onSubmit="return checkInput();">
              <tr> </tr>
              <tr>
                <td height="24" id="errMsg"><div align="center">
                      <% = errMsg %>
                </div></td>
              </tr>
              <tr>
                <td height="5" align="center"><span style="font-size:11px"></span>
                    <input name="item1" type="text" class="box" id="item1" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" value="<%=oabusyusername%>" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="5" align="center"><input name="item2" type="password" class="box" id="item2" style="ime-mode: disabled;" onFocus="this.select(); this.value='';" value="123" size="12" maxlength="16"></td>
              </tr>
              <tr>
                <td height="24" align="center">
                  <input type="submit" name="Submit" value="进入销售系统">
&nbsp;&nbsp; </td>
              </tr>
              <tr>
                <td height="24" align="center">&nbsp;</td>
              </tr>
            </form>
          </table>
        </div></td>
      </tr>
    </table></td>
  </tr>
</table>
<%
End Sub
%>
</body>
</html>
