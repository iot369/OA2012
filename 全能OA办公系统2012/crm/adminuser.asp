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
<!--#include file="Connections/conn.asp" -->
<!--登录权限判断，Session和MD5加密判断-->'
<%
''生成下拉列表
If request.cookies("oabusyname")=""  Then Response.Redirect("../default.asp")
If request.cookies("cook_allow_control_all_user")=""  Then Response.Redirect("../default.asp")
Function getList(i,sTable,iId,sValue,sName,selfValue)
    If i < 1 Or i > 2 Then
	    getList = ""
		Exit Function
	End If
	Dim strList
	Dim rs
	If i = 1 Then
	    strList = "<select name=""" & sName & """ selfValue=""" & selfValue & """>"
		strList = strList & "<option value="""">请选择</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From " & sTable & "",conn,3,1
		Do While Not rs.BOF And Not rs.EOF
		    strList = strList & "<option value=""" & rs(sValue) & """>" & rs(sValue) & "</option>"
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		getList = strList
	Else
	    strList = "<select name=""" & sName & """ selfValue=""" & selfValue & """>"
		strList = strList & "<option value="""">请选择</option>"
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From " & sTable & "",conn,3,1
		Do While Not rs.BOF And Not rs.EOF
		    strList = strList & "<option value=""" & rs(iId) & """>" & rs(sValue) & "</option>"
			rs.MoveNext
		Loop
		rs.Close
		Set rs = Nothing
		strList = strList & "</select>"
		getList = strList
	End If
End Function

Rem Session("CRM_account") 用户帐号
Rem Session("CRM_name") 用户名
Rem Session("CRM_level") 用户等级


Function getGroupName(gId)
    If Not IsNumeric(gId) Or gId < 0 Then
	    getGroupName = ""
	Else
	    Dim rs,gName
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gId = " & gId,conn,3,1
		If rs.RecordCount <> 1 Then
		    gName = ""
		Else
		    gName = rs("gName")
		End If
		rs.Close
		Set rs = Nothing
		getGroupName = gName
	End If
End Function

Function getLevelName(lId)
    If Not IsNumeric(lId) Or lId <= 0 Then
	    getLevelName = ""
	Else
	    Dim rs,lName
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_level Where lId = " & lId,conn,3,1
		If rs.RecordCount <> 1 Then
		    lName = ""
		Else
		    lName = rs("lName")
		End If
		rs.Close
		Set rs = Nothing
		getLevelName = lName
	End If
End Function

Dim strCounter,strToPrint

Dim rs,intTotalRecords,intTotalPages,intCurrentPage,intPageSize
intCurrentPage = CInt(ABS(Request("pageNum")))
If Not IsNumeric(intCurrentPage) Or intCurrentPage <= 0 Then intCurrentPage = 1
intPageSize = 10

Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select * From baidu_user Order By uId",conn,3,1
intTotalRecords = rs.RecordCount
rs.PageSize = intPageSize
intTotalPages = rs.PageCount
If intCurrentPage > intTotalPages Then intCurrentPage = intTotalPages
If intTotalRecords > 0 Then rs.AbsolutePage = intCurrentPage
strCounter = strCounter & "共 " & intTotalRecords & " 条记录 "
strCounter = strCounter & "共 " & intTotalPages & " 页 "
strCounter = strCounter & "当前第 " & intCurrentPage & " 页 "
If intCurrentPage <> 1 And intTotalRecords <> 0 Then
    strCounter = strCounter & "<a href=""?pageNum=1""><<首页</a> "
Else
    strCounter = strCounter & "<<首页 "
End If
If intCurrentPage > 1 Then
    strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage - 1 & """><上一页</a> "
Else
    strCounter = strCounter & "<上一页 "
End If
If intCurrentPage < intTotalPages Then
    strCounter = strCounter & "<a href=""?pageNum=" & intCurrentPage + 1 & """>下一页></a> "
Else
    strCounter = strCounter & "下一页> "
End If
If intCurrentPage <> intTotalPages Then
    strCounter = strCounter & "<a href=""?pageNum=" & intTotalPages & """>尾页>></a>"
Else
    strCounter = strCounter & "尾页>>"
End If

Dim i
i = 0
Do While Not rs.BOF And Not rs.EOF
    i = i + 1
	strToPrint = strToPrint & "        <tr>" & VBCrlf
	strToPrint = strToPrint & "          <td align=""center"">" & rs("uId") & "</td>" & VBCrlf
	If rs("uBlock") = False Then
	    strToPrint = strToPrint & "          <td>" & rs("uAccount") & "</td>" & VBCrlf
	Else
	    strToPrint = strToPrint & "          <td><font color=""#FF0000"">" & rs("uAccount") & "</font></td>" & VBCrlf
	End If
	strToPrint = strToPrint & "          <td>" & rs("uPassword") & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & rs("uName") & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & getGroupName(rs("uGroup")) & "</td>" & VBCrlf
	strToPrint = strToPrint & "          <td>" & getLevelName(rs("uLevel")) & "</td>" & VBCrlf
	If rs("uBlock") = False Then
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>修改</a>] [<a href=""?action=block&uId=" & rs("uId") & """>冻结</a>] [<a href=""?action=delete&uId=" & rs("uId") & """ onClick=""return confirm('确定删除该用户和相\r关的所有资料？');"">删除</a>]</td>" & VBCrlf
	Else
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>修改</a>] [<a href=""?action=block&uId=" & rs("uId") & """>解冻</a>] [<a href=""?action=delete&uId=" & rs("uId") & """ onClick=""return confirm('确定删除该用户和相\r关的所有资料？');"">删除</a>]</td>" & VBCrlf
	End If
	strToPrint = strToPrint & "        </tr>" & VBCrlf
    If i >= intPageSize Then Exit Do
	rs.MoveNext
Loop
rs.Close
Set rs = Nothing
%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>销售管理系统</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
<script language="JavaScript">
<!--
function showHideHead(strSrc)
{
	var strFile = strSrc.substring(strSrc.lastIndexOf("/"),strSrc.length);
    if (strFile == "/arrow_up.gif"){
	    oHead.style.display = "none";
		oHeadCtrl.src = "images/arrow_down.gif";
		oHeadCtrl.alt = "显示头部";
		oHeadBar.title = "显示头部";		
	}
	else {
	    oHead.style.display = "block";
		oHeadCtrl.src = "images/arrow_up.gif";
		oHeadCtrl.alt = "隐藏头部";
		oHeadBar.title = "隐藏头部";
	}
}

function checkInput(o)
{
    var oo = eval("document.all." + o);
    var num = oo.length;
    for(var i=0;i<num;i++){
	    if(oo[i].value == ""){
		    alert(oo[i].selfValue + "不能为空。");
			oo[i].focus();
			return false
			break;
		}
	}
}

if (this.location.href == top.location.href){
    top.location.href = "";
}
-->
</script>
</head>

<body >
<table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="778" border="0" cellspacing="0" cellpadding="0">
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
    <table width="100%" border="0" cellpadding="0" cellspacing="0">
      <tr> 
        <td height="5"><img src="images/null.gif" width="1" height="1"></td>
      </tr>
      <tr>
        <td bgcolor="#999999">&nbsp;</td>
      </tr>
    </table>
<%
Dim action
action = Trim(Request("action"))
Select Case action
Case "save"
    Call saveData()
Case "edit"
    Call editForm()
Case "saveEdit"
    Call saveEditData()
Case "delete"
    Call deleteData()
Case "block"
    Call block()
Case Else
    Call addForm()
End Select

Dim errMsg
errMsg = CInt(ABS(Request("errMsg")))
Select Case errMsg
Case 1
    errMsg = "<font color=""#FF0000"">请求的数据不存在。</font>"
Case 2
    errMsg = "<font color=""#FF0000"">发送的数据错误。</font>"
Case Else
    errMsg = ""
End Select

Sub addForm()
%>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="newUser" action="?action=save" method="post" onSubmit="return checkInput('newUser');">
		<tr> 
          <td align="right">&nbsp;</td>
          <td>新建账号 <% = errMsg %></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>账号：<input name="account" type="text" id="account" size="16" maxlength="16" selfValue="账号">
            真实姓名：<input name="name" type="text" id="name" size="16" maxlength="16" selfValue="真实姓名">
            用户等级：<% = getList(2,"baidu_level","lId","lName","level","用户等级") %></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>密码：<input name="password" type="password" id="password" size="16" maxlength="16" selfValue="密码">
            确认密码：<input name="confirmPWS" type="password" id="confirmPWS" size="16" maxlength="16" selfValue="确认密码">
            所属小组：<% = getList(2,"baidu_group","gId","gName","group","所属小组") %></td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="submit" name="Submit" value=" 建 立 "> &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" 重 置 "></td>
        </tr>
		</form>
      </table>
<%
End Sub

Sub saveData()
    Dim uAccount,uPassword,uConfirmPWS,uName,uLevel,uGroup
	uAccount = Trim(Request("account"))
	uPassword = Trim(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	uLevel = CInt(Request("level"))
	uGroup = CInt(Request("group"))
	If uAccount = "" Or uPassword = "" Or uPassword <> uConfirmPWS Or uName = "" Or uGroup = "" Then
	    Response.Redirect("adminuser.asp?errMsg=2")
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From baidu_user",conn,3,2
	rs.AddNew
	rs("uAccount") = uAccount
	rs("uPassword") = uPassword
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub editForm()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	Dim uAccount,uPassword,uName,uLevel,uGroup
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,1
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	uAccount = rs("uAccount")
	uPassword = rs("uPassword")
	uName = rs("uName")
	uLevel = rs("uLevel")
	uGroup = rs("uGroup")
%>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="newUser" action="?action=saveEdit" method="post" onSubmit="return checkInput('newUser');">
		<tr> 
          <td align="right">&nbsp;</td>
          <td>修改账号</td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>账号： 
              <input name="account" type="text" id="account" value="<% = uAccount %>" size="16" maxlength="16">
            真实姓名： 
              <input name="name" type="text" id="name" value="<% = uName %>" size="16" maxlength="16">
            用户等级：<% = getList(2,"baidu_level","lId","lName","level","用户等级") %>
              <input name="id" type="hidden" id="id" value="<% = uId %>"></td>
        </tr>
        <tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>密码： 
              <input name="password" type="password" id="password" value="<% = uPassword %>" size="16" maxlength="16">
            确认密码： 
              <input name="confirmPWS" type="password" id="confirmPWS" value="<% = uPassword %>" size="16" maxlength="16">
            所属小组：<% = getList(2,"baidu_group","gId","gName","group","所属小组") %></td>
        </tr>
        <tr> 
          <td align="right">&nbsp;</td>
          <td>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; 
            <input type="submit" name="Submit" value=" 编 辑 "> &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset" value=" 重 置 "></td>
        </tr>
		</form>
      </table>
<script language="JavaScript">
<!--
    var level = "<% = uLevel %>";
	var group = "<% = uGroup %>";
	for (i=0;i<document.newUser.level.options.length;i++){
        if (document.newUser.level.options[i].value == level)
            document.newUser.level.options[i].selected = true;
    }
	for (i=0;i<document.newUser.group.options.length;i++){
        if (document.newUser.group.options[i].value == group)
            document.newUser.group.options[i].selected = true;
    }
-->
</script>
<%
End Sub

Sub saveEditData()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("id")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
    Dim uAccount,uPassword,uConfirmPWS,uName,uLevel,uGroup
	uAccount = Trim(Request("account"))
	uPassword = Trim(Request("password"))
	uConfirmPWS = Trim(Request("confirmPWS"))
	uName = Trim(Request("name"))
	uLevel = CInt(Request("level"))
	uGroup = CInt(Request("group"))
	If uAccount = "" Or uPassword = "" Or uPassword <> uConfirmPWS Or uName = "" Or uGroup = "" Then
	    Response.Redirect("adminuser.asp?errMsg=2")
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	rs("uAccount") = uAccount
	rs("uPassword") = uPassword
	rs("uName") = uName
	rs("uLevel") = uLevel
	rs("uGroup") = uGroup
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub deleteData()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	''Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	rs.Delete
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub

Sub block()
    Dim uId
	uId = 0
	uId = CInt(ABS(Request("uId")))
	If Not IsNumeric(uId) Or uId <= 0 Then Response.Redirect("adminuser.asp?errMsg=1")
	''Dim rs
	Dim uBlock
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uId = " & uId,conn,3,2
	If rs.RecordCount <> 1 Then Response.Redirect("adminuser.asp?errMsg=1")
	uBlock = rs("uBlock")
	If uBlock = True Then
	    rs("uBlock") = False
	Else
	    rs("uBlock") = True
	End If
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Redirect("adminuser.asp")
End Sub
%>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#999999" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td> 
            <% = strCounter %>
          </td>
          <td align="right"> [<a href="adminuser.asp">用户列表</a>] [<a href="adminuser.asp">新建用户</a>]</td>
        </tr>
      </table> 
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr>
          <td width="40" align="center" bgcolor="menu">编号</td>
          <td width="80" align="center" bgcolor="menu">账号</td>
          <td width="80" align="center" bgcolor="menu">密码</td>
          <td width="80" align="center" bgcolor="menu">真实姓名</td>
          <td align="center" bgcolor="menu">所属小组</td>
          <td width="60" align="center" bgcolor="menu">用户等级</td>
          <td align="center" bgcolor="menu">操作</td><% = strToPrint %>
        </tr>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#999999"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
