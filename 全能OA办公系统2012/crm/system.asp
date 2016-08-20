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

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

If Session("CRM_level") <> 9 Then Response.Redirect("listAll.asp")

Function getGroupName(gId)
    If Not IsNumeric(gId) Or gId <= 0 Then
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

Function sql_AddNew(strSQL,objConn,intI,intJ,strFields,strValues,strTypes,strUrl)
    On Error Resume Next
	Dim rs,arrFields,arrValues,arrTypes,i
	If strFields <> "" Then arrFields = Split(strFields,",,")
	If strValues <> "" Then arrValues = Split(strValues,",,")
	If strTypes <> "" Then arrTypes = Split(strTypes,",,")
	If UBound(arrFields) <> UBound(arrValues) Then
		Exit Function
	End If
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open strSQL,objConn,intI,intJ
	rs.AddNew
	For i = 0 To UBound(arrFields)
	    rs(arrFields(i)) = arrValues(i)
	Next
	rs.Update
	rs.Close
	Set rs = Nothing
End Function

Dim strCounter,strToPrint,strForm,strList

Dim action,subAction
action = Trim(Request.QueryString("action"))
subAction = Trim(Request.QueryString("subAction"))
If action = "" Then Response.Redirect("?action=Level")
Select Case action
Case "Level"
    Select Case subAction
	'Case "new"
	Case "save"
	    Call saveLevel()
	Case "edit"
	    Call editLevel()
	Case "restore"
	    Call restoreLevel()
	Case "delete"
	    Call deleteLevel()
	Case "list"
	    Call listLevel()
	Case Else
	    listLevel()
	End Select
Case "Group"
    Select Case subAction
	'Case "new"
	Case "save"
	    Call saveGroup()
	Case "edit"
	    Call editGroup()
	Case "restore"
	    Call restoreGroup()
	Case "delete"
	    Call deleteGroup()
	Case "list"
	    Call listGroup()
	Case Else
	    Call listGroup()
	End Select
Case "ClientsType"
    Select Case subAction
	'Case "new"
	Case "save"
	    Call saveClientsType()
	Case "edit"
	    Call editClientsType()
	Case "restore"
	    Call restoreClientsType()
	Case "delete"
	    Call deleteClientsType()
	Case "list"
	    Call listClientsType()
	Case Else
	    Call listClientsType()
	End Select
Case "RecordsType"
    Select Case subAction
	'Case "new"
	Case "save"
	    Call saveRecordsType()
	Case "edit"
	    Call editRecordsType()
	Case "restore"
	    Call restoreRecordsType()
	Case "delete"
	    Call deleteRecordsType()
	Case "list"
	    Call listRecordsType()
	Case Else
	    Call listRecordsType()
	End Select
Case Else
    Response.Redirect("?action=Level&subAction=list")
End Select

Sub saveLevel()
    Dim lId,lName
	lId = CInt(Abs(Request.Form("lId")))
	lName = Trim(Request.Form("lName"))
	If lId = "" Or lName = "" Then Response.Write("<script>history.back();</script>")
	Dim rs,flag
	flag = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_level",conn,3,2
	Do While Not rs.BOF And Not rs.EOF
	    If rs("lId") = lId Or rs("lName") = gName Then
		    flag = 1
			Exit Do
		End If
		rs.MoveNext
	Loop
	If flag = 0 Then
	    rs.AddNew
		rs("lId") = lId
		rs("lName") = lName
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	If flag = 1 Then
	    Response.Write("<script>" & VBCrlf)
		Response.Write("alert(""编号或者名称已经存在。"")" & VBCrlf)
		Response.Write("history.back();" & VBCrlf)
		Response.Write("</script>")
	Else
	    Response.Redirect("?action=Level")
	End If
End Sub

Sub saveGroup()
    Dim gId,gName
	gId = CInt(Abs(Request.Form("lId")))
	gName = Trim(Request.Form("lName"))
	If gId = "" Or gName = "" Then Response.Write("<script>history.back();</script>")
	Dim rs,flag
	flag = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_group",conn,3,2
	Do While Not rs.BOF And Not rs.EOF
	    If rs("gId") = gId Or rs("gName") =gName Then
		    flag = 1
			Exit Do
		End If
		rs.MoveNext
	Loop
	If flag = 0 Then
	    rs.AddNew
		rs("gId") = gId
		rs("gName") = gName
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	If flag = 1 Then
	    Response.Write("<script>" & VBCrlf)
		Response.Write("alert(""编号或者名称已经存在。"")" & VBCrlf)
		Response.Write("history.back();" & VBCrlf)
		Response.Write("</script>")
	Else
	    Response.Redirect("?action=Group")
	End If
End Sub

Sub saveClientsType()
    Dim clientsType
	clientsType = Trim(Request.Form("clientsType"))
	If clientsType = "" Then Response.Write("<script>history.back();</script>")
	Dim rs,flag
	flag = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_clientsType",conn,3,2
	Do While Not rs.BOF And Not rs.EOF
	    If rs("clientsType") = clientsType Then
		    flag = 1
			Exit Do
		End If
		rs.MoveNext
	Loop
	If flag = 0 Then
	    rs.AddNew
		rs("clientsType") = clientsType
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	If flag = 1 Then
	    Response.Write("<script>" & VBCrlf)
		Response.Write("alert(""编号或者名称已经存在。"")" & VBCrlf)
		Response.Write("history.back();" & VBCrlf)
		Response.Write("</script>")
	Else
	    Response.Redirect("?action=ClientsType")
	End If
End Sub

Sub saveRecordsType()
    Dim recordsType
	recordsType = Trim(Request.Form("recordsType"))
	If recordsType = "" Then Response.Write("<script>history.back();</script>")
	Dim rs,flag
	flag = 0
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_recordsType",conn,3,2
	Do While Not rs.BOF And Not rs.EOF
	    If rs("recordsType") = recordsType Then
		    flag = 1
			Exit Do
		End If
		rs.MoveNext
	Loop
	If flag = 0 Then
	    rs.AddNew
		rs("recordsType") = recordsType
		rs.Update
	End If
	rs.Close
	Set rs = Nothing
	If flag = 1 Then
	    Response.Write("<script>" & VBCrlf)
		Response.Write("alert(""编号或者名称已经存在。"")" & VBCrlf)
		Response.Write("history.back();" & VBCrlf)
		Response.Write("</script>")
	Else
	    Response.Redirect("?action=RecordsType")
	End If
End Sub

Sub editLevel()
    Dim lNameOld
	lNameOld = Trim(Request.QueryString("lNameOld"))
	If lNameOld = "" Then Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_level Where lName = '" & lNameOld & "'",conn,3,1
	If rs.RecordCount = 1 Then
	    strForm = strForm & "    <table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0"" bgcolor=""#FFFFFF"">" & VBCrlf
        strForm = strForm & "      <form name=""levelForm"" action=""?action=Level&subAction=edit"" method=""post"">" & VBCrlf
		strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>添加用户级别：</td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr> " & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>用户级别名称：" & VBCrlf
        strForm = strForm & "          <input name=""lName"" type=""text"" id=""lName"" size=""16"" maxlength=""16"" value=""" & rs("lName") & """>" & VBCrlf
        strForm = strForm & "          级别编号：" & VBCrlf 
        strForm = strForm & "          <input name=""lId"" type=""text"" id=""lId"" size=""2"" maxlength=""2"" value=""" & rs("lId") & """>" & VBCrlf
        strForm = strForm & "          （1-9，数字越大级别越高，管理员为 9 级。）</td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"" align=""center"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td align=""center""> <input type=""submit"" name=""Submit"" value="" 提 交 "">" & VBCrlf 
        strForm = strForm & "          &nbsp;&nbsp; <input name=""Reset"" type=""reset"" id=""Reset"" value="" 重 置 "">" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
		strForm = strForm & "      </form>" & VBCrlf
        strForm = strForm & "    </table>" & VBCrlf
	End If
	rs.Close
	Set rs = Nothing
	strList = strList & list("level")	
End Sub

Sub editGroup()
    Dim gNameOld
	gNameOld = Trim(Request.QueryString("gNameOld"))
	If gNameOld = "" Then Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_group Where gName = '" & gNameOld & "'",conn,3,1
	If rs.RecordCount = 1 Then
	    strForm = strForm & "    <table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0"" bgcolor=""#FFFFFF"">" & VBCrlf
        strForm = strForm & "      <form name=""groupForm"" action=""?action=Group&subAction=edit"" method=""post"">" & VBCrlf
		strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>添加用户组：</td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr> " & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>用户组名称：" & VBCrlf
        strForm = strForm & "          <input name=""gName"" type=""text"" id=""gName"" size=""16"" maxlength=""16"" value=""" & rs("gName") & """>" & VBCrlf
        strForm = strForm & "          组编号：" & VBCrlf 
        strForm = strForm & "          <input name=""gId"" type=""text"" id=""gId"" size=""2"" maxlength=""2"" value=""" & rs("gId") & """>" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"" align=""center"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td align=""center""> <input type=""submit"" name=""Submit"" value="" 提 交 "">" & VBCrlf 
        strForm = strForm & "          &nbsp;&nbsp; <input name=""Reset"" type=""reset"" id=""Reset"" value="" 重 置 "">" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
		strForm = strForm & "      </form>" & VBCrlf
        strForm = strForm & "    </table>" & VBCrlf
	End If
	rs.Close
	Set rs = Nothing
	strList = strList & list("group")
End Sub

Sub editClientsType()
    Dim clientsTypeOld
	clientsTypeOld = Trim(Request.QueryString("clientsTypeOld"))
	If clientsTypeOld = "" Then Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_clientsType Where clientsType = '" & clientsTypeOld & "'",conn,3,1
	If rs.RecordCount = 1 Then
	    strForm = strForm & "    <table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0"" bgcolor=""#FFFFFF"">" & VBCrlf
        strForm = strForm & "      <form name=""clientsTypeForm"" action=""?action=ClientsType&subAction=edit"" method=""post"">" & VBCrlf
		strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>添加客户类型：</td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr> " & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>客户类型：" & VBCrlf
        strForm = strForm & "          <input name=""clientsType"" type=""text"" id=""clientsType"" size=""16"" maxlength=""16"" value=""" & rs("clientsType") & """>" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"" align=""center"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td align=""center""> <input type=""submit"" name=""Submit"" value="" 提 交 "">" & VBCrlf 
        strForm = strForm & "          &nbsp;&nbsp; <input name=""Reset"" type=""reset"" id=""Reset"" value="" 重 置 "">" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
		strForm = strForm & "      </form>" & VBCrlf
        strForm = strForm & "    </table>" & VBCrlf
	End If
	rs.Close
	Set rs = Nothing
	strList = strList & list("clientsType")
End Sub

Sub editRecordsType()
    Dim recordsTypeOld
	recordsTypeOld = Trim(Request.QueryString("recordsTypeOld"))
	If recordsTypeOld = "" Then Response.Redirect(Request.ServerVariables("HTTP_REFERER"))
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_recordsType Where recordsType = '" & clientsTypeOld & "'",conn,3,1
	If rs.RecordCount = 1 Then
	    strForm = strForm & "    <table width=""100%"" border=""0"" cellpadding=""3"" cellspacing=""0"" bgcolor=""#FFFFFF"">" & VBCrlf
        strForm = strForm & "      <form name=""recordsTypeForm"" action=""?action=RecordsType&subAction=edit"" method=""post"">" & VBCrlf
		strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>添加客户类型：</td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr> " & VBCrlf
        strForm = strForm & "        <td width=""60"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td>客户类型：" & VBCrlf
        strForm = strForm & "          <input name=""recordsType"" type=""text"" id=""recordsType"" size=""16"" maxlength=""16"" value=""" & rs("recordsType") & """>" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
        strForm = strForm & "      <tr>" & VBCrlf
        strForm = strForm & "        <td width=""60"" align=""center"">&nbsp;</td>" & VBCrlf
        strForm = strForm & "        <td align=""center""> <input type=""submit"" name=""Submit"" value="" 提 交 "">" & VBCrlf 
        strForm = strForm & "          &nbsp;&nbsp; <input name=""Reset"" type=""reset"" id=""Reset"" value="" 重 置 "">" & VBCrlf
        strForm = strForm & "        </td>" & VBCrlf
        strForm = strForm & "      </tr>" & VBCrlf
		strForm = strForm & "      </form>" & VBCrlf
        strForm = strForm & "    </table>" & VBCrlf
	End If
	rs.Close
	Set rs = Nothing
	strList = strList & list("recordsType")
End Sub

Sub restoreLevel()
End Sub

Sub restoreGroup()
End Sub

Sub restoreClientsType()
End Sub

Sub restoreRecordsType()
End Sub

Sub deleteLevel()
End Sub

Sub deleteGroup()
End Sub

Sub deleteClientsType()
End Sub

Sub deleteRecordsType()
End Sub

Sub listLevel()
    strForm = myForm("level")
	strList = list("level")
End Sub

Sub listGroup()
    strForm = myForm("group")
	strList = list("group")
End Sub

Sub listClientsType()
    strForm = myForm("clientsType")
	strList = list("clientsType")
End Sub

Sub listRecordsType()
    strForm = myForm("recordsType")
	strList = list("recordsType")
End Sub

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
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>修改</a>] [<a href=""?action=block&uId=" & rs("uId") & """>冻结</a>] [<a href=""?action=delete&uId=" & rs("uId") & """>删除</a>]</td>" & VBCrlf
	Else
	    strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=edit&uId=" & rs("uId") & """>修改</a>] [<a href=""?action=block&uId=" & rs("uId") & """>解冻</a>] [<a href=""?action=delete&uId=" & rs("uId") & """>删除</a>]</td>" & VBCrlf
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

function selectSystem(j)
{
    var num = 4;
	for(var i=1;i<=num;i++){
	    if(i != j){
		    document.all["block" + i].style.display = "none";
			document.all["block" + (i + 4)].style.display = "none";
		}
		else{
		    document.all["block" + i].style.display = "block";
		    document.all["block" + (i + 4)].style.display = "block";
		}
	}
}
-->
</script>
</head>

<body  >
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
          <table onclick="window.location.replace('listAll.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr > 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">查看所有数据</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('addData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr>   
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">添加数据</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>	  
          <table onclick="window.location.replace('advanceSearch.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">高级搜索</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('dataForm.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">数据报表</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('exportData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">数据导出</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
<% If Session("CRM_level") = 9 Then %>
          <table onclick="window.location.replace('transData.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
              <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
              <td background="images/tab_off_middle.gif">数据转移</td>
              <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
          <table onclick="window.location.replace('manageUser.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
            <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>                  
                <td background="images/tab_off_middle.gif">用户管理</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
            </tr>
          </table>
			<table onclick="window.location.replace('system_level.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="left">
              <tr> 
                <td><img src="images/tab_on_left.gif" width="16" height="24"></td>
                <td background="images/tab_on_middle.gif">系统设置</td>
                <td><img src="images/tab_on_right.gif" width="16" height="24"></td>
              </tr>
            </table>
<% End If %>
          <table onclick="window.location.replace('logout.asp')" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="right">
              <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">注销</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
              </tr>
            </table>
			<table onclick="window.location.reload();" style="cursor: hand;" border="0" cellspacing="0" cellpadding="0" align="right">
              <tr> 
                <td><img src="images/tab_off_left.gif" width="16" height="24"></td>
                <td background="images/tab_off_middle.gif">刷新</td>
                <td><img src="images/tab_off_right.gif" width="16" height="24"></td>
              </tr>
            </table></td>
      </tr>
    </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="#999999">&nbsp;</td>
        </tr>
        <tr id="block1" style="display: block;"> 
          <td><table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr> 
                <td width="60">&nbsp;</td>
                <td>添加用户级别：</td>
              </tr>
              <tr> 
                <td width="60">&nbsp;</td>
                <td>用户级别名称： 
                  <input name="lName" type="text" id="lName5" size="16" maxlength="16">
                  级别编号： 
                  <input name="lId" type="text" id="lId5" size="2" maxlength="2">
                  （1-9，数字越大级别越高，管理员为 9 级。） </td>
              </tr>
              <tr> 
                <td width="60" align="center">&nbsp;</td>
                <td align="center"> <input type="submit" name="Submit" value=" 提 交 "> 
                  &nbsp;&nbsp; <input name="Reset" type="reset" id="Reset5" value=" 重 置 "> 
                </td>
              </tr>
            </table>            
          </td>
        </tr>
        <tr id="block2" style="display: none;">
          <td><table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr> 
                <td width="60">&nbsp;</td>
                <td>添加用户组：</td>
              </tr>
              <tr> 
                <td width="60">&nbsp;</td>
                <td>用户组名称： 
                  <input name="gName" type="text" id="gName2" size="16" maxlength="16">
                  组编号： 
                  <input name="lId2" type="text" id="lId23" size="2" maxlength="2"> 
                </td>
              </tr>
              <tr> 
                <td width="60" align="center">&nbsp;</td>
                <td align="center"> <input type="submit" name="Submit2" value=" 提 交 "> 
                  &nbsp;&nbsp; <input name="Reset2" type="reset" id="Reset23" value=" 重 置 "> 
                </td>
              </tr>
            </table></td>
        </tr>
        <tr id="block3" style="display: none;">
          <td><table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr> 
                <td width="60">&nbsp;</td>
                <td>添加客户类型：</td>
              </tr>
              <tr> 
                <td width="60">&nbsp;</td>
                <td>客户类型名称： 
                  <input name="clientsType" type="text" id="clientsType2" size="16" maxlength="16"> 
                </td>
              </tr>
              <tr> 
                <td width="60" align="center">&nbsp;</td>
                <td align="center"> <input type="submit" name="Submit3" value=" 提 交 "> 
                  &nbsp;&nbsp; <input name="Reset3" type="reset" id="Reset33" value=" 重 置 "> 
                </td>
              </tr>
            </table></td>
        </tr>
        <tr id="block4" style="display: none;">
          <td><table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">
              <tr> 
                <td width="60">&nbsp;</td>
                <td>添加拜访记录类型：</td>
              </tr>
              <tr> 
                <td width="60">&nbsp;</td>
                <td>拜访记录名称： 
                  <input name="recordsType" type="text" id="recordsType" size="16" maxlength="16"> 
                </td>
              </tr>
              <tr> 
                <td width="60" align="center">&nbsp;</td>
                <td align="center"> <input type="submit" name="Submit4" value=" 提 交 "> 
                  &nbsp;&nbsp; <input name="Reset4" type="reset" id="Reset43" value=" 重 置 "> 
                </td>
              </tr>
            </table></td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#999999" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td>[<span onClick="selectSystem(1);" style="cursor: hand;">用户级别</span>] [<span onClick="selectSystem(2);" style="cursor: hand;">用户组</span>] [<span onClick="selectSystem(3);" style="cursor: hand;">客户类型</span>] [<span onClick="selectSystem(4);" style="cursor: hand;">拜访类型</span>]</td>
        </tr>
      </table>
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr id="block5" style="display: block;">
          <td><table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
              <tr> 
                <td width="120" align="center" bgcolor="menu">级别编号</td>
                <td align="center" bgcolor="menu">用户等级名称</td>
                <td width="120" align="center" bgcolor="menu">操作</td>
                <%' = strToPrint %>
              </tr>
            </table>
            
          </td>
        </tr>
        <tr id="block6" style="display: none;">
          <td><table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
              <tr> 
                <td width="120" align="center" bgcolor="menu">用户组编号</td>
                <td align="center" bgcolor="menu">用户组名称</td>
                <td width="120" align="center" bgcolor="menu">操作</td>
                <%' = strToPrint %>
              </tr>
            </table></td>
        </tr>
        <tr id="block7" style="display: none;">
          <td><table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
              <tr> 
                <td align="center" bgcolor="menu">客户类型</td>
                <td width="120" align="center" bgcolor="menu">操作</td>
                <%' = strToPrint %>
              </tr>
            </table></td>
        </tr>
        <tr id="block8" style="display: none;">
          <td><table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
              <tr> 
                <td align="center" bgcolor="menu">拜访类型</td>
                <td width="120" align="center" bgcolor="menu">操作</td>
                <%' = strToPrint %>
              </tr>
            </table></td>
        </tr>
      </table> </td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#999999"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
