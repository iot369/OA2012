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

<!--登录权限判断，Session和MD5加密判断-->
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
		rs.Open "Select * From baidu_records Where rId = " & rId,conn,3,1
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

Function list()
    Dim strToPrint
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
    rs.Open "Select * From baidu_records Order By rId",conn,3,1
    Do While Not rs.BOF And Not rs.EOF
    	strToPrint = strToPrint & "        <tr>" & VBCrlf
    	strToPrint = strToPrint & "          <td align=""center"">" & rs("cId") & "</td>" & VBCrlf
    	strToPrint = strToPrint & "          <td>" & rs("rType") & "</td>" & VBCrlf
    	strToPrint = strToPrint & "          <td align=""center"">[<a href=""?action=delete&lNameOld=" & rs("rType") & """ onClick=""return confirm('确定要删除该\r客户吗？');"">删除</a>]</td>" & VBCrlf
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	list = strToPrint
End Function
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
<style type="text/css"> 
.style7 {color: #2d4865}
.style8 {color: #0d79b3;
	font-weight: bold;
}
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
<table width="550" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;">
    <td height="1" valign="top"> 
      <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
        <tr> 
          <td bgcolor="#88ADDF">&nbsp;</td>
        </tr>
        <tr id="block1" style="display: block;"> 
          <td>
<%
Dim action
action = Trim(Request.QueryString("action"))
Select Case action
Case "add"
    Call addOrEdit()
Case "save"
    Call saveData()
Case "edit"
    Call addOrEdit()
Case "restore"
    Call restore()
Case "delete"
    Call deleteData()
Case Else
    Call addOrEdit()
End Select

Sub saveData()
    Dim lId,lName
	lId = CInt(Abs(Request.Form("lId")))
	lName = Trim(Request.Form("lName"))
	If lId = "" Or lName = "" Then
	    Response.Write("<div align=""center"">提交的数据不完整，请返回重新填写。<br>")
		Response.Write("<input name=""back"" type=""button"" value="" 返 回 "" onClick=""history.back();""></div>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From baidu_level Where lId = " & lId & " Or lName = '" & lName & "'",conn,3,2
	If rs.RecordCount > 0 Then
	    Response.Write("<div align=""center"">该名称或编号已经存在。<br>")
		Response.Write("<input name=""back"" type=""button"" value="" 返 回 "" onClick=""history.back();""></div>")
		rs.Close
		Set rs = Nothing
		Exit Sub
	Else
	    rs.AddNew
		rs("lId") = lId
		rs("lName") = lName
		rs.Update
		rs.Close
		Set rs = Nothing
		Response.Redirect("?")
	End If
End Sub

Sub restore()
    Dim lNameOld,lIdOld,lId,lName
	lNameOld = Trim(Request.Form("lNameOld"))
	lId = CInt(Abs(Request.Form("lId")))
	lName = Trim(Request.Form("lName"))
	If lNameOld = "" Or lId = "" Or lName = "" Then
	    Response.Write("<div align=""center"">提交的数据不完整，请返回重新填写。<br>")
		Response.Write("<input name=""back"" type=""button"" value="" 返 回 "" onClick=""history.back();""></div>")
		Exit Sub
	End If
    Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select  * From baidu_level Where lName <> '" & lNameOld & "'",conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If rs("lId") = lId Or rs("lName") = lName Then
	        Response.Write("<div align=""center"">该名称或编号已经存在。<br>")
		    Response.Write("<input name=""back"" type=""button"" value="" 返 回 "" onClick=""history.back();""></div>")
		    rs.Close
		    Set rs = Nothing
		    Exit Sub
		End If
		rs.MoveNext
	Loop
	rs.Close
	
	rs.Open "Select * From baidu_level Where lName = '" & lNameOld & "'",conn,3,2
	If rs.RecordCount = 1 Then
	    lIdOld = rs("lId")
	    rs("lId") = lId
		rs("lName") = lName
		rs.Update
		If lIdOld <> lId Then
		    Dim rss
			Set rss = Server.CreateObject("ADODB.Recordset")
			rss.Open "Select * From baidu_user Where uLevel = " & lIdOld,conn,3,2
			Do While Not rss.BOF And Not rss.EOF
			    rss("uLevel") = lId
				rss.Update
				rss.MoveNext
			Loop
			rss.Close
			Set rss = Nothing
		End If
	End If
	rs.Close
	Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub deleteData()
    Dim lNameOld,lIdOld
	lNameOld = Trim(Request("lNameOld"))
	If lNameOld = "" Then
	    Response.Write("<div align=""center"">提交的数据不完整，请返回重新填写。<br>")
		Response.Write("<input name=""back"" type=""button"" value="" 返 回 "" onClick=""history.back();""></div>")
		Exit Sub
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_records Where rType = '" & lNameOld & "'",conn,3,2
	If rs.RecordCount > 0 Then
	    lIdOld = rs("rId")
		rs.Delete
		rs.Update
	End If


	rs.Close
	
	If lIdOld <> "" Then
	    rs.Open "Select * From baidu_user Where uLevel = " & lIdOld,conn,3,2
		Do While Not rs.BOF And Not rs.EOF
		    rs.Delete
			rs.Update
			rs.MoveNext
		Loop
		rs.Close
	End If
	Set rs = Nothing
	Response.Redirect("?")
End Sub

Sub addOrEdit()
    Dim lId,lName,lNameOld,strOut,strAction
	If action = "edit" Then
	    Dim rs
		lNameOld = Trim(Request("lNameOld"))
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_level Where lName = '" & lNameOld & "'",conn,3,1
		If rs.RecordCount = 1 Then
		    lId = rs("lId")
			lName = rs("lName")
		End If
		rs.Close
		Set rs = Nothing
		strOut = "编辑用户级别："
		strAction = "?action=restore"
	Else
	    strOut = "客户数据删除"
		strAction = "?action=save"
	End If		    
%>
		    <table width="100%" border="0" cellpadding="3" cellspacing="0" bgcolor="#FFFFFF">

            </table>
<%
End Sub
%>
		  </td>
        </tr>
      </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"><table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr>
          <td height="40"><div align="center">[<a href="system_level.asp">用户级别</a>] [<a href="system_group.asp">用 户 组</a> ] [<a href="system_clientsType.asp">客户等级</a>] [<a href="system_clientsTrade.asp">行业类型</a>] [<a href="system_recordsType.asp">拜访类型</a>]<br>
[<a href="system_area.asp">业务区域</a>] [<a href="system_square.asp">业务小区</a>] [<a href="system_del1.asp">客户删除</a>] [<a href="system_del2.asp">记录删除</a>] [<a href="system_del3.asp">拜访删除</a>] </div></td>
        </tr>
      </table>
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr> 
          <td width="120" align="center" bgcolor="menu">客户编号</td>
          <td align="center" bgcolor="menu">行为内容</td>
          <td width="120" align="center" bgcolor="menu">操作</td>
          <% = list() %>
        </tr>
      </table> </td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
</body>
</html>
