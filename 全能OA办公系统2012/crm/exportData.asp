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
Rem Session("CRM_account") 用户帐号
Rem Session("CRM_name") 用户名
Rem Session("CRM_level") 用户等级

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Dim strCounter,strToPrint

Dim dataItem,beginData,endData
dataItem = Trim(Request("dataItem"))
beginData = Trim(Request("beginData"))
endData = Trim(Request("endData"))

If beginData = "起始数据" Then beginData = ""
If endData = "结束数据" Then endData = ""

Function getUserList(intLevel,intGroup)
    Dim rs,strUserList
	strUserList = "'" & Session("CRM_name") & "'"
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_user Where uLevel < " & intLevel & " And uGroup = " & intGroup,conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    If strUserList = "" Then
		    strUserList = "'" & rs("uName") & "'"
		Else
		    strUserList = strUserList & ",'" & rs("uName") & "'"
		End If
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	getUserList = strUserList
End Function

Function getItem(dataItem)
    If dataItem = "" Then
	    getItem = ""
	Else
	    Select Case dataItem
		Case "cEmail"
		    getItem = "电子信箱"
		Case Else
		    getItem = ""
		End Select
	End If
End Function

Dim flag
flag = 0
If dataItem <> "" Then
    strToPrint = strToPrint & "        <tr>" & VBCrlf
    strToPrint = strToPrint & "          <td width=""60"" align=""center"" bgcolor=""menu"">编号</td>" & VBCrlf
    strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">公司名称</td>" & VBCrlf
    strToPrint = strToPrint & "          <td align=""center"" bgcolor=""menu"">公司网址</td>" & VBCrlf
    strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">" & getItem(dataItem) & "</td>" & VBCrlf
    If Session("CRM_level") = 9 Then
        strToPrint = strToPrint & "          <td width=""80"" align=""center"" bgcolor=""menu"">业务员</td>" & VBCrlf
    End If
    strToPrint = strToPrint & "        </tr>" & VBCrlf
	
    Dim fso,f,fl
	Set fso = Server.CreateObject("Scripting.FileSystemObject")
	fl = Server.MapPath("data.CSV")
	If fso.FileExists(fl) Then
	    fso.DeleteFile(fl)
	    Set f = fso.CreateTextFile(fl)
		f.WriteLine("姓名,电子邮件地址,foxaddrID")
	Else
	    Set f = fso.CreateTextFile(fl)
		f.WriteLine("姓名,电子邮件地址,foxaddrID")
	End If
    Dim rs
    Set rs = Server.CreateObject("ADODB.Recordset")
	
    Select Case dataItem
    Case "cEmail"
        If beginData = "" And endData = "" Then
    	    If Session("CRM_level") = 9 Then
    	        rs.Open "Select * From baidu_client Order By cId Desc",conn,3,1
    		Else
    		    rs.Open "Select * From baidu_client Where cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
    		End If
    	Else
    	    If beginData = "" Then beginData = endData		
      	    If Session("CRM_level") = 9 Then
    	        rs.Open "Select * From baidu_client Where cEmail Like '%" & beginData & "%' Order By cId Desc",conn,3,1
    		Else
    		    rs.Open "Select * From baidu_client Where cEmail Like '%" & beginData & "%' And cUser = '" & Session("CRM_name") & "' Order By cId Desc",conn,3,1
    		End If
    	End If
    Case Else
    End Select
	
    Do While Not rs.BOF And Not rs.EOF
        strToPrint = strToPrint & "        <tr>" & VBCrlf
        strToPrint = strToPrint & "          <td width=""60"" align=""center"">" & rs("cId") & "</td>" & VBCrlf
        strToPrint = strToPrint & "          <td><a href=""view.asp?cId=" & rs("cId") & """>" & rs("cCompany") & "</a></td>" & VBCrlf
        strToPrint = strToPrint & "          <td><a href=""http://" & rs("cHomepage") & """ target=""_blank"">" & rs("cHomepage") & "</td>" & VBCrlf
        strToPrint = strToPrint & "          <td>" & rs(dataItem) & "</td>" & VBCrlf
    	If Session("CRM_level") = 9 Then
            strToPrint = strToPrint & "          <td>" & rs("cUser") & "</td>" & VBCrlf
    	End If
		f.WriteLine(rs("cLinkman") & "," & rs(dataItem) & "," & rs("cId"))
    	rs.MoveNext
    Loop
    rs.Close
    Set rs = Nothing
	f.CLose
	Set f = Nothing
	Set fso = Nothing
	flag = 1
End If
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

if (this.location.href == top.location.href){
    top.location.href = "";
}

function checkInput()
{
    if (document.exportForm.dataItem.value == ""){
	    alert("请选择要导出的数据种类。");
		document.exportForm.dataItem.focus();
		return false;
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
    </table>
      <table width="100%" border="0" cellpadding="5" cellspacing="0" bgcolor="#FFFFFF">
        <form name="exportForm" action="?" method="post" onSubmit="return checkInput();">
		<tr> 
          <td width="40" align="right">&nbsp;</td>
          <td>请选择导出数据项目：
            <select name="dataItem" id="dataItem">
              <option value="">请选择</option>
              <option value="cEmail">电子信箱</option>
            </select>
            <br>
            请选择导出数据范围： 
            <input name="beginData" type="text" id="beginData" value="起始数据" size="16" maxlength="36" onFocus="this.value='';">
            -
            <input name="endData" type="text" id="endData" value="结束数据" size="16" maxlength="36" onFocus="this.value='';">
            <input type="submit" name="Submit" value=" 导 出 ">
            <hr size="1" noshade>
            <span class="emRed">说明：</span><br>
            &nbsp;&nbsp;&nbsp;&nbsp;导出数据范围，可以帮助您限定一个数据列表范围，例如要导出包含有字符“abc”的电子信箱，在“起始数据”中输入“abc”即可，如果要导出2002年12月注册的域名，可以在“起始数据”中输入“2002-12-01”，在“结束数据”中输入“2002-12-31”。<br>
            &nbsp;&nbsp;&nbsp;&nbsp;如果“起始数据”和“结束数据”均留空，将导出全部数据。</td>
        </tr>
		</form>
      </table>
    </td>
  </tr>
  <tr>
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td>
    </td>
  </tr>
  <tr>
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;">
      <% = strCounter %> 
      <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF"><% = strToPrint %>
      </table></td>
  </tr>
  <tr>
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td>
	</td>
  </tr>
</table>
<%
If flag = 1 Then
    Response.Write("<script>window.open('downFile.asp?file=data.CSV','','');</script>")
End If
%>
</body>
</html>
