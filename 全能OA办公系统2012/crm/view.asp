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
Function getGroupName(gId)
    If gId = "" Then
	    getGroupName = ""
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_group Where gId = " & gId,conn,3,1
		If rs.RecordCount <> 1 Then
		    getGroupName = ""
		Else
		    getGroupName = rs("gName")
		End If
		rs.Close
		Set rs = Nothing
	End If
End Function

Function getList(i,sTable,iId,sValue)
    If i < 1 Or i > 2 Then
	    getList = ""
		Exit Function
	End If
	Dim strList
	Dim rs
	If i = 1 Then
	    strList = "<select name=""" & sValue & """>"
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
	    strList = "<select name=""" & sValue & """>"
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

Function listRecords(cId)
    Dim rs,strOut
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select * From baidu_records Where cId = " & cId,conn,3,1
	Do While Not rs.BOF And Not rs.EOF
	    strOut = strOut & "        <tr>" & VBCrlf
	    strOut = strOut & "          <td width=""120"" align=""center"">" & rs("rDate") & "</td>" & VBCrlf
	    strOut = strOut & "          <td width=""80"" align=""center"">" & rs("rType") & "</td>" & VBCrlf
	    strOut = strOut & "          <td>" & rs("rContent") & "</td>" & VBCrlf
	    strOut = strOut & "        </tr>" & VBCrlf
		rs.MoveNext
	Loop
	rs.Close
	Set rs = Nothing
	listRecords = strOut
End Function

Rem Session("CRM_account") 用户帐号
Rem Session("CRM_name") 用户名
Rem Session("CRM_level") 用户等级

If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

Session("CRM_url") = Request.ServerVariables("HTTP_REFERER")
If InStr(Session("CRM_url"),"addData.asp") > 0 Or InStr(Session("CRM_url"),"view.asp") > 0 Then
    Session("CRM_url") = "listAll.asp"
End If

Dim strNormal,strAdmin,strToPrint
strNormal = strNormal & "        <tr>" & VBCrlf
strNormal = strNormal & "          <td width=""60"" align=""center"" bgcolor=""menu"">编号</td>" & VBCrlf
strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">公司名称</td>" & VBCrlf
strNormal = strNormal & "          <td align=""center"" bgcolor=""menu"">公司网址</td>" & VBCrlf
strNormal = strNormal & "          <td width=""80"" align=""center"" bgcolor=""menu"">地区</td>" & VBCrlf
strNormal = strNormal & "        </tr>" & VBCrlf

strAdmin = strAdmin & "        <tr>" & VBCrlf
strAdmin = strAdmin & "          <td width=""60"" align=""center"" bgcolor=""menu"">编号</td>" & VBCrlf
strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">公司名称</td>" & VBCrlf
strAdmin = strAdmin & "          <td align=""center"" bgcolor=""menu"">公司网址</td>" & VBCrlf
strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">地区</td>" & VBCrlf
strAdmin = strAdmin & "          <td width=""80"" align=""center"" bgcolor=""menu"">业务员</td>" & VBCrlf
strAdmin = strAdmin & "        </tr>" & VBCrlf

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

function showHideBlock(s)
{
    if(eval("document.all." + s + ".style.display == \"none\"")){
	    eval("document.all." + s + ".style.display = \"block\"");
	}
	else{
	    eval("document.all." + s + ".style.display = \"none\"");	
	}
}

function checkInput()
{
    for(var i=0;i<arguments.length;i++){
	    var o = eval("document.all." + arguments[i]);
		if(o.value == ""){
		    alert("请输入内容。");
			o.focus();
			return false;
			break;
		}
	}
}

if (this.location.href == top.location.href){
    top.location.href = "";
}

function openModalDialog(thisUrl)
{
    var strUrl = thisUrl;
	//window.showModalDialog(strUrl,"","status:false;dialogWidth:600px;dialogHeight:450px;status:no;scroll=no;resizable=no;help=no;");
    window.open(strUrl,'','menubar=no,scrollbars=no,resizable=no,width=480,height=360');
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
<%
Dim action
action = Trim(Request("action"))
Select Case action
''Case "save"
''    Call saveData()
''Case "edit"
''    Call editForm()
Case "saveRecords"
    Call saveEditData()
Case Else
    Call editForm()
End Select

Sub saveEditData()
    Dim cId
	Dim rDate,rType,rContent
	cId = CInt(ABS(Request("cId")))	    
	rDate = Request.Form("rDate")
	rType = Request.Form("recordsType")
	rContent = htmlEncode2(Request.Form("rContent"))
	If Not IsNumeric(cId) Or cId <= 0 Or rType = "" Or Not IsDate(rDate) Or rContent = "" Then
	    Response.Write("<font color=""#FF0000""><b>数据请求错误。</b></font><br><br>")
		Response.Write("<input type=""button"" value="" 返 回 "" onClick=""history.back();"">")
	Else
    	Dim rs
    	Set rs = Server.CreateObject("ADODB.Recordset")
    	rs.Open "Select Top 1 * From baidu_records",conn,3,2
		rs.AddNew
		rs("cId") = cId
		rs("rDate") = rDate
		rs("rType") = rType
		rs("rContent") = rContent
    	rs.Update
    	rs.Close
    	Set rs = Nothing
    	Response.Redirect("view.asp?cId=" & cId)
	End If
End Sub

Sub editform()
    Dim cId
	cId = CInt(ABS(Request("cId")))
	If Not IsNumeric(cId) Or cId <= 0 Then
	    Response.Write("<font color=""#FF0000""><b>数据请求错误。</b></font><br><br>")
		Response.Write("<input type=""button"" value="" 返 回 "" onClick=""location.replace('listAll.asp');"">")
	Else
	    Dim rs
		Set rs = Server.CreateObject("ADODB.Recordset")
		rs.Open "Select * From baidu_client Where cId = " & cId,conn,3,1
		If rs.RecordCount <> 1 Then		
	        Response.Write("<font color=""#FF0000""><b>数据请求错误。</b></font><br><br>")
		    Response.Write("<input type=""button"" value="" 返 回 "" onClick=""location.replace('listAll.asp');"">")
			Response.End()
		End If
		Dim cCompany,cLinkman,cHomepage,cTel,cEmail,cUser
	    Dim cArea,cLocal,cType,cStatus,cAddress
		Dim cGroup,cTrade,cSquare
	    'Dim cDomainBegin,cDomainEnd
	    'Dim cSpaceBegin,cSpaceEnd
	    'Dim cOprationBegin,cOprationEnd
	    Dim cInfo
		cCompany = rs("cCompany")
		cLinkman = rs("cLinkman")
		cHomepage = rs("cHomepage")
		cTel = Replace(rs("cTel"),"|"," ")
		cEmail = rs("cEmail")
		cUser = rs("cUser")
		cArea = rs("cArea")
		cLocal = rs("cLocal")
		cType = rs("cType")
		cStatus = rs("cStatus")
		cAddress = rs("cAddress")
		cGroup = rs("cGroup")
		cTrade = rs("cTrade")
		cSquare = rs("cSquare")
		'cDomainBegin = rs("cDomainBegin")
		'If cDomainBegin = "1900-12-31" Then cDomainBegin = ""
		'cDomainEnd = rs("cDomainEnd")
		'If cDomainEnd = "1900-12-31" Then cDomainEnd = ""
		'cSpaceBegin = rs("cSpaceBegin")
		'If cSpaceBegin = "1900-12-31" Then cSpaceBegin = ""
		'cSpaceEnd = rs("cSpaceEnd")
		'If cSpaceEnd = "1900-12-31" Then cSpaceEnd = ""
		'cOprationBegin = rs("cOprationBegin")
		'If cOprationBegin = "1900-12-31" Then cOprationBegin = ""
		'cOprationEnd = rs("cOprationEnd")
		'If cOprationEnd = "1900-12-31" Then cOprationEnd = ""
		cInfo = rs("cInfo")
		rs.Close
		Set rs = Nothing
%>
<table width="550" height="100%" border="0" align="center" cellpadding="0" cellspacing="0">
  <tr id="oHead" style="display: block;"> 
    <td height="1" valign="top"> <table width="550" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table>
      <table width="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td height="5"><img src="images/null.gif" width="1" height="1"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="16" align="center" bgcolor="#88ADDF" id="oHeadBar" style="cursor: hand;" title="隐藏头部" onClick="return showHideHead(document.all.oHeadCtrl.src);"> 
      <img src="images/arrow_up.gif" alt="隐藏头部" width="16" height="16" align="absmiddle" id="oHeadCtrl">&nbsp;</td></td>
  </tr>
  <tr> 
    <td height="1" align="center" valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> 
      <table width="100%" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td width="100" height="25" align="right">公司名称：</td>
          <td width="93"> <% = cCompany %> <input name="id" type="hidden" id="id" value="<% = cId %>"> 
          </td>
          <td width="91" align="right">联系人：</td>
          <td width="222"><% = cLinkman %></td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">网址：</td>
          <td width="93"> <% = cHomepage %></td>
          <td width="91" align="right">联系电话：</td>
          <td><% = cTel %></td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">电子信箱：</td>
          <td width="93"> <% = cEmail %></td>
          <td width="91" align="right">公司地址：</td>
          <td> 
            <% = cAddress %>
          </td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">地区：</td>
          <td width="93"> <% = cArea %></td>
          <td width="91" align="right">所属小区：</td>
          <td><% = cSquare %></td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">客户类型：</td>
          <td>
            <% = cType %>
          </td>
          <td align="right">行业类型：</td>
          <td><% = cTrade %></td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">业务员：</td>
          <td width="93"> 
            <% = cUser %>
          </td>
          <td width="91" align="right">业务组：</td>
          <td><% = getGroupName(cGroup) %></td>
        </tr>
        <tr> 
          <td width="100" height="25" align="right">客户描述：</td>
          <td colspan="3"><% = cInfo %></td>
        </tr>
        <tr> 
          <td colspan="4" align="center"> <input name="Reset" type="button" id="Reset" value=" 修 改 " onClick="location.href='addData.asp?action=edit&cId=<% = cId %>';"> 
            &nbsp;&nbsp; <input name="Back" type="button" id="Back" value=" 返回列表 " onClick="location.href='<% = Session("CRM_url") %>';"></td>
        </tr>
      </table></td>
  </tr>
  <tr> 
    <td height="16" align="center" bgcolor="#88ADDF" style="cursor: hand;" onClick="return showHideBlock('addRecords');"><span style="color: #FFFFFF; font-weight: bold;">[添加拜访纪录]</span></td>
  </tr>
  <tr id="addRecords" style="display: block;"> 
    <td height="1" bgcolor="#FFFFFF" style="padding: 10px;"> <table width="100%" border="0" cellspacing="0" cellpadding="3">
        <form name="recordsForm" action="?action=saveRecords" method="post" onSubmit="return checkInput('rDate','recordsType','cId','rContent');">
          <tr> 
            <td>添加拜访纪录：</td>
          </tr>
          <tr> 
            <td>拜访日期： 
              <input name="rDate" type="text" id="rDate3" value="<% = Date() %>" size="16" maxlength="12">
              拜访类型： 
              <% = getList(1,"baidu_recordsType",,"recordsType") %> <input name="cId" type="hidden" id="cId" value="<% = cId %>"></td>
          </tr>
          <tr> 
            <td valign="top">拜访记录： 
              <textarea name="rContent" cols="80" rows="4" id="rContent"></textarea></td>
          </tr>
          <tr> 
            <td align="center"><input type="submit" name="Submit" value=" 提 交 ">
              &nbsp;&nbsp;
              <input name="AddPlan" type="button" id="AddPlan" value=" 添加拜访计划 " onClick="openModalDialog('addPlan_records.asp?cId=<% = cId %>');"></td>
          </tr>
        </form>
      </table></td>
  </tr>
  <tr> 
    <td height="16" align="center" bgcolor="#88ADDF" style="cursor: hand;" onClick="return showHideBlock('listRecords');"><span style="color: #FFFFFF; font-weight: bold;">[拜访纪录列表]</span></td>
  </tr>
  <tr id="listRecords" style="display: block;"> 
    <td valign="top" bgcolor="#FFFFFF" style="padding: 10px;"> <table width="100%" align="center" border="1" cellpadding="3" cellspacing="0" bordercolor="#DCDCDC" bordercolordark="#FFFFFF">
        <tr> 
          <td width="120" align="center" bgcolor="menu">拜访时间</td>
          <td width="80" align="center" bgcolor="menu">拜访类别</td>
          <td align="center" bgcolor="menu">拜访记录</td>
        </tr>
        <% = listRecords(cId) %>
      </table></td>
  </tr>
  <tr> 
    <td bgcolor="#FFFFFF">&nbsp;</td>
  </tr>
  <tr> 
    <td height="16" align="right" bgcolor="#88ADDF"><a href="#top"><img src="images/arrow_up.gif" alt="返回顶部" width="16" height="16" border="0" align="absmiddle"></a>&nbsp;</td></td>
  </tr>
</table>
<%
    End If
End Sub
%>
</body>
</html>
