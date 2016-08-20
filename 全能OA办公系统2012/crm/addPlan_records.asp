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
<%
If Session("CRM_account") = "" Or Session("CRM_name") = "" Or Session("CRM_level") <= 0 Then Response.Redirect("login.asp")

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

Sub addDate()
    Dim cId,rDate,rType,rContent,rUser,rDay
	cId = CInt(Abs(Request.Form("cId")))
	rDate = Trim(Request.Form("date"))
	rType = Trim(Request.Form("recordsType"))
	rContent = htmlEncode2(Request.Form("content"))
	rDay = CInt(Abs(Request.Form("day")))
	If rDay = 0 Then rDay = 2
	rUser = Session("CRM_name")
	If cId <= 0 Or rDate = "" Or rType = "" Or rContent = "" Or rUser = "" Then
	    Response.Write("<div align=""center"">数据错误</div>")
		Response.End()
	End If
	Dim rs
	Set rs = Server.CreateObject("ADODB.Recordset")
	rs.Open "Select Top 1 * From baidu_recordsPlan",conn,3,2
	rs.AddNew
	rs("cId") = cId
	rs("rDate") = rDate
	rs("rType") = rType
	rs("rContent") = rContent
	rs("rUser") = rUser
	rs("rDay") = rDay
	rs.Update
	rs.Close
	Set rs = Nothing
	Response.Write("<script>window.close();</script>")
	Response.End()
End Sub
Dim action
action = Trim(Request.QueryString("action"))
If action = "add" Then Call addDate()

Dim cId
cId = 0 
cId = CInt(Abs(Request.QueryString("cId")))
If cId <= 0 Then Response.Write("<script>window.close();</script>")
Dim rs
Set rs = Server.CreateObject("ADODB.Recordset")
rs.Open "Select cCompany,cType,cTrade From baidu_client Where cId = " & cId,conn,3,1
If rs.RecordCount = 1 Then
    Dim cCompany,cType,cTrade
	cCompany = rs("cCompany")
	cType = rs("cType")
	cTrade = rs("cTrade")
End If
rs.Close
Set rs = Nothing

%>
<html>
<head>
<meta http-equiv="Content-Type" content="text/html; charset=gb2312">
<title>添加拜访计划</title>
<link href="myStyle.css" rel="stylesheet" type="text/css">
</head>

<body  >
<table width="480" height="360" border="0" align="center" cellpadding="3" cellspacing="0">
  <tr valign="top"> 
    <td width="140" height="1"> <fieldset style="padding: 10px;">
      <legend>日期</legend>
      <table width="140" border="0" cellspacing="0" cellpadding="0">
        <tr> 
          <td><select name="mm" onChange="changeMonth();">
              <option value="1">一月</option>
              <option value="2">二月</option>
              <option value="3">三月</option>
              <option value="4">四月</option>
              <option value="5">五月</option>
              <option value="6">六月</option>
              <option value="7">七月</option>
              <option value="8">八月</option>
              <option value="9">九月</option>
              <option value="10">十月</option>
              <option value="11">十一月</option>
              <option value="12">十二月</option>
            </select> </td>
          <td align="right"> <select name="yyyy" size="1" onChange="changeYear();">
              <option value="2000">2000 </option>
              <option value="2001">2001 </option>
              <option value="2002">2002 </option>
              <option value="2003">2003 </option>
            </select> </td>
        </tr>
        <tr> 
          <td colspan="2"><br> <table width="140%" border="0" cellspacing="0" cellpadding="0" style="border-left: 1px solid #808080; border-top: 1px solid #808080; border-right: 1px solid #FFFFFF; border-bottom: 1px solid #FFFFFF;">
              <tr> 
                <td id="calendar">&nbsp;</td>
              </tr>
            </table></td>
        </tr>
      </table>
      </fieldset></td>
    <td rowspan="2"><fieldset style="padding: 10px;">
      <legend>计划内容</legend>
      <table width="100%" height="100%" border="0" cellpadding="3" cellspacing="0">
        <form name="addPlan" action="?action=add" method="post">
          <tr> 
            <td height="20"> 计划日期： 
              <input name="date" type="text" id="date" value="<% = Date() %>" size="12" maxlength="12"> 
              <br>
              拜访类型： 
              <% = getList(1,"baidu_recordsType",,"recordsType") %> <input name="cId" type="hidden" id="cId" value="<% = cId %>"> 
            </td>
          </tr>
          <tr> 
            <td><textarea name="content" rows="11" id="content" style="width: 100%;"></textarea></td>
          </tr>
          <tr>
            <td height="20">提前
              <input name="day" type="text" id="day" size="4" maxlength="4" value="2">
              天提醒我</td>
          </tr>
          <tr> 
            <td height="20" align="center"><input type="submit" name="Submit" value=" 提交计划 "></td>
          </tr>
        </form>
      </table>
      </fieldset></td>
  </tr>
  <tr valign="top"> 
    <td><fieldset style="padding: 10px;">
      <legend>客户资料</legend>
      <table width="100%" height="100%" border="0" cellpadding="0" cellspacing="0">
        <tr> 
          <td><strong>客户名称：</strong></td>
        </tr>
        <tr> 
          <td><% = cCompany %></td>
        </tr>
        <tr> 
          <td><strong>客户等级：</strong></td>
        </tr>
        <tr> 
          <td><% = cType %></td>
        </tr>
        <tr> 
          <td><strong>行业类型：</strong></td>
        </tr>
        <tr> 
          <td><% = cTrade %></td>
        </tr>
      </table>
      </fieldset></td>
  </tr>
  <tr valign="top">
    <td>&nbsp;</td>
    <td align="center">
<input type="button" name="Submit2" value=" 关 闭 " onClick="window.close();">
    </td>
  </tr>
</table>
</body>
</html>
<script language="JavaScript">
<!--
function changeDate()
{
    var yyyy = document.all.yyyy.value;
	var mm = document.all.mm.value;
	
	var strOut;
	var DateArray = ["日","一","二","三","四","五","六"];	
	var arrDate = GetDay(yyyy,mm).TDate();
    strOut = "<table width=\"140\" border=\"0\" cellpadding=\"2\" cellspacing=\"0\" bgcolor=\"#FFFFFF\">";
    strOut = strOut + "<tr>";
	for(var i=0;i<7;i++){
	    strOut = strOut + "<td bgcolor=\"#999999\">" + DateArray[i] + "</td>";
	}		
    strOut = strOut + "</tr>";
	
	for(var i=0;i<6;i++){
        strOut = strOut + "<tr align='center'>";
        for(var j=0;j<7;j++){
		    strOut = strOut + "<td style=\"cursor: hand;\" onClick=\"SetDate('" + arrDate[i * 7 + j] + "');\">" + arrDate[i * 7 + j] + "</td>";
		}
        strOut = strOut + "</tr>";
    }
    strOut = strOut + "</table>";
	
    document.all.calendar.innerHTML = strOut;
	
}

function changeYear()
{
    changeDate();
}

function changeMonth()
{
    changeDate();
}
function SetDate(d)
{
    var y = document.all.yyyy.value;
	var m = document.all.mm.value;
	document.all.date.value = y + "-" + m + "-" + d;
}
function GetDay(y,m)
{
    this.TDate = function()
	{
        this.DayArray = [];
        for(var i=0;i<42;i++)this.DayArray[i] = "&nbsp;";
        for(var i=0;i<new Date(y,m,0).getDate();i++)this.DayArray[i+new Date(y,m-1,1).getDay()] = i + 1;
        return this.DayArray;
    }
    return this;
}
var today = new Date();
var year = today.getYear();
var month = today.getMonth() + 1;
for(var i=0;i<document.all.yyyy.options.length;i++){
    if(document.all.yyyy.options[i].value == year){
	    document.all.yyyy.options[i].selected = true;
	}
}
for(var i=0;i<document.all.mm.options.length;i++){
    if(document.all.mm.options[i].value == month){
	    document.all.mm.options[i].selected = true;
	}
}
changeDate();
-->
</script>